import io
import json
import datetime
import hashlib

from openpyxl import load_workbook, Workbook
from openpyxl.utils import column_index_from_string
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.styles.colors import Color


# ─── Color helpers ────────────────────────────────────────────────────────────

def _color(obj) -> str | None:
    if obj is None:
        return None
    if obj.type == "rgb":
        return f"#{obj.rgb}"
    if obj.type == "theme":
        return f"theme:{obj.theme}"
    if obj.type == "indexed":
        return f"indexed:{obj.indexed}"
    return None


def _parse_color(s: str | None) -> Color | None:
    if not s:
        return None
    if s.startswith("#"):
        return Color(rgb=s[1:])
    if s.startswith("theme:"):
        return Color(theme=int(s.split(":")[1]))
    if s.startswith("indexed:"):
        return Color(indexed=int(s.split(":")[1]))
    return None


# ─── Excel → JSON ─────────────────────────────────────────────────────────────

def _extract_style(cell) -> dict:
    """Only non-default properties — nothing null, nothing zero."""
    style = {}

    f = cell.font
    fp = {}
    if f.name:      fp["name"] = f.name
    if f.size:      fp["size"] = f.size
    if f.bold:      fp["bold"] = True
    if f.italic:    fp["italic"] = True
    if f.underline: fp["underline"] = f.underline
    if f.strike:    fp["strike"] = True
    fc = _color(f.color)
    if fc:          fp["color"] = fc
    if f.vertAlign: fp["vertAlign"] = f.vertAlign
    if fp:          style["font"] = fp

    fill = cell.fill
    if fill.fill_type and fill.fill_type not in (None, "none"):
        fp2 = {"type": fill.fill_type}
        fg = _color(fill.fgColor)
        if fg and fg != "#00000000": fp2["fgColor"] = fg
        bg = _color(fill.bgColor)
        if bg and bg != "#00000000": fp2["bgColor"] = bg
        style["fill"] = fp2

    b = cell.border
    bp = {}
    for name in ("left", "right", "top", "bottom"):
        side = getattr(b, name)
        if side and side.border_style:
            sp = {"style": side.border_style}
            sc = _color(side.color)
            if sc and not sc.startswith("indexed:"): sp["color"] = sc
            bp[name] = sp
    if b.diagonal and b.diagonal.border_style:
        dp = {"style": b.diagonal.border_style}
        dc = _color(b.diagonal.color)
        if dc: dp["color"] = dc
        bp["diagonal"] = dp
    if b.diagonalUp:   bp["diagonalUp"] = True
    if b.diagonalDown: bp["diagonalDown"] = True
    if bp: style["border"] = bp

    a = cell.alignment
    ap = {}
    if a.horizontal:    ap["horizontal"] = a.horizontal
    if a.vertical:      ap["vertical"] = a.vertical
    if a.wrap_text:     ap["wrapText"] = True
    if a.shrink_to_fit: ap["shrinkToFit"] = True
    if a.indent:        ap["indent"] = a.indent
    if a.text_rotation: ap["textRotation"] = a.text_rotation
    if ap: style["alignment"] = ap

    nf = cell.number_format
    if nf and nf != "General": style["number_format"] = nf

    return style


def _fp(style: dict) -> str:
    return hashlib.md5(json.dumps(style, sort_keys=True).encode()).hexdigest()[:8]


def _cell_value(cell):
    val = cell.value
    if isinstance(val, (datetime.datetime, datetime.date, datetime.time)):
        return val.isoformat()
    return val


def _parse_sheet(ws) -> dict:
    cols   = {c: d.width  for c, d in ws.column_dimensions.items() if d.width}
    rows_h = {str(i): d.height for i, d in ws.row_dimensions.items() if d.height}
    merged = [str(r) for r in ws.merged_cells.ranges]

    styles: dict[str, dict] = {}
    cells:  list[dict]      = []

    for row in ws.iter_rows():
        for cell in row:
            val   = _cell_value(cell)
            style = _extract_style(cell)

            if val is None and not style:
                continue

            entry: dict = {"c": cell.coordinate}
            if val is not None:
                entry["v"] = val
            if style:
                sid = _fp(style)
                styles[sid] = style
                entry["s"] = sid
            if cell.hyperlink:
                entry["href"] = str(cell.hyperlink)
            if cell.comment:
                entry["comment"] = cell.comment.text

            cells.append(entry)

    return {
        "name":         ws.title,
        "tab_color":    _color(ws.sheet_properties.tabColor) if ws.sheet_properties.tabColor else None,
        "freeze_panes": str(ws.freeze_panes) if ws.freeze_panes else None,
        "zoom":         ws.sheet_view.zoomScale if ws.sheet_view else None,
        "cols":         cols,
        "rows_h":       rows_h,
        "merged":       merged,
        "styles":       styles,
        "cells":        cells,
    }


def excel_to_json(file_bytes: bytes) -> dict:
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    return {
        "filename_hint": None,
        "active_sheet":  wb.active.title if wb.active else None,
        "sheets":        [_parse_sheet(wb[name]) for name in wb.sheetnames],
    }


# ─── JSON → Excel ─────────────────────────────────────────────────────────────

def _make_font(d: dict | None) -> Font | None:
    if not d: return None
    kw = {k: d.get(k) for k in ("name", "size", "bold", "italic", "underline", "strike", "vertAlign")}
    c = _parse_color(d.get("color"))
    if c: kw["color"] = c
    return Font(**kw)


def _make_fill(d: dict | None) -> PatternFill | None:
    if not d or not d.get("type"): return None
    kw: dict = {"fill_type": d["type"]}
    fg = _parse_color(d.get("fgColor"))
    if fg: kw["fgColor"] = fg
    bg = _parse_color(d.get("bgColor"))
    if bg: kw["bgColor"] = bg
    return PatternFill(**kw)


def _make_side(d: dict | None) -> Side:
    if not d or not d.get("style"): return Side()
    c = _parse_color(d.get("color"))
    return Side(border_style=d["style"], color=c) if c else Side(border_style=d["style"])


def _make_border(d: dict | None) -> Border | None:
    if not d: return None
    return Border(
        left=_make_side(d.get("left")),   right=_make_side(d.get("right")),
        top=_make_side(d.get("top")),     bottom=_make_side(d.get("bottom")),
        diagonal=_make_side(d.get("diagonal")),
        diagonalUp=d.get("diagonalUp", False),
        diagonalDown=d.get("diagonalDown", False),
    )


def _make_alignment(d: dict | None) -> Alignment | None:
    if not d: return None
    return Alignment(
        horizontal=d.get("horizontal"),   vertical=d.get("vertical"),
        wrap_text=d.get("wrapText"),      shrink_to_fit=d.get("shrinkToFit"),
        indent=d.get("indent") or 0,      text_rotation=d.get("textRotation") or 0,
    )


def _apply_style(cell, style: dict) -> None:
    font = _make_font(style.get("font"))
    if font: cell.font = font
    fill = _make_fill(style.get("fill"))
    if fill: cell.fill = fill
    border = _make_border(style.get("border"))
    if border: cell.border = border
    alignment = _make_alignment(style.get("alignment"))
    if alignment: cell.alignment = alignment
    if style.get("number_format"):
        cell.number_format = style["number_format"]


def json_to_excel(data: dict) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    for sheet_data in data.get("sheets", []):
        ws = wb.create_sheet(title=sheet_data["name"])

        if sheet_data.get("tab_color"):
            c = _parse_color(sheet_data["tab_color"])
            if c: ws.sheet_properties.tabColor = c
        if sheet_data.get("freeze_panes"):
            ws.freeze_panes = sheet_data["freeze_panes"]
        if sheet_data.get("zoom"):
            ws.sheet_view.zoomScale = sheet_data["zoom"]

        for col, width in sheet_data.get("cols", {}).items():
            ws.column_dimensions[col].width = width
        for row_str, height in sheet_data.get("rows_h", {}).items():
            ws.row_dimensions[int(row_str)].height = height

        styles = sheet_data.get("styles", {})

        for cd in sheet_data.get("cells", []):
            coord = cd["c"]
            col_letter = "".join(ch for ch in coord if ch.isalpha())
            row_num    = int("".join(ch for ch in coord if ch.isdigit()))
            cell = ws.cell(row=row_num, column=column_index_from_string(col_letter))
            if "v" in cd:
                cell.value = cd["v"]
            if cd.get("s") and cd["s"] in styles:
                _apply_style(cell, styles[cd["s"]])
            if cd.get("href"):
                cell.hyperlink = cd["href"]

        for merge_str in sheet_data.get("merged", []):
            try:
                ws.merge_cells(merge_str)
            except Exception:
                pass

    active = data.get("active_sheet")
    if active and active in wb.sheetnames:
        wb.active = wb[active]
    elif wb.sheetnames:
        wb.active = wb[wb.sheetnames[0]]

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()
