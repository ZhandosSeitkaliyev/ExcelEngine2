import json
import io
from typing import Any
from urllib.parse import quote
from fastapi import FastAPI, File, UploadFile, HTTPException, Body, Request
from fastapi.responses import JSONResponse, StreamingResponse
from excel_parser import excel_to_json, json_to_excel

app = FastAPI(
    title="ExcelEngine",
    description="Конвертирует Excel файлы в JSON с полным сохранением форматирования",
    version="1.0.0",
)

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _disposition(filename: str) -> str:
    encoded = quote(filename.encode("utf-8"), safe="")
    return f"attachment; filename*=UTF-8''{encoded}"


def _excel_response(data: dict, fallback_name: str = "restored.xlsx") -> StreamingResponse:
    try:
        xlsx_bytes = json_to_excel(data)
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"Ошибка сборки Excel: {e}")

    hint = data.get("filename_hint") or fallback_name
    output_name = hint.rsplit(".", 1)[0] + ".xlsx"
    return StreamingResponse(
        io.BytesIO(xlsx_bytes),
        media_type=XLSX_MIME,
        headers={"Content-Disposition": _disposition(output_name)},
    )


async def _get_upload_file(request: Request) -> UploadFile:
    """Извлекает файл из form-data независимо от имени поля."""
    form = await request.form()
    file = next((v for v in form.values() if hasattr(v, "filename")), None)
    if not file:
        raise HTTPException(status_code=400, detail="Файл не найден в запросе.")
    if not file.filename.endswith((".xlsx", ".xlsm", ".xltx", ".xltm")):
        raise HTTPException(status_code=400, detail="Поддерживаются только .xlsx / .xlsm файлы.")
    return file


@app.post("/convert", summary="Конвертировать Excel → JSON (ответ в теле)")
async def convert(request: Request):
    file = await _get_upload_file(request)
    content = await file.read()
    try:
        result = excel_to_json(content)
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"Ошибка парсинга файла: {e}")
    result["filename_hint"] = file.filename
    return JSONResponse(content=result)


@app.post("/convert/download", summary="Конвертировать Excel → скачать .json файл")
async def convert_download(request: Request):
    file = await _get_upload_file(request)
    content = await file.read()
    try:
        result = excel_to_json(content)
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"Ошибка парсинга файла: {e}")
    result["filename_hint"] = file.filename
    json_bytes = json.dumps(result, ensure_ascii=False, indent=2).encode("utf-8")
    output_name = file.filename.rsplit(".", 1)[0] + ".json"
    return StreamingResponse(
        io.BytesIO(json_bytes),
        media_type="application/json",
        headers={"Content-Disposition": _disposition(output_name)},
    )


@app.post("/restore", summary="Восстановить Excel из JSON (тело запроса)")
async def restore(data: Any = Body(...)):
    return _excel_response(data)


@app.post("/restore/upload", summary="Восстановить Excel из загруженного .json файла")
async def restore_upload(file: UploadFile = File(...)):
    if not file.filename.endswith(".json"):
        raise HTTPException(status_code=400, detail="Ожидается .json файл.")
    try:
        data = json.loads(await file.read())
    except Exception:
        raise HTTPException(status_code=400, detail="Файл не является валидным JSON.")
    return _excel_response(data, fallback_name=file.filename)


@app.get("/", summary="Health check")
def root():
    return {"status": "ok", "docs": "/docs"}
