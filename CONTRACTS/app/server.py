# -*- coding: utf-8 -*-
"""Web app: tải lên Onboarding_List.xlsx, xem danh sách, chọn nhân sự và xuất hợp đồng."""
import io
import time
import uuid
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from . import contract_spec as spec
from .excel import LoiFileExcel, dien_cho_trong, doc_file
from .render_docx import ensure_templates, render_docx
from .render_pdf import build_pdf
from .util import ten_file

ROOT = Path(__file__).resolve().parents[1]
STATIC = Path(__file__).resolve().parent / "static"
MAU_EXCEL = ROOT / "Onboarding_List.xlsx"

HAN_PHIEN = 6 * 3600          # giữ file đã tải lên trong 6 giờ
GIOI_HAN_MB = 20

app = FastAPI(title="QCD - Xuất hợp đồng")
_store: dict[str, dict] = {}


@app.middleware("http")
async def khong_cache(request, call_next):
    """Giao diện chạy tại chỗ nên luôn lấy bản mới nhất, tránh phải xóa cache thủ công."""
    res = await call_next(request)
    if not request.url.path.startswith("/api/"):
        res.headers["Cache-Control"] = "no-store, must-revalidate"
    return res


def _don_dep():
    cu = [k for k, v in _store.items() if time.time() - v["ts"] > HAN_PHIEN]
    for k in cu:
        _store.pop(k, None)


def _phien(token):
    _don_dep()
    ses = _store.get(token)
    if not ses:
        raise HTTPException(404, "Phiên làm việc đã hết hạn. Vui lòng tải lại file Excel.")
    return ses


def _tai_ve(data: bytes, filename: str, media: str):
    ascii_name = ten_file(filename)
    return StreamingResponse(
        io.BytesIO(data),
        media_type=media,
        headers={"Content-Disposition":
                 f"attachment; filename=\"{ascii_name}\"; filename*=UTF-8''{quote(filename)}"},
    )


@app.on_event("startup")
def _khoi_dong():
    ensure_templates()


@app.get("/api/health")
def health():
    return {"ok": True, "contracts": {k: v["ten"] for k, v in spec.CONTRACTS.items()}}


@app.get("/api/mau-excel")
def mau_excel():
    if not MAU_EXCEL.exists():
        raise HTTPException(404, "Chưa có file mẫu trên máy chủ.")
    return FileResponse(MAU_EXCEL, filename="Onboarding_List.xlsx",
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.post("/api/upload")
async def upload(file: UploadFile = File(...)):
    if not file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(400, "Chỉ nhận file Excel định dạng .xlsx hoặc .xlsm.")
    data = await file.read()
    if len(data) > GIOI_HAN_MB * 1024 * 1024:
        raise HTTPException(400, f"File vượt quá {GIOI_HAN_MB} MB.")
    try:
        parsed = doc_file(data)
    except LoiFileExcel as e:
        raise HTTPException(400, str(e))

    token = uuid.uuid4().hex
    _don_dep()
    _store[token] = {"parsed": parsed, "ts": time.time()}
    people = [{k: v for k, v in p.items() if k != "ctx"} for p in parsed["people"]]
    return {
        "token": token,
        "filename": file.filename,
        "company": parsed["company"],
        "people": people,
    }


class YeuCauXuat(BaseModel):
    token: str
    rows: list[int]
    format: str = "pdf"
    merge: bool = False


@app.post("/api/export")
def export(req: YeuCauXuat):
    ses = _phien(req.token)
    chon = [p for p in ses["parsed"]["people"] if p["row"] in set(req.rows)]
    if not chon:
        raise HTTPException(400, "Chưa chọn nhân sự nào để xuất.")
    thieu_loai = [p["name"] for p in chon if not p["type"]]
    if thieu_loai:
        raise HTTPException(400, "Các nhân sự sau chưa chọn loại hợp đồng trong Excel: "
                                 + ", ".join(thieu_loai))

    items = [(p["type"], dien_cho_trong(p["ctx"], p["type"])) for p in chon]

    # Gộp nhiều hợp đồng chỉ áp dụng cho PDF; còn lại giao diện gọi từng người một
    # để mỗi nhân sự ra một file mang đúng tên của họ, không bọc trong file nén.
    if req.merge and req.format == "pdf" and len(chon) > 1:
        buf = io.BytesIO()
        build_pdf(items, buf)
        return _tai_ve(buf.getvalue(), f"Hop_dong_{len(chon)}_nhan_su.pdf", "application/pdf")

    if len(chon) > 1:
        raise HTTPException(400, "Mỗi lần tải chỉ xuất một hợp đồng để file mang đúng tên nhân sự.")

    p, item = chon[0], items[0]
    ten = f'{spec.CONTRACTS[p["type"]]["ten_file"]}-{p["name"].upper()}'
    if req.format == "docx":
        return _tai_ve(render_docx(*item), ten + ".docx",
                       "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    buf = io.BytesIO()
    build_pdf([item], buf)
    return _tai_ve(buf.getvalue(), ten + ".pdf", "application/pdf")


@app.post("/api/preview")
def preview(req: YeuCauXuat):
    """Xem trước ngay trên trình duyệt (PDF mở inline thay vì tải về)."""
    ses = _phien(req.token)
    chon = [p for p in ses["parsed"]["people"] if p["row"] in set(req.rows) and p["type"]]
    if not chon:
        raise HTTPException(400, "Không có hợp đồng nào để xem trước.")
    items = [(p["type"], dien_cho_trong(p["ctx"], p["type"])) for p in chon]
    buf = io.BytesIO()
    build_pdf(items, buf)
    return StreamingResponse(io.BytesIO(buf.getvalue()), media_type="application/pdf",
                             headers={"Content-Disposition": "inline; filename=\"preview.pdf\""})


app.mount("/", StaticFiles(directory=STATIC, html=True), name="static")
