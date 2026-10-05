# -*- coding: utf-8 -*-
"""Cầu nối giữa JavaScript và lõi Python khi chạy thẳng trong trình duyệt (Pyodide).

Thay cho server.py ở bản chạy trên máy. Không có mạng, không có máy chủ: file
Excel được đọc và hợp đồng được dựng ngay trong tab trình duyệt, nên dữ liệu
nhân sự không bao giờ rời khỏi máy người dùng.
"""
import io
import json

from . import contract_spec as spec
from .excel import dien_cho_trong, doc_file
from .render_docx import render_docx
from .render_pdf import build_pdf

_phien = {}


def nap_excel(data):
    """data: bytes của file .xlsx (hoặc Uint8Array do JS truyền sang)."""
    if hasattr(data, "to_py"):
        data = data.to_py()
    parsed = doc_file(bytes(data))
    _phien["people"] = parsed["people"]
    return json.dumps({
        "company": parsed["company"],
        "people": [{k: v for k, v in p.items() if k != "ctx"} for p in parsed["people"]],
    }, ensure_ascii=False)


def _chon(rows):
    can = set(rows)
    ds = [p for p in _phien.get("people", []) if p["row"] in can and p["type"]]
    if not ds:
        raise ValueError("Không có hợp đồng nào để xuất.")
    return ds


def _item(p):
    return p["type"], dien_cho_trong(p["ctx"], p["type"])


def ten_hop_dong(row):
    p = _chon([row])[0]
    return f'{spec.CONTRACTS[p["type"]]["ten_file"]}-{p["name"].upper()}'


def xuat_pdf(rows):
    buf = io.BytesIO()
    build_pdf([_item(p) for p in _chon(rows)], buf)
    return buf.getvalue()


def xuat_docx(row):
    return render_docx(*_item(_chon([row])[0]))
