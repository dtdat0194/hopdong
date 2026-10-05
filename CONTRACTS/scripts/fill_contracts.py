#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Sinh hợp đồng hàng loạt từ Onboarding_List.xlsx (bản dòng lệnh của web app).

    python3 scripts/fill_contracts.py                     # PDF cho mọi dòng đánh dấu "x"
    python3 scripts/fill_contracts.py --format docx       # xuất file Word
    python3 scripts/fill_contracts.py "Hà Tuấn Anh"       # chỉ một người
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app import contract_spec as spec                                  # noqa: E402
from app.excel import LoiFileExcel, dien_cho_trong, doc_file           # noqa: E402
from app.render_docx import render_docx                                # noqa: E402
from app.render_pdf import build_pdf                                   # noqa: E402
from app.util import khong_dau                                         # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
XLSX = ROOT / "Onboarding_List.xlsx"
OUT = ROOT / "output"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("ten", nargs="*", help="Chỉ sinh hợp đồng cho (các) nhân sự này")
    ap.add_argument("--format", choices=["pdf", "docx"], default="pdf")
    ap.add_argument("--merge", action="store_true", help="Gộp tất cả vào một file PDF")
    args = ap.parse_args()

    try:
        data = doc_file(str(XLSX))
    except LoiFileExcel as e:
        sys.exit(str(e))

    loc = {khong_dau(t).lower() for t in args.ten}
    if loc:
        chon = [p for p in data["people"] if khong_dau(p["name"]).lower() in loc]
    else:
        chon = [p for p in data["people"] if p["marked"]]

    for p in chon:
        if not p["type"]:
            print(f'  !  Bỏ qua {p["name"]}: chưa chọn loại hợp đồng trong Excel')
    chon = [p for p in chon if p["type"]]

    if not chon:
        sys.exit("Không có nhân sự nào hợp lệ. Kiểm tra cột 'Tạo HĐ (x)' và 'Loại hợp đồng'.")

    OUT.mkdir(exist_ok=True)
    items = [(p["type"], dien_cho_trong(p["ctx"], p["type"])) for p in chon]

    if args.merge and args.format == "pdf":
        dest = OUT / f"Hop_dong_{len(chon)}_nhan_su.pdf"
        with open(dest, "wb") as f:
            build_pdf(items, f)
        print(f"  OK  {len(chon)} hợp đồng -> output/{dest.name}")
    else:
        for p, item in zip(chon, items):
            name = f'{spec.CONTRACTS[p["type"]]["ten_file"]}-{p["name"].upper()}.{args.format}'
            dest = OUT / name
            if args.format == "docx":
                dest.write_bytes(render_docx(*item))
            else:
                with open(dest, "wb") as f:
                    build_pdf([item], f)
            print(f'  OK  {p["type"]}  {p["name"]:<24} -> output/{name}')

    thieu = [(p["name"], p["missing"]) for p in chon if p["missing"]]
    if thieu:
        print("\nCẦN BỔ SUNG (chỗ thiếu để dấu chấm lửng trong file xuất ra):")
        for ten, m in thieu:
            print(f"  ! {ten}: {', '.join(m)}")


if __name__ == "__main__":
    main()
