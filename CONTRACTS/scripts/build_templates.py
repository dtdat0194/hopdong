#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Sinh lại template Word từ app/contract_spec.py.

Template .docx mới là bản gốc: bạn sửa trực tiếp trong Word và PDF xuất theo đó.
Script này chỉ dùng khi muốn dựng lại từ đầu, và sẽ KHÔNG ghi đè nếu template đã
tồn tại trừ khi chạy kèm --force (lúc đó bản cũ vẫn được sao lưu sang templates/backup).
"""
import argparse
import shutil
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app import contract_spec as spec             # noqa: E402
from app.render_docx import TEMPLATE_DIR, build_template   # noqa: E402


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--force", action="store_true",
                    help="Ghi đè template đang có (sao lưu bản cũ trước khi ghi)")
    args = ap.parse_args()

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    for loai in spec.CONTRACTS:
        dest = TEMPLATE_DIR / spec.CONTRACTS[loai]["template"]
        if dest.exists():
            if not args.force:
                print(f"  BỎ QUA  {dest.name} (đã có sẵn - thêm --force nếu thật sự muốn ghi đè)")
                continue
            backup = TEMPLATE_DIR / "backup" / f"{dest.stem}-{stamp}.docx"
            backup.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(dest, backup)
            print(f"  SAO LƯU {backup.relative_to(TEMPLATE_DIR.parent)}")
        print(f"  GHI     {build_template(loai, dest)}")


if __name__ == "__main__":
    main()
