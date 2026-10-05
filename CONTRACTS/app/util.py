# -*- coding: utf-8 -*-
"""Tiện ích dùng chung: chuẩn hóa giá trị, định dạng tiền, đọc số thành chữ."""
import datetime as dt
import re
import unicodedata

DON_VI = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"]
HANG = ["", " nghìn", " triệu", " tỷ"]


def _ba_chu_so(n, day_du):
    tram, chuc, dv = n // 100, (n // 10) % 10, n % 10
    out = []
    if tram > 0 or day_du:
        out.append(DON_VI[tram] + " trăm")
        if chuc == 0 and dv > 0:
            out.append("lẻ")
    if chuc > 1:
        out.append(DON_VI[chuc] + " mươi")
    elif chuc == 1:
        out.append("mười")
    if dv > 0:
        if chuc > 1 and dv == 1:
            out.append("mốt")
        elif chuc >= 1 and dv == 5:
            out.append("lăm")
        else:
            out.append(DON_VI[dv])
    return " ".join(out)


def doc_so_tien(n):
    """15000000 -> 'Mười lăm triệu đồng'"""
    n = int(round(float(n)))
    if n == 0:
        return "Không đồng"
    nhom = []
    while n > 0:
        nhom.append(n % 1000)
        n //= 1000
    phan = []
    for i in range(len(nhom) - 1, -1, -1):
        if nhom[i] == 0:
            continue
        phan.append(_ba_chu_so(nhom[i], day_du=(i != len(nhom) - 1)) + HANG[i % 4])
    s_ = re.sub(r"\s+", " ", " ".join(phan).strip())
    return s_[0].upper() + s_[1:] + " đồng"


def s(v):
    """Chuẩn hóa mọi giá trị ô Excel thành chuỗi hiển thị được."""
    if v is None:
        return ""
    if isinstance(v, dt.datetime):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, dt.date):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def ngay(v):
    if isinstance(v, dt.datetime):
        return v
    if isinstance(v, dt.date):
        return dt.datetime(v.year, v.month, v.day)
    txt = s(v)
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(txt, fmt)
        except ValueError:
            pass
    return None


def tien(v):
    """15000000 -> '15.000.000'"""
    return f"{int(round(float(v))):,}".replace(",", ".")


def khong_dau(txt):
    txt = unicodedata.normalize("NFD", txt)
    txt = "".join(c for c in txt if unicodedata.category(c) != "Mn")
    return txt.replace("đ", "d").replace("Đ", "D")


def ten_file(txt):
    return re.sub(r"[^\w\s.-]", "", khong_dau(txt)).strip() or "khong-ten"
