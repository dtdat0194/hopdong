# -*- coding: utf-8 -*-
"""Đo bề rộng chữ để cột nhãn trong bảng thông tin hai Bên không bị xuống dòng."""
from reportlab.pdfbase import pdfmetrics

from . import contract_spec as spec


def text_width_cm(text, bold=False, size=spec.SIZE):
    from .render_pdf import _face
    return pdfmetrics.stringWidth(text, _face({"font": spec.FONT, "b": bold}), size) / 28.3464567


CELL_MARGIN = 0.19 * 2        # lề trái + phải mặc định bên trong ô của Word


def label_width_cm(rows, minimum):
    """Bề rộng cột nhãn vừa đủ cho nhãn dài nhất (nhãn hàng đầu được in đậm)."""
    widest = max(text_width_cm(label, bold=(i == 0)) for i, (label, _v) in enumerate(rows))
    content = spec.PAGE["width"] - spec.PAGE["left"] - spec.PAGE["right"]
    return min(max(widest + CELL_MARGIN + 0.3, minimum), content * 0.55)
