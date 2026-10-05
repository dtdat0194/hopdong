# -*- coding: utf-8 -*-
"""Điền dữ liệu vào template Word bằng docxtpl.

Template trong templates/ là bản gốc duy nhất - bạn sửa trực tiếp trong Word.
build_template() chỉ dùng để dựng bản đầu tiên khi file chưa tồn tại; nó không
bao giờ tự ghi đè template đang có (xem scripts/build_templates.py).
"""
import io
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docxtpl import DocxTemplate

from . import contract_spec as spec

TEMPLATE_DIR = Path(__file__).resolve().parents[1] / "templates"
ALIGN = {
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
}


def _font(run, size=spec.SIZE):
    run.font.name = spec.FONT
    run.font.size = Pt(size)
    run._element.rPr.rFonts.set(qn("w:eastAsia"), spec.FONT)


def _new_doc():
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(spec.PAGE["width"]), Cm(spec.PAGE["height"])
    sec.top_margin, sec.bottom_margin = Cm(spec.PAGE["top"]), Cm(spec.PAGE["bottom"])
    sec.left_margin, sec.right_margin = Cm(spec.PAGE["left"]), Cm(spec.PAGE["right"])
    st = doc.styles["Normal"]
    st.font.name = spec.FONT
    st.font.size = Pt(spec.SIZE)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), spec.FONT)
    st.paragraph_format.space_after = Pt(4)
    st.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    return doc


def _para(doc, b, text=None):
    par = doc.add_paragraph()
    par.alignment = ALIGN[b.get("align", "justify")]
    pf = par.paragraph_format
    pf.space_before, pf.space_after = Pt(b.get("before", 0)), Pt(b.get("after", 4))
    pf.keep_with_next = bool(b.get("keep"))
    if b.get("left") or b.get("hanging"):
        pf.left_indent = Cm(b.get("left", 0))
        if b.get("hanging"):
            pf.first_line_indent = Cm(-b["hanging"])
    body = b["text"] if text is None else text
    if body:
        r = par.add_run(body)
        r.bold, r.italic, r.underline = b.get("bold", False), b.get("italic", False), b.get("underline", False)
        _font(r, b.get("size", spec.SIZE))
    return par


def _party_table(doc, b):
    from .measure import label_width_cm
    t = doc.add_table(rows=len(b["rows"]), cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    t.autofit = False
    content_w = spec.PAGE["width"] - spec.PAGE["left"] - spec.PAGE["right"]
    label_w = label_width_cm(b["rows"], b["label_w"])
    for i, (label, value) in enumerate(b["rows"]):
        for j, (txt, w) in enumerate(((label, label_w), (": " + value, content_w - label_w))):
            cell = t.cell(i, j)
            cell.width = Cm(w)
            par = cell.paragraphs[0]
            par.paragraph_format.space_before, par.paragraph_format.space_after = Pt(0), Pt(2)
            r = par.add_run(txt)
            r.bold = i == 0
            _font(r)
    return t


def _sign_table(doc):
    doc.add_paragraph()
    t = doc.add_table(rows=4, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    data = [
        ("ĐẠI DIỆN BÊN A", "BÊN B"),
        ("(Ký, ghi rõ họ tên)", "(Ký, ghi rõ họ tên)"),
        ("", ""),
        ("{{ dai_dien }}", "{{ ho_ten }}"),
    ]
    for i, pair in enumerate(data):
        for j, txt in enumerate(pair):
            cell = t.cell(i, j)
            cell.width = Cm(8)
            par = cell.paragraphs[0]
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER
            par.paragraph_format.space_after = Pt(36 if i == 2 else 2)
            r = par.add_run(txt)
            r.bold, r.italic = i in (0, 3), i == 1
            _font(r)
    par = t.cell(3, 0).add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = par.add_run("{{ chuc_vu_dai_dien }}")
    r.bold = True
    _font(r)


def build_template(loai, dest=None):
    """Sinh file template .docx còn nguyên placeholder {{ ... }} cho docxtpl."""
    doc = _new_doc()
    for b in spec.CONTRACTS[loai]["blocks"]():
        if b["k"] == "table":
            _party_table(doc, b)
        elif b["k"] == "sign":
            _sign_table(doc)
        else:
            if b.get("cond"):
                _para(doc, dict(b, bold=False, italic=False), text="{%p if " + b["cond"] + " %}")
                _para(doc, b)
                _para(doc, dict(b, bold=False, italic=False), text="{%p endif %}")
            else:
                _para(doc, b)
    dest = dest or TEMPLATE_DIR / spec.CONTRACTS[loai]["template"]
    dest.parent.mkdir(parents=True, exist_ok=True)
    doc.save(dest)
    return dest


def ensure_templates():
    for loai in spec.CONTRACTS:
        path = TEMPLATE_DIR / spec.CONTRACTS[loai]["template"]
        if not path.exists():
            build_template(loai, path)
    return TEMPLATE_DIR


def render_docx(loai, ctx):
    """Trả về bytes của file Word đã điền dữ liệu."""
    ensure_templates()
    tpl = DocxTemplate(TEMPLATE_DIR / spec.CONTRACTS[loai]["template"])
    tpl.render(ctx)
    buf = io.BytesIO()
    tpl.save(buf)
    return buf.getvalue()
