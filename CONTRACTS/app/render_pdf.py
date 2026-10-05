# -*- coding: utf-8 -*-
"""Xuất hợp đồng PDF từ file Word đã điền dữ liệu (không cần Word hay LibreOffice).

Luồng: template .docx -> docxtpl điền dữ liệu -> docx_read đọc định dạng -> PDF.
Định dạng lấy nguyên từ file Word nên sửa template trong Word là PDF đổi theo.
"""
import struct
from html import escape
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (KeepTogether, PageBreak, Paragraph, SimpleDocTemplate,
                                Spacer, Table, TableStyle)

from . import docx_read
from .render_docx import render_docx

FONT_DIRS = [
    Path("/System/Library/Fonts/Supplemental"),
    Path("/Library/Fonts"),
    Path.home() / "Library/Fonts",
    Path("C:/Windows/Fonts"),
    Path("/usr/share/fonts/truetype/msttcorefonts"),
    # Chốt chặn cuối: Tinos đi kèm repo, trùng khít metric với Times New Roman
    # (cùng ascent, descent, lineGap và bề rộng từng chữ). Bản chạy trong trình
    # duyệt không với tới font hệ thống nhưng vẫn dàn trang y hệt bản trên máy.
    Path(__file__).resolve().parents[1] / "fonts",
]
FAMILIES = {
    "Times New Roman": {
        "": ["Times New Roman.ttf", "times.ttf", "Times_New_Roman.ttf", "Tinos-Regular.ttf"],
        "b": ["Times New Roman Bold.ttf", "timesbd.ttf", "Tinos-Bold.ttf"],
        "i": ["Times New Roman Italic.ttf", "timesi.ttf", "Tinos-Italic.ttf"],
        "bi": ["Times New Roman Bold Italic.ttf", "timesbi.ttf", "Tinos-BoldItalic.ttf"],
    },
    "Arial": {
        "": ["Arial.ttf", "arial.ttf"],
        "b": ["Arial Bold.ttf", "arialbd.ttf"],
        "i": ["Arial Italic.ttf", "ariali.ttf"],
        "bi": ["Arial Bold Italic.ttf", "arialbi.ttf"],
    },
}
FALLBACK = {
    "": ["Arial Unicode.ttf", "Arial.ttf", "DejaVuSerif.ttf"],
    "b": ["Arial Bold.ttf", "Arial Unicode.ttf", "DejaVuSerif-Bold.ttf"],
    "i": ["Arial Italic.ttf", "Arial Unicode.ttf", "DejaVuSerif-Italic.ttf"],
    "bi": ["Arial Bold Italic.ttf", "Arial Unicode.ttf", "DejaVuSerif-BoldItalic.ttf"],
}
DEFAULT_FAMILY = "Times New Roman"
ALIGN = {"justify": TA_JUSTIFY, "left": TA_LEFT, "center": TA_CENTER, "right": TA_RIGHT}
VALIGN = {"top": "TOP", "center": "MIDDLE", "bottom": "BOTTOM", "both": "MIDDLE"}
_ready = set()
_single = {}                     # font -> hệ số giãn dòng "single"


def _line_factor(path):
    """Giãn dòng 'single' của Word = (ascent - descent + lineGap) / unitsPerEm của chính font đó.

    Times New Roman ra 1.1499; đoán bừa một con số là bản PDF sẽ dài ngắn khác Word.
    """
    try:
        data = Path(path).read_bytes()
        tabs = {}
        for i in range(struct.unpack(">H", data[4:6])[0]):
            o = 12 + 16 * i
            tabs[data[o:o + 4].decode("latin-1")] = struct.unpack(">I", data[o + 8:o + 12])[0]
        upm = struct.unpack(">H", data[tabs["head"] + 18:tabs["head"] + 20])[0]
        asc, desc, gap = struct.unpack(">hhh", data[tabs["hhea"] + 4:tabs["hhea"] + 10])
        return (asc - desc + gap) / upm
    except Exception:
        return 1.15


def _find(names):
    for d in FONT_DIRS:
        for n in names:
            if (f := d / n).exists():
                return str(f)
    return None


def register_family(name):
    """Nạp 4 kiểu chữ của một font; font lạ thì dùng tạm font mặc định."""
    if name in _ready:
        return name
    faces = FAMILIES.get(name)
    if faces is None:
        return register_family(DEFAULT_FAMILY)
    paths = {k: _find(v) or _find(FALLBACK[k]) for k, v in faces.items()}
    if not paths[""]:
        raise RuntimeError(f"Không tìm thấy font {name}. Cần cài Times New Roman hoặc Arial.")
    for suffix, path in paths.items():
        pdfmetrics.registerFont(TTFont(f"{name}{suffix and '-' + suffix}",
                                       path or paths[""]))
    pdfmetrics.registerFontFamily(name, normal=name, bold=f"{name}-b",
                                  italic=f"{name}-i", boldItalic=f"{name}-bi")
    _single[name] = _line_factor(paths[""])
    _ready.add(name)
    return name


def _face(f):
    """Tên font cụ thể theo đậm/nghiêng, dùng để đo bề rộng chuỗi."""
    fam = register_family(f.get("font") or DEFAULT_FAMILY)
    suffix = ("b" if f.get("b") else "") + ("i" if f.get("i") else "")
    return f"{fam}-{suffix}" if suffix else fam


def _run_markup(r):
    txt = escape(r["t"]).replace("\t", "&nbsp;" * 4).replace("\n", "<br/>")
    if r.get("b"):
        txt = f"<b>{txt}</b>"
    if r.get("i"):
        txt = f"<i>{txt}</i>"
    if r.get("u"):
        txt = f"<u>{txt}</u>"
    if r.get("sup") == "superscript":
        txt = f"<super>{txt}</super>"
    elif r.get("sup") == "subscript":
        txt = f"<sub>{txt}</sub>"
    if r.get("color"):
        txt = f'<font color="{r["color"]}">{txt}</font>'
    return txt


# Word lưu dấu đầu dòng bằng ô ký tự riêng của Wingdings/Symbol; Times New Roman
# không có các glyph đó nên phải đổi sang ký tự Unicode tương đương.
DAU_DONG = {0xF0B7: "•", 0xF0A7: "▪", 0xF0A8: "▫", 0xF06C: "●", 0xF071: "□",
            0xF075: "◆", 0xF0D8: "➢", 0xF0FC: "✓", 0xF0FE: "▪", 0xF02D: "–"}


def _bullet(mark, face):
    co = pdfmetrics.getFont(face).face.charToGlyph
    out = ""
    for ch in mark:
        if 0xF000 <= ord(ch) <= 0xF0FF:
            ch = DAU_DONG.get(ord(ch), "•")
        out += ch if ord(ch) in co else "•"
    return out


def _marker(b, base):
    """Số thứ tự Word tự sinh, chèn kèm khoảng trắng cho chữ thẳng với lề treo."""
    mark = b.get("marker")
    if not mark:
        return ""
    f = b["runs"][0] if b["runs"] else base
    face, size = _face(f), f.get("size", base["size"])
    mark = _bullet(mark, face)
    gap = -b.get("first", 0) * cm - pdfmetrics.stringWidth(mark, face, size)
    space = pdfmetrics.stringWidth("\u00a0", face, size) or size / 4
    pad = "&nbsp;" * max(1, round(gap / space)) if gap > space / 2 else "&nbsp;"
    return _run_markup(dict(f, t=mark)) + pad


class _GiuVoiDoanSau(KeepTogether):
    """Giữ tiêu đề cùng trang với đoạn ngay sau nó, theo cách của Word.

    KeepTogether gốc của ReportLab đẩy cả cụm sang trang mới nếu đoạn sau không
    vừa trọn vẹn, nên mỗi "Điều ..." dài đều làm hụt gần nửa trang. Word chỉ đòi
    tiêu đề đứng cùng vài dòng đầu của đoạn kế tiếp, phần còn lại được tràn trang.
    """

    TOI_THIEU = 2          # số dòng đầu của đoạn sau phải ở lại cùng tiêu đề

    def wrap(self, aW, aH):
        kq = KeepTogether.wrap(self, aW, aH)
        cuoi = self._content[-1]
        if len(self._content) > 1 and isinstance(cuoi, Paragraph):
            can = cuoi.style.leading * self.TOI_THIEU
            self._H -= max(cuoi.height - can, 0)
        return kq


class _Bordered(Paragraph):
    """Đoạn văn có đường kẻ viền (w:pBdr) - Word dùng kiểu này để gạch dưới tiêu ngữ."""

    def __init__(self, text, style, borders):
        Paragraph.__init__(self, text, style)
        self._bd = borders

    def draw(self):
        Paragraph.draw(self)
        x0, x1 = self.style.leftIndent, self.width - self.style.rightIndent
        c = self.canv
        c.saveState()
        for side in ("top", "bottom", "left", "right"):
            if not self._bd.get(side):
                continue
            w, color, space = self._bd[side]
            lo, hi = -space, self.height + space
            pts = {"bottom": (x0, lo, x1, lo), "top": (x0, hi, x1, hi),
                   "left": (x0, lo, x0, hi), "right": (x1, lo, x1, hi)}[side]
            c.setLineWidth(w)
            c.setStrokeColor(colors.HexColor(color))
            c.line(*pts)
        c.restoreState()


def _para(b, width):
    runs = b["runs"]
    base = runs[0] if runs else {"size": 13.0, "font": DEFAULT_FAMILY}
    size = max([r["size"] for r in runs] or [base["size"]])
    fam = register_family(base.get("font") or DEFAULT_FAMILY)
    single = size * _single.get(fam, 1.15)
    leading = b["line"] * (single if b["line_rule"] == "auto" else 1.0)
    if b["line_rule"] == "atLeast":
        leading = max(leading, single)

    st = ParagraphStyle(
        "p", fontName=fam, fontSize=size, leading=leading,
        alignment=ALIGN[b["align"]],
        spaceBefore=b["before"], spaceAfter=b["after"],
        leftIndent=b["left"] * cm, rightIndent=b["right"] * cm,
        firstLineIndent=b["first"] * cm,
        allowWidows=0, allowOrphans=0,
    )
    st.keepWithNext = b["keep"]

    body = "".join(
        f'<font name="{register_family(r.get("font") or fam)}" size="{r["size"]}">'
        f'{_run_markup(r)}</font>' for r in runs)
    text = _marker(b, base) + body
    bd = {k: v for k, v in b["borders"].items() if v}
    if not text.strip() and not bd:
        return Spacer(1, leading + b["before"] + b["after"])
    return _Bordered(text, st, bd) if bd else Paragraph(text, st)


def _cell_flowables(cell, width):
    out = [_block(x, width) for x in cell["blocks"]]
    return [x for x in out if x is not None] or [Spacer(0, 0)]


def _columns(b, width):
    """Word ghi bề rộng ở cả tblGrid lẫn từng ô; bề rộng của ô mới là cái có hiệu lực."""
    ncols = max([len(b["grid"])]
                + [sum(c["span"] for c in row) for row in b["rows"]] or [1])
    grid = (b["grid"] + [0.0] * ncols)[:ncols]
    for row in b["rows"]:
        col = 0
        for cell in row:
            if cell["span"] == 1 and cell["w"]:
                grid[col] = cell["w"]
            col += cell["span"]
    missing = [i for i, g in enumerate(grid) if not g]
    if missing:
        spare = max(width / cm - sum(grid), 0) / len(missing)
        for i in missing:
            grid[i] = spare or width / cm / ncols
    return grid


def _table(b, width):
    grid = _columns(b, width)
    total = sum(grid)
    if total > 0 and total * cm > width + 1:
        grid = [g * (width / cm) / total for g in grid]

    pad = b.get("pad", {})
    inner = [g * cm - (pad.get("left", 0) + pad.get("right", 0)) * cm for g in grid]
    tb = b.get("borders", {})

    data, cmds, vstart = [], [], {}
    for r, row in enumerate(b["rows"]):
        line, col = [], 0
        for cell in row:
            span = cell["span"]
            avail = sum(inner[col:col + span]) or width
            if cell["vmerge"] == "continue" and col in vstart:
                r0 = vstart[col]
                cmds.append(("SPAN", (col, r0), (col + span - 1, r)))
                line.append("")
            else:
                if cell["vmerge"] == "restart":
                    vstart[col] = r
                if span > 1:
                    cmds.append(("SPAN", (col, r), (col + span - 1, r)))
                line.append(_cell_flowables(cell, avail))
            cmds.append(("VALIGN", (col, r), (col, r), VALIGN.get(cell["valign"], "TOP")))
            if cell.get("fill"):
                cmds.append(("BACKGROUND", (col, r), (col + span - 1, r),
                             colors.HexColor(cell["fill"])))
            # Cạnh ngoài lấy viền bảng, cạnh trong lấy insideH/insideV, ô tự đặt thì ưu tiên ô.
            ria = {"top": r == 0, "bottom": r == len(b["rows"]) - 1,
                   "left": col == 0, "right": col + span >= len(grid)}
            for side, cmd in (("top", "LINEABOVE"), ("bottom", "LINEBELOW"),
                              ("left", "LINEBEFORE"), ("right", "LINEAFTER")):
                trong = {"bottom": "insideH", "left": "insideV"}.get(side)
                mac_dinh = tb.get(side) if ria[side] else (tb.get(trong) if trong else None)
                if (nét := cell["borders"].get(side, mac_dinh)):
                    cmds.append((cmd, (col, r), (col + span - 1, r),
                                 nét[0], colors.HexColor(nét[1])))
            line += [""] * (span - 1)
            col += span
        data.append(line or [""])

    ncols = max(len(x) for x in data)
    for line in data:
        line += [""] * (ncols - len(line))
    widths = [g * cm for g in grid[:ncols]]
    widths += [width / max(1, ncols)] * (ncols - len(widths))

    t = Table(data, colWidths=widths, rowHeights=[h and h * cm for h in b["heights"]],
              hAlign={"center": "CENTER", "right": "RIGHT"}.get(b["align"], "LEFT"))
    t.setStyle(TableStyle(cmds + [
        ("LEFTPADDING", (0, 0), (-1, -1), pad.get("left", 0) * cm),
        ("RIGHTPADDING", (0, 0), (-1, -1), pad.get("right", 0) * cm),
        ("TOPPADDING", (0, 0), (-1, -1), pad.get("top", 0) * cm),
        ("BOTTOMPADDING", (0, 0), (-1, -1), pad.get("bottom", 0) * cm),
    ]))
    return t


def _block(b, width):
    return _table(b, width) if b["k"] == "table" else _para(b, width)


def flowables(doc, width):
    out = []
    for b in doc["blocks"]:
        if b["k"] == "p" and b.get("break_before") and out:
            out.append(PageBreak())
        item = _block(b, width)
        if b["k"] == "p" and b.get("keep_lines") and b["runs"]:
            item = KeepTogether(item)
        out.append(item)
        if b["k"] == "p" and b.get("page_break"):
            out.append(PageBreak())
    return out


def build_pdf(items, stream):
    """items = [(loai, ctx), ...] - nhiều hợp đồng nối tiếp nhau, mỗi bản sang trang mới."""
    register_family(DEFAULT_FAMILY)
    docs = [docx_read.read(render_docx(loai, ctx)) for loai, ctx in items]
    page = docs[0]["page"]
    width = (page["width"] - page["left"] - page["right"]) * cm

    doc = SimpleDocTemplate(
        stream,
        pagesize=(page["width"] * cm, page["height"] * cm),
        leftMargin=page["left"] * cm, rightMargin=page["right"] * cm,
        topMargin=page["top"] * cm, bottomMargin=page["bottom"] * cm,
        # Tên công ty lấy từ sheet CongTy trong file Excel, không ghi cứng ở đây:
        # đổi tên công ty trong Excel là siêu dữ liệu PDF đổi theo.
        title="Hợp đồng", author=items[0][1].get("cong_ty", ""),
    )
    doc.keepTogetherClass = _GiuVoiDoanSau
    story = []
    for i, d in enumerate(docs):
        if i:
            story.append(PageBreak())
        story += flowables(d, width)
    doc.build(story)
    return stream
