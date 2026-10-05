# -*- coding: utf-8 -*-
"""Đọc file .docx thành cấu trúc trung gian để dựng PDF.

Template Word là bản gốc duy nhất. Mọi thứ chỉnh trong Word - thụt lề, cỡ chữ,
đánh số tự động, giãn dòng, in đậm/nghiêng, bảng - đều được đọc ra ở đây, nên
PDF xuất ra giống hệt file Word chứ không dựng lại theo mặc định của code.
"""
from docx import Document
from docx.oxml.ns import qn

TWIP_CM = 2.54 / 1440.0          # 1 twip = 1/1440 inch
EMU_CM = 2.54 / 914400.0
JC = {"both": "justify", "distribute": "justify", "center": "center",
      "right": "right", "end": "right", "left": "left", "start": "left"}


def _find(el, tag):
    return None if el is None else el.find(qn(tag))


def _val(el, tag, attr="w:val"):
    c = _find(el, tag)
    return None if c is None else c.get(qn(attr))


def _attr(el, tag, attr):
    c = _find(el, tag)
    return None if c is None else c.get(qn(attr))


def _onoff(el, tag):
    """Thẻ kiểu <w:b/> nghĩa là bật, <w:b w:val="0"/> nghĩa là tắt."""
    c = _find(el, tag)
    if c is None:
        return None
    v = c.get(qn("w:val"))
    return True if v is None else v not in ("0", "false", "off")


def _num(v, default=None):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


ROMAN = [(1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"),
         (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")]


def _roman(n):
    out = []
    for v, s in ROMAN:
        while n >= v:
            out.append(s)
            n -= v
    return "".join(out)


def _letter(n):
    out = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        out = chr(97 + r) + out
    return out


def fmt_num(n, fmt):
    if fmt == "lowerLetter":
        return _letter(n)
    if fmt == "upperLetter":
        return _letter(n).upper()
    if fmt == "lowerRoman":
        return _roman(n)
    if fmt == "upperRoman":
        return _roman(n).upper()
    if fmt == "decimalZero":
        return "%02d" % n
    if fmt in ("bullet", "none"):
        return ""
    return str(n)


class DocxReader:
    """Phân giải định dạng thực tế của từng đoạn (kế thừa style + numbering)."""

    def __init__(self, document):
        self.d = document
        self.styles, self.def_p, self.def_r = {}, None, None
        self.base_style, self.table_style = None, None
        self.abstract, self.num = {}, {}
        self.counters = {}
        self._load_styles()
        self._load_numbering()

    # ---------------------------------------------------------------- styles
    def _load_styles(self):
        root = self.d.styles.element
        for st in root.findall(qn("w:style")):
            sid = st.get(qn("w:styleId"))
            if sid:
                self.styles[sid] = st
            if st.get(qn("w:default")) in ("1", "true", "on"):
                if st.get(qn("w:type")) == "paragraph":
                    self.base_style = sid
                elif st.get(qn("w:type")) == "table":
                    self.table_style = sid
        dd = _find(root, "w:docDefaults")
        if dd is not None:
            self.def_p = _find(_find(dd, "w:pPrDefault"), "w:pPr")
            self.def_r = _find(_find(dd, "w:rPrDefault"), "w:rPr")

    def _chain(self, style_id):
        """Style cha trước, style con sau - để cái sau ghi đè cái trước."""
        out, seen = [], set()
        while style_id and style_id not in seen:
            seen.add(style_id)
            st = self.styles.get(style_id)
            if st is None:
                break
            out.append(st)
            style_id = _val(st, "w:basedOn")
        return list(reversed(out))

    # ------------------------------------------------------------- numbering
    def _load_numbering(self):
        try:
            root = self.d.part.numbering_part.element
        except Exception:
            return
        for an in root.findall(qn("w:abstractNum")):
            self.abstract[an.get(qn("w:abstractNumId"))] = {
                lv.get(qn("w:ilvl")): lv for lv in an.findall(qn("w:lvl"))}
        for n in root.findall(qn("w:num")):
            aid = _val(n, "w:abstractNumId")
            ov = {}
            for o in n.findall(qn("w:lvlOverride")):
                ov[o.get(qn("w:ilvl"))] = (_val(o, "w:startOverride"), _find(o, "w:lvl"))
            self.num[n.get(qn("w:numId"))] = (aid, ov)

    def _lvl(self, num_id, ilvl):
        ent = self.num.get(num_id)
        if not ent:
            return None, None
        aid, ov = ent
        o = ov.get(ilvl)
        if o and o[1] is not None:
            return o[1], aid
        return self.abstract.get(aid, {}).get(ilvl), aid

    def number(self, num_id, ilvl):
        """Trả về (chuỗi số đã đếm, phần tử w:lvl) cho đoạn đánh số tự động."""
        lvl, aid = self._lvl(num_id, ilvl)
        if lvl is None:
            return None, None
        fmt = _val(lvl, "w:numFmt") or "decimal"
        text = _val(lvl, "w:lvlText") or ""
        if fmt == "none":
            return "", lvl
        if fmt == "bullet":
            return text, lvl

        i = int(ilvl or 0)
        start = int(_num(_val(lvl, "w:start"), 1))
        ov = self.num.get(num_id, (None, {}))[1].get(ilvl)
        if ov and ov[0] and (aid, i) not in self.counters:
            start = int(ov[0])
        self.counters[(aid, i)] = self.counters.get((aid, i), start - 1) + 1
        for key in [k for k in self.counters if k[0] == aid and k[1] > i]:
            del self.counters[key]

        for n in range(1, 10):
            ph = "%%%d" % n
            if ph not in text:
                continue
            sub, _ = self._lvl(num_id, str(n - 1))
            sfmt = (_val(sub, "w:numFmt") if sub is not None else None) or "decimal"
            val = self.counters.get((aid, n - 1))
            if val is None:
                val = int(_num(_val(sub, "w:start") if sub is not None else None, 1))
            text = text.replace(ph, fmt_num(val, sfmt))
        return text, lvl

    # ------------------------------------------------------------ paragraphs
    def para(self, p):
        pPr = _find(p, "w:pPr")
        sid = _val(pPr, "w:pStyle") or self.base_style
        stack = [x for x in [self.def_p] if x is not None]
        stack += [spp for st in self._chain(sid)
                  if (spp := _find(st, "w:pPr")) is not None]

        num_pr = _find(pPr, "w:numPr")
        if num_pr is None:
            for spp in reversed(stack):
                if (n := _find(spp, "w:numPr")) is not None:
                    num_pr = n
                    break
        marker, lvl_el = None, None
        if num_pr is not None:
            num_id = _val(num_pr, "w:numId")
            ilvl = _val(num_pr, "w:ilvl") or "0"
            if num_id and num_id != "0":
                marker, lvl_el = self.number(num_id, ilvl)
        if lvl_el is not None and (lp := _find(lvl_el, "w:pPr")) is not None:
            stack.append(lp)
        if pPr is not None:
            stack.append(pPr)

        b = {"k": "p", "align": "left", "left": 0.0, "right": 0.0, "first": 0.0,
             "before": 0.0, "after": 0.0, "line": 1.0, "line_rule": "auto",
             "keep": False, "keep_lines": False, "break_before": False,
             "borders": {}, "marker": marker or ""}
        for el in stack:
            if (jc := _val(el, "w:jc")) is not None:
                b["align"] = JC.get(jc, "left")
            ind = _find(el, "w:ind")
            if ind is not None:
                for key, names in (("left", ("w:left", "w:start")),
                                   ("right", ("w:right", "w:end"))):
                    for nm in names:
                        if (v := ind.get(qn(nm))) is not None:
                            b[key] = _num(v, 0) * TWIP_CM
                if (v := ind.get(qn("w:hanging"))) is not None:
                    b["first"] = -_num(v, 0) * TWIP_CM
                elif (v := ind.get(qn("w:firstLine"))) is not None:
                    b["first"] = _num(v, 0) * TWIP_CM
            sp = _find(el, "w:spacing")
            if sp is not None:
                if (v := sp.get(qn("w:before"))) is not None:
                    b["before"] = _num(v, 0) / 20.0
                if (v := sp.get(qn("w:after"))) is not None:
                    b["after"] = _num(v, 0) / 20.0
                if (v := sp.get(qn("w:line"))) is not None:
                    rule = sp.get(qn("w:lineRule")) or "auto"
                    b["line_rule"] = rule
                    b["line"] = _num(v, 240) / (240.0 if rule == "auto" else 20.0)
            for flag, tag in (("keep", "w:keepNext"), ("keep_lines", "w:keepLines"),
                              ("break_before", "w:pageBreakBefore")):
                if (v := _onoff(el, tag)) is not None:
                    b[flag] = v
            if (bd := _find(el, "w:pBdr")) is not None:
                b["borders"] = self._borders(bd)

        b["runs"], b["page_break"] = self.runs(p, sid, pPr)
        return b

    # ------------------------------------------------------------------ runs
    def runs(self, p, style_id, pPr):
        base = [x for x in [self.def_r] if x is not None]
        base += [srp for st in self._chain(style_id)
                 if (srp := _find(st, "w:rPr")) is not None]
        mark = _find(pPr, "w:rPr")
        if mark is not None:
            base.append(mark)

        out, page_break = [], False

        def walk(node):
            nonlocal page_break
            for ch in node:
                tag = ch.tag.split("}")[-1]
                if tag == "r":
                    txt, brk = self._run_text(ch)
                    page_break = page_break or brk
                    if txt:
                        out.append(dict(self._run_fmt(ch, base), t=txt))
                elif tag in ("hyperlink", "ins", "smartTag", "sdtContent", "bookmarkStart"):
                    walk(ch)
                elif tag == "sdt":
                    walk(ch)

        walk(p)
        return out, page_break

    @staticmethod
    def _run_text(r):
        if _onoff(_find(r, "w:rPr"), "w:vanish"):
            return "", False
        parts, page_break = [], False
        for ch in r:
            tag = ch.tag.split("}")[-1]
            if tag == "t":
                parts.append(ch.text or "")
            elif tag == "tab":
                parts.append("\t")
            elif tag in ("cr", "br"):
                if ch.get(qn("w:type")) == "page":
                    page_break = True
                else:
                    parts.append("\n")
            elif tag == "noBreakHyphen":
                parts.append("\u2011")
        return "".join(parts), page_break

    def _run_fmt(self, r, base):
        rPr = _find(r, "w:rPr")
        stack = list(base)
        if (cs := _val(rPr, "w:rStyle")):
            stack += [srp for st in self._chain(cs)
                      if (srp := _find(st, "w:rPr")) is not None]
        if rPr is not None:
            stack.append(rPr)

        f = {"b": False, "i": False, "u": False, "size": 13.0,
             "font": None, "sup": None, "color": None}
        for el in stack:
            for key, tag in (("b", "w:b"), ("i", "w:i")):
                if (v := _onoff(el, tag)) is not None:
                    f[key] = v
            if (v := _val(el, "w:u")) is not None:
                f["u"] = v not in ("none", "0")
            if (v := _val(el, "w:sz")) is not None:
                f["size"] = _num(v, 26) / 2.0
            if (v := _attr(el, "w:rFonts", "w:ascii")) is not None:
                f["font"] = v
            if (v := _val(el, "w:vertAlign")) is not None:
                f["sup"] = v if v in ("superscript", "subscript") else None
            if (v := _val(el, "w:color")) is not None:
                f["color"] = None if v in ("auto", "000000") else "#" + v
        return f

    # ---------------------------------------------------------------- tables
    def table(self, tbl):
        grid = [_num(gc.get(qn("w:w")), 0) * TWIP_CM
                for gc in _find(tbl, "w:tblGrid").findall(qn("w:gridCol"))] \
            if _find(tbl, "w:tblGrid") is not None else []
        tblPr = _find(tbl, "w:tblPr")
        align = JC.get(_val(tblPr, "w:jc") or "left", "left")
        pad = self._cell_margins(tblPr)

        # Viền bảng có thể đến từ style (vd "Table Grid") rồi mới bị tblPr ghi đè.
        tbl_borders = {}
        for st in self._chain(_val(tblPr, "w:tblStyle") or self.table_style):
            tbl_borders.update(self._borders(_find(_find(st, "w:tblPr"), "w:tblBorders")))
        tbl_borders.update(self._borders(_find(tblPr, "w:tblBorders")))

        rows, heights = [], []
        for tr in tbl.findall(qn("w:tr")):
            trPr = _find(tr, "w:trPr")
            h = _num(_val(trPr, "w:trHeight"), 0) * TWIP_CM
            rule = _attr(trPr, "w:trHeight", "w:hRule")
            heights.append(h if h and rule == "exact" else None)
            cells = []
            for tc in tr.findall(qn("w:tc")):
                tcPr = _find(tc, "w:tcPr")
                span = int(_num(_val(tcPr, "w:gridSpan"), 1))
                vm = _find(tcPr, "w:vMerge")
                vmerge = (vm.get(qn("w:val")) or "continue") if vm is not None else None
                borders = self._borders(_find(tcPr, "w:tcBorders"))
                w = _attr(tcPr, "w:tcW", "w:w")
                dxa = _attr(tcPr, "w:tcW", "w:type") in (None, "dxa")
                fill = _attr(tcPr, "w:shd", "w:fill")
                cells.append({"blocks": self.body(tc), "span": span,
                              "vmerge": vmerge, "borders": borders,
                              "w": _num(w, 0) * TWIP_CM if w and dxa else None,
                              "fill": None if fill in (None, "auto", "FFFFFF") else "#" + fill,
                              "valign": _val(tcPr, "w:vAlign") or "top"})
            rows.append(cells)
        return {"k": "table", "rows": rows, "grid": grid, "align": align,
                "heights": heights, "pad": pad, "borders": tbl_borders}

    def _cell_margins(self, tblPr):
        """Lề trong ô: style bảng quy định trước, tblPr của bảng ghi đè."""
        sources = []
        sid = _val(tblPr, "w:tblStyle") or self.table_style
        for st in self._chain(sid):
            sources.append(_find(_find(st, "w:tblPr"), "w:tblCellMar"))
        sources.append(_find(tblPr, "w:tblCellMar"))

        pad = {"top": 0.0, "bottom": 0.0, "left": 0.19, "right": 0.19}
        for mar in sources:
            if mar is None:
                continue
            for side in pad:
                if (v := _attr(mar, "w:" + side, "w:w")) is not None:
                    pad[side] = _num(v, 0) * TWIP_CM
        return pad

    @staticmethod
    def _borders(el):
        """{cạnh: (độ dày pt, màu, khoảng cách pt)}; None nghĩa là cạnh đó không kẻ."""
        out = {}
        if el is None:
            return out
        for side in ("top", "bottom", "left", "right", "start", "end",
                     "between", "insideH", "insideV"):
            b = _find(el, "w:" + side)
            if b is None:
                continue
            key = {"start": "left", "end": "right"}.get(side, side)
            if b.get(qn("w:val")) in (None, "nil", "none"):
                out[key] = None
                continue
            color = b.get(qn("w:color"))
            out[key] = (max(_num(b.get(qn("w:sz")), 4) / 8.0, 0.25),
                        "#000000" if color in (None, "auto") else "#" + color,
                        _num(b.get(qn("w:space")), 0))
        return out

    # ------------------------------------------------------------------ body
    def body(self, parent):
        out = []
        for ch in parent:
            tag = ch.tag.split("}")[-1]
            if tag == "p":
                out.append(self.para(ch))
            elif tag == "tbl":
                out.append(self.table(ch))
        return out


def page_setup(document):
    """Khổ giấy và lề lấy từ chính file Word, đổi lề trong Word là PDF đổi theo."""
    sec = document.sections[0]
    get = lambda v, d: (v.cm if v is not None else d)  # noqa: E731
    return {"width": get(sec.page_width, 21.0), "height": get(sec.page_height, 29.7),
            "top": get(sec.top_margin, 2.0), "bottom": get(sec.bottom_margin, 2.0),
            "left": get(sec.left_margin, 3.0), "right": get(sec.right_margin, 2.0)}


def read(src):
    """src: đường dẫn, bytes hoặc file-like của một file .docx đã điền dữ liệu."""
    import io
    doc = Document(io.BytesIO(src) if isinstance(src, (bytes, bytearray)) else src)
    reader = DocxReader(doc)
    return {"page": page_setup(doc), "blocks": reader.body(doc.element.body)}
