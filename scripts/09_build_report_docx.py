"""Build the Word version of the technical report (report/PathSnap_Technical_Report.docx).

The DOCX is generated from report/PathSnap_Technical_Report.html (written by scripts/08_build_report.py),
so the Word and PDF versions always contain exactly the same text, tables and figures.
Formatting: A4, 2.54 cm margins, Times New Roman throughout (body 12 pt, 1.15 line spacing, justified),
numbered section headings, captioned figures, tables with a repeated shaded header row, page numbers.

    python scripts/08_build_report.py        # (re)generate the HTML first
    python scripts/09_build_report_docx.py
"""
import _bootstrap  # noqa: F401

import re
from html.parser import HTMLParser
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor

from pathsnap.config import ROOT

REP = ROOT / "report"
SRC = REP / "PathSnap_Technical_Report.html"
OUT = REP / "PathSnap_Technical_Report.docx"
FONT = "Times New Roman"
TEXT_WIDTH_IN = 6.27          # A4 (8.27 in) minus 2 x 1 in margins
VOID = {"img", "br", "meta", "hr", "link"}


# --------------------------------------------------------------------------- tiny HTML -> tree
class Node:
    def __init__(self, tag, attrs=None, parent=None):
        self.tag, self.attrs, self.parent, self.children = tag, dict(attrs or []), parent, []

    def text(self):
        return "".join(c if isinstance(c, str) else c.text() for c in self.children)


class TreeBuilder(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.root = Node("root"); self.cur = self.root; self.skip = 0

    def handle_starttag(self, tag, attrs):
        if tag in ("style", "title", "head"):
            self.skip += 1; return
        n = Node(tag, attrs, self.cur); self.cur.children.append(n)
        if tag not in VOID:
            self.cur = n

    def handle_endtag(self, tag):
        if tag in ("style", "title", "head"):
            self.skip -= 1; return
        node = self.cur
        while node is not self.root and node.tag != tag:
            node = node.parent
        if node is not self.root:
            self.cur = node.parent

    def handle_data(self, data):
        if not self.skip:
            self.cur.children.append(data)


# --------------------------------------------------------------------------- docx helpers
def set_font(run, size=None, bold=None, italic=None, color=None, mono=False):
    run.font.name = FONT
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts"); rpr.append(rfonts)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(a), FONT)
    if size: run.font.size = Pt(size)
    if bold is not None: run.font.bold = bold
    if italic is not None: run.font.italic = italic
    if color: run.font.color.rgb = RGBColor.from_string(color)


def style_document(doc):
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, side, Cm(2.54))
    st = doc.styles["Normal"]
    st.font.name = FONT; st.font.size = Pt(12)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    pf = st.paragraph_format
    pf.line_spacing = 1.15; pf.space_after = Pt(6); pf.space_before = Pt(0)
    for name, size, before, after in [("Heading 1", 16, 18, 8), ("Heading 2", 14, 14, 6), ("Heading 3", 12, 10, 4),
                                      ("Title", 22, 0, 12)]:
        s = doc.styles[name]
        s.font.name = FONT; s.font.size = Pt(size); s.font.bold = True; s.font.italic = False
        s.font.color.rgb = RGBColor(0, 0, 0)
        rpr = s.element.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts"); rpr.append(rf)
        for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            rf.set(qn(a), FONT)
        for a in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
            if rf.get(qn(a)) is not None:
                del rf.attrib[qn(a)]
        s.paragraph_format.space_before = Pt(before); s.paragraph_format.space_after = Pt(after)
        s.paragraph_format.keep_with_next = True
    # footer page numbers
    p = sec.footer.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(); set_font(r, 10)
    for kind, txt in (("begin", None), (None, "PAGE"), ("end", None)):
        if kind:
            el = OxmlElement("w:fldChar"); el.set(qn("w:fldCharType"), kind)
        else:
            el = OxmlElement("w:instrText"); el.set(qn("xml:space"), "preserve"); el.text = txt
        r._r.append(el)


def shade(cell, hex_fill):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hex_fill)
    tcpr.append(shd)


def repeat_header(row):
    trpr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader"); el.set(qn("w:val"), "true"); trpr.append(el)


def clean(s):
    return re.sub(r"\s+", " ", s)


def add_inline(par, node, size=12, bold=False, italic=False, color=None):
    """Append the inline content of `node` (text, b, i, code, span, br) to paragraph `par`."""
    for c in node.children:
        if isinstance(c, str):
            t = clean(c)
            if t:
                r = par.add_run(t); set_font(r, size, bold or None, italic or None, color)
        elif c.tag == "br":
            par.add_run().add_break()
        elif c.tag in ("b", "strong"):
            add_inline(par, c, size, True, italic, color)
        elif c.tag in ("i", "em"):
            add_inline(par, c, size, bold, True, color)
        elif c.tag == "code":
            add_inline(par, c, size, bold, italic, "1F3864")
        elif c.tag == "span":
            red = "missing" in c.attrs.get("class", "")
            add_inline(par, c, size, True if red else bold, italic, "C00000" if red else color)
        else:
            add_inline(par, c, size, bold, italic, color)


def first_paragraph_fix(par):
    """Strip a leading space created by HTML whitespace."""
    if par.runs and par.runs[0].text.startswith(" "):
        par.runs[0].text = par.runs[0].text.lstrip()


def add_table(doc, node):
    rows = [tr for sec in node.children if not isinstance(sec, str) for tr in
            ([sec] if sec.tag == "tr" else [x for x in sec.children if not isinstance(x, str) and x.tag == "tr"])]
    grid = [[c for c in tr.children if not isinstance(c, str) and c.tag in ("th", "td")] for tr in rows]
    grid = [g for g in grid if g]
    if not grid:
        return
    ncol = max(len(g) for g in grid)
    font = 10 if ncol <= 4 else 9 if ncol <= 7 else 8 if ncol <= 10 else 7

    # column widths: every column gets at least its longest header word (numeric columns: their longest
    # value) so nothing is broken mid-word; the remaining width goes to columns with longer text.
    def minimum_widths(font):
        char_in = font * 0.58 / 72      # average character width in inches at this font size (TNR, incl. bold)
        mins, extra = [], []
        for j in range(ncol):
            head = clean(grid[0][j].text()).replace("_", " ") if j < len(grid[0]) else ""
            word = max((len(w) for w in head.split()), default=3)
            vals = [clean(g[j].text()) for g in grid[1:] if j < len(g)]
            numeric = bool(vals) and all(re.fullmatch(r"[-+\d.,%NA/ ]*", v) for v in vals)
            vmax = max((len(v) for v in vals), default=3)
            mchars = max(word, vmax if numeric else min(vmax, 8), 3)
            mins.append(mchars * char_in + 0.18)   # + left/right cell margins
            typical = sorted(len(v) for v in vals)[int(0.8 * (len(vals) - 1))] if vals else 3
            extra.append(0 if numeric else max(min(typical, 60) - mchars, 0))
        return mins, extra

    mins, extra = minimum_widths(font)
    avail = TEXT_WIDTH_IN
    landscape = sum(mins) > TEXT_WIDTH_IN * 1.02
    if landscape:                         # too wide for portrait: put this table on a landscape page
        font = 8
        mins, extra = minimum_widths(font)
        avail = 11.69 - 2.0
        set_orientation(doc, landscape=True)
    t = doc.add_table(rows=len(grid), cols=ncol)
    t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER; t.autofit = False
    free = avail - sum(mins)
    if free < 0:
        widths_in = [m * avail / sum(mins) for m in mins]
    else:
        tot_extra = sum(extra) or 1
        widths_in = [m + free * e / tot_extra if sum(extra) else m + free / ncol for m, e in zip(mins, extra)]
    widths = [Inches(w) for w in widths_in]
    # fixed layout + grid column widths so Word respects them
    tblpr = t._tbl.tblPr
    lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); tblpr.append(lay)
    for gc, w in zip(t._tbl.tblGrid.findall(qn("w:gridCol")), widths):
        gc.set(qn("w:w"), str(int(w.inches * 1440)))
    for i, g in enumerate(grid):
        header = all(c.tag == "th" for c in g) or i == 0 and g[0].tag == "th"
        for j in range(ncol):
            cell = t.cell(i, j); cell.width = widths[j]
            par = cell.paragraphs[0]; par.paragraph_format.space_after = Pt(0); par.paragraph_format.line_spacing = 1.0
            if j < len(g):
                if header:   # let header names wrap at word boundaries
                    for s in g[j].children:
                        if isinstance(s, str):
                            g[j].children[g[j].children.index(s)] = s.replace("_", " ")
                add_inline(par, g[j], font, bold=header)
                first_paragraph_fix(par)
            if header:
                shade(cell, "D9E2F3")
        trpr = t.rows[i]._tr.get_or_add_trPr()
        cs = OxmlElement("w:cantSplit"); cs.set(qn("w:val"), "true"); trpr.append(cs)
        if header:
            repeat_header(t.rows[i])
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    if landscape:
        set_orientation(doc, landscape=False)


def set_orientation(doc, landscape):
    """Start a new section (new page) in landscape or portrait A4; margins and footer carry over."""
    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    sec.orientation = WD_ORIENT.LANDSCAPE if landscape else WD_ORIENT.PORTRAIT
    sec.page_width, sec.page_height = (Cm(29.7), Cm(21.0)) if landscape else (Cm(21.0), Cm(29.7))
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, side, Cm(2.54))


def add_figure(doc, node):
    img = next((c for c in node.children if not isinstance(c, str) and c.tag == "img"), None)
    cap = next((c for c in node.children if not isinstance(c, str) and c.tag == "figcaption"), None)
    if img is None:
        return
    path = (REP / img.attrs["src"]).resolve()
    if not path.exists():
        doc.add_paragraph(f"[missing figure {path.name}]"); return
    pct = float(re.search(r"width:\s*([\d.]+)%", node.attrs.get("style", "width:90%")).group(1))
    w = 6.2 if pct >= 85 else 5.4 if pct >= 60 else 4.8
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next = True; p.paragraph_format.space_after = Pt(2)
    p.add_run().add_picture(str(path), width=Inches(w))
    if cap is not None:
        c = doc.add_paragraph(); c.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_inline(c, cap, 10, italic=True); first_paragraph_fix(c)
        c.paragraph_format.space_after = Pt(10)


def add_block(doc, node, state):
    if isinstance(node, str):
        if node.strip():
            p = doc.add_paragraph(); r = p.add_run(clean(node).strip()); set_font(r, 12)
        return
    tag = node.tag
    if tag in ("html", "body", "root"):
        for c in node.children: add_block(doc, c, state)
    elif tag == "div" and "cover" in node.attrs.get("class", ""):
        for _ in range(6): doc.add_paragraph()
        for c in node.children:
            if isinstance(c, str):
                continue
            p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if c.tag == "h1":
                add_inline(p, c, 22, bold=True); p.paragraph_format.space_after = Pt(24)
            else:
                add_inline(p, c, 13); p.paragraph_format.space_after = Pt(14)
            first_paragraph_fix(p)
        doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    elif tag == "h2":
        if state["h2_seen"] and "first" not in node.attrs.get("class", ""):
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        state["h2_seen"] = True
        h = doc.add_heading(level=1); add_inline(h, node, 16, bold=True); first_paragraph_fix(h)
    elif tag == "h3":
        h = doc.add_heading(level=2); add_inline(h, node, 14, bold=True); first_paragraph_fix(h)
    elif tag == "h4":
        h = doc.add_heading(level=3); add_inline(h, node, 12, bold=True); first_paragraph_fix(h)
    elif tag == "p":
        if not node.text().strip():
            return
        p = doc.add_paragraph()
        # justify normal prose; left-align paragraphs with long unbreakable tokens (paths, settings)
        long_token = any(len(tok) > 28 for tok in node.text().split())
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT if long_token else WD_ALIGN_PARAGRAPH.JUSTIFY
        add_inline(p, node, 12); first_paragraph_fix(p)
    elif tag in ("ol", "ul"):
        for i, li in enumerate([c for c in node.children if not isinstance(c, str) and c.tag == "li"], 1):
            p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.left_indent = Cm(0.9); p.paragraph_format.first_line_indent = Cm(-0.6)
            r = p.add_run(f"{i}. " if tag == "ol" else "• "); set_font(r, 12)
            add_inline(p, li, 12)
    elif tag == "table":
        add_table(doc, node)
    elif tag == "figure":
        add_figure(doc, node)
    elif tag == "pre":
        for line in node.text().splitlines():
            p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(0); p.paragraph_format.line_spacing = 1.0
            r = p.add_run(line); set_font(r, 8)
        doc.add_paragraph()
    else:
        for c in node.children: add_block(doc, c, state)


def main():
    tb = TreeBuilder(); tb.feed(SRC.read_text(encoding="utf-8"))
    doc = Document()
    style_document(doc)
    doc.core_properties.title = "PathSnap Technical Report"
    doc.core_properties.author = "Team PT-044"
    add_block(doc, tb.root, {"h2_seen": False})
    doc.save(OUT)
    print("saved", OUT)


if __name__ == "__main__":
    main()
