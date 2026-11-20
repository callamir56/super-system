"""سازنده فایل Word پژوهش — A4، راست‌به‌چپ، با فهرست خودکار، جدول و شکل."""
import os
import sys
import math
import importlib

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
FIG = os.path.join(ROOT, "assets", "figures")

NAVY = RGBColor(0x1C, 0x37, 0x5B)
TEAL = RGBColor(0x2A, 0x82, 0x97)
GRAY = RGBColor(0x3A, 0x42, 0x4E)
MUTED = RGBColor(0x5B, 0x66, 0x74)
PERSIAN_FONT = "B Nazanin"
LATIN_FONT = "Times New Roman"

# ---------------------------------------------------------------- هندسه صفحه
PAGE_W, PAGE_H = 21.0, 29.7            # A4 (cm)
MARGIN_TOP = MARGIN_BOTTOM = 2.5
MARGIN_RIGHT, MARGIN_LEFT = 3.0, 2.5   # در سند راست‌به‌چپ، حاشیه صحافی سمت راست
TEXT_W = PAGE_W - MARGIN_RIGHT - MARGIN_LEFT
TEXT_H = PAGE_H - MARGIN_TOP - MARGIN_BOTTOM
BODY_PT = 14.0
_FA = str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹")


def fa_num(x):
    return str(x).translate(_FA)

LINE_CM = BODY_PT * 1.5 * 1.18 / 28.35  # ارتفاع تقریبی هر خط با فاصله ۱٫۵
CHARS_PER_LINE = 74


# ---------------------------------------------------------------- ابزار XML
def el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn("w:" + k), str(v))
    return e


def para_set_rtl(p):
    pPr = p._p.get_or_add_pPr()
    if pPr.find(qn("w:bidi")) is None:
        pPr.insert_element_before(el("w:bidi"),
                                  "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
                                  "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap",
                                  "w:jc", "w:textDirection", "w:textAlignment",
                                  "w:textboxTightWrap", "w:outlineLvl", "w:divId",
                                  "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange")


def run_set_fonts(run, cs=PERSIAN_FONT, size=None, bold=None, color=None):
    run.font.name = LATIN_FONT
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = el("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), LATIN_FONT)
    rFonts.set(qn("w:hAnsi"), LATIN_FONT)
    rFonts.set(qn("w:cs"), cs)
    if rPr.find(qn("w:rtl")) is None:
        rPr.insert_element_before(el("w:rtl"), "w:cs", "w:em", "w:lang",
                                  "w:eastAsianLayout", "w:specVanish", "w:oMath")
    if size is not None:
        run.font.size = Pt(size)
        szCs = el("w:szCs", val=int(size * 2))
        rPr.append(szCs)
    if bold is not None:
        run.font.bold = bold
        b = el("w:bCs")
        if bold:
            b.set(qn("w:val"), "1")
        rPr.append(b)
    if color is not None:
        run.font.color.rgb = color


def add_field(p, instr, cached="1", size=None, color=None, bold=None):
    r1 = p.add_run()
    r1._r.append(el("w:fldChar", fldCharType="begin"))
    r2 = p.add_run()
    it = el("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = f" {instr} "
    r2._r.append(it)
    r3 = p.add_run()
    r3._r.append(el("w:fldChar", fldCharType="separate"))
    r4 = p.add_run(cached)
    r5 = p.add_run()
    r5._r.append(el("w:fldChar", fldCharType="end"))
    for r in (r1, r2, r3, r4, r5):
        run_set_fonts(r, size=size, bold=bold, color=color)
    return r4


def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(el("w:shd", val="clear", color="auto", fill=hexcolor))


def cell_margins(table, top=40, bottom=40, left=90, right=90):
    tblPr = table._tbl.tblPr
    mar = el("w:tblCellMar")
    for side, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        mar.append(el("w:" + side, w=val, type="dxa"))
    tblPr.append(mar)


def table_borders(table, color="B9C3D0", sz=6):
    tblPr = table._tbl.tblPr
    borders = el("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        borders.append(el("w:" + side, val="single", sz=sz, space=0, color=color))
    tblPr.append(borders)


def keep_with_next(p):
    p.paragraph_format.keep_with_next = True


# ---------------------------------------------------------------- پیکربندی سند
def setup_document():
    doc = Document()

    # --- Normal
    st = doc.styles["Normal"]
    st.font.name = LATIN_FONT
    st.font.size = Pt(BODY_PT)
    rPr = st.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = el("w:rFonts"); rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), LATIN_FONT)
    rFonts.set(qn("w:hAnsi"), LATIN_FONT)
    rFonts.set(qn("w:cs"), PERSIAN_FONT)
    rPr.append(el("w:szCs", val=int(BODY_PT * 2)))
    rPr.append(el("w:rtl"))
    pPr = st.element.get_or_add_pPr()
    pPr.insert_element_before(el("w:bidi"),
                              "w:spacing", "w:ind", "w:contextualSpacing", "w:jc", "w:rPr")
    st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    st.paragraph_format.line_spacing = 1.5
    st.paragraph_format.space_before = Pt(0)
    st.paragraph_format.space_after = Pt(6)

    # --- عنوان‌ها
    specs = {
        "Heading 1": dict(size=18, color=NAVY, before=0, after=14, page_break=True),
        "Heading 2": dict(size=14.5, color=TEAL, before=14, after=6, page_break=False),
        "Heading 3": dict(size=13.5, color=GRAY, before=10, after=4, page_break=False),
    }
    for name, s in specs.items():
        h = doc.styles[name]
        h.font.size = Pt(s["size"])
        h.font.bold = True
        h.font.color.rgb = s["color"]
        h.font.name = LATIN_FONT
        hrPr = h.element.get_or_add_rPr()
        hrf = hrPr.find(qn("w:rFonts"))
        if hrf is None:
            hrf = el("w:rFonts"); hrPr.insert(0, hrf)
        hrf.set(qn("w:ascii"), LATIN_FONT)
        hrf.set(qn("w:hAnsi"), LATIN_FONT)
        hrf.set(qn("w:cs"), PERSIAN_FONT)
        h.element.get_or_add_pPr().insert_element_before(
            el("w:bidi"), "w:spacing", "w:ind", "w:jc", "w:rPr")
        h.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        h.paragraph_format.space_before = Pt(s["before"])
        h.paragraph_format.space_after = Pt(s["after"])
        h.paragraph_format.line_spacing = 1.25
        h.paragraph_format.keep_with_next = True
        if s["page_break"]:
            h.paragraph_format.page_break_before = True

    # --- شرح جدول/شکل
    cap = doc.styles["Caption"]
    cap.font.size = Pt(10.5)
    cap.font.bold = False
    cap.font.italic = False
    cap.font.color.rgb = MUTED
    cap.font.name = LATIN_FONT
    caprPr = cap.element.get_or_add_rPr()
    caprf = caprPr.find(qn("w:rFonts"))
    if caprf is None:
        caprf = el("w:rFonts"); caprPr.insert(0, caprf)
    caprf.set(qn("w:ascii"), LATIN_FONT)
    caprf.set(qn("w:hAnsi"), LATIN_FONT)
    caprf.set(qn("w:cs"), PERSIAN_FONT)
    cap.element.get_or_add_pPr().insert_element_before(
        el("w:bidi"), "w:spacing", "w:ind", "w:jc", "w:rPr")
    cap.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_before = Pt(4)
    cap.paragraph_format.space_after = Pt(10)

    # --- صفحه
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(PAGE_W), Cm(PAGE_H)
    sec.top_margin, sec.bottom_margin = Cm(MARGIN_TOP), Cm(MARGIN_BOTTOM)
    sec.right_margin, sec.left_margin = Cm(MARGIN_RIGHT), Cm(MARGIN_LEFT)
    sec.header_distance, sec.footer_distance = Cm(1.1), Cm(1.1)
    sectPr = sec._sectPr
    sectPr.append(el("w:bidi"))
    sectPr.append(el("w:titlePg"))
    sec.different_first_page_header_footer = True
    return doc


def add_header_footer(doc, running_title):
    sec = doc.sections[0]
    # سرصفحه
    hp = sec.header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para_set_rtl(hp)
    r = hp.add_run(running_title)
    run_set_fonts(r, size=9.5, color=MUTED)
    pPr = hp._p.get_or_add_pPr()
    bdr = el("w:pBdr")
    bdr.append(el("w:bottom", val="single", sz=4, space=4, color="B9C3D0"))
    pPr.append(bdr)

    # پاصفحه
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para_set_rtl(fp)
    r = fp.add_run("پژوهش گروهی پایه دهم تجربی  |  صفحه ")
    run_set_fonts(r, size=9.5, color=MUTED)
    add_field(fp, "PAGE", "1", size=9.5, color=MUTED)
    r = fp.add_run(" از ")
    run_set_fonts(r, size=9.5, color=MUTED)
    add_field(fp, "NUMPAGES", "50", size=9.5, color=MUTED)


# ---------------------------------------------------------------- اجزای محتوا
def add_para(doc, text, style=None, size=None, bold=False, color=None,
             align=None, space_after=None, space_before=None, first_line=0.7):
    p = doc.add_paragraph(style=style)
    para_set_rtl(p)
    if align is not None:
        p.alignment = align
    if space_after is not None:
        p.paragraph_format.space_after = Pt(space_after)
    if space_before is not None:
        p.paragraph_format.space_before = Pt(space_before)
    if first_line:
        p.paragraph_format.first_line_indent = Cm(first_line)
    if text:
        r = p.add_run(text)
        run_set_fonts(r, size=size, bold=bold, color=color)
    return p


def add_bullets(doc, items, marker="•"):
    for it in items:
        p = doc.add_paragraph()
        para_set_rtl(p)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.right_indent = Cm(0.4)
        p.paragraph_format.left_indent = Cm(0.6)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.4
        r = p.add_run(f"{marker} ")
        run_set_fonts(r, bold=True, color=TEAL)
        r = p.add_run(it)
        run_set_fonts(r)


def add_heading(doc, text, level):
    p = doc.add_paragraph(style=f"Heading {level}")
    para_set_rtl(p)
    r = p.add_run(text)
    return p


def add_caption(doc, text):
    p = doc.add_paragraph(style="Caption")
    para_set_rtl(p)
    r = p.add_run(text)
    run_set_fonts(r, size=10.5, color=MUTED)
    return p


def add_figure(doc, filename, caption, width_cm=14.5):
    p = doc.add_paragraph()
    para_set_rtl(p)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(os.path.join(FIG, filename), width=Cm(width_cm))
    add_caption(doc, caption)


def add_table(doc, caption, header, rows, widths=None, font_size=11, note=None):
    add_caption(doc, caption)
    t = doc.add_table(rows=1, cols=len(header))
    t.alignment = 1  # center
    t.autofit = False
    tblPr = t._tbl.tblPr
    tblPr.append(el("w:bidiVisual"))
    table_borders(t)
    cell_margins(t)
    total = TEXT_W
    if widths is None:
        widths = [total / len(header)] * len(header)
    hdr = t.rows[0].cells
    for i, h in enumerate(header):
        cell = hdr[i]
        cell.width = Cm(widths[i])
        shade(cell, "E7EDF4")
        p = cell.paragraphs[0]
        para_set_rtl(p)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.15
        r = p.add_run(h)
        run_set_fonts(r, size=font_size, bold=True, color=NAVY)
    trPr = t.rows[0]._tr.get_or_add_trPr()
    trPr.append(el("w:tblHeader"))
    for row in rows:
        cells = t.add_row().cells
        for i, val in enumerate(row):
            cell = cells[i]
            cell.width = Cm(widths[i])
            p = cell.paragraphs[0]
            para_set_rtl(p)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i > 0 else WD_ALIGN_PARAGRAPH.RIGHT
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.line_spacing = 1.15
            r = p.add_run(str(val))
            run_set_fonts(r, size=font_size)
    for row in t.rows:
        for i, c in enumerate(row.cells):
            c.width = Cm(widths[i])
    if note:
        p = add_para(doc, note, size=9.5, color=MUTED,
                     align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=10, first_line=0)
        p.paragraph_format.line_spacing = 1.15
    else:
        doc.add_paragraph().paragraph_format.space_after = Pt(4)


# ---------------------------------------------------------------- صفحه عنوان
def build_title_page(doc, meta):
    def centered(text, size, color, bold=True, after=8, before=0):
        p = doc.add_paragraph()
        para_set_rtl(p)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(after)
        p.paragraph_format.space_before = Pt(before)
        p.paragraph_format.line_spacing = 1.2
        r = p.add_run(text)
        run_set_fonts(r, size=size, bold=bold, color=color)
        return p

    centered(meta["bismillah"], 13, TEAL, bold=False, before=18, after=52)
    # خط تزئینی
    p = doc.add_paragraph()
    para_set_rtl(p)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(14)
    pPr = p._p.get_or_add_pPr()
    bdr = el("w:pBdr")
    bdr.append(el("w:bottom", val="single", sz=12, space=1, color="2A8297"))
    pPr.append(bdr)

    centered(meta["doctype"], 15, MUTED, after=30)
    centered(meta["title"], 24, NAVY, after=6)
    centered(meta["subtitle"], 13, GRAY, bold=False, after=44)

    for line in meta["credits"]:
        centered(line, 12.5, GRAY, bold=False, after=4)
    centered(meta["teacher"], 12.5, GRAY, bold=False, after=4)
    centered(meta["year"], 12.5, TEAL, bold=False, before=28)


# ---------------------------------------------------------------- فهرست خودکار
def build_toc(doc, entries):
    """فهرست مطالب با فیلد TOC و ورودی‌های ذخیره‌شده (Word با F9 به‌روزش می‌کند)."""
    p = doc.add_paragraph()
    para_set_rtl(p)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(18)
    r = p.add_run("فهرست مطالب")
    run_set_fonts(r, size=18, bold=True, color=NAVY)

    first = True
    for (level, title, page) in entries:
        p = doc.add_paragraph()
        para_set_rtl(p)
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.space_after = Pt(1.5)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.right_indent = Cm(0 if level == 1 else 0.7)
        p.paragraph_format.tab_stops.add_tab_stop(
            Cm(TEXT_W - 0.1), WD_TAB_ALIGNMENT.LEFT, WD_TAB_LEADER.DOTS)
        if first:
            r = p.add_run()
            r._r.append(el("w:fldChar", fldCharType="begin"))
            r = p.add_run()
            it = el("w:instrText")
            it.set(qn("xml:space"), "preserve")
            it.text = ' TOC \\o "1-2" \\h \\z \\u '
            r._r.append(it)
            r = p.add_run()
            r._r.append(el("w:fldChar", fldCharType="separate"))
            first = False
        r = p.add_run(title)
        run_set_fonts(r, size=12 if level == 1 else 11.5,
                      bold=(level == 1), color=NAVY if level == 1 else GRAY)
        r = p.add_run("\t")
        r = p.add_run(fa_num(page))
        run_set_fonts(r, size=11.5, bold=(level == 1), color=NAVY if level == 1 else GRAY)

    p = doc.add_paragraph()
    r = p.add_run()
    r._r.append(el("w:fldChar", fldCharType="end"))


def enable_update_fields(doc):
    settings = doc.settings.element
    settings.append(el("w:updateFields", val="true"))


# ---------------------------------------------------------------- تخمین صفحه
class Pager:
    def __init__(self):
        self.page = 1
        self.y = 0.0  # سانتی‌متر مصرف‌شده در صفحه جاری

    def _advance(self, cm):
        # اگر بلوک از یک صفحه بلندتر باشد، به‌درستی روی چند صفحه پخش می‌شود
        while self.y + cm > TEXT_H:
            cm -= (TEXT_H - self.y)
            self.page += 1
            self.y = 0.0
            if cm > TEXT_H:
                full = int(cm // TEXT_H)
                self.page += full
                cm -= full * TEXT_H
        self.y += cm

    def _new_page(self):
        self.page += 1
        self.y = 0.0

    def block(self, kind, data=None):
        if kind == "h1":
            self._new_page()
            self.y = 0
            self._advance(2.2)
            return self.page
        if kind == "h2":
            self._advance(1.6)
            return self.page
        if kind == "h3":
            self._advance(1.3)
            return self.page
        if kind == "p":
            n = max(1, math.ceil(len(data) / CHARS_PER_LINE))
            self._advance(n * LINE_CM + 0.22)
            return self.page
        if kind == "bullets":
            for it in data:
                n = max(1, math.ceil(len(it) / (CHARS_PER_LINE - 4)))
                self._advance(n * LINE_CM + 0.12)
            return self.page
        if kind == "figure":
            h = data.get("height_cm", 9.0)
            self._advance(h + 1.5)
            return self.page
        if kind == "table":
            self._advance(1.5)                      # شرح جدول
            for _ in range(len(data["rows"]) + 1):  # سطرها، یک‌به‌یک برای شکست صفحه
                self._advance(1.25)
            self._advance(1.0)                      # فاصله پس از جدول
            return self.page
        return self.page


# ---------------------------------------------------------------- ساخت سند
def render_block(doc, block):
    kind = block[0]
    if kind == "h1":
        add_heading(doc, block[1], 1)
    elif kind == "h2":
        add_heading(doc, block[1], 2)
    elif kind == "h3":
        add_heading(doc, block[1], 3)
    elif kind == "p":
        add_para(doc, block[1])
    elif kind == "lead":
        add_para(doc, block[1], size=14, color=GRAY)
    elif kind == "bullets":
        add_bullets(doc, block[1])
    elif kind == "numbered":
        add_bullets(doc, block[1], marker="◾")
    elif kind == "figure":
        d = block[1]
        add_figure(doc, d["file"], d["caption"], d.get("width", 14.5))
    elif kind == "table":
        d = block[1]
        add_table(doc, d["caption"], d["header"], d["rows"],
                  widths=d.get("widths"), font_size=d.get("font_size", 11),
                  note=d.get("note"))
    elif kind == "refs":
        for item in block[1]:
            p = add_para(doc, item, size=12, align=WD_ALIGN_PARAGRAPH.RIGHT,
                         space_after=6, first_line=0)
            p.paragraph_format.left_indent = Cm(0.8)
            p.paragraph_format.line_spacing = 1.3
    elif kind == "pagebreak":
        doc.add_page_break()


def build():
    sys.path.insert(0, BASE)
    meta = importlib.import_module("content.ch_front").META
    modules = ["ch_front", "ch1", "ch2", "ch3", "ch4", "ch5", "ch6", "ch7", "ch_back"]
    blocks = []
    for m in modules:
        blocks += importlib.import_module(f"content.{m}").BLOCKS

    # --- محاسبه شماره صفحه‌ها برای فهرست
    pager = Pager()
    pager.page = 1                      # صفحه عنوان = ۱
    pager.y = TEXT_H * 0.25
    toc_entries = []
    toc_pages = 2                       # دو صفحه فهرست
    for _ in range(toc_pages):
        pager._new_page()
    for b in blocks:
        p = pager.block(b[0], b[1] if len(b) > 1 else None)
        if b[0] == "h1":
            toc_entries.append((1, b[1], p))
        elif b[0] == "h2":
            toc_entries.append((2, b[1], p))

    # --- ساخت سند
    doc2 = setup_document()
    add_header_footer(doc2, meta["running_title"])
    build_title_page(doc2, meta)
    doc2.add_page_break()
    build_toc(doc2, toc_entries)
    for blk in blocks:
        render_block(doc2, blk)

    cp = doc2.core_properties
    cp.title = meta["title"]
    cp.subject = "پژوهش علمی دانش‌آموزی — پایه دهم تجربی"
    cp.author = "گروه پژوهش: گالره مرادپور، ندا بانی‌الشکی، مانیا خزاعی، ترنم حق‌داد رضایی، کوثر صیادی"
    cp.last_modified_by = "گروه پژوهش"
    cp.category = "پژوهش مدرسه‌ای"
    cp.keywords = "هوش مصنوعی؛ بیماری‌های قلبی؛ نوار قلب؛ یادگیری ماشین؛ ارزیابی مدل"
    cp.comments = "نسخه بازنویسی‌شده و گسترش‌یافته پژوهش، با فصل‌بندی، جدول‌ها، شکل‌ها و ارجاع درون‌متنی."
    enable_update_fields(doc2)
    out = os.path.join(ROOT, "پژوهش_قلب_و_هوش_مصنوعی_ویرایش_کامل.docx")
    doc2.save(out)
    import normalize_docx
    print("reordered elements:", normalize_docx.normalize(out))
    print("saved:", out)
    print("blocks:", len(blocks), "toc entries:", len(toc_entries))
    print("estimated pages:", pager.page)
    return out


if __name__ == "__main__":
    build()
