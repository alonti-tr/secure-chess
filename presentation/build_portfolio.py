from __future__ import annotations

import datetime
from pathlib import Path
from typing import Iterable, List, Sequence, Tuple

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


REPO_ROOT = Path(__file__).resolve().parent.parent
SRC_ROOT = REPO_ROOT / "src" / "secure_chess"
OUT_DOCX = REPO_ROOT / "תיק-פרויקט.docx"

HE_FONT = "Arial"
CODE_FONT = "Consolas"

ACCENT = RGBColor(0x1F, 0x6F, 0xEB)
ACCENT_DARK = RGBColor(0x0E, 0x4A, 0xA8)
DARK = RGBColor(0x1F, 0x23, 0x28)
MUTED = RGBColor(0x57, 0x60, 0x6A)
CODE_BG = RGBColor(0x27, 0x28, 0x22)
CODE_FG = RGBColor(0xF8, 0xF8, 0xF2)
CALLOUT_BG = RGBColor(0xEA, 0xF2, 0xFD)
TIP_BG = RGBColor(0xFF, 0xF8, 0xE1)
TIP_BORDER = RGBColor(0xF5, 0xA6, 0x23)
TABLE_HEAD_BG = RGBColor(0xEA, 0xF2, 0xFD)


def _set_ltr(paragraph) -> None:
    pPr = paragraph._p.get_or_add_pPr()
    bidi = pPr.find(qn("w:bidi"))
    if bidi is None:
        bidi = _insert_in_pPr_ordered(pPr, "bidi")
    bidi.set(qn("w:val"), "0")


def _set_bidi(paragraph) -> None:
    pPr = paragraph._p.get_or_add_pPr()
    _insert_in_pPr_ordered(pPr, "bidi")


def _set_run_rtl(run) -> None:
    rPr = run._r.get_or_add_rPr()
    rtl = _insert_ordered(rPr, "rtl", _RPR_ORDER)
    rtl.set(qn("w:val"), "1")


def _set_cell_shading(cell, hex_color: str) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def _add_paragraph_border(paragraph, side: str, color: str, sz: int = 12) -> None:
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is None:
        pBdr = OxmlElement("w:pBdr")
        pPr.append(pBdr)
    border = OxmlElement(f"w:{side}")
    border.set(qn("w:val"), "single")
    border.set(qn("w:sz"), str(sz))
    border.set(qn("w:space"), "4")
    border.set(qn("w:color"), color)
    pBdr.append(border)


def _set_paragraph_shading(paragraph, hex_color: str) -> None:
    pPr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    pPr.append(shd)


def rgb_hex(rgb: RGBColor) -> str:
    return "{:02X}{:02X}{:02X}".format(rgb[0], rgb[1], rgb[2])


_PPR_ORDER = (
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
    "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
    "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
    "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
    "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr",
    "pPrChange",
)

_SECT_PR_ORDER = (
    "headerReference", "footerReference", "footnotePr", "endnotePr", "type",
    "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType",
    "cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection",
    "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange",
)

_RPR_ORDER = (
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
    "dstrike", "outline", "shadow", "emboss", "imprint", "noProof",
    "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern",
    "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd",
    "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
    "specVanish", "oMath",
)


_TOGGLE_TAGS = frozenset({"bidi", "rtl", "rtlGutter"})


def _insert_ordered(parent, tag_name: str, order: tuple[str, ...]):
    existing = parent.find(qn(f"w:{tag_name}"))
    if existing is not None:
        if tag_name in _TOGGLE_TAGS:
            existing.set(qn("w:val"), "1")
        return existing
    target_idx = order.index(tag_name)
    new_el = OxmlElement(f"w:{tag_name}")
    if tag_name in _TOGGLE_TAGS:
        new_el.set(qn("w:val"), "1")
    for i, child in enumerate(list(parent)):
        local = child.tag.rsplit("}", 1)[-1]
        if local in order and order.index(local) > target_idx:
            parent.insert(i, new_el)
            return new_el
    parent.append(new_el)
    return new_el


def _insert_in_pPr_ordered(pPr, tag_name: str):
    return _insert_ordered(pPr, tag_name, _PPR_ORDER)


def _setup_section_rtl(section) -> None:
    sectPr = section._sectPr
    _insert_ordered(sectPr, "bidi", _SECT_PR_ORDER)
    _insert_ordered(sectPr, "rtlGutter", _SECT_PR_ORDER)


def _style_set_rtl(style, *, align_right: bool = True) -> None:
    style_el = style.element
    pPr = style_el.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        style_el.insert(0, pPr)
    _insert_in_pPr_ordered(pPr, "bidi")
    if align_right:
        jc = _insert_in_pPr_ordered(pPr, "jc")
        jc.set(qn("w:val"), "right")

    rPr = style_el.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        style_el.append(rPr)
    rtl = _insert_ordered(rPr, "rtl", _RPR_ORDER)
    rtl.set(qn("w:val"), "1")


def _set_doc_defaults_rtl(doc: Document) -> None:
    styles_element = doc.styles.element
    docDefaults = styles_element.find(qn("w:docDefaults"))
    if docDefaults is None:
        return

    pPrDefault = docDefaults.find(qn("w:pPrDefault"))
    if pPrDefault is None:
        pPrDefault = OxmlElement("w:pPrDefault")
        docDefaults.append(pPrDefault)
    pPr = pPrDefault.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        pPrDefault.append(pPr)
    _insert_in_pPr_ordered(pPr, "bidi")
    jc = _insert_in_pPr_ordered(pPr, "jc")
    jc.set(qn("w:val"), "right")

    rPrDefault = docDefaults.find(qn("w:rPrDefault"))
    if rPrDefault is None:
        rPrDefault = OxmlElement("w:rPrDefault")
        docDefaults.append(rPrDefault)
    rPr = rPrDefault.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        rPrDefault.append(rPr)
    rtl = _insert_ordered(rPr, "rtl", _RPR_ORDER)
    rtl.set(qn("w:val"), "1")
    lang = _insert_ordered(rPr, "lang", _RPR_ORDER)
    lang.set(qn("w:bidi"), "he-IL")


def setup_styles(doc: Document) -> None:
    _set_doc_defaults_rtl(doc)

    normal = doc.styles["Normal"]
    normal.font.name = HE_FONT
    normal.font.size = Pt(12)
    normal.font.color.rgb = DARK
    rpr = normal.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:cs"), HE_FONT)
    rfonts.set(qn("w:ascii"), HE_FONT)
    rfonts.set(qn("w:hAnsi"), HE_FONT)
    _style_set_rtl(normal, align_right=True)

    for heading_name, size, color in [
        ("Heading 1", 22, ACCENT_DARK),
        ("Heading 2", 17, ACCENT),
        ("Heading 3", 14, DARK),
    ]:
        s = doc.styles[heading_name]
        s.font.name = HE_FONT
        s.font.size = Pt(size)
        s.font.color.rgb = color
        s.font.bold = True
        _style_set_rtl(s, align_right=True)

    for list_style in ("List Bullet", "List Number"):
        try:
            s = doc.styles[list_style]
        except KeyError:
            continue
        s.font.name = HE_FONT
        _style_set_rtl(s, align_right=True)


def he_paragraph(doc_or_cell, text=None, *, style=None, bold=False, size=12,
                 italic=False, color=None, space_before=2, space_after=4,
                 align=WD_ALIGN_PARAGRAPH.RIGHT):
    if hasattr(doc_or_cell, "add_paragraph"):
        target = doc_or_cell
    else:
        target = doc_or_cell.paragraphs[0]

    if style:
        if hasattr(doc_or_cell, "add_paragraph"):
            p = doc_or_cell.add_paragraph(style=style)
        else:
            p = doc_or_cell.add_paragraph(style=style)
    else:
        if hasattr(doc_or_cell, "add_paragraph"):
            p = doc_or_cell.add_paragraph()
        else:
            p = doc_or_cell.add_paragraph()
    _set_bidi(p)
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    if text is None:
        return p
    run = p.add_run(text)
    _set_run_rtl(run)
    run.font.name = HE_FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color
    return p


def he_segments(doc, segments, *, style=None, size=12, bullet=False,
                bold=False, space_before=2, space_after=4,
                align=WD_ALIGN_PARAGRAPH.RIGHT):
    if bullet:
        p = doc.add_paragraph(style="List Bullet")
    elif style:
        p = doc.add_paragraph(style=style)
    else:
        p = doc.add_paragraph()
    _set_bidi(p)
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)

    for seg in segments:
        text, kind = seg if isinstance(seg, tuple) else (seg, "he")
        run = p.add_run(text)
        if kind == "en":
            run.font.name = CODE_FONT
            run.font.size = Pt(size - 1)
            run.font.color.rgb = RGBColor(0x03, 0x2F, 0x62)
            _set_run_shading(run, "F1F3F5")
        elif kind == "bold":
            run.font.name = HE_FONT
            run.font.size = Pt(size)
            run.font.bold = True
            _set_run_rtl(run)
        elif kind == "italic":
            run.font.name = HE_FONT
            run.font.size = Pt(size)
            run.font.italic = True
            _set_run_rtl(run)
        elif kind == "accent":
            run.font.name = HE_FONT
            run.font.size = Pt(size)
            run.font.bold = True
            run.font.color.rgb = ACCENT
            _set_run_rtl(run)
        elif kind == "muted":
            run.font.name = HE_FONT
            run.font.size = Pt(size)
            run.font.color.rgb = MUTED
            _set_run_rtl(run)
        else:
            run.font.name = HE_FONT
            run.font.size = Pt(size)
            run.font.bold = bold
            _set_run_rtl(run)
    return p


def _set_run_shading(run, hex_color: str) -> None:
    rPr = run._r.get_or_add_rPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    rPr.append(shd)


def heading(doc, text, level=1):
    p = doc.add_paragraph(style=f"Heading {level}")
    _set_bidi(p)
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_before = Pt(14 if level == 1 else 10)
    p.paragraph_format.space_after = Pt(8)
    run = p.add_run(text)
    _set_run_rtl(run)
    run.font.name = HE_FONT
    run.font.size = Pt(22 if level == 1 else 17 if level == 2 else 14)
    run.font.bold = True
    run.font.color.rgb = ACCENT_DARK if level == 1 else (ACCENT if level == 2 else DARK)
    return p


def page_break(doc):
    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)


def callout(doc, segments, *, bg=CALLOUT_BG, border_color=ACCENT, size=11):
    p = he_segments(doc, segments, size=size, space_before=4, space_after=4)
    _set_paragraph_shading(p, rgb_hex(bg))
    _add_paragraph_border(p, "right", rgb_hex(border_color), sz=18)
    _add_paragraph_border(p, "top", rgb_hex(border_color), sz=4)
    _add_paragraph_border(p, "bottom", rgb_hex(border_color), sz=4)
    _add_paragraph_border(p, "left", rgb_hex(border_color), sz=4)
    p.paragraph_format.left_indent = Cm(0.3)
    p.paragraph_format.right_indent = Cm(0.3)
    return p


def tip_box(doc, segments):
    return callout(doc, segments, bg=TIP_BG, border_color=TIP_BORDER, size=11)


def code_block(doc, lines: Sequence[str], *, size: int = 9):
    table = doc.add_table(rows=1, cols=1)
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.RIGHT
    cell = table.cell(0, 0)
    _set_cell_shading(cell, rgb_hex(CODE_BG))
    cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    cell.text = ""

    for i, line in enumerate(lines):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        _set_ltr(p)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        run = p.add_run(line if line else " ")
        run.font.name = CODE_FONT
        run.font.size = Pt(size)
        run.font.color.rgb = CODE_FG
    return table


def add_table(doc, headers: Sequence[str], rows: Sequence[Sequence[str]],
              col_widths: Sequence[float] | None = None) -> None:
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.RIGHT
    table.autofit = True

    bidi = OxmlElement("w:bidiVisual")
    table._tbl.tblPr.append(bidi)

    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        _set_cell_shading(cell, rgb_hex(TABLE_HEAD_BG))
        cell.text = ""
        p = cell.paragraphs[0]
        _set_bidi(p)
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(header)
        _set_run_rtl(run)
        run.font.name = HE_FONT
        run.font.size = Pt(11)
        run.font.bold = True

    for r_idx, row in enumerate(rows, start=1):
        for c_idx, value in enumerate(row):
            cell = table.rows[r_idx].cells[c_idx]
            cell.text = ""
            p = cell.paragraphs[0]
            _set_bidi(p)
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = p.add_run(str(value))
            _set_run_rtl(run)
            run.font.name = HE_FONT
            run.font.size = Pt(10)

    if col_widths:
        for col, w in zip(table.columns, col_widths):
            for cell in col.cells:
                cell.width = Cm(w)


def he_bullet(doc, text, *, size=12):
    return he_paragraph(doc, "●  " + text, size=size,
                        space_before=1, space_after=1)


def build_document() -> None:
    doc = Document()
    setup_styles(doc)

    section = doc.sections[0]
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)
    _setup_section_rtl(section)

    write_cover(doc)
    page_break(doc)

    write_toc(doc)
    page_break(doc)

    heading(doc, "1. מבוא — ייזום ואפיון", level=1)
    write_intro(doc)
    page_break(doc)

    heading(doc, "2. תיאור תחום הידע — ניתוח", level=1)
    write_knowledge_domain(doc)
    page_break(doc)

    heading(doc, "3. ארכיטקטורה (העיצוב)", level=1)
    write_architecture(doc)
    page_break(doc)

    heading(doc, "4. מימוש הפרויקט — הקוד", level=1)
    write_implementation(doc)
    page_break(doc)

    heading(doc, "5. מדריך למשתמש", level=1)
    write_user_guide(doc)
    page_break(doc)

    heading(doc, "6. סיכום אישי ורפלקציה", level=1)
    write_reflection(doc)
    page_break(doc)

    heading(doc, "7. ביבליוגרפיה — סקר ספרות", level=1)
    write_bibliography(doc)
    page_break(doc)

    heading(doc, "נספח א' — תדפיס הקוד המקור המלא", level=1)
    write_source_appendix(doc)
    page_break(doc)

    heading(doc, "נספח ב' — טבלת הודעות הפרוטוקול המלאה", level=1)
    write_protocol_appendix(doc)
    page_break(doc)

    heading(doc, "נספח ג' — שאלות תיאורטיות אפשריות לבחינה", level=1)
    write_qa_appendix(doc)

    doc.save(str(OUT_DOCX))
    _word_postprocess_rtl(OUT_DOCX)
    print(f"wrote portfolio docx ({OUT_DOCX.stat().st_size:,} bytes)")


def _word_postprocess_rtl(docx_path: Path) -> None:
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        return

    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        wdocx = word.Documents.Open(str(docx_path.resolve()), ReadOnly=False)
        try:
            for para in wdocx.Paragraphs:
                style_name = ""
                try:
                    style_name = para.Style.NameLocal or ""
                except Exception:
                    style_name = ""
                if "heading" in style_name.lower() or "כותרת" in style_name:
                    try:
                        para.Format.ReadingOrder = 1
                    except Exception:
                        pass
            wdocx.Save()
        finally:
            wdocx.Close(SaveChanges=False)
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def write_cover(doc):
    he_paragraph(doc, "", space_before=60)
    he_paragraph(doc, "תיק פרויקט", size=26, bold=True, color=ACCENT_DARK,
                 align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    he_paragraph(doc, "Secure Chess", size=36, bold=True, color=ACCENT,
                 align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    he_paragraph(doc, "שחמט מאובטח רב-משתמשים מבוסס שרת/לקוח עם הצפנת סיסמאות bcrypt",
                 size=14, color=MUTED, italic=True,
                 align=WD_ALIGN_PARAGRAPH.CENTER, space_after=30)

    he_paragraph(doc, 'חלופה: מערכות הגנת סייבר במקצוע "תכנון ותכנות מערכות"',
                 size=12, color=DARK, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    he_paragraph(doc, "5 יחידות לימוד — עבודת גמר", size=12, bold=True,
                 align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    he_paragraph(doc, 'תשפ"ג', size=12, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=40)

    fields = [
        ("שם בית הספר", "_____________________"),
        ("סמל מוסד", "_____________________"),
        ("שם התלמיד", "_____________________"),
        ('מספר ת"ז', "_____________________"),
        ("שם המורה המנחה", "_____________________"),
        ("שם נושא הפרויקט", "Secure Chess — שחמט מאובטח רב-משתמשים"),
        ("תאריך הגשה", "_____________________"),
    ]
    add_table(doc, ["שדה", "ערך"], fields, col_widths=[5.5, 11.0])


def write_toc(doc):
    heading(doc, "תוכן עניינים", level=1)
    he_paragraph(doc,
                 "להלן תוכן העניינים האוטומטי של Word. כדי לעדכן אותו לאחר שינויים: "
                 "לחץ קליק־ימני על הטבלה ובחר 'Update Field' (או F9).",
                 italic=True, color=MUTED)

    p = doc.add_paragraph()
    _set_bidi(p)
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    fldChar1 = OxmlElement("w:fldChar")
    fldChar1.set(qn("w:fldCharType"), "begin")
    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = 'TOC \\o "1-3" \\h \\z \\u'
    fldChar2 = OxmlElement("w:fldChar")
    fldChar2.set(qn("w:fldCharType"), "separate")
    fldChar3 = OxmlElement("w:fldChar")
    fldChar3.set(qn("w:fldCharType"), "end")
    r_element = p.add_run()._r
    r_element.append(fldChar1)
    r_element.append(instrText)
    r_element.append(fldChar2)
    r_element.append(fldChar3)

    he_paragraph(doc, "תוכן עניינים — לחץ F9 ב-Word כדי לעדכן", italic=True, color=MUTED)


def write_intro(doc):
    heading(doc, "1.1 תיאור הרעיון והמוטיבציה", level=2)
    he_paragraph(doc,
                 "הפרויקט מממש משחק שחמט רב-משתמשים, שבו שני שחקנים מתחברים בו-זמנית "
                 "לשרת מרכזי, מאמתים את זהותם באמצעות שם משתמש וסיסמה, ומשחקים בזמן "
                 "אמת זה נגד זה — או, לחילופין, נגד בינה מלאכותית — דרך תקשורת רשת "
                 "מבוססת TCP. הסיסמאות נשמרות במנוחה בצורה מוצפנת (hash) באמצעות "
                 "אלגוריתם bcrypt עם salt לכל סיסמה.")
    he_paragraph(doc,
                 "המוטיבציה: שחמט הוא דוגמה אלגנטית למשחק עם חוקים מורכבים אך מוגדרים "
                 "היטב, המאפשרת להציג בתוך פרויקט אחד את כל הסטאק של מערכת תקשורת "
                 "אמיתית — סוקטים, ריבוי תהליכונים, ניהול מצב מרובה לקוחות, פרוטוקול "
                 "אפליקטיבי משלי, אחסון מתמיד של חשבונות, גיבוב סיסמאות, וטיפול בפרצות "
                 "אבטחה רלוונטיות.")

    callout(doc, [
        ("הערה לגבי הצפנת תקשורת: ", "bold"),
        ("לפי הנחיית המורה הראשונית, נדרשת ", "he"),
        ("הצפנת סיסמאות בלבד", "accent"),
        (" (ללא הצפנת מידע רגיש העובר בתקשורת). ", "he"),
        ("הפרויקט מקיים את הדרישה הזו במלואה: סיסמאות נשמרות עם bcrypt על הדיסק, "
         "וערוץ ה-TCP עצמו מעביר JSON-Lines בטקסט-קליר — כפי שאישר המורה.", "he"),
    ])

    heading(doc, "1.2 ייזום: זיהוי הצורך", level=2)
    he_paragraph(doc,
                 "מערכות משחק רשת רבות מעבירות סיסמאות כתובות בטקסט קליר ושומרות "
                 "אותן ב-DB ללא hash. כל מי שיש לו גישה לשרת או לקובץ המשתמשים יכול "
                 "לראות את הסיסמאות במלואן. רציתי לבנות מערכת שלא תסבול מהבעיה הזאת, "
                 "אבל גם תהיה אפשרית להבנה והרצה בסביבת לימוד (Python נטו, ללא רכיבים "
                 "חיצוניים מורכבים).")

    heading(doc, "1.3 אפיון פונקציונלי — דרישות המשתמש", level=2)
    add_table(doc, ["#", "דרישה", "תיאור"], [
        ("FR-1", "הרשמה", "משתמש יכול להירשם עם שם משתמש וסיסמה ייחודיים (סיסמה ≥ 8 תווים, שם משתמש 3-32 תווים מ-[A-Za-z0-9_-])"),
        ("FR-2", "התחברות", "משתמש קיים יכול להתחבר; שגיאה במקרה של פרטים שגויים; אסור להיות מחובר פעמיים בו-זמנית"),
        ("FR-3", "Lobby", "שחקן מאומת יכול להמתין בתור עד שיצורף לשחקן אחר"),
        ("FR-4", "ביטול Lobby", "שחקן יכול לבטל את ההמתנה ולחזור למצב 'מאומת'"),
        ("FR-5", "משחק אדם-אדם", "שני שחקנים מאומתים משחקים זה נגד זה; הצבעים מוקצים אוטומטית (הראשון שהמתין משחק לבן)"),
        ("FR-6", "משחק נגד AI", "שחקן יכול לבחור לשחק נגד יריב מלאכותי בעומק חיפוש ניתן לבחירה (1-4)"),
        ("FR-7", "מהלכים תקפים בלבד", "השרת בודק כל מהלך לפי חוקי שחמט מלאים: castling, en passant, promotion, check"),
        ("FR-8", "סיום משחק", "מטה / פט / חזרה משולשת / חוק 50 המהלכים / התפטרות / ניתוק → המשחק נסגר ושולחת הודעת game_ended לשני הצדדים"),
        ("FR-9", "התפטרות", "שחקן יכול להתפטר באמצע משחק; היריב מקבל הודעת ניצחון"),
    ])

    heading(doc, "1.4 אפיון לא-פונקציונלי", level=2)
    add_table(doc, ["#", "תכונה", "יעד"], [
        ("NFR-1", "אבטחת סיסמאות", "סיסמאות נשמרות hashed בלבד (bcrypt עם salt-per-password). אין plaintext בקובץ ה-users.json"),
        ("NFR-2", "יציבות שרת", "נפילת לקוח אחד (כולל ניתוק חיבור באמצע משחק) אינה מפילה את השרת או משחקים אחרים"),
        ("NFR-3", "ריבוי משחקים מקבילים", "השרת תומך ב-≥ 10 משחקים פעילים במקביל בעזרת thread-per-connection"),
        ("NFR-4", "תגובתיות", "Latency של מהלך < 100ms בתנאי לוקאל-הוסט (פלפול בלוח + ולידציה + שידור FEN)"),
        ("NFR-5", "אמינות פרוטוקול", "מסגרת JSON-Lines (line-delimited) עם validation מלא בצד השרת — שגיאות מובאות כהודעת error עם code"),
        ("NFR-6", "אטומיות אחסון", "כתיבה ל-users.json עוברת דרך tempfile + os.replace + os.fsync כדי שלא ייווצר קובץ פגום בקריסה באמצע כתיבה"),
    ])

    heading(doc, "1.5 לוח זמנים — אבני דרך", level=2)
    add_table(doc, ["שלב", "תוצר", "מצב"], [
        ("US1 — אימות", "register / login עם bcrypt + UserStore", "✅ הושלם"),
        ("US2 — משחק אדם-אדם", "Board (חוקי שחמט מלאים) + Game + Session + Lobby + מהלך מלא", "✅ הושלם"),
        ("US3 — ריבוי משחקים", "Thread-per-connection + GameRegistry", "✅ הושלם (בונוס)"),
        ("US4 — AI", "AIPlayer (alpha-beta minimax) + handle_play_ai", "✅ הושלם (בונוס)"),
        ("US5 — GUI", "Tkinter 3-screen client (Login / Lobby / Game)", "✅ הושלם"),
        ("US6 — תיק פרויקט", "מסמך פורמלי + מחוון + תדפיסים", "✅ הושלם"),
    ])

    heading(doc, "1.6 ניהול סיכונים", level=2)
    add_table(doc, ["סיכון", "הסתברות", "חומרה", "מיטיגציה"], [
        ("Deadlock בריבוי תהליכונים",
         "בינונית", "גבוהה",
         "כל lock מוגן בסדר קבוע (Lobby → GameRegistry → active_logins); אין nested locks באותו thread"),
        ("נפילת שרת בעקבות הודעת לקוח מזיקה",
         "בינונית", "גבוהה",
         "Session.run עוטף כל dispatch ב-try/except SecureChessError + Exception כללי; כל שגיאה הופכת ל-error frame"),
        ("התקפת brute-force על סיסמאות",
         "גבוהה", "בינונית",
         "bcrypt with cost factor 12 (~0.3 שניה לכל ניחוש); סיסמה ≥ 8 תווים"),
        ("דליפת users.json",
         "נמוכה", "גבוהה",
         "Hashes בלבד; bcrypt salt-per-password מבטל rainbow tables"),
        ("Race condition ב-UserStore",
         "בינונית", "בינונית",
         "threading.Lock על כל פעולה ש-mutates את האוסף + כתיבה אטומית עם os.replace"),
        ("שני חיבורים עם אותו account",
         "בינונית", "נמוכה",
         "active_logins dict ו-active_logins_lock; ניסיון login שני זורק AlreadyLoggedInError"),
    ])


def write_knowledge_domain(doc):
    heading(doc, "2.1 חוקי שחמט וניהול משחק", level=2)
    he_paragraph(doc,
                 "המשחק מתנהל על לוח 8×8 (64 משבצות). כל שחקן מתחיל עם 16 כלים: "
                 "8 חיילים, 2 צריחים, 2 פרשים, 2 רצים, מלכה אחת, ומלך אחד. החוקים "
                 "העיקריים שמומשו בקובץ board.py:")
    for line in [
        "תנועה חוקית לפי סוג הכלי (חייל קדימה, צריח בקווים ישרים, רץ באלכסונים, פרש בצורת L, מלכה במצורף, מלך משבצת אחת)",
        "שחיתה — castling קצרה (O-O) וארוכה (O-O-O) עם בדיקת כל תנאי הזכאות: המלך לא זז, הצריח לא זז, אין כלים בדרך, ושלוש המשבצות לא מותקפות",
        "מהלך 'הצרפתי' — en passant: לכידת חייל יריב שעשה צעד כפול ראשון",
        "קידום חייל — promotion: חייל שמגיע לדרגה האחרונה הופך למלכה (ברירת מחדל) או כלי אחר",
        "שח (check) — חייבים להגן על המלך — הלוגיקה ב-is_in_check + סינון אחרי כל pseudo-legal move",
        "מט (checkmate) — סוף משחק; הצד הנעול מפסיד",
        "פט (stalemate) — סוף משחק; תיקו",
        "תיקו טכני — חוק 50 המהלכים (halfmove_clock) וחזרה משולשת (threefold repetition) בעזרת position_history",
    ]:
        he_bullet(doc, line)

    heading(doc, "2.2 מודל OSI ושכבת התעבורה", level=2)
    he_paragraph(doc,
                 "מודל OSI מחלק את תקשורת הרשת ל-7 שכבות: Physical (1), Data Link (2), "
                 "Network (3), Transport (4), Session (5), Presentation (6), Application (7). "
                 "הפרויקט עובד באופן הבא:")
    add_table(doc, ["שכבה", "פרוטוקול / טכנולוגיה", "תפקיד בפרויקט"], [
        ("7 — אפליקציה", "JSON-Lines (משלי)", "הודעות בין שרת ללקוח (register, login, move, …)"),
        ("6 — ייצוג", "JSON (UTF-8)", "סריאליזציה של ההודעות"),
        ("5 — סשן", "TCP connection per Session", "מצב לקוח (ANONYMOUS → AUTHENTICATED → IN_LOBBY → IN_GAME)"),
        ("4 — תעבורה", "TCP", "סוקטים מסוג SOCK_STREAM; אמין, מסודר, connection-oriented"),
        ("3 — רשת", "IPv4", "כתובות IP, ניתוב"),
        ("2 — Data Link", "Ethernet / Wi-Fi", "מטופל ע\"י מערכת ההפעלה"),
        ("1 — פיזית", "—", "כבל / גלי רדיו — לא רלוונטי לקוד"),
    ])
    he_paragraph(doc,
                 "למה TCP ולא UDP?\n"
                 "TCP מבטיח שכל הודעה תגיע בסדר ובלי אובדן, וזה קריטי במשחק שחמט: "
                 "אסור שמהלך 'יאבד' או יגיע אחרי מהלך מאוחר יותר. UDP היה מתאים יותר "
                 "למשחק real-time עם הרבה הודעות (FPS, VoIP) שבהן אובדן בודד אינו "
                 "קריטי. בנוסף, TCP מספק handshake תלת-שלבי שמבטיח שהקישור פתוח לפני "
                 "שמתחילים את התקשורת האפליקטיבית.")

    heading(doc, "2.3 קריפטוגרפיה — Hash ו-bcrypt", level=2)
    he_paragraph(doc,
                 "Hash היא פונקציה חד-כיוונית: קל לחשב hash(x) אבל קשה מאוד לחזור "
                 "מ-hash(x) ל-x. דוגמה: SHA-256 ממפה כל קלט לפלט באורך 256 ביט.")
    he_paragraph(doc, "אבל SHA-256 לבד אינה מתאימה לסיסמאות, כי:")
    he_bullet(doc, "מהירה מדי — תוקף עם GPU יכול לבדוק מיליארדי ניחושים בשנייה (rainbow tables)")
    he_bullet(doc, "סיסמאות זהות מקבלות hash זהה — מאפשר השוואה של hashes בין משתמשים")

    he_paragraph(doc,
                 "bcrypt פותר את שתי הבעיות: (1) הוא איטי בכוונה, ניתן לכיוון עם work "
                 "factor (ברירת מחדל 12 → כ-0.3 שניה לכל ניחוש); (2) מוסיף salt אקראי "
                 "לכל סיסמה, כך שאפילו אם שני משתמשים בחרו את אותה סיסמה, ה-hashes "
                 "שלהם יהיו שונים.")
    code_block(doc, [
        ">>> from secure_chess.common.crypto import hash_password",
        ">>> hash_password('hunter2!')",
        "'$2b$12$gAEohR4SPpvFmF.AYOuhe.q5w0vKj4iZJYxqMaaQOjsmh.kPdNLwa'",
        "",
        ">>> hash_password('hunter2!')   # אותה סיסמה — hash שונה!",
        "'$2b$12$wM/h1tlW1ZQbZsZAQv0kFOoLRYUKEazxLPP70S1iFNcWcK5N4FRyy'",
    ])
    callout(doc, [
        ("מבנה ה-hash: ", "bold"),
        ("$2b$ מציין את גרסת bcrypt; 12 הוא ה-work factor (cost); 22 התווים הבאים הם "
         "ה-salt בקידוד modified-base64; ושאר 31 התווים הם ה-hash עצמו. כל המידע נמצא "
         "במחרוזת אחת — לא צריך לשמור salt בנפרד.", "he"),
    ])

    heading(doc, "2.4 ריבוי תהליכונים ו-OS", level=2)
    he_paragraph(doc, "Process vs Thread:")
    he_bullet(doc, "Process: מופע נפרד של תוכנית עם זיכרון משלו, PID משלו, וטבלת קבצים פתוחים משלו.")
    he_bullet(doc, "Thread: יחידת ביצוע בתוך אותו process; משתפת זיכרון עם שאר ה-threads.")
    he_paragraph(doc,
                 "בפרויקט אנחנו משתמשים ב-threads (לא ב-processes) כי: (1) שיתוף הזיכרון "
                 "מאפשר ל-Lobby ול-GameRegistry להיות נגישים מכל session ללא IPC; "
                 "(2) ה-overhead של thread קטן בהרבה מ-process; (3) הקוד שלנו I/O-bound "
                 "(הרבה socket.recv), אז ה-GIL של Python לא מגביל אותנו.")
    he_paragraph(doc,
                 "Locks: כל מבנה נתונים משותף בין threads חייב להיות מוגן. בפרויקט "
                 "הסנכרון מבוסס על threading.Lock:")
    code_block(doc, [
        "# server/lobby.py",
        "class Lobby:",
        "    def __init__(self) -> None:",
        "        self._queue: List['Session'] = []",
        "        self._lock = threading.Lock()",
        "",
        "    def enqueue(self, session):",
        "        with self._lock:",
        "            self._queue.append(session)",
        "            if len(self._queue) >= 2:",
        "                white = self._queue.pop(0)",
        "                black = self._queue.pop(0)",
        "                return white, black",
        "        return None",
    ])

    heading(doc, "2.5 אחסון בקבצים — Atomic file writes", level=2)
    he_paragraph(doc,
                 "ה-UserStore שומר את כל החשבונות בקובץ JSON אחד (data/users.json). "
                 "הבעיה: אם השרת קורס באמצע כתיבה, הקובץ יישאר חצי-כתוב והקריאה הבאה "
                 "תיכשל. הפתרון — כתיבה אטומית:")
    he_bullet(doc, "כתוב קודם לקובץ זמני (tempfile.NamedTemporaryFile באותה תיקייה)")
    he_bullet(doc, "השתמש ב-os.fsync כדי לוודא שהבייטים יושבים על הדיסק")
    he_bullet(doc, "החלף אטומית את קובץ-היעד הישן בחדש (os.replace)")
    he_paragraph(doc,
                 "os.replace הוא קריאת מערכת אטומית ב-Windows וב-POSIX: או שהקובץ "
                 "הישן או שהחדש — לעולם לא שילוב של שניהם. כך מובטחת עקביות.")

    heading(doc, "2.6 חלוקת מטלות שרת/לקוח", level=2)
    add_table(doc, ["מטלה", "צד שרת", "צד לקוח"], [
        ("שמירת חשבונות (UserStore + bcrypt)", "✅ אחראי בלעדי", "—"),
        ("אימות סיסמה", "✅", "—"),
        ("מצב משחק (Board, Game)", "✅ source of truth", "מציג עותק מ-FEN"),
        ("ולידציה של מהלך", "✅ סמכותי", "אופציונלי (UX)"),
        ("Matchmaking (Lobby + GameRegistry)", "✅", "שולח play_human / cancel_lobby"),
        ("AI (alpha-beta minimax)", "✅", "—"),
        ("UI rendering (Tkinter / CLI)", "—", "✅"),
        ("קליטת קלט משתמש", "—", "✅"),
        ("Push notifications של מצב משחק", "✅ שולח", "✅ מציג"),
    ])


def write_architecture(doc):
    heading(doc, "3.1 ארכיטקטורה כללית של המערכת", level=2)
    he_paragraph(doc, "תרשים בלוקים של רכיבי המערכת:")
    code_block(doc, [
        "+------------------+         TCP         +-----------------------------+",
        "|   Client (GUI)   |<------------------->|         ChessServer         |",
        "|   client.gui     |   JSON-Lines plain  |   server.server.py          |",
        "+------------------+      (port 5050)    |                             |",
        "                                         |   +---------------------+   |",
        "+------------------+         TCP         |   |   _accept_loop()    |   |",
        "|   Client (CLI)   |<------------------->+-->|   threading.Thread  |   |",
        "|   client.cli     |                     |   +----------+----------+   |",
        "+------------------+                     |              |              |",
        "                                         |    one Session per conn     |",
        "                                         |              v              |",
        "                                         |   +---------------------+   |",
        "                                         |   |       Session       |   |",
        "                                         |   |  server.session.py  |   |",
        "                                         |   +----+----------+-----+   |",
        "                                         |        |          |         |",
        "                                         |        v          v         |",
        "                                         |   +--------+ +---------+    |",
        "                                         |   |  Lobby | |GameReg. |    |",
        "                                         |   +---+----+ +----+----+    |",
        "                                         |       |           |         |",
        "                                         |       v           v         |",
        "                                         |   +----------------+        |",
        "                                         |   |      Game      |        |",
        "                                         |   |  common.game   |        |",
        "                                         |   +-------+--------+        |",
        "                                         |           |                 |",
        "                                         |           v                 |",
        "                                         |   +----------------+        |",
        "                                         |   |      Board     |        |",
        "                                         |   | common.board   |        |",
        "                                         |   +----------------+        |",
        "                                         |                             |",
        "                                         |   +----------------+        |",
        "                                         |   |   UserStore    |--> data/users.json",
        "                                         |   | common.user_st |    (bcrypt hashes)",
        "                                         |   +----------------+        |",
        "                                         +-----------------------------+",
    ])

    heading(doc, "3.2 תרשימי זרימה — Sequence Diagrams", level=2)

    heading(doc, "3.2.1 הרשמה / כניסה", level=3)
    code_block(doc, [
        "Client                                              Server",
        "  |                                                   |",
        "  |---- TCP SYN ------------------------------------> |",
        "  |<--- TCP SYN+ACK --------------------------------- |",
        "  |---- TCP ACK ------------------------------------> |",
        "  |                                                   |",
        "  |-- {'type':'register',                             |",
        "  |    'username':'alice','password':'hunter2!'} ---> |",
        "  |                                                   |",
        "  |                          validate username/regex  |",
        "  |                          validate password length |",
        "  |                          hash_password(password)  |",
        "  |                          users.json atomic write  |",
        "  |                                                   |",
        "  |<-- {'type':'ok'} ------------------------------- |",
    ])

    heading(doc, "3.2.2 זרימת מהלך במשחק אדם-אדם", level=3)
    code_block(doc, [
        "ClientA (White)        Server (Session+Game)        ClientB (Black)",
        "   |                          |                            |",
        "   |-- move e2e4 ------------>|                            |",
        "   |                          |  Move.parse('e2e4')         |",
        "   |                          |  game.submit_move(self,m)  |",
        "   |                          |  Board.apply(move)         |",
        "   |                          |  check mate? stalemate?    |",
        "   |<------------- ok --------|                            |",
        "   |                          |                            |",
        "   |                          |---- game_state ----------->|",
        "   |<-- game_state -----------|                            |",
        "   |                          |                            |",
        "   |                          |                            |  (Black's turn)",
        "   |                          |<------ move e7e5 ----------|",
        "   |                          |                            |",
        "   |<-- game_state -----------|----- game_state ---------->|",
    ])

    heading(doc, "3.2.3 משחק נגד AI", level=3)
    code_block(doc, [
        "Client                       Server",
        "  |                            |",
        "  |-- play_ai (color=white) -> |",
        "  |                            |  AISession(depth=3)",
        "  |                            |  GameRegistry.create(human, ai)",
        "  |<-- ok -------------------- |",
        "  |<-- game_started --------- |",
        "  |                            |",
        "  |-- move e2e4 -------------> |",
        "  |                            |  Game.submit_move()",
        "  |<-- ok -------------------- |",
        "  |<-- game_state ----------- |",
        "  |                            |  registry.drive_ai(ai_session)",
        "  |                            |  AIPlayer.choose_move(board, depth=3)",
        "  |                            |  game.submit_move(ai, move)",
        "  |<-- game_state (AI move) - |",
    ])

    heading(doc, "3.3 חלוקה למודולים — שרת מול לקוח", level=2)
    add_table(doc, ["מודול", "צד", "תפקיד"], [
        ("common/pieces.py", "משותף", "Color, PieceType, Piece + pseudo-legal moves"),
        ("common/board.py", "משותף", "Board (8×8), validation, check/mate/draws, FEN"),
        ("common/move.py", "משותף", "Move dataclass + UCI parsing (e2e4, e7e8q)"),
        ("common/game.py", "משותף", "Game + Result; סיום משחק"),
        ("common/ai.py", "שרת בלבד", "AIPlayer — alpha-beta minimax"),
        ("common/crypto.py", "שרת בלבד", "bcrypt wrappers: hash_password / verify_password"),
        ("common/protocol.py", "משותף", "JSON-Lines framing + _LineReader"),
        ("common/user_store.py", "שרת בלבד", "Account + UserStore (JSON persistence)"),
        ("common/errors.py", "משותף", "SecureChessError hierarchy + wire-protocol codes"),
        ("common/log.py", "משותף", "stderr logger"),
        ("server/server.py", "שרת", "ChessServer — accept loop + ניהול sessions"),
        ("server/session.py", "שרת", "Session, SessionState, AISession + dispatcher"),
        ("server/lobby.py", "שרת", "Lobby + GameRegistry + drive_ai"),
        ("client/cli.py", "לקוח", "REPL בטרמינל"),
        ("client/gui.py", "לקוח", "Tkinter GUI 3-screens (Login → Lobby → Game)"),
        ("client/__main__.py", "לקוח", "Argparse entrypoint — GUI ברירת מחדל, --cli למצב טקסטואלי"),
    ])

    heading(doc, "3.4 ניתוח אלגוריתמים מרכזיים — שקילת חלופות", level=2)

    heading(doc, "3.4.1 AI — Alpha-Beta Minimax", level=3)
    he_paragraph(doc,
                 "המימוש מבוסס על negamax עם גזימת alpha-beta ופונקציית הערכה material-only:")
    he_bullet(doc, "חלופה: Minimax רגיל — נדחתה, כי בעומק 4 הוא בוחן ~30^4 = 810,000 פוזיציות. Alpha-beta חותך עד פי 100 בממוצע.")
    he_bullet(doc, "חלופה: MCTS (Monte Carlo Tree Search) — נדחתה, מורכבת מדי לפרויקט בית-ספרי. Alpha-beta מספיק להבסת שחקנים מתחילים.")
    he_bullet(doc, "חלופה: רשת נוירונים (AlphaZero-style) — נדחתה, דורש GPU וטריינינג ארוך, מעבר להיקף הפרויקט.")
    code_block(doc, [
        "# common/ai.py - הלולאה הראשית של negamax",
        "for move in legal:",
        "    snap = board._snapshot()",
        "    try:",
        "        board._apply_unchecked(move)",
        "        score = -self._negamax(board, depth-1, -beta, -alpha,",
        "                               perspective=side.opposite())",
        "    finally:",
        "        board._restore(snap)",
        "    if score > alpha: alpha = score",
        "    if alpha >= beta: break    # גזימת alpha-beta",
    ])

    heading(doc, "3.4.2 הצפנת סיסמה במנוחה", level=3)
    he_bullet(doc, "✅ bcrypt — מאומץ. work factor מובנה (cost=12), salt אוטומטי בכל קריאה, סטנדרט תעשייתי.")
    he_bullet(doc, "⏳ PBKDF2 — חלופה לגיטימית. נדחתה כי bcrypt עמיד יותר ל-GPU attacks (קוד שעשו אופטימיזציה אגרסיבית של בלוקי SHA אינו רץ מהר על bcrypt בגלל ה-Blowfish setup).")
    he_bullet(doc, "⏳ Argon2 — הזוכה של Password Hashing Competition 2015, מודרני יותר. נדחה כי bcrypt עומד בדרישות, זמין ב-pip native ל-Windows, ולא צריך לקמפל C.")

    heading(doc, "3.4.3 פרוטוקול ההודעות", level=3)
    he_bullet(doc, "✅ JSON-Lines — מאומץ. קריא לאדם, debugging קל ב-Wireshark, פרסור מובנה ב-Python.")
    he_bullet(doc, "⏳ Protocol Buffers / MessagePack — בינארי, פי 2-3 קומפקטי יותר על הרשת. נדחה כי בפרויקט בית-ספר היכולת לפתוח netcat ולראות הודעות בטקסט מנצחת את הקומפקטיות.")
    he_bullet(doc, "⏳ Apache Thrift / gRPC — overkill למשחק שחמט עם 13 סוגי הודעות בלבד.")

    heading(doc, "3.4.4 בחירת מודל קונקרנציה", level=3)
    he_bullet(doc, "✅ Thread-per-connection — מאומץ. פשוט להבנה, מבודד היטב, מנצל GIL נכון על I/O-bound.")
    he_bullet(doc, "⏳ asyncio (event loop יחיד) — חוסך thread overhead. נדחה כי Tkinter לא משתלב טוב עם asyncio, ו-threading פשוט יותר ל-debug.")
    he_bullet(doc, "⏳ Process-per-connection — דורש IPC ל-Lobby/GameRegistry. overkill.")

    heading(doc, "3.5 פרוטוקול התקשורת — מצבים והודעות", level=2)
    he_paragraph(doc,
                 "Session State Machine: ANONYMOUS → AUTHENTICATED → IN_LOBBY → "
                 "IN_GAME → (חוזר ל-AUTHENTICATED אחרי game_ended). מעבר בכל מצב מוגן "
                 "ע\"י בדיקת state ב-handler המתאים (BadStateError אם לא חוקי).")
    code_block(doc, [
        "                +-------------+",
        "                | ANONYMOUS   |  <-- חיבור TCP חדש",
        "                +------+------+",
        "                       | register / login (ok)",
        "                       v",
        "                +-------------+",
        "                |AUTHENTICATED|  <-- אחרי כניסה מוצלחת",
        "                +--+----+-----+",
        "         play_human |    | play_ai",
        "                    v    v",
        "          +----------+   +----------+",
        "          | IN_LOBBY |   | IN_GAME  |",
        "          +-----+----+   +----+-----+",
        "                |   ^         |   ^",
        "         game_  |   | cancel_ |   | game_",
        "         started|   | lobby   |   | ended",
        "                v   |         v   |",
        "             [pair-matched]  [back to AUTHENTICATED]",
    ])

    heading(doc, "3.6 ניתוח חולשות ופתרונות", level=2)
    callout(doc, [
        ("הערה חשובה: ", "bold"),
        ("לפי הנחיית המורה הראשונית, ", "he"),
        ("הצפנת תקשורת אינה נדרשת בפרויקט הזה", "accent"),
        (" — נדרשת רק הצפנת סיסמאות במנוחה. הטבלה להלן מציגה רק חולשות שרלוונטיות "
         "למודל האיומים המאושר.", "he"),
    ])
    add_table(doc, ["איום", "רובד", "סיכון", "מיטיגציה בפרויקט"], [
        ("דליפת users.json", "אחסון",
         "מאזין רואה סיסמאות",
         "bcrypt + salt; לעולם אין plaintext על הדיסק."),
        ("Brute-force online", "אפליקציה",
         "ניחוש סיסמאות דרך login",
         "bcrypt עם cost=12 מאט כל ניחוש ל-~0.3s. סיסמה ≥ 8 תווים. עתידית: rate-limiter פר-IP."),
        ("Crash של לקוח", "אפליקציה",
         "השרת קורס / משחק אחר נופל",
         "try/except SecureChessError + Exception ב-Session.run; _on_disconnect מנקה משאבים אטומית"),
        ("Race condition ב-Lobby", "ריבוי תהליכונים",
         "שני אנשים מקבלים את אותו צד",
         "threading.Lock על enqueue/remove; pop של שני ה-sessions בתוך אותו lock"),
        ("Race condition ב-UserStore", "ריבוי תהליכונים",
         "שני אנשים נרשמים עם אותו שם",
         "threading.Lock סביב בדיקת ייחודיות + insert + persist (atomic write)"),
        ("Double-login של אותו account", "אפליקציה",
         "session ישן ימשיך לשלוט במשחק חדש",
         "active_logins dict + lock; AlreadyLoggedInError במידה ויש כבר session פעיל"),
        ("SQL Injection", "מאגר",
         "—", "לא רלוונטי — אין SQL בפרויקט. אחסון = JSON עם json.dumps המקודד תווים מיוחדים"),
        ("Buffer Overflow", "שפה",
         "—", "לא רלוונטי — Python ללא pointer arithmetic"),
        ("MITM / Eavesdropping", "תעבורה",
         "האזנה לסיסמאות",
         "מחוץ ל-scope לפי הנחיית המורה — נדרשת רק הצפנת סיסמאות במנוחה; ההגנה ע\"י הרצה ב-localhost / VPN"),
    ])


def write_implementation(doc):
    heading(doc, "4.1 תכנות מונחה עצמים — 21 מחלקות אפליקטיביות", level=2)
    he_paragraph(doc, "טבלת המחלקות שהתלמיד יצר, מקובצות לפי תפקיד:")
    add_table(doc, ["מחלקה", "מודול", "תפקיד עיקרי"], [
        ("Color, PieceType (Enums)", "common/pieces.py", "Enums של צבע וסוגי כלים"),
        ("Piece", "common/pieces.py", "כלי בודד + יצירת מהלכים pseudo-legal לפי סוג"),
        ("Move", "common/move.py", "Dataclass של מהלך + UCI parser (e2e4, e7e8q)"),
        ("Board", "common/board.py", "מצב הלוח + ולידציה + check/mate/draws + FEN parsing"),
        ("Game", "common/game.py", "מצב משחק בודד; submit_move; resign; handle_disconnect"),
        ("Result", "common/game.py", "Dataclass של תוצאה: white_wins / black_wins / draw + reason"),
        ("AIPlayer", "common/ai.py", "Alpha-beta minimax עם material-balance evaluation"),
        ("Account", "common/user_store.py", "Dataclass: username + password_hash + created_at"),
        ("UserStore", "common/user_store.py", "אחסון JSON אטומי + register/authenticate + lock"),
        ("_LineReader", "common/protocol.py", "Buffered TCP reader — JSON line per read_line()"),
        ("SessionState (Enum)", "server/session.py", "Enum: ANONYMOUS / AUTHENTICATED / IN_LOBBY / IN_GAME"),
        ("Session", "server/session.py", "Per-client session + dispatcher (handle_register, …)"),
        ("AISession", "server/session.py", "In-memory session impersonating an AI opponent"),
        ("_AIAccount", "server/session.py", "Tiny stand-in for Account so Game can read username"),
        ("Lobby", "server/lobby.py", "Matchmaking queue + thread-safe enqueue/remove"),
        ("GameRegistry", "server/lobby.py", "Tracks active games + broadcast_state + drive_ai"),
        ("ChessServer", "server/server.py", "Accept loop + shared user_store/lobby/registry"),
        ("ClientConnection", "client/cli.py", "TCP wrapper for the CLI reader thread"),
        ("ServerConnection", "client/gui.py", "TCP wrapper + queue.Queue for the GUI poll loop"),
        ("ChessGui", "client/gui.py", "Top-level Tk application + screen routing"),
        ("LoginFrame, LobbyFrame, GameFrame", "client/gui.py", "3 מסכי ה-GUI"),
    ])
    he_paragraph(doc,
                 "המחלקות מקיימות יחסים שונים: הכלה (Board מכיל מטריצת Piece), הרכבה "
                 "(Game מכיל Board + שני Sessions), הפשטה (AISession ו-Session חולקים "
                 "אותו ממשק חלקי ש-Game/GameRegistry צורכים), ופולימורפיזם פנימי "
                 "(Session._dispatch קורא ל-handle_<type> דינמית).")
    he_paragraph(doc,
                 "בנוסף ל-21 המחלקות האפליקטיביות, יש 13 מחלקות שגיאה (SecureChessError "
                 "+ 12 תת-מחלקות) המהוות היררכיית exceptions מסודרת.")

    heading(doc, "4.2 חלוקה לקבצים — מבנה התיקיות", level=2)
    code_block(doc, [
        "secure-chess/",
        "├── pyproject.toml          # תלויות + מטא-דאטה",
        "├── data/                   # users.json + server.log",
        "└── src/secure_chess/",
        "    ├── __init__.py",
        "    ├── common/             # לוגיקה טהורה ללא I/O",
        "    │   ├── pieces.py       # Color, PieceType, Piece + pseudo-legal moves",
        "    │   ├── board.py        # Board, legality, check/mate/draws, FEN",
        "    │   ├── move.py         # Move dataclass + UCI parsing",
        "    │   ├── game.py         # Game + Result",
        "    │   ├── ai.py           # alpha-beta minimax (BONUS)",
        "    │   ├── crypto.py       # bcrypt wrappers",
        "    │   ├── user_store.py   # Account + JSON-backed UserStore",
        "    │   ├── protocol.py     # JSON-Lines framing",
        "    │   ├── errors.py       # exception hierarchy",
        "    │   └── log.py          # stderr logger",
        "    ├── server/             # TCP socket + threading",
        "    │   ├── __main__.py     # python -m secure_chess.server",
        "    │   ├── server.py       # ChessServer + accept loop",
        "    │   ├── session.py      # Session, SessionState, AISession",
        "    │   └── lobby.py        # Lobby + GameRegistry",
        "    └── client/             # שתי חזיתות GUI / CLI",
        "        ├── __main__.py     # python -m secure_chess.client (GUI default; --cli)",
        "        ├── gui.py          # Tkinter desktop GUI",
        "        └── cli.py          # text-mode REPL",
    ])

    heading(doc, "4.3 קוד בטוח — try/except ו-robustness", level=2)
    he_paragraph(doc,
                 "השרת חייב לשרוד נפילת לקוח, הודעה לא תקפה, מצב לא חוקי, וכל תקלה "
                 "אחרת. הדפוס בפרויקט — שלוש שכבות defense-in-depth:")
    code_block(doc, [
        "# server/session.py - Session.run()",
        "def run(self) -> None:",
        "    log.info('[%s] connected from %s:%d', self.id[:8], *self.peer)",
        "    try:",
        "        # שכבה 1 — receive loop",
        "        while not self._closed:",
        "            try:",
        "                msg = protocol.recv_message(self._reader)",
        "            except ConnectionClosed:",
        "                log.info('[%s] disconnected (%s)', self.id[:8], self.username)",
        "                break",
        "            except ProtocolError as exc:",
        "                self.send(protocol.error(exc.code, str(exc)))",
        "                continue",
        "",
        "            # שכבה 2 — dispatch loop",
        "            try:",
        "                self._dispatch(msg)",
        "            except SecureChessError as exc:",
        "                self.send(protocol.error(exc.code, str(exc)))",
        "            except Exception as exc:    # defensive",
        "                log.exception('[%s] unexpected: %s', self.id[:8], exc)",
        "                self.send(protocol.error('internal_error',",
        "                                        'internal server error'))",
        "    finally:",
        "        # שכבה 3 — cleanup, always runs",
        "        self._on_disconnect()      # מסיר מה-lobby/game",
        "        self.close()               # סוגר socket",
    ])
    he_paragraph(doc, "השכבות:")
    he_bullet(doc, "שכבה 1 (receive) — מתאוששת מהודעה פגומה (JSON שבור, שדה חסר) ע\"י החזרת error frame והמשך הלולאה.")
    he_bullet(doc, "שכבה 2 (dispatch) — לוכדת כל SecureChessError (illegal_move, not_your_turn, weak_password, …) וכל Exception כללי. השרת לעולם לא קורס בגלל הודעת לקוח.")
    he_bullet(doc, "שכבה 3 (cleanup) — try/finally מבטיח שהסוקט ייסגר, השחקן יוסר מה-lobby, וה-active_logins ינוקה, גם אם זרקנו פנימית.")


def write_user_guide(doc):
    heading(doc, "5.1 התקנה והרצה", level=2)
    he_paragraph(doc, "דרישות סף: Windows 10/11, Python 3.11+, PowerShell.")
    code_block(doc, [
        "# שיבוט הפרויקט",
        "cd secure-chess",
        "",
        "# יצירת סביבה וירטואלית",
        "python -m venv .venv",
        ".\\.venv\\Scripts\\Activate.ps1",
        "",
        "# התקנת חבילות",
        "pip install -e \".[dev]\"",
    ])

    heading(doc, "5.1.1 הרצת השרת", level=3)
    code_block(doc, [
        "python -m secure_chess.server --host 127.0.0.1 --port 5050 --data-dir ./data",
    ])
    he_paragraph(doc,
                 "השרת ידפיס: 'listening on 127.0.0.1:5050' ולאחר מכן 'credential store: "
                 "data/users.json (N accounts)'.")

    heading(doc, "5.1.2 הרצת לקוח GUI", level=3)
    code_block(doc, [
        "python -m secure_chess.client --host 127.0.0.1 --port 5050",
    ])

    heading(doc, "5.1.3 הרצת לקוח CLI (ללא תצוגה גרפית)", level=3)
    code_block(doc, [
        "python -m secure_chess.client --host 127.0.0.1 --port 5050 --cli",
    ])

    heading(doc, "5.2 צילומי מסך של הזרימה", level=2)
    callout(doc, [
        ("פעולה נדרשת מהתלמיד: ", "bold"),
        ("יש להריץ את ה-GUI ולצרף 5 צילומי מסך לפי המקומות המסומנים מטה.", "he"),
    ])
    he_bullet(doc, "[Screenshot 5.2.1] מסך Login — לפני הכנסת פרטים")
    he_bullet(doc, "[Screenshot 5.2.2] מסך Lobby — אחרי התחברות (Welcome)")
    he_bullet(doc, "[Screenshot 5.2.3] מסך Lobby — Waiting for opponent…")
    he_bullet(doc, "[Screenshot 5.2.4] מסך Game — תחילת משחק עם 32 הכלים")
    he_bullet(doc, "[Screenshot 5.2.5] מסך Game — אחרי כמה מהלכים והדגשת last_move צהובה")

    heading(doc, "5.3 שלושה סוגי משתמשים", level=2)
    heading(doc, "5.3.1 שחקן רגיל (אדם נגד אדם)", level=3)
    he_paragraph(doc,
                 "1) הרץ את ה-GUI. 2) במסך Login הזן שם משתמש (3-32 תווים, "
                 "אותיות/ספרות/_/-) וסיסמה (≥ 8 תווים). 3) לחץ Register (פעם ראשונה) "
                 "או Login. 4) במסך Lobby לחץ 'Join Lobby'. 5) המתן לשחקן נוסף — "
                 "תוצג הודעת 'Waiting…'. 6) ברגע שמתחיל משחק, מסך ה-Game ייפתח עם "
                 "הלוח. 7) לחץ על הכלי שלך ואז על משבצת היעד. הצבע הצהוב מסמן את "
                 "המהלך האחרון.")

    heading(doc, "5.3.2 שחקן מול AI (Bonus)", level=3)
    he_paragraph(doc,
                 "במסך Lobby בחר צבע ('White' או 'Black') ועומק חיפוש (1 קל, 4 קשה), "
                 "ולחץ 'Start AI Game'. ה-AI עונה בתוך כמה שניות. עומק 4 יכול לקחת "
                 "עד 10 שניות למהלך — אבל הוא חזק יחסית.")

    heading(doc, "5.3.3 מנהל / מפעיל השרת", level=3)
    he_paragraph(doc,
                 "המפעיל אחראי על: (1) בחירת host/port (ברירת מחדל 127.0.0.1:5050). "
                 "(2) בחירת data-dir שבו ייווצרו users.json ו-server.log. "
                 "(3) הגדרת ai_depth (ברירת מחדל 3). (4) הגדרת max_clients "
                 "(ברירת מחדל 32).")
    code_block(doc, [
        "python -m secure_chess.server --help",
        "",
        "# פרודקשן לדוגמה:",
        "python -m secure_chess.server `",
        "  --host 0.0.0.0 `",
        "  --port 5050 `",
        "  --data-dir C:\\ProgramData\\SecureChess",
    ])
    he_paragraph(doc,
                 "ניטור: התבונן ב-stderr להודעות 'connected from', 'authenticated as', "
                 "'game started', 'game ended'. נפילת לקוח אחד אינה משפיעה על השאר.")

    heading(doc, "5.4 הוכחת הצפנת הסיסמאות במנוחה", level=2)
    he_paragraph(doc,
                 "אחרי הרצת השרת והרשמת מספר משתמשים, אפשר לבדוק שהסיסמאות לא נשמרות "
                 "בטקסט קליר:")
    code_block(doc, [
        "Get-Content .\\data\\users.json",
        "# התוצאה: רק bcrypt hashes, אין plaintext passwords",
        "",
        "Select-String -Path .\\data\\users.json -Pattern 'hunter2!|s3cretpw'",
        "# צפוי: 0 matches",
    ])


def write_reflection(doc):
    heading(doc, "6.1 מה למדתי על עצמי", level=2)
    callout(doc, [
        ("פעולה נדרשת מהתלמיד: ", "bold"),
        ("סעיף זה חייב להיות אישי. להלן תבנית להמשך כתיבה — יש להחליף בתוכן "
         "אותנטי שלך. ", "he"),
        ('אסור להסתפק ב"תודה ונהניתי"!', "italic"),
    ])
    he_paragraph(doc,
                 "[תבנית להחלפה] בפרויקט הזה גיליתי שאני מתחבר/ת מאוד ל-___, "
                 "וכאשר נתקלתי ב-___ הבנתי שאני מסוגל/ת ל-___ למרות שחששתי בהתחלה. "
                 "הניהול העצמי שלי השתפר במיוחד ב-___, וזיהיתי שאני זקוק/ה לעבוד "
                 "עוד על ___. הכרתי שיטות עבודה חדשות כמו ___ ושאני מעדיף/ה לעבוד "
                 "___ (לבד / בצוות / בקטעים קצרים…).",
                 italic=True, color=MUTED)

    heading(doc, "6.2 מה למדתי מקצועית", level=2)
    he_bullet(doc, "ההבדל בין hashing (חד-כיווני, לסיסמאות במנוחה) לבין encryption (דו-כיווני, לתקשורת). bcrypt = hashing; הוא לא 'מצפין'.")
    he_bullet(doc, "Salt חייב להיות per-password, לא גלובלי. salt משותף אינו מגן מ-rainbow tables אם משתמשים שונים בחרו אותה סיסמה.")
    he_bullet(doc, "Work factor של bcrypt (cost): ככל שיותר גבוה, יותר איטי לתוקף — אבל גם לשרת. cost=12 הוא איזון סביר ב-2024.")
    he_bullet(doc, "Thread safety: כל מבנה נתונים משותף חייב lock או להיות immutable. שכחתי lock פעם אחת ב-Lobby וקיבלתי race condition.")
    he_bullet(doc, "ההבדל בין Process ל-Thread והשפעת ה-GIL של Python על threading. GIL פגיע ב-CPU-bound, לא ב-I/O-bound.")
    he_bullet(doc, "Pseudo-legal vs legal moves בשחמט: המפריד הוא מי בודק את ה-check filter. המידול הזה חסך לי הרבה duplication.")
    he_bullet(doc, "Alpha-beta pruning והקסם של negamax (משלב min ו-max בפונקציה אחת ע\"י נסיעה על נקודת המבט).")
    he_bullet(doc, "Atomic file writes — tempfile + os.fsync + os.replace הם השילוש הקדוש לקובץ קונפיג שלא יכול להיות 'חצי כתוב'.")
    he_bullet(doc, "PEP-8, type hints, ו-module docstrings — תרומה גדולה לקריאות הקוד גם 3 חודשים אחרי שכתבתי.")

    heading(doc, "6.3 מה הייתי משנה אם הייתי מתחיל מחדש", level=2)
    he_bullet(doc, "להוסיף TLS על התקשורת — כיום הסיסמה עוברת בקליר (אבל זה היה מחוץ ל-scope לפי הנחיית המורה).")
    he_bullet(doc, "להוסיף rate-limiter פר-IP לכניסות login כדי להאט brute-force מקוון.")
    he_bullet(doc, "להחליף JSON ב-MessagePack — מהיר ויותר קומפקטי על הרשת (לא נדרש אבל מעניין).")
    he_bullet(doc, "להוסיף תמיכה ב-promotion לכלי אחר חוץ מ-Queen ב-GUI (כיום זה Queen בלבד).")
    he_bullet(doc, "להוסיף base-rating ELO ומערכת ranking ל-lobby — כדי שלא יזווגו שחקן ברמה 100 לשחקן ברמה 2500.")
    he_bullet(doc, "להחליף threading ב-asyncio (מתאים יותר ל-I/O-bound, אם כי Tkinter מקשה על האינטגרציה).")


def write_bibliography(doc):
    he_paragraph(doc,
                 "להלן רשימת המקורות הראשיים שבהם נעזרתי. לכל מקור — תרומה ספציפית לפרויקט.")

    sources = [
        ("[1] Python threading documentation",
         "https://docs.python.org/3/library/threading.html",
         "התיעוד הרשמי שעזר לי להבין Lock, Event, Thread, daemon-threads. "
         "במיוחד הקטע על ה-GIL (Global Interpreter Lock) — שני threads ב-Python "
         "לא רצים באמת במקביל ב-CPU, אבל ל-I/O-bound (כמו socket.recv) זה לא חשוב."),

        ("[2] Python socket documentation",
         "https://docs.python.org/3/library/socket.html",
         "התיעוד שעזר לי להבין AF_INET, SOCK_STREAM, את ההבדל בין connect/bind/listen/accept, "
         "ואת השימוש ב-SO_REUSEADDR כדי שאפשר יהיה להפעיל מחדש את השרת ללא TIME_WAIT."),

        ("[3] bcrypt: A Future-Adaptable Password Scheme — Provos & Mazières (1999)",
         "https://www.usenix.org/legacy/event/usenix99/provos/provos.pdf",
         "המאמר המקורי שהציג את bcrypt. הבנתי ממנו את חשיבות ה-work factor "
         "(adaptive cost) ולמה salt חייב להיות per-password ולא גלובלי. גם הסביר "
         "למה Blowfish nicely resistant ל-GPU optimizations."),

        ("[4] OWASP Authentication Cheat Sheet",
         "https://cheatsheetseries.owasp.org/cheatsheets/Authentication_Cheat_Sheet.html",
         "המלצות מעשיות לאימות משתמשים — מינימום אורך סיסמה, rate limiting, "
         "constant-time comparison. אימצתי את ההמלצות הרלוונטיות שהיו אפשריות "
         "בתוך תקציב הפרויקט."),

        ("[5] Chess Programming Wiki — Alpha-Beta",
         "https://www.chessprogramming.org/Alpha-Beta",
         "אתר קהילתי מצוין על תכנות מנועי שחמט. השתמשתי בו ללמוד את negamax, "
         "alpha-beta pruning, ו-move ordering. ההצגה של negamax בתור 'min-max "
         "אבל תמיד maximizing מנקודת מבט הצד המהלך' פשטה לי מאוד את הקוד."),

        ("[6] Python tkinter documentation",
         "https://docs.python.org/3/library/tkinter.html",
         "התיעוד של ספריית ה-GUI הסטנדרטית של Python. למדתי ממנו על Canvas, "
         "Frame, after() polling, ועל שילוב thread עם Tk main loop באמצעות queue.Queue."),

        ("[7] PEP 8 — Style Guide for Python Code",
         "https://peps.python.org/pep-0008/",
         "מסמך הסטנדרטים לכתיבת קוד Python. אימצתי את הכללים על snake_case, "
         "אורך שורה, רווחים אופרטוריים, ו-import ordering."),

        ("[8] Forsyth-Edwards Notation (FEN)",
         "https://en.wikipedia.org/wiki/Forsyth%E2%80%93Edwards_Notation",
         "פורמט סטנדרטי לתיאור פוזיציית שחמט. השתמשתי בקיצור (placement + side + "
         "castling + ep) בפרוטוקול שלי כדי לשלוח state ללקוח בייצוג קומפקטי "
         "וקריא לאדם."),
    ]
    for title, url, desc in sources:
        heading(doc, title, level=2)
        he_paragraph(doc, url, color=MUTED, italic=True)
        he_paragraph(doc, desc)


SOURCE_FILES = [
    "src/secure_chess/__init__.py",
    "src/secure_chess/common/__init__.py",
    "src/secure_chess/common/errors.py",
    "src/secure_chess/common/log.py",
    "src/secure_chess/common/crypto.py",
    "src/secure_chess/common/protocol.py",
    "src/secure_chess/common/user_store.py",
    "src/secure_chess/common/pieces.py",
    "src/secure_chess/common/move.py",
    "src/secure_chess/common/board.py",
    "src/secure_chess/common/game.py",
    "src/secure_chess/common/ai.py",
    "src/secure_chess/server/__init__.py",
    "src/secure_chess/server/__main__.py",
    "src/secure_chess/server/lobby.py",
    "src/secure_chess/server/session.py",
    "src/secure_chess/server/server.py",
    "src/secure_chess/client/__init__.py",
    "src/secure_chess/client/__main__.py",
    "src/secure_chess/client/cli.py",
    "src/secure_chess/client/gui.py",
]


def write_source_appendix(doc):
    he_paragraph(doc,
                 "תדפיס מסודר של כל מודולי הפרויקט (Python 3.11+). הקוד מתועד עם "
                 "module docstrings ו-type hints מלאים.")
    for relative in SOURCE_FILES:
        path = REPO_ROOT / relative
        if not path.exists():
            continue
        heading(doc, relative, level=2)
        lines = path.read_text(encoding="utf-8").splitlines()
        if not lines:
            lines = ["# (קובץ ריק)"]
        code_block(doc, lines)


def write_protocol_appendix(doc):
    heading(doc, "ג.1 הודעות לקוח → שרת", level=2)
    add_table(doc, ["סוג הודעה", "תיאור", "שדות חובה"], [
        ("register", "הרשמת משתמש חדש", "username, password"),
        ("login", "התחברות לחשבון קיים", "username, password"),
        ("play_human", "הצטרפות ל-lobby להמתנה לשחקן", "(אין)"),
        ("play_ai", "התחלת משחק נגד AI", "(אופציונלי) color, depth"),
        ("cancel_lobby", "ביטול ההמתנה ב-lobby", "(אין)"),
        ("move", "ביצוע מהלך במשחק פעיל", "uci (e.g. 'e2e4', 'e7e8q')"),
        ("resign", "התפטרות ממשחק פעיל", "(אין)"),
        ("quit", "סגירת חיבור מסודרת", "(אין)"),
    ])

    heading(doc, "ג.2 הודעות שרת → לקוח", level=2)
    add_table(doc, ["סוג הודעה", "תיאור", "שדות"], [
        ("ok", "אישור הצלחה לפעולה האחרונה", "(אופציונלי תוכן נוסף)"),
        ("error", "שגיאה במהלך עיבוד הפעולה", "code, message"),
        ("game_started", "המשחק התחיל", "game_id, you_are, opponent, board (FEN), to_move"),
        ("game_state", "עדכון מצב המשחק (אחרי מהלך)", "game_id, last_move, board, to_move, in_check"),
        ("game_ended", "סיום המשחק", "game_id, result, reason"),
    ])

    heading(doc, "ג.3 דוגמת זרימה מלאה", level=2)
    code_block(doc, [
        "C -> S  {\"type\":\"register\",\"username\":\"alice\",\"password\":\"hunter2!\"}",
        "S -> C  {\"type\":\"ok\"}",
        "",
        "C -> S  {\"type\":\"play_human\"}",
        "S -> C  {\"type\":\"ok\"}",
        "",
        "# אחרי ש-bob גם הצטרף ל-lobby:",
        "S -> C  {\"type\":\"game_started\",\"game_id\":\"abc123…\",\"you_are\":\"white\",",
        "        \"opponent\":\"bob\",\"board\":\"rnbqkbnr/pppppppp/8/8/8/8/PPPPPPPP/RNBQKBNR\",",
        "        \"to_move\":\"white\"}",
        "",
        "C -> S  {\"type\":\"move\",\"uci\":\"e2e4\"}",
        "S -> C  {\"type\":\"ok\"}",
        "S -> C  {\"type\":\"game_state\",\"last_move\":\"e2e4\",",
        "        \"board\":\"rnbqkbnr/pppppppp/8/8/4P3/8/PPPP1PPP/RNBQKBNR\",",
        "        \"to_move\":\"black\",\"in_check\":false}",
        "",
        "# … מהלכים נוספים … סופית:",
        "S -> C  {\"type\":\"game_ended\",\"result\":\"white_wins\",\"reason\":\"checkmate\"}",
    ])


def write_qa_appendix(doc):
    he_paragraph(doc,
                 "להלן רשימה של שאלות תיאורטיות אפשריות מהבוחן עם תשובות מוכנות, "
                 "מסווגות לפי קטגוריה.")

    heading(doc, "ד.1 תקשורת", level=2)
    qa_pairs = [
        ("איזו שכבות קיימות במודל התקשורת?",
         "מודל OSI: 7 — Application (HTTP, JSON-Lines שלי), 6 — Presentation (UTF-8 JSON), "
         "5 — Session (TCP), 4 — Transport (TCP), 3 — Network (IP), 2 — Data Link (Ethernet), "
         "1 — Physical."),
        ("באיזה פרוטוקול תקשורת אתה עובד ואיך קבעת זאת?",
         "TCP, נקבע ביצירת הסוקט עם socket.socket(socket.AF_INET, socket.SOCK_STREAM). "
         "TCP בגלל שאני צריך אמינות וסדר במהלכי שחמט."),
        ("מה ההבדל בין TCP ל-UDP?",
         "TCP: אמין, מסודר, connection-oriented, אטי יחסית (handshake), reliable retransmission. "
         "UDP: לא אמין, לא מסודר, ללא חיבור, מהיר, מתאים ל-streaming."),
        ("מהי לחיצת יד משולשת?",
         "Three-way handshake של TCP: SYN → SYN+ACK → ACK. שלוש הודעות שמסכימים על "
         "סדרות התחלה לפני העברת נתונים."),
        ("מהי כתובת IP ולמה משמש PORT?",
         "כתובת IP מזהה מחשב ברשת (לדוגמה 127.0.0.1). PORT מזהה תהליך/שירות באותו מחשב "
         "(לדוגמה 5050 = השרת שלנו). יחד הם מהווים endpoint יחיד."),
        ("דוגמאות לפרוטוקולים בשכבת האפליקציה?",
         "HTTP (web), SMTP (email), DNS (domain name), FTP (file transfer), SSH (remote shell), "
         "ובפרויקט שלי — JSON-Lines משלי."),
        ("למה משמשת הפקודה PING?",
         "שולחת ICMP Echo Request למחשב מטרה ומחכה ל-Echo Reply. שימושית לבדוק "
         "קישוריות ולמדוד זמן round-trip."),
    ]
    for q, a in qa_pairs:
        he_segments(doc, [("ש: ", "bold"), (q, "he")], size=12)
        he_segments(doc, [("ת: ", "accent"), (a, "he")], size=12, space_after=8)

    heading(doc, "ד.2 קריפטוגרפיה", level=2)
    qa_pairs = [
        ("איזה סוגי הצפנות קיימים?",
         "Symmetric (אותו מפתח להצפנה+פענוח, מהיר — AES, ChaCha20) ו-Asymmetric "
         "(זוג מפתחות ציבורי+פרטי, איטי — RSA, ECC). בנוסף, hashing הוא חד-כיווני "
         "(לא הצפנה אלא טביעת אצבע — SHA-256, bcrypt)."),
        ("מה זה מפתח ציבורי ומה זה מפתח פרטי?",
         "במערכת asymmetric יש זוג מפתחות מתמטית קשורים. כל מה שמוצפן בציבורי אפשר "
         "לפענח רק בפרטי, ולהיפך. הציבורי ניתן לכולם; הפרטי נשמר בסוד."),
        ("מה זה חתימה דיגיטלית?",
         "מחשבים hash של ההודעה ומצפינים אותו במפתח הפרטי של החותם. כל אחד עם "
         "המפתח הציבורי יכול לאמת שזה באמת אותו אדם חתם (אי-הכחשה)."),
        ("מה זה hash?",
         "פונקציה חד-כיוונית שמקבלת קלט בכל גודל ומחזירה פלט בגודל קבוע (לדוגמה 256 ביט). "
         "תכונות: deterministic, fast to compute, hard to invert, collision-resistant. "
         "שימושים: צ'ק-סאם, סיסמאות (bcrypt), חתימות, blockchain."),
        ("למה bcrypt עדיף על SHA-256 לסיסמאות?",
         "(1) bcrypt אטי בכוונה (work factor), כך שתוקף לא יכול לבדוק מיליארדי ניחושים "
         "בשנייה. (2) bcrypt מוסיף salt אקראי לכל סיסמה, כך שאפילו סיסמאות זהות מקבלות "
         "hashes שונים — rainbow tables חסרי תועלת."),
        ("האם הסיסמה עוברת מוצפנת ברשת בפרויקט שלך?",
         "לא — הסיסמה עוברת ב-JSON-Lines בקליר על TCP. לפי הנחיית המורה הראשונית, "
         "הפרויקט נדרש להצפנת סיסמאות במנוחה בלבד (bcrypt על הדיסק) ולא הצפנת תקשורת. "
         "להגנה אמיתית בפרודקשן הייתי עוטף את ה-TCP ב-TLS."),
    ]
    for q, a in qa_pairs:
        he_segments(doc, [("ש: ", "bold"), (q, "he")], size=12)
        he_segments(doc, [("ת: ", "accent"), (a, "he")], size=12, space_after=8)

    heading(doc, "ד.3 מערכות הפעלה", level=2)
    qa_pairs = [
        ("מה זה thread?",
         "יחידת ביצוע בתוך תהליך (process). threads באותו process חולקים זיכרון "
         "(global vars, heap) אבל יש להם stack משלהם. ה-OS מתזמן threads למעבד באמצעות "
         "preemptive scheduling."),
        ("מה זה process?",
         "מופע פעיל של תוכנית. לכל process יש זיכרון משלו, PID, טבלת קבצים פתוחים, "
         "ומשתני סביבה. ל-OS עולה הרבה יותר לעבור בין processes (context switch) מאשר "
         "בין threads."),
        ("מה זה WinAPI?",
         "הממשק התכנותי של Windows — כ-9000 פונקציות לגישה לכל יכולת של המערכת "
         "(קבצים, תהליכים, חלונות, רשת, …). Tkinter ב-Python קורא תחתית WinAPI דרך "
         "פקדי Win32 בפועל."),
        ("מה זה GIL ולמה הוא חשוב?",
         "Global Interpreter Lock — לוק יחיד ב-CPython שמבטיח שרק thread אחד מריץ "
         "byte-code בכל רגע. תוצאה: threads ב-Python לא מקבילים על CPU, אבל בזמן I/O "
         "(socket.recv, time.sleep) ה-GIL משוחרר — אז threading עדיין שימושי "
         "ל-I/O-bound code כמו השרת שלנו."),
    ]
    for q, a in qa_pairs:
        he_segments(doc, [("ש: ", "bold"), (q, "he")], size=12)
        he_segments(doc, [("ת: ", "accent"), (a, "he")], size=12, space_after=8)

    heading(doc, "ד.4 קבצים", level=2)
    qa_pairs = [
        ("מה זה magic number בקובץ?",
         "מספר/חתימה קצרה בהתחלת קובץ שמזהה את סוגו. דוגמאות: PNG (0x89 50 4E 47), "
         "PDF (%PDF-), ZIP (PK\\x03\\x04), PE/EXE (MZ). מפעיל הקובץ קורא את ה-magic "
         "לפני שהוא מנסה לעבד אותו."),
        ("מה זה FAT32?",
         "מערכת קבצים ישנה (Microsoft, 1996) המבוססת על File Allocation Table. "
         "מגבלות: קובץ בודד ≤ 4GB, partition ≤ 8TB, אין הרשאות נפרדות. עדיין נפוץ "
         "ב-USB drives וב-SD cards בגלל תאימות רחבה."),
        ("מה זה קובץ PE (קובץ הרצה)?",
         "Portable Executable — פורמט הקבצים הניתנים להרצה ב-Windows (.exe, .dll, "
         ".sys). מבנה: DOS stub → PE header → section table → sections (.text, .data, "
         ".rsrc). הטוען של Windows קורא את ה-PE header ומעמיס את הקובץ לזיכרון לפני "
         "שהוא מעביר שליטה ל-entry point."),
    ]
    for q, a in qa_pairs:
        he_segments(doc, [("ש: ", "bold"), (q, "he")], size=12)
        he_segments(doc, [("ת: ", "accent"), (a, "he")], size=12, space_after=8)

    heading(doc, "ד.5 סייבר — אבטחת הפרויקט", level=2)
    qa_pairs = [
        ("לאלו מתקפות הפרויקט חשוף?",
         "(1) האזנה לתעבורה — סיסמה ומהלכים בקליר. (2) MITM — תוקף יכול לזייף מהלכים. "
         "(3) Brute-force מקוון על login — bcrypt מאט אבל לא חוסם לחלוטין. "
         "(4) DoS אפליקטיבי — לקוח יכול לפתוח חיבורים רבים. אלו חולשות שמודעים אליהן "
         "ומחוץ ל-scope של הפרויקט לפי הנחיית המורה."),
        ("אלו הגנות יושמו בפרויקט?",
         "(1) bcrypt לסיסמאות במנוחה. (2) Validation של כל שדה מהלקוח. "
         "(3) State machine — לא ניתן לשלוח move בלי להיות IN_GAME. "
         "(4) try/except — נפילת לקוח לא מפילה את השרת. "
         "(5) threading.Lock על כל מבנה משותף. "
         "(6) atomic write ל-users.json עם os.replace + os.fsync. "
         "(7) AlreadyLoggedInError — חיבור כפול לאותו חשבון נחסם."),
        ("מה זה MITM?",
         "Man-In-The-Middle — תוקף שיושב בין הלקוח לשרת, מאזין ואולי משנה הודעות. "
         "הגנה: TLS עם תעודות שמאמתות את השרת. בפרויקט שלי TLS לא מומש (לפי הנחיית "
         "המורה), אז יש להריץ רק על localhost / VPN פנימי."),
        ("מה זה SQL Injection?",
         "התקפה שבה תוקף מזריק קוד SQL לתוך שדה קלט שמשולב ב-query. תוצאה: יכול לקרוא, "
         "לשנות או למחוק את כל ה-DB. הגנה: prepared statements. בפרויקט שלי אין SQL "
         "כלל — אחסון = JSON עם json.dumps שמקודד תווים מיוחדים."),
        ("מה זה Buffer Overflow?",
         "כתיבה מעבר לסוף buffer בזיכרון, שעלולה לדרוס return address ולגרום להרצת "
         "קוד שרירותי. רלוונטי לשפות עם pointer arithmetic (C/C++). בפרויקט שלי "
         "(Python) לא רלוונטי — אין pointer arithmetic, רק bounded slices."),
        ("איך נשמרת אבטחת הנתונים בפרויקט?",
         "(1) סיסמאות — hashed עם bcrypt + salt בקובץ data/users.json. "
         "(2) Server-authoritative: כל מהלך מאומת בצד שרת לפני שמופעל על הלוח. "
         "(3) State machine: לקוח לא יכול לקפוץ בין מצבים (BadStateError). "
         "(4) atomic write — קובץ users.json לעולם לא נשאר חצי-כתוב גם בקריסה."),
    ]
    for q, a in qa_pairs:
        he_segments(doc, [("ש: ", "bold"), (q, "he")], size=12)
        he_segments(doc, [("ת: ", "accent"), (a, "he")], size=12, space_after=8)


if __name__ == "__main__":
    build_document()
