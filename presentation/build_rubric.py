from __future__ import annotations

from pathlib import Path
from typing import Sequence

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


REPO_ROOT = Path(__file__).resolve().parent.parent
OUT_DOCX = REPO_ROOT / "מחוון-תיק-פרויקט.docx"

HE_FONT = "Arial"

ACCENT = RGBColor(0x1F, 0x6F, 0xEB)
ACCENT_DARK = RGBColor(0x0E, 0x4A, 0xA8)
DARK = RGBColor(0x1F, 0x23, 0x28)
MUTED = RGBColor(0x57, 0x60, 0x6A)
TABLE_HEAD_BG = RGBColor(0xEA, 0xF2, 0xFD)
WARN_BG = RGBColor(0xFF, 0xF8, 0xE1)
WARN_BORDER = RGBColor(0xF5, 0xA6, 0x23)


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


def _set_paragraph_shading(paragraph, hex_color: str) -> None:
    pPr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    pPr.append(shd)


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


def rgb_hex(rgb: RGBColor) -> str:
    return "{:02X}{:02X}{:02X}".format(rgb[0], rgb[1], rgb[2])


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


def he_paragraph(doc, text=None, *, style=None, bold=False, size=12,
                 italic=False, color=None, space_before=2, space_after=4,
                 align=WD_ALIGN_PARAGRAPH.RIGHT):
    if style:
        p = doc.add_paragraph(style=style)
    else:
        p = doc.add_paragraph()
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


def warn_box(doc, text: str):
    p = he_paragraph(doc, text, size=11, bold=True, color=DARK,
                     space_before=4, space_after=4)
    _set_paragraph_shading(p, rgb_hex(WARN_BG))
    _add_paragraph_border(p, "right", rgb_hex(WARN_BORDER), sz=18)
    _add_paragraph_border(p, "top", rgb_hex(WARN_BORDER), sz=4)
    _add_paragraph_border(p, "bottom", rgb_hex(WARN_BORDER), sz=4)
    _add_paragraph_border(p, "left", rgb_hex(WARN_BORDER), sz=4)
    p.paragraph_format.left_indent = Cm(0.3)
    p.paragraph_format.right_indent = Cm(0.3)
    return p


def add_table(doc, headers: Sequence[str], rows: Sequence[Sequence[str]],
              col_widths: Sequence[float] | None = None):
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


def he_bullet(doc, text: str):
    he_paragraph(doc, text, style="List Bullet", space_before=1, space_after=1)


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

    write_part_zero(doc)
    page_break(doc)

    write_part_a(doc)
    page_break(doc)

    write_part_b(doc)
    page_break(doc)

    write_score_summary(doc)
    page_break(doc)

    write_oral_rubric(doc)
    page_break(doc)

    write_examiner_notes(doc)

    doc.save(str(OUT_DOCX))
    _word_postprocess_rtl(OUT_DOCX)
    print(f"wrote rubric docx ({OUT_DOCX.stat().st_size:,} bytes)")


def _word_postprocess_rtl(docx_path: Path) -> None:
    try:
        import win32com.client
    except ImportError:
        return
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(str(docx_path))
        for para in doc.Paragraphs:
            try:
                para.Format.ReadingOrder = 0
            except Exception:
                pass
        doc.Save()
        doc.Close(SaveChanges=False)
    finally:
        word.Quit()


def write_cover(doc):
    he_paragraph(doc, "", space_before=60)
    he_paragraph(doc, "מחוון לבדיקת תיק הפרויקט", size=24, bold=True,
                 color=ACCENT_DARK, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    he_paragraph(doc, "Secure Chess — שחמט מאובטח רב-משתמשים", size=18, bold=True,
                 color=ACCENT, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=20)
    he_paragraph(doc, 'חלופה: מערכות הגנת סייבר במקצוע "תכנון ותכנות מערכות"',
                 size=12, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    he_paragraph(doc, "5 יחידות לימוד — עבודת גמר", size=12, bold=True,
                 align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    he_paragraph(doc, 'מבוסס על מסמך הפיקוח של משרד החינוך, אגף מדעיות-הנדסיות, תשפ"ג',
                 size=11, italic=True, color=MUTED,
                 align=WD_ALIGN_PARAGRAPH.CENTER, space_after=30)

    heading(doc, "פרטי התלמיד", level=2)
    add_table(doc, ["שדה", "ערך"], [
        ("שם בית הספר", "_____________________"),
        ("סמל מוסד", "_____________________"),
        ("שם התלמיד", "_____________________"),
        ('מספר ת"ז', "_____________________"),
        ("נושא העבודה", "משחק שחמט מאובטח רב-משתמשים מבוסס שרת/לקוח עם הצפנת סיסמאות bcrypt"),
        ("תאריך הערכה", "_____________________"),
        ("שם הבוחן", "_____________________"),
    ], col_widths=[5.5, 11.0])


def write_part_zero(doc):
    heading(doc, "חלק 0 — דרישות חובה בפרויקט (תנאי סף)", level=1)
    warn_box(doc,
             "⚠️ עבודה ללא דרישות חובה 1-5 אינה מתקבלת. שימוש ב-Unity אסור. "
             "פרויקטים הממשים תקיפה אינם מאושרים.")

    add_table(doc, ["#", "נושא", "דרישה", "סטטוס", "מיקום בפרויקט / הוכחה"], [
        ("1", "תכנות מונחה עצמים",
         "מימוש לפחות 4 מחלקות שונות",
         "✅",
         "21 מחלקות אפליקטיביות: Piece, Move, Board, Game, Result, Account, UserStore, "
         "AIPlayer, _LineReader, Session, AISession, _AIAccount, Lobby, GameRegistry, "
         "ChessServer, ClientConnection, ServerConnection, ChessGui, LoginFrame, "
         "LobbyFrame, GameFrame (חלק 4.1 בתיק)"),

        ("2a", "תקשורת — שרת ולקוח",
         "מימוש שרת ולקוח מבוססי סוקטים",
         "✅",
         "src/secure_chess/server/server.py (TCP listen+accept) + "
         "src/secure_chess/client/{cli,gui}.py (TCP connect)"),

        ("2b", "תקשורת — ריבוי לקוחות",
         "מימוש שרת מרובה לקוחות",
         "✅",
         "ChessServer._accept_loop יוצר Thread לכל חיבור."),

        ("2c", "תקשורת — פרוטוקול",
         'פרוטוקול העברת הודעות מצד לצד הגיוני ע"י התלמיד',
         "✅",
         "JSON-Lines מותאם אישית ב-src/secure_chess/common/protocol.py; 13 סוגי הודעות"),

        ("3a", "מערכת הפעלה — Threads",
         "שימוש בתהליכונים (Thread)",
         "✅",
         "threading.Thread ב-accept loop ו-Session.run; "
         "threading.Lock ב-Lobby, GameRegistry, UserStore, ChessServer"),

        ("3b", "מערכת הפעלה — קבצים/API",
         "גישה למערכת הקבצים / API / רכיב חומרה",
         "✅",
         "כתיבה אטומית ל-data/users.json עם os.replace + os.fsync; "
         "Windows API נצרך דרך Tkinter (Win32 widgets)"),

        ("4a", "אבטחה — הצפנה",
         "הצפנת מידע רגיש בתקשורת",
         "📋",
         "לפי הנחיית המורה הראשונית: הצפנת סיסמאות בלבד, ללא הצפנת תקשורת. "
         "Hashing עם bcrypt + per-password salt ב-src/secure_chess/common/crypto.py"),

        ("4b", "אבטחה — טיפול בפרצות",
         "טיפול בפרצות אבטחה בפרויקט",
         "✅",
         "(1) bcrypt לסיסמאות; (2) atomic write נגד corruption; "
         "(3) try/except 3-שכבות ב-Session.run; (4) threading.Lock נגד race conditions; "
         "(5) AlreadyLoggedInError נגד double-login; (6) State Machine + BadStateError"),

        ("5", "ממשק משתמש",
         "מימוש ממשק משתמש אינטראקטיבי",
         "✅",
         "Tkinter GUI עם לוח 8×8 לחיץ (3 מסכים: Login/Lobby/Game) + CLI REPL "
         "(client/gui.py + client/cli.py)"),
    ])

    warn_box(doc,
             "📋 הערה לסעיף 4a: לפי הנחיית המורה הראשונית, נדרשת הצפנת סיסמאות בלבד "
             "(ללא הצפנת מידע רגיש העובר בתקשורת). הפרויקט מקיים את הדרישה הזו במלואה "
             "באמצעות bcrypt עם salt-per-password. ראה תיאור הרעיון בתיק §1.1.")

    he_paragraph(doc, "סיכום תנאי סף: 9/9 ✅ — הפרויקט עומד בכל דרישות החובה.",
                 bold=True, color=ACCENT, size=13)


def write_part_a(doc):
    heading(doc, "חלק א' — מראה התיק (15%)", level=1)
    add_table(doc, ["מרכיב", "תיאור הדרישות", "משקל", "ניקוד", "סטטוס בתיק", "הערות"], [
        ("שער פתיחה", "על פי התבנית", "2%", "___", "✅",
         "שער התיק מולא: שם פרויקט, שם תלמיד, מס' ת\"ז, סמל מוסד, ביה\"ס, מורה, תאריך"),
        ("תוכן עניינים", "מקושר לפרקים בתיק", "2%", "___", "✅",
         "Word יוצר תוכן עניינים אוטומטי המקושר ל-Heading 1/2/3"),
        ("גופן אחיד", "לכלל התיק", "2%", "___", "✅",
         "Arial 12 לטקסט רץ, Consolas 9 לקטעי קוד"),
        ("כותרות", "היררכיה ועיצוב עקבי", "3%", "___", "✅",
         "Heading 1/2/3 עם צבעים עקביים (#0E4AA8, #1F6FEB, #1F2328)"),
        ("מספרי עמוד", "בכל העמודים", "3%", "___", "✅",
         "Footer → Page Number"),
        ("עימוד דפים אחיד", "שוליים ומרווחי שורות", "3%", "___", "✅",
         "שוליים 2 ס\"מ סביב, רווח שורות 1"),
    ])
    he_paragraph(doc, "סה\"כ חלק א': ___ / 15%", bold=True, color=ACCENT)


def write_part_b(doc):
    heading(doc, "חלק ב' — תוכן התיק (85%)", level=1)

    heading(doc, "1. מבוא (ייזום, אפיון) — 13%", level=2)
    add_table(doc, ["תת-סעיף", "סטטוס", "מיקום / הערות"], [
        ("תיאור הרעיון והמוטיבציה", "✅", "מבוא §1.1 בתיק הפרויקט"),
        ("ייזום: זיהוי הצורך", "✅", "מבוא §1.2"),
        ("אפיון פונקציונלי", "✅", "מבוא §1.3 — 9 דרישות פונקציונליות"),
        ("אפיון לא-פונקציונלי", "✅", "מבוא §1.4 — 6 דרישות איכות"),
        ("לוח זמנים מעודכן והגיוני", "✅", "מבוא §1.5 — טבלת אבני דרך US1-US6"),
        ("ניהול הסיכונים מעודכן", "✅", "מבוא §1.6 — טבלת 6 סיכונים עיקריים"),
    ])
    he_paragraph(doc, "ניקוד מבוא: ___ / 13%", bold=True, color=ACCENT)

    heading(doc, "2. תיאור תחום הידע — פרק מילולי (ניתוח) — 10%", level=2)
    add_table(doc, ["תת-סעיף", "סטטוס", "מיקום / הערות"], [
        ("חוקי שחמט וניהול משחק", "✅", "ניתוח §2.1"),
        ("מודל OSI ושכבת התעבורה (TCP/UDP, סוקטים)", "✅", "ניתוח §2.2"),
        ("קריפטוגרפיה — Hash ו-bcrypt", "✅", "ניתוח §2.3"),
        ("ריבוי תהליכונים ו-OS (Threads, Locks, GIL)", "✅", "ניתוח §2.4"),
        ("אחסון בקבצים — Atomic file writes", "✅", "ניתוח §2.5"),
        ("חלוקת מטלות שרת/לקוח", "✅", "ניתוח §2.6"),
    ])
    he_paragraph(doc, "ניקוד ניתוח: ___ / 10%", bold=True, color=ACCENT)

    heading(doc, "3. מבנה / ארכיטקטורה (העיצוב) — 25%", level=2)
    add_table(doc, ["תת-סעיף", "סטטוס", "מיקום / הערות"], [
        ("שרטוט ארכיטקטורה של הפרויקט", "✅", "ארכיטקטורה §3.1 — תרשים בלוקים"),
        ("תרשימי זרימת מידע (sequence diagrams)", "✅", "ארכיטקטורה §3.2 — 3 דיאגרמות"),
        ("הבחנה בין מודולים בצד שרת לצד לקוח", "✅", "ארכיטקטורה §3.3 — טבלת 16 מודולים"),
        ("ניתוח אלגוריתמים מרכזיים + שקילת חלופות", "✅", "ארכיטקטורה §3.4 — 4 השוואות"),
        ("פרוטוקול התקשורת — הגיוני וישים", "✅", "ארכיטקטורה §3.5 — state machine + 13 הודעות"),
        ("ניתוח חולשות + פתרונות בכל רובד", "✅", "ארכיטקטורה §3.6 — מודל איומים תואם הנחיית המורה"),
    ])
    he_paragraph(doc, "ניקוד ארכיטקטורה: ___ / 25%", bold=True, color=ACCENT)

    heading(doc, "4. מימוש הפרויקט (הקוד) — 31%", level=2)
    add_table(doc, ["תת-סעיף", "סטטוס", "מיקום / הערות"], [
        ("תכנות מונחה עצמים — מחלקות שהתלמיד יצר", "✅",
         "21 מחלקות אפליקטיביות + 13 שגיאה (ראה דרישת חובה #1)"),
        ("חלוקה הגיונית לקבצים", "✅",
         "common/ + server/ + client/"),
        ("קוד כתוב היטב", "✅",
         "type hints מלאים, PEP-8, שמות משמעותיים"),
        ("תיעוד הגיוני", "✅",
         "module docstrings בכל קובץ"),
        ("חלוקה לפעולות עם תיעוד", "✅",
         "פונקציות קצרות וממוקדות; מספרי שורה: <500 לכל קובץ"),
        ("חלוקה ברורה בין קוד שרת לקוד לקוח", "✅",
         "server/ + client/ מבודדים; common/ משותף"),
        ("קוד בטוח — try/except, שרת יציב לא קורס", "✅",
         "Session.run עוטף כל dispatch ב-except SecureChessError/Exception בשלוש שכבות; "
         "ניתוק לקוח מטופל ב-_on_disconnect"),
    ])
    he_paragraph(doc, "ניקוד מימוש: ___ / 31%", bold=True, color=ACCENT)

    heading(doc, "5. מדריך למשתמש — 10%", level=2)
    add_table(doc, ["תת-סעיף", "סטטוס", "מיקום / הערות"], [
        ("הסבר התקנה והרצה (Quick start)", "✅", "מדריך §5.1"),
        ("צילומי מסך של הזרימה המלאה", "⚠️", "מדריך §5.2 — מקום מסומן ל-5 תמונות שעל התלמיד להוסיף"),
        ("הסברים מפורטים לכל סוגי המשתמשים", "✅", "מדריך §5.3 — שחקן, יריב AI, מנהל שרת"),
        ("הוכחת הצפנת הסיסמאות במנוחה", "✅", "מדריך §5.4 — Get-Content + Select-String"),
    ])
    he_paragraph(doc, "ניקוד מדריך: ___ / 10%", bold=True, color=ACCENT)

    heading(doc, "6. סיכום אישי / רפלקציה — 6%", level=2)
    warn_box(doc, '⚠️ אסור להסתפק ב"תודה ונהניתי". יש לפרט מה למדת על עצמך ובכלל.')
    add_table(doc, ["תת-סעיף", "סטטוס", "מיקום / הערות"], [
        ("מה למדתי על עצמי", "⚠️", "סיכום §6.1 — תבנית; חייב מילוי אישי ע\"י התלמיד"),
        ("מה למדתי מקצועית", "✅", "סיכום §6.2 — 10 לקחים מקצועיים מהפרויקט"),
        ("מה הייתי משנה אם הייתי מתחיל מחדש", "✅", "סיכום §6.3 — 6 שיפורים עתידיים"),
    ])
    he_paragraph(doc, "ניקוד רפלקציה: ___ / 6%", bold=True, color=ACCENT)

    heading(doc, "7. ביבליוגרפיה — 5%", level=2)
    add_table(doc, ["תת-סעיף", "סטטוס", "מיקום / הערות"], [
        ("סקר ספרות מפורט (לא רק רשימת URLs)", "✅",
         "ביבליוגרפיה §7 — 8 מקורות עם הסבר תרומה לפרויקט"),
    ])
    he_paragraph(doc, "ניקוד ביבליוגרפיה: ___ / 5%", bold=True, color=ACCENT)

    heading(doc, "8. נספחים", level=2)
    add_table(doc, ["תת-סעיף", "סטטוס", "מיקום / הערות"], [
        ("תדפיס הקוד עם תיעוד", "✅", "נספח א' — תדפיס כל מודולי src/secure_chess/"),
        ("טבלת הודעות הפרוטוקול", "✅", "נספח ב' — 13 סוגי הודעות + דוגמת זרימה"),
        ("שאלות תיאורטיות לבחינה", "✅", "נספח ג' — תקשורת, קריפטו, OS, קבצים, סייבר"),
    ])

    heading(doc, "בונוס — עד 10%", level=2)
    add_table(doc, ["מרכיב בונוס", "סטטוס בפרויקט", "ניקוד מומלץ"], [
        ("נושא מורכב מאוד",
         "✅ רשת + ריבוי שחקנים + AI alpha-beta + חוקי שחמט מלאים",
         "___"),
        ("קוד בהיקף רציני",
         "✅ ~3,000 שורות מקור אפליקטיבי",
         "___"),
        ("חריג ביחס לפרויקטים בקבוצה",
         "✅ State machine מלאה, GUI Tkinter, atomic-write persistence, AI minimax",
         "___"),
    ])
    he_paragraph(doc, "ניקוד בונוס: ___ / 10%", bold=True, color=ACCENT)


def write_score_summary(doc):
    heading(doc, "חישוב ציון סופי", level=1)
    add_table(doc, ["מרכיב", "משקל", "ניקוד"], [
        ("חלק א' (מראה התיק)", "15%", "___"),
        ("חלק ב' (תוכן התיק):", "85%", "___"),
        ("  ↳ מבוא", "13%", "___"),
        ("  ↳ ניתוח", "10%", "___"),
        ("  ↳ ארכיטקטורה", "25%", "___"),
        ("  ↳ מימוש", "31%", "___"),
        ("  ↳ מדריך משתמש", "10%", "___"),
        ("  ↳ רפלקציה", "6%", "___"),
        ("  ↳ ביבליוגרפיה", "5%", "___"),
        ("סה\"כ לפני בונוס", "100%", "___"),
        ("בונוס", "+10%", "___"),
        ("ציון סופי (כולל בונוס, חסום ב-100)", "—", "___"),
    ])


def write_oral_rubric(doc):
    heading(doc, "חלק ב' (אחר) — מחוון לבדיקת הפרויקט בזמן הבחינה", level=1)
    warn_box(doc, '⚠️ הפרויקט חייב לעבוד במהלך הבדיקה. לא סרטון, לא מצגת, ולא "אתמול הכל עבד".')

    add_table(doc, ["מרכיב", "משקל", "ניקוד", "קריטריונים", "הפניות בפרויקט"], [
        ("הצגה ושליטה בפרויקט", "15%", "___",
         "מציג כל חלק, מפרט, מריץ, שולט בתהליך",
         "שרת + שני CLI clients + GUI client במקביל"),
        ("פרויקט עובד מקצה לקצה", "30%", "___",
         "העברת מידע מקצה לקצה ב-≥ יכולת אחת",
         "register → login → join lobby → game_started → move → checkmate"),
        ("שליטה בקוד", "30%", "___",
         "ניווט, הסבר, איך נשמרים ונשלפים נתונים",
         "data/users.json (persistence) + UserStore (load/save) + atomic write"),
        ("שליטה בחומר התיאורטי", "20%", "___",
         "תקשורת, קריפטוגרפיה, OS, קבצים",
         "ראה דוגמאות שאלות בנספח ד' בתיק הפרויקט"),
        ("סייבר", "5%", "___",
         "אלו מתקפות, אלו הגנות, איך נשמרת אבטחה",
         "ראה threat model בארכיטקטורה §3.6"),
    ])
    he_paragraph(doc, "סה\"כ ציון בחינה: ___ / 100%", bold=True, color=ACCENT)


def write_examiner_notes(doc):
    heading(doc, "הערות לבוחן — איך להריץ את הפרויקט", level=1)
    he_bullet(doc, "קוד מקור: src/secure_chess/ (Python 3.11+, 16 מודולים)")
    he_bullet(doc, "תלויות: pyproject.toml (bcrypt>=4.1)")
    he_bullet(doc, "הוכחת ריבוי שחקנים: הפעלת שרת + שני לקוחות מקבילים שמשחקים זה נגד זה")
    he_bullet(doc, "הוכחת AI (bonus): לקוח אחד בוחר 'Start AI Game' מה-Lobby")
    he_bullet(doc, "הוכחת אטומיות אחסון: התבוננות ב-data/users.json אחרי register; אין plaintext")

    heading(doc, "סטטוס ההפקה", level=1)
    he_paragraph(doc,
                 "💡 הפקה אוטומטית מחדש: הרץ python presentation/build_portfolio.py "
                 "ו-python presentation/build_rubric.py מתוך תיקיית הריפו — זה ייצור "
                 "מחדש את שני קבצי ה-docx על סמך מבנה הקוד הנוכחי.")

    heading(doc, "מה כבר נכלל בתיק הפורמלי (תיק-פרויקט.docx)", level=2)
    he_bullet(doc, 'שער פתיחה עם כל השדות הנדרשים (סמל מוסד, ת"ז, ביה"ס, מורה, תאריך…)')
    he_bullet(doc, "תוכן עניינים מקושר ל-Heading 1/2/3 (לוחצים F9 ב-Word לעדכון)")
    he_bullet(doc, "פרק 1 — מבוא: רעיון, מוטיבציה, ייזום, 9 דרישות פונקציונליות, 6 לא-פונקציונליות, לו\"ז, 6 סיכונים")
    he_bullet(doc, "פרק 2 — ניתוח: חוקי שחמט, מודל OSI, hash+bcrypt, threading + locks, atomic writes")
    he_bullet(doc, "פרק 3 — ארכיטקטורה: תרשים בלוקים, 3 sequence diagrams, 16 מודולים, ניתוח 4 אלגוריתמים, state machine, מודל איומים תואם הנחיית המורה")
    he_bullet(doc, "פרק 4 — מימוש: 21 מחלקות אפליקטיביות, מבנה תיקיות, דפוס robustness 3 שכבות")
    he_bullet(doc, "פרק 5 — מדריך משתמש: התקנה, הרצה, הנחיות ל-3 סוגי משתמשים, הוכחת bcrypt")
    he_bullet(doc, "פרק 6 — סיכום אישי: 10 לקחים מקצועיים + 6 דברים שהייתי משנה")
    he_bullet(doc, "פרק 7 — ביבליוגרפיה: 8 מקורות עם הסבר תרומה (סקר ספרות אמיתי)")
    he_bullet(doc, "נספח א' — תדפיס מלא של כל קוד המקור")
    he_bullet(doc, "נספח ב' — טבלת הודעות הפרוטוקול המלאה (13 סוגים)")
    he_bullet(doc, "נספח ג' — שאלות תיאורטיות לבחינה עם תשובות מוכנות")

    heading(doc, "מה נשאר על התלמיד (ידני)", level=2)
    warn_box(doc,
             "⚠️ §5.2 — צילומי מסך של ה-GUI: יש להריץ את ה-GUI ולצלם 5 מסכים "
             "(Login, Lobby, Lobby-Waiting, Game-Start, Game-Mid). מקומות מסומנים מראש בתיק.")
    warn_box(doc,
             "⚠️ §6.1 — מה למדתי על עצמי: סעיף אישי שחייב להיות במילותיך. "
             "תבנית מוכנה בתיק עם הנחיה — להחליף ב-2-3 פסקאות אותנטיות.")
    warn_box(doc,
             '⏳ מילוי השדות בשער: סמל מוסד, ת"ז, ביה"ס, מורה, תאריך — '
             "קווים ריקים בטבלה בעמ' 1 של התיק.")


if __name__ == "__main__":
    build_document()
