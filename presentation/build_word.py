from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


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



def _set_bidi(paragraph) -> None:
    pPr = paragraph._p.get_or_add_pPr()
    bidi = pPr.find(qn("w:bidi"))
    if bidi is None:
        bidi = OxmlElement("w:bidi")
        pPr.append(bidi)


def _set_run_rtl(run) -> None:
    rPr = run._r.get_or_add_rPr()
    rtl = rPr.find(qn("w:rtl"))
    if rtl is None:
        rtl = OxmlElement("w:rtl")
        rtl.set(qn("w:val"), "1")
        rPr.append(rtl)


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



def setup_styles(doc: Document) -> None:
    style = doc.styles["Normal"]
    style.font.name = HE_FONT
    style.font.size = Pt(12)
    style.font.color.rgb = DARK
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:cs"), HE_FONT)
    rfonts.set(qn("w:ascii"), HE_FONT)
    rfonts.set(qn("w:hAnsi"), HE_FONT)

    for heading_name, size, color, bold in [
        ("Heading 1", 22, ACCENT_DARK, True),
        ("Heading 2", 17, ACCENT, True),
        ("Heading 3", 14, DARK, True),
    ]:
        s = doc.styles[heading_name]
        s.font.name = HE_FONT
        s.font.size = Pt(size)
        s.font.color.rgb = color
        s.font.bold = bold



def he_paragraph(doc, text=None, *, style=None, bold=False, size=12,
                 italic=False, color=None, space_before=2, space_after=4,
                 align=WD_ALIGN_PARAGRAPH.RIGHT):
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
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
            _set_shading(run, "F1F3F5")
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


def _set_shading(run, hex_color: str) -> None:
    rPr = run._r.get_or_add_rPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    rPr.append(shd)


PY_KEYWORDS = {
    "def", "class", "if", "elif", "else", "for", "while", "try", "except",
    "finally", "raise", "return", "with", "as", "import", "from", "in",
    "is", "not", "and", "or", "lambda", "yield", "None", "True", "False",
    "pass", "break", "continue", "assert", "global", "self",
}


def code_block(doc, lines, *, size=10):
    table = doc.add_table(rows=1, cols=1)
    table.autofit = False
    cell = table.cell(0, 0)
    _set_cell_shading(cell, rgb_hex(CODE_BG))

    cell.text = ""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    for i, line in enumerate(lines):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        _highlight_code_line(p, line, size)
    return table


def _highlight_code_line(paragraph, text: str, size: int) -> None:
    if not text:
        run = paragraph.add_run(" ")
        run.font.name = CODE_FONT
        run.font.size = Pt(size)
        run.font.color.rgb = CODE_FG
        return

    leading = len(text) - len(text.lstrip(" "))
    if leading:
        _add_code_run(paragraph, " " * leading, size, CODE_FG)
        text = text[leading:]

    comment_pos = _find_comment(text)
    if comment_pos == 0:
        _add_code_run(paragraph, text, size, RGBColor(0x75, 0x71, 0x5E))
        return
    code_part = text if comment_pos < 0 else text[:comment_pos]
    comment_part = "" if comment_pos < 0 else text[comment_pos:]
    _emit_code_tokens(paragraph, code_part, size)
    if comment_part:
        _add_code_run(paragraph, comment_part, size, RGBColor(0x75, 0x71, 0x5E))


def _find_comment(text: str) -> int:
    in_str = None
    i = 0
    while i < len(text):
        ch = text[i]
        if in_str:
            if ch == "\\":
                i += 2
                continue
            if ch == in_str:
                in_str = None
        else:
            if ch in ('"', "'"):
                in_str = ch
            elif ch == "#":
                return i
        i += 1
    return -1


def _emit_code_tokens(paragraph, code: str, size: int) -> None:
    i = 0
    buf = ""

    def flush():
        nonlocal buf
        if buf:
            _add_code_run(paragraph, buf, size, CODE_FG)
            buf = ""

    while i < len(code):
        ch = code[i]
        if ch in ('"', "'"):
            flush()
            quote = ch
            j = i + 1
            while j < len(code) and code[j] != quote:
                if code[j] == "\\":
                    j += 2
                    continue
                j += 1
            j = min(j + 1, len(code))
            _add_code_run(paragraph, code[i:j], size, RGBColor(0xE6, 0xDB, 0x74))
            i = j
            continue
        if ch.isalpha() or ch == "_":
            j = i
            while j < len(code) and (code[j].isalnum() or code[j] == "_"):
                j += 1
            word = code[i:j]
            if word in PY_KEYWORDS:
                flush()
                _add_code_run(paragraph, word, size, RGBColor(0xF9, 0x26, 0x72))
            else:
                buf += word
            i = j
            continue
        buf += ch
        i += 1
    flush()


def _add_code_run(paragraph, text: str, size: int, color: RGBColor) -> None:
    run = paragraph.add_run(text)
    run.font.name = CODE_FONT
    run.font.size = Pt(size)
    run.font.color.rgb = color


def callout(doc, segments, *, bg=CALLOUT_BG, border_color=ACCENT, size=11):
    p = he_segments(doc, segments, size=size,
                    space_before=4, space_after=4)
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
    run.font.color.rgb = ACCENT_DARK if level == 1 else ACCENT
    return p


def page_break(doc):
    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)


def quote_in_he(text: str) -> str:
    return f'"{text}"'



def build_document(out_path: Path) -> None:
    doc = Document()
    setup_styles(doc)

    section = doc.sections[0]
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)
    sectPr = section._sectPr
    bidi = OxmlElement("w:bidi")
    sectPr.append(bidi)

    write_cover(doc)
    page_break(doc)

    write_how_to_use(doc)
    page_break(doc)

    heading(doc, "חלק א'  —  הרקע שאתה חייב להבין לפני שמציגים")
    he_paragraph(doc,
                 "לפני שאתה ניגש לבוחן, ודא שאתה מבין את התשעה מושגי הליבה הבאים. "
                 "כל אחד מסביר חלק אחר בפרויקט. אם הבוחן ישאל \"מה זה X?\" "
                 "ולא תוכל להסביר במילים שלך, יורידו לך נקודות. אל תשנן — תבין.",
                 italic=True)
    write_background_section(doc)
    page_break(doc)

    heading(doc, "חלק ב'  —  מדריך שקופית-אחר-שקופית")
    he_paragraph(doc,
                 "לכל אחת מ‑30 השקופיות במצגת יש כאן: (1) רעיון מרכזי, "
                 "(2) טקסט שתגיד בערך, (3) פרטים בקוד שצריך להראות באצבע, "
                 "(4) שאלה שצפויה לעלות מהבוחן בנקודה הזאת.",
                 italic=True)
    write_slide_guide(doc)
    page_break(doc)

    heading(doc, "חלק ג'  —  תסריט להדגמה החיה")
    write_demo_script(doc)
    page_break(doc)

    heading(doc, "חלק ד'  —  שאלות צפויות מהבוחן ותשובות מוכנות")
    write_qa(doc)
    page_break(doc)

    heading(doc, "חלק ה'  —  מפת הקוד  (אילו קבצים לפתוח אם הבוחן יבקש)")
    write_files_map(doc)
    page_break(doc)

    heading(doc, "חלק ו'  —  מילון מונחים")
    write_glossary(doc)
    page_break(doc)

    heading(doc, "חלק ז'  —  דף תזכורת מהיר  (להדפיס ולקחת איתך)")
    write_cheat_sheet(doc)

    doc.save(str(out_path))
    print(f"wrote {out_path}")



def write_cover(doc):
    he_paragraph(doc, "", space_before=80)
    p = he_paragraph(doc, "Secure Two-Player Chess",
                     size=32, bold=True, color=ACCENT_DARK,
                     align=WD_ALIGN_PARAGRAPH.CENTER,
                     space_before=20, space_after=8)
    he_paragraph(doc, "משחק שחמט מאובטח בין שני שחקנים — Python, TCP, bcrypt, AI",
                 size=16, color=MUTED,
                 align=WD_ALIGN_PARAGRAPH.CENTER, space_after=30)
    p = he_paragraph(doc, "מדריך מקיף לתלמיד — איך להציג את הפרויקט לבוחן",
                     size=20, bold=True, color=ACCENT,
                     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)
    he_paragraph(doc,
                 "מסמך זה מיועד לתלמיד שלא בנה את הפרויקט בעצמו. "
                 "הוא מסביר את הקוד מהיסוד, נותן תסריט להצגה, "
                 "ומכין אותך לכל שאלה שהבוחן עלול לשאול.",
                 size=12, italic=True, color=MUTED,
                 align=WD_ALIGN_PARAGRAPH.CENTER, space_after=40)

    box_segments = [
        ("מה יש כאן: ", "bold"),
        ("רקע טכני (9 מושגים), מדריך ל‑30 שקופיות, תסריט הדגמה, "
         "20 שאלות+תשובות, מפת קוד, ומילון מונחים.", "he"),
    ]
    callout(doc, box_segments)


def write_how_to_use(doc):
    heading(doc, "איך להשתמש במסמך הזה")
    he_segments(doc, [
        ("השלב הראשון: ", "bold"),
        ("קרא את חלק א' (הרקע) פעם אחת בשלמותו. אל תדלג. ", "he"),
        ("הבוחן יזהה מיד אם אתה מסתמך על שינון.", "italic"),
    ])
    he_segments(doc, [
        ("השלב השני: ", "bold"),
        ("עבור על חלק ב' (השקופיות) פעמיים — פעם תוך כדי הסתכלות במצגת, "
         "ופעם נוספת תוך הסתכלות בקוד עצמו.", "he"),
    ])
    he_segments(doc, [
        ("השלב השלישי: ", "bold"),
        ("שחק את ההדגמה (חלק ג') לבד 3–4 פעמים, עד שאתה יכול לבצע אותה "
         "מבלי להסתכל במסמך.", "he"),
    ])
    he_segments(doc, [
        ("השלב הרביעי: ", "bold"),
        ("בקש ממישהו לשאול אותך את השאלות מחלק ד'. הוא לא צריך להבין "
         "תכנות — מספיק שיש לו את התשובות מולו.", "he"),
    ])
    he_segments(doc, [
        ("ביום של הבחינה: ", "bold"),
        ("הדפס רק את חלק ז' (דף תזכורת). תוכל לעיין בו לפני שאתה נכנס.", "he"),
    ])

    tip_box(doc, [
        ("עיקרון זהב: ", "bold"),
        ("אם הבוחן שואל ואתה לא יודע — תגיד \"אני לא בטוח, אבל לפי "
         "הקוד נראה לי ש‑X\". זה הרבה יותר טוב מלהמציא, וזה מראה שאתה מבין "
         "איך לחקור קוד.", "he"),
    ])


def write_background_section(doc):

    heading(doc, "1. מה זה הפרויקט, ב‑3 משפטים", level=2)
    he_paragraph(doc,
                 "זאת מערכת שמאפשרת לשני שחקנים אנושיים להירשם בשרת מרכזי "
                 "עם שם משתמש וסיסמה, להתחבר ולשחק נגד השני משחק שחמט מלא. "
                 "השרת מאמת את הסיסמאות, מאמת את חוקיות המהלכים, "
                 "ומזהה סוף משחק (שחמט, פט, תיקו).")
    he_paragraph(doc,
                 "מעבר לדרישות הבסיסיות, יש שתי תוספות שעושות את הפרויקט "
                 "ראוי לציון גבוה: (א) השרת תומך בכמה זוגות שחקנים בו זמנית, "
                 "(ב) במקום לשחק נגד אדם אחר, אפשר לשחק נגד יריב מחושב (AI).")

    heading(doc, "2. \"לקוח‑שרת\" (Client–Server) — איך זה עובד?", level=2)
    he_paragraph(doc,
                 "תוכנית \"שרת\" היא תוכנית אחת שרצה כל הזמן ומחכה שמישהו יתחבר "
                 "אליה. תוכנית \"לקוח\" היא תוכנית שעולה, מתחברת לשרת, מבקשת ממנו "
                 "דברים, ומפסיקה בסיום העבודה.")
    he_paragraph(doc,
                 "בפרויקט שלנו השרת רץ במחשב אחד (לרוב על אותו מחשב, על "
                 "כתובת 127.0.0.1). שני הלקוחות מתחברים אליו דרך הרשת. "
                 "השרת מתווך ביניהם — כשמישהו מבצע מהלך, השרת בודק ושולח "
                 "ללקוח השני את המצב המעודכן.")
    callout(doc, [
        ("למה השרת חייב להיות באמצע, במקום שהלקוחות ידברו ישירות? ", "bold"),
        ("כדי שלא יוכלו לרמות. השרת הוא ה‑", "he"),
        ("source of truth", "en"),
        (" — הוא היחיד שבודק חוקיות מהלכים, ", "he"),
        ("רק", "italic"),
        (" הוא שומר את הסיסמאות. הלקוח הוא רק \"מסך\" שמראה את התוצאות.", "he"),
    ])

    heading(doc, "3. TCP — איך נשלחים הביטים בין הלקוח לשרת", level=2)
    he_paragraph(doc,
                 "TCP זה פרוטוקול קישוריות ברמה נמוכה (השכבה הרביעית במודל "
                 "OSI). הוא נותן \"זרם בתים אמין\" — כל מה שצד אחד שולח יגיע "
                 "בסדר הנכון לצד השני, או שתיווצר שגיאת \"חיבור נפל\".")
    he_segments(doc, [
        ("בקוד שלנו השרת קורא ל‑", "he"),
        ("socket.socket()", "en"),
        (" כדי ליצור סוקט, ל‑", "he"),
        ("bind()", "en"),
        (" כדי לתפוס פורט (5050), ול‑", "he"),
        ("listen()", "en"),
        (" כדי להתחיל לחכות. כשלקוח מתחבר, ", "he"),
        ("accept()", "en"),
        (" מחזיר אובייקט סוקט חדש שמדבר ספציפית מול הלקוח הזה.", "he"),
    ])
    callout(doc, [
        ("נקודה שצפויה לעלות: ", "bold"),
        ("TCP מבטיח שהבייטים יגיעו בסדר, אבל הוא ", "he"),
        ("לא", "italic"),
        (" יודע איפה הודעה אחת נגמרת והשנייה מתחילה. לכן בנינו מעליו "
         "שכבת framing (ראה סעיף 4).", "he"),
    ])

    heading(doc, "4. JSON-Lines — איך כל הודעה מובדלת מהבאה", level=2)
    he_paragraph(doc,
                 "JSON זה פורמט טקסטואלי לתיאור אובייקטים. JSON-Lines זה "
                 "פשוט אוסף של אובייקטי JSON — אחד בכל שורה, מופרדים על‑ידי "
                 "תו ירידת שורה (\\n).")
    he_paragraph(doc, "דוגמה לחילופי הודעות בין הלקוח לשרת:")
    code_block(doc, [
        'C -> S  {"type":"register","username":"alice","password":"hunter2!"}',
        'S -> C  {"type":"ok"}',
        'C -> S  {"type":"move","uci":"e2e4"}',
        'S -> C  {"type":"ok"}',
        'S -> C  {"type":"game_state", "board": "...", "to_move": "black"}',
    ])
    he_segments(doc, [
        ("הקוד שמטפל בקבלה נמצא במחלקה ", "he"),
        ("_LineReader", "en"),
        (" בקובץ ", "he"),
        ("common/protocol.py", "en"),
        (". הוא קורא מהסוקט לתוך ", "he"),
        ("bytearray", "en"),
        (" עד שהוא רואה ", "he"),
        ("\\n", "en"),
        (", ואז חותך משם הודעה אחת ומחזיר אותה.", "he"),
    ])

    heading(doc, "5. הצפנת סיסמאות עם bcrypt — מה זה ולמה דווקא הוא", level=2)
    he_paragraph(doc,
                 "כשמשתמש נרשם בשרת אנחנו לא יכולים לשמור את הסיסמה בקובץ "
                 "כטקסט גלוי — אם מישהו ייחדור לשרת הוא יוכל לקרוא את כל "
                 "הסיסמאות. הפתרון: לשמור hash במקום הסיסמה.")
    he_paragraph(doc,
                 "Hash זה ערך שמחושב מהסיסמה בכיוון אחד — קל לחשב מהסיסמה "
                 "למה ה‑hash, אבל בלתי אפשרי לחשב מה‑hash בחזרה לסיסמה.")
    he_segments(doc, [
        ("bcrypt", "en"),
        (" הוא אלגוריתם hash מיוחד שתוכנן במיוחד לסיסמאות. שני דברים מבדילים "
         "אותו מ‑", "he"),
        ("SHA-256", "en"),
        (" \"רגיל\":", "he"),
    ])
    he_segments(doc, [
        ("(א) הוא ", "he"),
        ("איטי בכוונה", "italic"),
        (" — קל לחשב פעם אחת, אבל ניסיון לחשב מיליארדי סיסמאות "
         "בכוח הזרוע (", "he"),
        ("brute-force", "en"),
        (") הופך לבלתי מעשי.", "he"),
    ])
    he_segments(doc, [
        ("(ב) הוא מכיל ", "he"),
        ("salt", "en"),
        (" מובנה — מחרוזת אקראית שמוצמדת לסיסמה לפני ה‑hash. "
         "כך שני משתמשים שונים שבחרו אותה סיסמה יקבלו hash שונה.", "he"),
    ])
    callout(doc, [
        ("מה לחפש בקוד אם הבוחן ישאל: ", "bold"),
        ("הקובץ ", "he"),
        ("common/crypto.py", "en"),
        (" — שתי פונקציות, ", "he"),
        ("hash_password", "en"),
        (" ו‑", "he"),
        ("verify_password", "en"),
        (". יש שם רק 14 שורות קוד בסך הכל.", "he"),
    ])

    heading(doc, "6. Threading — איך השרת מטפל בכמה שחקנים בו זמנית", level=2)
    he_paragraph(doc,
                 "Thread זה \"חוט ביצוע\" עצמאי בתוך אותה תוכנית. כשפותחים "
                 "thread חדש, הוא רץ במקביל לקוד הראשי — שניהם חולקים זיכרון "
                 "אבל יכולים לעשות פעולות באותו רגע.")
    he_paragraph(doc,
                 "בפרויקט: השרת פותח thread חדש לכל לקוח שמתחבר. ככה כשלקוח "
                 "אחד עסוק בשליחה של הודעה, ה‑thread שלו לא חוסם את הלקוחות "
                 "האחרים.")
    he_segments(doc, [
        ("הסכנה: ", "bold"),
        ("שני threads יכולים לנסות לערוך את אותה רשימה באותו רגע ולהרוס אותה. "
         "כדי למנוע את זה, הקוד עוטף כל פעולה רגישה ב‑", "he"),
        ("with self._lock:", "en"),
        (" — בלוק שמאפשר רק ל‑", "he"),
        ("thread", "en"),
        (" אחד להיכנס בכל רגע נתון.", "he"),
    ])
    callout(doc, [
        ("איפה זה בקוד: ", "bold"),
        ("בכל אחת מהמחלקות ", "he"),
        ("UserStore", "en"),
        (", ", "he"),
        ("Lobby", "en"),
        (", ", "he"),
        ("GameRegistry", "en"),
        (" יש ", "he"),
        ("self._lock = threading.Lock()", "en"),
        (" ב‑", "he"),
        ("__init__", "en"),
        (".", "he"),
    ])

    heading(doc, "7. AI — איך המחשב בוחר מהלך טוב", level=2)
    he_paragraph(doc,
                 "אלגוריתם השחמט שלנו נקרא \"מינימקס עם גיזום אלפא-בטא\" "
                 "(alpha-beta minimax). הרעיון: לבחור את המהלך שייתן לי "
                 "את המצב הכי טוב, בהנחה שהיריב יבחר אחר כך את המהלך הכי "
                 "טוב בשבילו (וזה הכי גרוע בשבילי).")
    he_paragraph(doc,
                 "מעמיקים לעומק קבוע (במקרה שלנו, ברירת מחדל = 3 מהלכים "
                 "קדימה). בכל עומק מחשבים \"ציון\" של המצב לפי איזון "
                 "החומר (כמה כלים יש לכל צד).")
    he_segments(doc, [
        ("ערכי הכלים: ", "bold"),
        ("רגלי = 100, סוס/רץ = 300, צריח = 500, מלכה = 900, "
         "מלך = אינסוף (כי שחמט נטופל בנפרד עם הציון ", "he"),
        ("CHECKMATE_SCORE", "en"),
        (").", "he"),
    ])
    he_segments(doc, [
        ("גיזום אלפא-בטא: ", "bold"),
        ("טכניקה שמאפשרת לקצר את החיפוש. אם כבר מצאתי מהלך שמבטיח לי "
         "ציון ≥ 5, אין טעם להמשיך לחקור ענפים שהיריב יוכל להגביל ל‑3 — "
         "הוא פשוט יבחר את הענף הזה ואני לעולם לא אגיע. אז אני \"גוזם\" "
         "וחוסך זמן.", "he"),
    ])
    callout(doc, [
        ("הטריק היפה ב‑AI שלנו: ", "bold"),
        ("המחלקה ", "he"),
        ("AISession", "en"),
        (" מחקה את הממשק של ", "he"),
        ("Session", "en"),
        (" אמיתי (אותו ", "he"),
        ("send", "en"),
        (", אותו ", "he"),
        ("account", "en"),
        (", אותו ", "he"),
        ("current_game", "en"),
        ("). לכן ", "he"),
        ("Game", "en"),
        (" אפילו לא יודע שזה לא בן-אדם — וגם לא צריך לדעת.", "he"),
    ])

    heading(doc, "8. \"מכונת מצבים\" של החיבור (Session State Machine)", level=2)
    he_paragraph(doc,
                 "לכל לקוח שמחובר לשרת יש \"מצב\" שאומר מה הוא עושה כרגע. "
                 "ארבעת המצבים האפשריים:")
    for state, desc in [
        ("ANONYMOUS",
         "התחבר לשרת אבל עוד לא הזדהה — יכול רק לרשום משתמש או להיכנס"),
        ("AUTHENTICATED",
         "הזדהה בהצלחה — יכול להיכנס ללובי או להתחיל משחק נגד AI"),
        ("IN_LOBBY",
         "מחכה ליריב — יכול לבטל את ההמתנה"),
        ("IN_GAME",
         "באמצע משחק חי — יכול לשלוח מהלכים או להיכנע"),
    ]:
        he_segments(doc, [
            (state, "en"),
            ("  —  ", "he"),
            (desc, "he"),
        ], bullet=True)
    he_paragraph(doc,
                 "כשמגיעה הודעה מהלקוח, השרת בודק שהפקודה חוקית במצב הנוכחי. "
                 "אם לא — נזרק BadStateError והלקוח מקבל הודעת שגיאה במקום "
                 "ש‑\"משהו יתפוצץ\".",
                 space_after=8)

    heading(doc, "9. בדיקות אוטומטיות עם pytest", level=2)
    he_paragraph(doc,
                 "pytest זאת ספריית בדיקות סטנדרטית של Python. כל קובץ "
                 "שמתחיל ב‑test_ ופונקציה שמתחילה ב‑test_ נחשבים בדיקה. "
                 "אם הפונקציה רצה בלי לזרוק חריגה — הבדיקה עוברת.")
    he_paragraph(doc,
                 "בפרויקט יש 60+ בדיקות, מפוצלות לשני סוגים: בדיקות יחידה "
                 "(בודקות מחלקה אחת) ובדיקות אינטגרציה (מפעילות את השרת "
                 "ומתחברות אליו כלקוח אמיתי דרך TCP).")
    callout(doc, [
        ("שתי הבדיקות הכי חשובות להראות: ", "bold"),
        ("test_parallel_games.py", "en"),
        (" (מוכיחה שהבונוס של ריבוי משחקים עובד) ו‑", "he"),
        ("test_play_ai.py", "en"),
        (" (מוכיחה שה‑AI עובד).", "he"),
    ])



SLIDE_GUIDE = [
    (1, "שער",
     "פתיחה — אומרים בקצרה מה הפרויקט.",
     "\"שלום, אני אציג היום את הפרויקט שלי — Secure Two-Player Chess. "
     "מערכת לקוח-שרת בשפת Python שמאפשרת לשני שחקנים לשחק שחמט מרחוק, "
     "כאשר הסיסמאות שלהם מוגנות בהצפנה. בנוסף יש שתי תוספות בונוס: "
     "השרת תומך בכמה משחקים במקביל, ויש יריב AI שמשתמש באלגוריתם "
     "מינימקס עם גיזום אלפא-בטא.\"",
     None,
     "אין שאלות כאן בדרך כלל."),

    (2, "סדר הצגה לבוחן (אג'נדה)",
     "מראים שיש תוכנית מסודרת. רץ מהר על השקופית.",
     "\"במצגת אעבור על שלוש עשרה נקודות עיקריות — מהמטרה הכללית, דרך "
     "הארכיטקטורה, ההצפנה והפרוטוקול, ועד להדגמה חיה ולשאלות. בכל שלב "
     "אראה גם את הקוד הרלוונטי.\"",
     None,
     "אין."),

    (3, "מה הפרויקט עושה?",
     "מסבירים את התרחיש מקצה לקצה במילים פשוטות.",
     "\"שני שחקנים נכנסים לשרת מרכזי עם שם משתמש וסיסמה. הסיסמאות "
     "נשמרות רק כ-hash מאובטח על הדיסק — לא טקסט גלוי. אחרי הזדהות "
     "הם נכנסים ללובי, השרת מצמיד שניים זמינים ופותח משחק. כל מהלך "
     "עובר דרך השרת שבודק חוקיות ומשדר את המצב לשני השחקנים.\"",
     None,
     "\"איך אתה בודק שמשהו נשמר באמת רק כ-hash?\" — תפתח את "
     "data/users.json ותראה לבוחן."),

    (4, "טבלת כיסוי דרישות המטלה",
     "מראים שכל דרישת המטלה מסומנת. השקופית הזאת היא ה-\"חוזה\".",
     "\"דרשו ממני שתי מחלקות לפחות — יש לי תשע. דרשו לקוח-שרת — יש. "
     "ריבוי לקוחות וAI הם בונוסים שעשיתי. הצפנת הסיסמאות בלבד, לפי הדרישה, "
     "כי דרשו שהתקשורת תהיה ב-clear-text.\"",
     None,
     "\"איפה במטלה כתוב שהתקשורת לא מוצפנת?\" — תגיד שזה כתוב במטלה "
     "המקורית ושאתה מציית בדיוק. ההצפנה היא רק על הסיסמאות."),

    (5, "ארכיטקטורה — מבט על",
     "השקופית הכי חשובה במצגת. כאן מסבירים את המבנה הכללי.",
     "\"שני לקוחות מתחברים לאותו שרת מרכזי דרך TCP. השרת מחזיק שלושה "
     "מאגרים: UserStore לסיסמאות (על הדיסק), Lobby — תור המתנה בזיכרון, "
     "ו-GameRegistry — מילון של כל המשחקים הפעילים. כל חיבור לקוח רץ "
     "ב-thread עצמאי בתוך השרת.\"",
     ["UserStore — נשמר בקובץ users.json",
      "Lobby — בזיכרון, מנוקה כשהשרת נסגר",
      "GameRegistry — מילון Dict[game_id, Game]"],
     "\"למה השרת באמצע ולא peer-to-peer?\" — כדי שאף לקוח לא יוכל לרמות. "
     "השרת הוא source of truth."),

    (6, "מבנה התיקיות",
     "סקירה מהירה של הקבצים.",
     "\"הקוד מחולק לשלוש חבילות: common (לוגיקה טהורה, ללא רשת), "
     "server (שכבת הרשת והניהול), וclient (הממשק עם המשתמש). "
     "ההפרדה הזאת מאפשרת לבדוק את ה-common בלי להפעיל את השרת.\"",
     None,
     "\"למה הפרדת בין common ל-server?\" — מאפשר בדיקות קלות, "
     "שינוי של פרוטוקול לא ישבור את החוקיות, ושימוש חוזר אפשרי "
     "(לדוגמה אם רוצים אחר כך לעשות גם לקוח web)."),

    (7, "9 מחלקות עיקריות",
     "ממלאים את דרישת המטלה \"לפחות 2 מחלקות\" 4.5 פעמים.",
     "\"המטלה דרשה לפחות שתי מחלקות עם עצמים ופעולות לא טריוויאליות. "
     "אצלי יש תשע: Piece, Board, Move, Game — הלוגיקה הטהורה של השחמט. "
     "Account ו-UserStore — אחסון משתמשים. Session, Lobby, GameRegistry — "
     "ניהול השרת. כל אחת מחזיקה state ויש לה פעולות שמשנות אותו.\"",
     None,
     "\"איזו מחלקה הכי מורכבת?\" — Board (כ-315 שורות) כי היא מכילה "
     "את כל חוקי השחמט: תנועות, רוקדה, en-passant, הכתרה, זיהוי שח."),

    (8, "פרוטוקול JSON-Lines",
     "מסבירים את \"חוקי המשחק\" של התקשורת.",
     "\"כל הודעה היא אובייקט JSON אחד בשורה, מסתיים ב-\\n. ככה גם הלקוח "
     "וגם השרת תמיד יודעים איפה הודעה אחת נגמרת והבאה מתחילה. הסיבה "
     "שצריך את זה: TCP הוא זרם בייטים בלי גבולות, אז בלי framing לא היה "
     "אפשר לדעת איפה כל הודעה מסתיימת.\"",
     None,
     "\"למה לא protobuf או XML?\" — JSON זה ספרייה סטנדרטית של Python, "
     "אין תלות חיצונית, וקל לקרוא ידנית לצורכי דיבאג."),

    (9, "קוד הפרוטוקול — protocol.py",
     "מראים את שתי הפונקציות העיקריות: send_message ו-_LineReader.",
     "\"השליחה היא קצרה — קודדים את ה-JSON, מוסיפים \\n ושולחים. הקבלה "
     "יותר מעניינת: _LineReader מחזיק bytearray פנימי. בכל קריאה הוא "
     "ממשיך לקרוא מהסוקט עד שמופיע \\n, ואז חותך משם הודעה אחת ומחזיר.\"",
     ["שורות 48-53: send_message",
      "שורות 56-79: _LineReader עם buffer",
      "שורה 73: chunk ריק = peer ניתק"],
     "\"מה קורה אם מגיע חצי הודעה ב-recv?\" — ה-buffer שלנו ממשיך לקבל "
     "עד שיש לנו \\n. אז אנחנו חוסמים על recv עד שהקטע הבא מגיע."),

    (10, "הצפנת סיסמאות — bcrypt",
     "השקופית הכי חשובה לדרישת ההצפנה.",
     "\"הדרישה הייתה להצפין את הסיסמאות אבל לא את התקשורת. השתמשתי "
     "ב-bcrypt — סטנדרט תעשייתי. שתי פונקציות בלבד: hash_password מקבל "
     "סיסמה גלויה ומחזיר hash מוכן לשמירה. verify_password מקבל סיסמה "
     "ו-hash, מחזיר True/False אם הם תואמים.\"",
     ["bcrypt.gensalt() — salt אקראי בכל פעם",
      "bcrypt.checkpw — השוואה בזמן קבוע (מונע timing attacks)"],
     "\"למה bcrypt ולא MD5 או SHA-256?\" — bcrypt איטי בכוונה, מה "
     "שהופך תקיפת brute-force לבלתי מעשית. בנוסף הוא מכיל salt מובנה."),

    (11, "UserStore — שמירה לדיסק",
     "מראים את לוגיקת הרישום, הוולידציה, והנעילה.",
     "\"כל פעולה שמשנה את ה-store נעולה ב-threading.Lock כי כמה לקוחות "
     "יכולים להירשם בו זמנית. הכתיבה לדיסק היא אטומית — כותבים לקובץ "
     "זמני ואז עושים os.replace, כך אם המחשב נופל באמצע, או שיש את "
     "הקובץ הישן או את החדש, אבל לא חצי.\"",
     ["ולידציית username: 3-32 chars [A-Za-z0-9_-]",
      "ולידציית password: 8 תווים מינימום",
      "_persist_locked: NamedTemporaryFile + os.replace"],
     "\"מה קורה אם שני משתמשים נרשמים בו זמנית עם אותו שם?\" — "
     "הראשון תופס את ה-lock, רושם בהצלחה, מציג DuplicateUserError לשני."),

    (12, "הוכחה שאין סיסמה גלויה בדיסק",
     "השקופית להדגמה מול הבוחן.",
     "\"אחרי שמשתמש נרשם אני יכול לפתוח את data/users.json. כפי שתראו "
     "יש שם רק password_hash בפורמט bcrypt — שמתחיל ב-$2b$12$. הסיסמה "
     "המקורית לעולם לא נשמרת.\"",
     None,
     "\"איך אני יודע ש-$2b$12$XYZ זה באמת hash ולא הסיסמה?\" — אפשר "
     "להריץ Select-String לחפש את הסיסמה המקורית בקובץ — והוא לא ימצא."),

    (13, "השרת — לולאת קבלת חיבורים",
     "מסבירים איך השרת מקבל לקוחות חדשים.",
     "\"ChessServer.start פותח סוקט מאזין על port 5050 ומריץ ברקע "
     "thread שמחכה לחיבורים. כשמגיע חיבור חדש — accept מחזיר סוקט "
     "ייעודי לאותו לקוח, יוצרים אובייקט Session, ומפעילים אותו "
     "ב-thread נפרד. ככה כל לקוח עצמאי לחלוטין.\"",
     ["SO_REUSEADDR — מאפשר reset של השרת בלי לחכות ל-TIME_WAIT",
      "daemon=True — ה-threads מתות אוטומטית כשהתוכנית הראשית נסגרת",
      "max_clients — הגבלה רכה למניעת התקפת DoS"],
     "\"מה קורה אם 50 לקוחות מתחברים בו זמנית?\" — listen(max_clients) "
     "מקבל את הראשונים, השאר מקבלים ServerFullError."),

    (14, "Session — מכונת המצבים",
     "מסבירים את ארבעת המצבים והמעברים ביניהם.",
     "\"כל חיבור מתחיל במצב ANONYMOUS. אחרי register/login עובר ל-"
     "AUTHENTICATED. משם או IN_LOBBY (מחכה ליריב) או ישר IN_GAME "
     "(נגד AI). השרת בודק לפני כל פקודה שהיא חוקית במצב הנוכחי, "
     "ואם לא — זורק BadStateError.\"",
     None,
     "\"למה צריך state machine?\" — כדי למנוע פעולות לא חוקיות, "
     "למשל לקוח לא רשום שמנסה לזוז במשחק שלא קיים."),

    (15, "לולאת ה-dispatch של Session",
     "מסבירים איך הודעות מנותבות ל-handlers.",
     "\"Session.run קורא הודעה אחר הודעה. הוא קורא ל-_dispatch שלוקח "
     "את ה-type ומחפש לפי שם שיטה: handle_register, handle_login, "
     "handle_move. אם יש שגיאת domain — נתפסת אוטומטית ונשלחת ללקוח "
     "כ-error message מסודרת.\"",
     ["getattr(self, f\"handle_{kind}\", None) — דרך אלגנטית להוסיף handlers",
      "SecureChessError — כל השגיאות שלי יורשות מבסיס משותף עם code"],
     "\"מה אם הלקוח שולח type שלא מוכר?\" — handler יוחזר None ונזרק "
     "BadStateError עם הודעה מתאימה."),

    (16, "הלובי",
     "תור FIFO פשוט שמצמיד שני שחקנים.",
     "\"Lobby זאת רשימה פשוטה תחת lock. כשלקוח שולח play_human — "
     "enqueue מוסיף אותו לרשימה. אם יש שניים בתור, יוצא הזוג הראשון "
     "ונפתח להם משחק חדש דרך GameRegistry.\"",
     None,
     "\"איך אתה מתמודד עם שחקן שמתנתק בזמן שהוא ממתין?\" — "
     "_on_disconnect ב-Session קורא ל-Lobby.remove."),

    (17, "GameRegistry — שידור מצב המשחק",
     "מסבירים איך שני השחקנים מקבלים עדכון בו זמנית.",
     "\"אחרי כל מהלך מוצלח, broadcast_state בונה הודעת game_state "
     "אחת ושולח אותה לשני השחקנים. הלוח מועבר בפורמט FEN — הסטנדרט "
     "של עולם השחמט לייצוג מצב לוח.\"",
     None,
     "\"מה זה FEN?\" — Forsyth-Edwards Notation, ייצוג טקסטואלי "
     "קומפקטי של מצב הלוח. הלקוח קורא ומצייר ממנו את הלוח."),

    (18, "זרימת מהלך — מקצה לקצה",
     "השקופית מסבירה את ה-flow המרכזי במשחק.",
     "\"כשהלקוח שולח move, handle_move עושה: 1. בודק שאני באמצע משחק, "
     "2. ממיר את ה-UCI לאובייקט Move, 3. submit_move ב-Game בודק "
     "שזה התור שלי, 4. Board.apply בודק חוקיות מלאה ומבצע, 5. אם "
     "נגמר משחק — שולחים game_ended, אחרת broadcast_state.\"",
     ["Move.parse — מפענח 'e2e4' לאובייקט",
      "Board.apply — המקום היחיד שמשנה לוח",
      "isinstance(next_session, AISession) — מפעיל את ה-AI אוטומטית"],
     "\"איפה החוקיות נאכפת?\" — רק ב-Board.apply. הלקוח לא בודק כלום."),

    (19, "Game.submit_move — חוקיות מהלכים",
     "מציגים את הקוד המרכזי שמטפל בכל המהלכים.",
     "\"submit_move מבצע שלוש בדיקות בסדר: שהמשחק עוד לא נגמר, שזה "
     "התור שלי, שהמהלך חוקי. אם המהלך חוקי הלוח מתעדכן, ואז בודקים "
     "ארבעה תנאי סיום: שחמט, פט, חזרה שלוש פעמים, חוק 50 המהלכים.\"",
     None,
     "\"מה זה threefold repetition?\" — אם אותו מצב לוח חזר 3 פעמים "
     "במשחק — תיקו אוטומטי. נשמר ב-position_history של Board."),

    (20, "AI — alpha-beta minimax",
     "מציגים את האלגוריתם המרכזי של ה-AI.",
     "\"AIPlayer.choose_move מנסה את כל המהלכים החוקיים. לכל מהלך הוא "
     "קורא ל-negamax עם עומק קטן ב-1. negamax זה minimax אבל בצורה "
     "סימטרית — אותה פונקציה מנקודת מבט של הצד הנוכחי. אלפא-בטא חותך "
     "ענפים שלא יכולים לשפר את התוצאה.\"",
     ["_snapshot/_restore — בודקים מהלך ומחזירים את הלוח",
      "depth=3 — מסתכלים 3 מהלכים קדימה",
      "random.choice בין מהלכים שווי-ערך — שלא יהיה צפוי"],
     "\"כמה זמן לוקח חישוב מהלך?\" — תלוי בעומק. בעומק 3, על המחשב שלי "
     "(לפי הקוד) זה כמה עשרות מילישניות במצב פתיחה, יותר במצב סבוך."),

    (21, "פונקציית הערכת מצב",
     "מציגים את הפונקציה שמחזירה ציון למצב לוח.",
     "\"evaluate סופר את החומר על הלוח: חיובי לבן, שלילי שחור. "
     "הערכים בסנטיפיון: רגלי 100, סוס/רץ 300, צריח 500, מלכה 900. "
     "המלך לא נספר כי שחמט מטופל בנפרד עם הציון CHECKMATE_SCORE.\"",
     None,
     "\"אין הערכה של עמדה? מיקום הכלים?\" — הדרישה היא שימוש ב-AI, "
     "לא מנוע ברמת Stockfish. למימוש פשוט material-only מספיק לנצח "
     "שחקן מתחיל."),

    (22, "שילוב ה-AI בפרוטוקול",
     "מסבירים את הקסם של AISession.",
     "\"AISession מחקה את הממשק של Session אמיתי — אותו send, "
     "אותו account, אותו state. לכן Game אפילו לא יודע שזה לא בן-אדם. "
     "כשבא תור ה-AI, registry.drive_ai מחשב מהלך וקורא ל-submit_move "
     "בדיוק כמו שלקוח אמיתי היה עושה.\"",
     None,
     "\"זה לא duplication?\" — להפך, זה duck typing. שתי המחלקות "
     "מקיימות חוזה זהה, כך שאפשר להחליף ביניהן."),

    (23, "בונוס — ריבוי משחקים במקביל",
     "מסבירים איך השרת מטפל בכמה זוגות בו זמנית.",
     "\"כל Session ב-thread משלו, כל GameRegistry פעולה תחת lock. "
     "אז שני זוגות יכולים לשחק ממש באותו רגע — כל אחד עם הלוח שלו, "
     "עדכוני state רק על המשחק שלהם. יש לי בדיקת אינטגרציה "
     "test_parallel_games.py שמוכיחה את זה.\"",
     None,
     "\"מה אם שני messages מגיעים בו זמנית?\" — כל Session "
     "מטפל ב-messages שלו לפי הסדר, כי TCP מבטיח סדר."),

    (24, "ממשק Tkinter GUI",
     "מסבירים את הלקוח הגרפי.",
     "\"Tkinter זה הספרייה הסטנדרטית של Python לממשק גרפי. שלושה "
     "מסכים בתוך אותו חלון: Login, Lobby, Game. הלוח עצמו מצויר על "
     "Canvas — 64 משבצות, כלים בסמלי Unicode. הקליק על משבצת מתורגם "
     "ל-UCI string ונשלח לשרת.\"",
     ["queue.Queue — thread של רשת דוחף, ה-UI שולף",
      "root.after(50, ...) — polling כל 50 מילישניות",
      "שום שינוי UI לא קורה מ-thread של הרשת"],
     "\"למה Tkinter ולא PyQt?\" — Tkinter כלול ב-Python סטנדרטי, "
     "אין תלות חיצונית. גם Pygame אפשרי אבל הגזמה למשחק לוח."),

    (25, "ממשק CLI (טקסט)",
     "מסבירים שיש גם ממשק טקסטואלי.",
     "\"במקביל ל-GUI יש גם ממשק טקסט. שני הממשקים מדברים בדיוק את "
     "אותו פרוטוקול ולכן GUI יכול לשחק נגד CLI. ה-CLI הוא REPL — "
     "thread אחד קורא משרת ומדפיס, השני קורא מ-stdin ושולח.\"",
     None,
     "\"לא יותר טוב רק GUI?\" — ה-CLI שימושי לאוטומציה ולבדיקות "
     "(test_full_game.py מריץ סשנים כאלה)."),

    (26, "היררכיית שגיאות",
     "מציגים את ה-error handling.",
     "\"כל שגיאה בפרויקט יורשת מ-SecureChessError. כל אחת מוגדרת "
     "עם code שמועבר ללקוח. ה-dispatcher תופס SecureChessError אחד "
     "וממיר אוטומטית להודעת error בפרוטוקול. ככה שגיאות מקצוע "
     "(לא חוקי, סיסמה חלשה) הופכות להודעות נקיות במקום stack trace.\"",
     None,
     "\"איך לקוח יודע איזו שגיאה זאת?\" — code הוא string זיהוי "
     "(\"auth_failed\", \"illegal_move\") שהלקוח יכול לבדוק תכנותית."),

    (27, "בדיקות אוטומטיות",
     "סוקרים את חבילת הבדיקות.",
     "\"60+ בדיקות, מחולקות לשתי קטגוריות. בדיקות יחידה (unit) — "
     "כל מחלקה ב-common נבדקת בנפרד. בדיקות אינטגרציה — מפעילות "
     "שרת אמיתי ומתחברות אליו דרך TCP אמיתי. שתי בדיקות חשובות "
     "במיוחד: test_parallel_games (בונוס המקבילות) ו-test_play_ai "
     "(בונוס ה-AI).\"",
     None,
     "\"מה אחוז הכיסוי?\" — תגיד שאתה לא מדדת בדיוק אבל כל המחלקות "
     "ב-common מכוסות ושלוש הזרימות המרכזיות (register, full game, "
     "AI game) נבדקות באינטגרציה."),

    (28, "הדגמה חיה",
     "מתחילים את ההדגמה. תסריט מפורט בחלק ג'.",
     "\"כדי להראות שזה עובד באמת, אבצע הדגמה חיה: שרת + שני לקוחות, "
     "כמה מהלכים, אראה את הקובץ users.json עם ה-hashes, ואז אדגים גם "
     "AI ושני משחקים במקביל.\"",
     None,
     "ראה חלק ג'."),

    (29, "שאלות צפויות + תשובות",
     "השקופית עצמה היא כיסוי מקדים לשאלות צפויות.",
     "\"לפני שנסיים — הנה שש שאלות שכנראה תרצה לשאול, וכל תשובה כתובה. "
     "אעבור עליהן בקצרה לחיסכון בזמן.\"",
     None,
     "ראה חלק ד'."),

    (30, "סיכום",
     "סוגרים את המצגת.",
     "\"לסיכום: ארכיטקטורה נקייה, פרוטוקול JSON-Lines פשוט אבל עמיד, "
     "סיסמאות מוגנות ב-bcrypt, ריבוי משחקים תחת threading, ויריב AI. "
     "הכל מכוסה ב-60+ בדיקות. תודה — אשמח לשאלות.\"",
     None,
     "אין."),
]


def write_slide_guide(doc):
    for n, title, summary, script, key_lines, expected_q in SLIDE_GUIDE:
        heading(doc, f"שקופית {n}: {title}", level=2)

        he_segments(doc, [
            ("רעיון מרכזי: ", "bold"),
            (summary, "he"),
        ])
        he_segments(doc, [
            ("מה לומר (בערך): ", "bold"),
        ])
        he_paragraph(doc, script, italic=True, space_after=6)

        if key_lines:
            he_segments(doc, [
                ("נקודות בקוד שכדאי להצביע עליהן:", "bold"),
            ])
            for line in key_lines:
                he_paragraph(doc, f"• {line}", size=11, space_after=2)

        callout(doc, [
            ("שאלה צפויה: ", "bold"),
            (expected_q, "he"),
        ])



def write_demo_script(doc):
    he_paragraph(doc,
                 "ההדגמה היא הרגע הקריטי במצגת. תכין הכל מראש — אסור "
                 "להתחיל לחפש קבצים מול הבוחן. הנה ההכנה והתסריט:",
                 italic=True)

    heading(doc, "הכנה לפני הבחינה", level=2)
    for step in [
        "פתח 4 חלונות PowerShell בתיקיית secure-chess עם ה-venv מופעל.",
        "ודא ש-port 5050 פנוי (אם תפוס — שנה ל-5051 בכל הפקודות).",
        "ודא ש-data/users.json קיים (אם לא — תיווצר אוטומטית).",
        "פתח גם את הקובץ data/users.json ב-Notepad או VS Code, מוכן להראות.",
        "ודא שגם Tkinter עובד (python -c \"import tkinter; tkinter.Tk()\").",
    ]:
        he_paragraph(doc, f"• {step}", space_after=3)

    heading(doc, "השלב 1: הפעלת השרת (טרמינל 1)", level=2)
    code_block(doc, [
        "cd C:\\Users\\Alonti\\Documents\\GitHub\\secure-chess",
        ".\\.venv\\Scripts\\Activate.ps1",
        "python -m secure_chess.server --host 127.0.0.1 --port 5050 --data-dir ./data",
    ])
    he_paragraph(doc,
                 "תופיע ההודעה \"listening on 127.0.0.1:5050\". להגיד "
                 "לבוחן: \"השרת עלה ומחכה לחיבורים.\"")

    heading(doc, "השלב 2: לקוח ראשון — Alice (טרמינל 2)", level=2)
    code_block(doc, [
        "python -m secure_chess.client --host 127.0.0.1 --port 5050",
    ])
    he_paragraph(doc,
                 "בחלון ה-GUI שנפתח: ללחוץ Register, להכניס שם משתמש "
                 "ייחודי (לדוגמה demo_alice) וסיסמה (לדוגמה chess1234), "
                 "ואז Join Lobby. להגיד: \"היא ממתינה ליריב.\"")

    heading(doc, "השלב 3: לקוח שני — Bob (טרמינל 3)", level=2)
    code_block(doc, [
        "python -m secure_chess.client --host 127.0.0.1 --port 5050",
    ])
    he_paragraph(doc,
                 "באותה צורה: Register עם demo_bob ו-chess5678, אז "
                 "Join Lobby. בו רגע השרת מצמיד את שניהם, נפתח לוח "
                 "אצל שניהם, ו-Alice משחקת לבן.")

    heading(doc, "השלב 4: לבצע מהלכים", level=2)
    he_paragraph(doc,
                 "ב-Alice: ללחוץ על e2 ואז על e4 — המהלך נשלח. "
                 "להראות שה-לוח של Bob התעדכן אוטומטית. ב-Bob: ללחוץ "
                 "e7 ואז e5. אצל שניהם הלוח התעדכן.")
    tip_box(doc, [
        ("הדגמה ", "bold"),
        ("ויזואלית", "italic"),
        (" של עדכון בזמן אמת היא ה-WOW של המצגת. תוודא שיש זמן לזה.", "he"),
    ])

    heading(doc, "השלב 5: הוכחת ההצפנה", level=2)
    he_paragraph(doc,
                 "להעביר את החלון ל-Notepad עם data/users.json. להראות "
                 "שתי הרשומות (demo_alice, demo_bob) מכילות רק password_hash "
                 "שמתחיל ב-$2b$12$ — לא הסיסמה המקורית. כאן נמצא המקום "
                 "הכי טוב לאמירה: \"כפי שאתם רואים, גם אם הקובץ הזה ידלוף, "
                 "לאף אחד אין סיסמה אמיתית.\"")

    heading(doc, "השלב 6: בונוס AI (טרמינל 4)", level=2)
    code_block(doc, [
        "python -m secure_chess.client --host 127.0.0.1 --port 5050 --cli",
        "> register demo_carol pass1234",
        "ok",
        "> play_ai",
        "== game started ==",
        "you are: white   opponent: AI(depth=3)",
        "> move e2e4",
    ])
    he_paragraph(doc,
                 "תוך כמה מילישניות יחזור מהלך של ה-AI. להסביר: "
                 "\"כפי שאתם רואים, האלגוריתם בחר תגובה. הוא חיפש "
                 "עד עומק 3, כלומר שקל את שלושת המהלכים הבאים.\"")

    heading(doc, "השלב 7: בונוס ריבוי משחקים", level=2)
    he_paragraph(doc,
                 "המשחקים של Alice-Bob ושל Carol-AI שניהם רצים בו זמנית "
                 "על אותו שרת. להראות זאת על-ידי ביצוע מהלך בכל אחד "
                 "מהמשחקים בזה אחר זה — שניהם ממשיכים בלי לחסום.")

    heading(doc, "השלב 8: בדיקות אוטומטיות", level=2)
    he_paragraph(doc,
                 "לסגור את שאר החלונות (כדי שה-port יתפנה), ובטרמינל "
                 "חדש להריץ:")
    code_block(doc, ["pytest -v"])
    he_paragraph(doc,
                 "אחרי כ-30-60 שניות יתקבל \"60+ passed\" בירוק. "
                 "להגיד: \"אלה כל הבדיקות האוטומטיות. כל אחת בודקת "
                 "תרחיש אחר. הירוק מוכיח שהקוד עובד כמצופה.\"")



QA = [
    ("למה לא הצפנת גם את התקשורת בנוסף לסיסמאות?",
     "ההצפנה של התקשורת לא נדרשה במטלה — דווקא ההפך, המטלה הגדירה "
     "שהפרוטוקול יהיה plain-text. כדי להוסיף הצפנת תקשורת היה אפשר "
     "להחליף את socket.socket ב-ssl.wrap_socket עם תעודה.\n"
     "במציאות בייצור — כן, חובה. כאן לצורך התרגיל הקובץ users.json "
     "הוא הנכס היחיד שצריך להגן עליו, וההצפנה שלו בלבד מספיקה."),

    ("מה ההבדל בין hash ל-encryption?",
     "Encryption הוא דו-כיווני — אפשר להצפין ולפענח חזרה. Hash הוא "
     "חד-כיווני — קל לחשב מהמקור, בלתי אפשרי לחשב את המקור מהתוצאה.\n"
     "סיסמאות צריכות hash (לא encryption) כי גם השרת לא צריך לדעת "
     "את הסיסמה — מספיק לו לדעת אם הסיסמה שהוגשה מייצרת את אותו hash."),

    ("מה זה salt ולמה הוא חשוב?",
     "salt זה מחרוזת אקראית שנוסיפה לסיסמה לפני ה-hash. בלי salt, שני "
     "משתמשים שבחרו אותה סיסמה יקבלו אותו hash, ו-rainbow tables "
     "(טבלאות hash מחושבות מראש) היו פותחות הכל. עם salt — לכל משתמש "
     "hash אחר גם אם הסיסמה זהה."),

    ("מה קורה אם לקוח מתנתק באמצע משחק?",
     "Session.run מזהה ConnectionClosed ויוצא מהלולאה. _on_disconnect "
     "בודק את ה-state: אם היה IN_GAME, קוראים ל-Game.handle_disconnect "
     "שמסמן ניצחון לצד השני (reason='disconnect'). אם היה IN_LOBBY — "
     "מורידים מהתור."),

    ("איך אתה מוודא ששני threads לא מתנגשים?",
     "כל מבנה נתונים שניגשים אליו מ-threads שונים עטוף ב-threading.Lock "
     "ופעולות עליו תחת with self._lock:. זה כולל UserStore._accounts, "
     "Lobby._queue, GameRegistry._games. בנוסף, כל Session ב-thread "
     "משלו, אז הם לא נוגעים אחד בשני."),

    ("איפה החוקיות של מהלכי השחמט נאכפת?",
     "רק במחלקה Board, ספציפית בפונקציה apply. הלקוח לא בודק כלום — "
     "הוא רק שולח את הקלט של המשתמש (\"e2e4\") לשרת. הפרדה הזאת קריטית: "
     "לקוח זדוני יוכל להריץ Move שלא חוקי, אבל השרת יזרוק IllegalMoveError."),

    ("איך ה-AI עובד ולמה הוא משחק מהר?",
     "אלגוריתם minimax — מנסה את כל המהלכים שלי, ועבור כל אחד מניח שהיריב "
     "יבחר את התשובה הכי טובה בשבילו. הולכים לעומק 3 כברירת מחדל. "
     "Alpha-beta pruning חותך ענפים שלא יכולים להשפיע על התוצאה הסופית "
     "— זה מצמצם את החיפוש פי 10-100."),

    ("למה בחרת JSON-Lines ולא XML או Protocol Buffers?",
     "JSON-Lines קל לקרוא ידנית בדיבאג, מובנה ב-Python (json module "
     "סטנדרטי), והקבצים קטנים יחסית. XML יותר verbose ומיותר. "
     "Protobuf דורש קומפילציה מ-.proto, ובמטלה לא נדרש ביצוע מקסימלי."),

    ("איך אתה ימנע ש-2 שחקנים יקבלו את אותו ID משחק?",
     "Game.id הוא uuid.uuid4().hex — מזהה אקראי של 128 ביט. "
     "מתמטית הסיכוי להתנגשות אפסי."),

    ("מה אם השרת נופל באמצע כתיבת users.json?",
     "אני כותב לקובץ זמני (.users.XXX.tmp) ואז עושה os.replace שזה "
     "פעולה אטומית במערכת הקבצים. אז או שיש את הקובץ הישן או את החדש, "
     "אבל לעולם לא חצי."),

    ("מה זה FEN ולמה השתמשת בו?",
     "Forsyth-Edwards Notation — סטנדרט עולמי לייצוג טקסטואלי של מצב "
     "לוח שחמט. דוגמה: 'rnbqkbnr/pppppppp/8/8/8/8/PPPPPPPP/RNBQKBNR w' "
     "אומר לוח מצב התחלתי, לבן לזוז. השתמשתי כי זה compact, תקני, "
     "ויש מימושים מוכנים."),

    ("מה זה UCI?",
     "Universal Chess Interface — תקן של ייצוג מהלכים: 'e2e4' זה רגלי "
     "מ-e2 ל-e4. 'e7e8q' זה הכתרה למלכה. השתמשתי כי קל לפענח (4-5 תווים) "
     "וזה הסטנדרט בעולם השחמט."),

    ("מה היה הקושי הכי גדול בפרויקט?",
     "(תשובה אישית): \"אכיפת חוקי השחמט עצמם בקובץ board.py הייתה הכי "
     "מאתגרת — לטפל בכל המקרים המיוחדים: רוקדה, en-passant, הכתרה, "
     "סינון מהלכים שמשאירים את המלך בשח. זה גם הקובץ הארוך ביותר.\""),

    ("מה היית עושה אחרת אם היית מתחיל מההתחלה?",
     "(תשובה אישית): \"אולי הייתי משתמש ב-asyncio במקום threads — "
     "זה ייתן ביצועים טובים יותר ופחות סיבוכיות סביב locks. אבל "
     "threading היה פשוט יותר להבנה ראשונית.\""),

    ("איך תרחיב את הפרויקט בעתיד?",
     "מספר אפשרויות: (1) WebSocket במקום TCP גולמי + לקוח web — "
     "ככה אפשר לשחק מהדפדפן. (2) שמירת היסטוריית משחקים ב-SQLite. "
     "(3) דירוג ELO אוטומטי. (4) AI חזק יותר — להוסיף הערכת מיקום, "
     "Quiescence search, transposition table."),

    ("איך זה משתלב עם מערכת ההפעלה?",
     "Python 3.11+ עובד בלי שינוי בכל ה-OS (Windows, Linux, macOS). "
     "תלות יחידה היא bcrypt — מותקנת אוטומטית עם pip install -e. "
     "אין שום קוד OS-specific."),

    ("יש לך בדיקות יחידה לקלאסות?",
     "כן — tests/unit/. כל מחלקה ב-common נבדקת: test_pieces.py, "
     "test_board_rules.py, test_crypto.py, test_user_store.py, "
     "test_ai.py. בנוסף יש tests/integration/ עם שרת אמיתי על TCP."),

    ("איך אתה מטפל ב-edge cases?",
     "כמה דוגמאות: (1) סיסמה ריקה → WeakPasswordError לפני ה-hash. "
     "(2) JSON לא תקין → ProtocolError → error message. "
     "(3) move לא חוקי → IllegalMoveError. (4) שחקן זר שמנסה לזוז "
     "במשחק לא שלו → NotYourTurnError."),

    ("שפת התכנות שבחרת מתאימה לפרויקט הזה?",
     "Python מתאים מאוד: יש לו socket מובנה, json מובנה, threading "
     "מובנה, ו-Tkinter מובנה. כל מה שצריך לפרויקט הזה — בלי תלויות "
     "חיצוניות חוץ מ-bcrypt. שפה כמו C++ הייתה מהירה יותר אבל גם "
     "ארוכה יותר ב-3-4 פעמים."),

    ("איך אני בוחר את הסיסמה הראשונה כשאני נרשם?",
     "בלקוח: ב-GUI יש שדה password עם show='*' (לא מציג את הסיסמה). "
     "ב-CLI: המשתמש מקליד אותה בפקודה. בשני המקרים הסיסמה מועברת "
     "ל-handle_register בשרת שמייצר את ה-hash ושומר."),
]


def write_qa(doc):
    for i, (q, a) in enumerate(QA, start=1):
        he_paragraph(doc, f"שאלה {i}: {q}", bold=True, size=13,
                     color=ACCENT, space_before=8, space_after=2)
        for paragraph in a.split("\n"):
            he_paragraph(doc, paragraph, size=12, space_after=4)



def write_files_map(doc):
    he_paragraph(doc,
                 "אם הבוחן יבקש \"תראה לי איפה X בקוד\" — הנה רשימה של "
                 "כל הקבצים החשובים. אל תפתח שום קובץ בלי לדעת בערך מה "
                 "אומר בו:")

    files = [
        ("common/crypto.py",
         "הצפנת סיסמאות — שתי פונקציות בלבד, ~25 שורות."),
        ("common/user_store.py",
         "אחסון משתמשים על הדיסק — Account, UserStore, וולידציות."),
        ("common/protocol.py",
         "JSON-Lines framing — send_message ו-_LineReader."),
        ("common/board.py",
         "הלוח ומחלקות חוקיות השחמט — הקובץ הארוך ביותר (~315 שורות)."),
        ("common/move.py",
         "Move + פירוק UCI."),
        ("common/pieces.py",
         "Piece, PieceType, Color, ותנועות פסאודו לכל סוג כלי."),
        ("common/game.py",
         "Game + Result — מנהל משחק חי בין שני שחקנים."),
        ("common/ai.py",
         "AIPlayer — alpha-beta minimax (בונוס)."),
        ("common/errors.py",
         "היררכיית שגיאות — SecureChessError ויורשיו."),
        ("server/server.py",
         "ChessServer — accept loop ולולאת הקבלה."),
        ("server/session.py",
         "Session, SessionState, AISession — מכונת המצבים והניהול."),
        ("server/lobby.py",
         "Lobby ו-GameRegistry — תור המתנה ומאגר משחקים."),
        ("client/gui.py",
         "ChessGui — Tkinter, שלושה מסכים, polling של queue."),
        ("client/cli.py",
         "ChessCli — REPL טקסטואלי."),
        ("tests/unit/",
         "11 קבצי בדיקות יחידה — אחד למחלקה."),
        ("tests/integration/",
         "5 קבצי בדיקות אינטגרציה — מפעילות שרת אמיתי."),
        ("data/users.json",
         "ה-store של הסיסמאות — bcrypt hashes בלבד."),
        ("pyproject.toml",
         "הגדרת החבילה ותלויות (bcrypt, pytest)."),
    ]
    table = doc.add_table(rows=len(files) + 1, cols=2)
    table.style = "Light Grid Accent 1"
    head_cells = table.rows[0].cells
    head_cells[0].text = ""
    head_cells[1].text = ""
    _set_cell_shading(head_cells[0], rgb_hex(TABLE_HEAD_BG))
    _set_cell_shading(head_cells[1], rgb_hex(TABLE_HEAD_BG))
    for cell, label in zip(head_cells, ["קובץ", "מה יש בו"]):
        p = cell.paragraphs[0]
        _set_bidi(p)
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(label)
        _set_run_rtl(run)
        run.font.name = HE_FONT
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = ACCENT

    for i, (path, desc) in enumerate(files, start=1):
        row = table.rows[i].cells
        row[0].text = ""
        p = row[0].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(path)
        run.font.name = CODE_FONT
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(0x03, 0x2F, 0x62)

        row[1].text = ""
        p2 = row[1].paragraphs[0]
        _set_bidi(p2)
        p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run2 = p2.add_run(desc)
        _set_run_rtl(run2)
        run2.font.name = HE_FONT
        run2.font.size = Pt(11)



GLOSSARY = [
    ("TCP",
     "פרוטוקול תקשורת ברמה נמוכה שמבטיח זרם בייטים אמין ובסדר בין שני "
     "מחשבים. בנוי על שכבת IP."),
    ("Socket",
     "ה-API ברמת מערכת ההפעלה לתקשורת רשת. סוקט הוא נקודת קצה — או "
     "מאזין (server) או מחובר (client)."),
    ("Port",
     "מספר (1-65535) שמזהה אפליקציה ספציפית על מחשב. השרת שלנו "
     "תופס port 5050."),
    ("Thread",
     "חוט ביצוע עצמאי בתוך תהליך. שני threads רצים במקביל, חולקים זיכרון."),
    ("Lock / Mutex",
     "מנגנון סנכרון שמאפשר רק ל-thread אחד להיכנס לבלוק קוד מסוים בזמן נתון. "
     "מונע race conditions."),
    ("Race Condition",
     "מצב שגוי שקורה כששני threads ניגשים לאותו משאב בלי סנכרון, "
     "והתוצאה תלויה בסדר ביצוע אקראי."),
    ("Hash",
     "פונקציה חד-כיוונית מקלט גדול לפלט בגודל קבוע. קל לחשב, "
     "בלתי אפשרי להפוך."),
    ("bcrypt",
     "אלגוריתם hash מותאם לסיסמאות. איטי בכוונה ומכיל salt מובנה."),
    ("Salt",
     "מחרוזת אקראית שמתווספת לסיסמה לפני ה-hash, כדי שאותה סיסמה "
     "לא תיתן את אותו hash אצל משתמשים שונים."),
    ("Rainbow Table",
     "טבלת hashes מחושבים מראש שמשמשת תוקפים לפענח hashes ידועים. "
     "salt מנטרל את ההתקפה הזאת."),
    ("Brute-Force",
     "התקפה שמנסה את כל הסיסמאות האפשריות עד למציאת ההתאמה. "
     "bcrypt איטי במיוחד כדי שזה ייקח אינסוף זמן."),
    ("JSON",
     "JavaScript Object Notation — פורמט טקסטואלי לתיאור מבני נתונים. "
     "נקרא על-ידי אדם, ספרייה סטנדרטית בכל שפה."),
    ("JSON-Lines",
     "פורמט שבו כל שורה היא אובייקט JSON עצמאי, מופרדים ב-\\n. "
     "נוח לזרמים (streaming)."),
    ("Framing",
     "השכבה שמפרידה הודעות בתוך זרם רציף של בייטים. אצלנו = "
     "תו \\n בסוף כל הודעה."),
    ("UCI",
     "Universal Chess Interface — תקן לייצוג מהלכים כ-4-5 תווים, "
     "למשל e2e4 או e7e8q (הכתרה)."),
    ("FEN",
     "Forsyth-Edwards Notation — מחרוזת קומפקטית שמתארת מצב לוח שחמט "
     "מלא: מיקום הכלים, תור מי, זכות רוקדה, en-passant."),
    ("Castling / רוקדה",
     "מהלך מיוחד שמזיז גם מלך וגם צריח. דורש שהמלך לא זז עוד, "
     "הצריח לא זז, ואין כלים בדרך."),
    ("En-passant",
     "מהלך מיוחד שבו רגלי לוכד רגלי יריב שזז שני צעדים. תקף רק "
     "במהלך הבא אחרי המהלך הכפול."),
    ("Checkmate / שחמט",
     "המלך תחת איום ואין מהלך חוקי שיציל אותו. סוף המשחק — מנצח הצד "
     "שהאיים."),
    ("Stalemate / פט",
     "תור הצד שלי, אבל אין לי שום מהלך חוקי, והמלך שלי לא בשח. "
     "תיקו אוטומטי."),
    ("Minimax",
     "אלגוריתם לבחירת מהלך טוב במשחקי שני שחקנים. אני מנסה למקסם, "
     "היריב מנסה למזער."),
    ("Alpha-Beta Pruning",
     "אופטימיזציה של minimax — חיתוך ענפים שכבר ידוע שלא יכולים "
     "להשפיע על התוצאה. חוסך זמן חישוב."),
    ("Centipawn / סנטיפיון",
     "יחידת מידה של ציון בשחמט = 1/100 מערך של רגלי. השתמשתי "
     "בערכים: רגלי 100, סוס/רץ 300, צריח 500, מלכה 900."),
    ("Duck Typing",
     "עיקרון תכנותי: אם משהו נראה כמו ברווז ועושה קולות כמו ברווז "
     "— נתייחס אליו כמו ברווז. אצלנו: AISession מתנהג כמו Session "
     "וגם זה מספיק."),
    ("Daemon Thread",
     "thread שמת אוטומטית כשהתוכנית הראשית נסגרת. בשרת שלנו כל ה-"
     "session threads הם daemon."),
    ("State Machine",
     "מודל שאומר שלאובייקט יש state, ולפי ה-state הנוכחי נכנסות "
     "פעולות חוקיות שונות. אצלנו: Session עם 4 states."),
    ("Atomic Operation",
     "פעולה שלא ניתן להפסיק באמצע. os.replace היא אטומית — או "
     "הקובץ הישן או החדש, לא חצי."),
    ("Pytest",
     "ספריית בדיקות סטנדרטית של Python. הריץ אותה ב-pytest -v "
     "כדי להריץ את כל הבדיקות."),
    ("Integration Test",
     "בדיקה שמפעילה כמה רכיבים יחד (אצלנו: שרת + לקוח על TCP אמיתי), "
     "בניגוד ל-unit test שבודק רכיב אחד מבודד."),
]


def write_glossary(doc):
    he_paragraph(doc,
                 "המילון מסודר לפי הופעה במצגת, לא לפי א-ב. אם הבוחן "
                 "ישאל \"מה זה X?\" — תמצא X כאן ותענה במילים שלך, לא "
                 "במילים של המילון.", italic=True)

    for term, defn in GLOSSARY:
        he_segments(doc, [
            (term, "accent"),
            ("  —  ", "he"),
            (defn, "he"),
        ], space_after=3)



def write_cheat_sheet(doc):
    he_paragraph(doc,
                 "דף אחד שמרכז את הכל. הדפס וקח איתך לבחינה.",
                 italic=True, color=MUTED)

    heading(doc, "תקציר ב-30 שניות", level=2)
    he_paragraph(doc,
                 "\"Secure Two-Player Chess — מערכת לקוח-שרת בפייתון. "
                 "שני שחקנים נרשמים בשרת מרכזי, הסיסמאות שלהם נשמרות "
                 "כ-bcrypt hashes, והם משחקים שחמט מלא דרך פרוטוקול "
                 "JSON-Lines מעל TCP. בונוסים: ריבוי משחקים מקבילי "
                 "ו-AI מבוסס alpha-beta minimax.\"")

    heading(doc, "המספרים שצריך לזכור", level=2)
    for label, val in [
        ("מספר מחלקות בפרויקט", "9 (Piece, Board, Move, Game, Account, "
                                "UserStore, Session, Lobby, GameRegistry)"),
        ("מספר בדיקות אוטומטיות", "60+ ב-pytest"),
        ("פורט ברירת מחדל", "5050"),
        ("עומק חיפוש AI", "3 מהלכים קדימה"),
        ("אורך מינימלי לסיסמה", "8 תווים"),
        ("מספר תווים מקסימלי בשם משתמש", "32"),
        ("גודל buffer ברשת", "4096 בייט בכל recv"),
    ]:
        he_segments(doc, [
            (f"• {label}: ", "bold"),
            (val, "he"),
        ], space_after=2)

    heading(doc, "פקודות שכדאי לדעת בעל-פה", level=2)
    code_block(doc, [
        "# הפעלת שרת",
        "python -m secure_chess.server --host 127.0.0.1 --port 5050 --data-dir ./data",
        "",
        "# הפעלת לקוח GUI",
        "python -m secure_chess.client --host 127.0.0.1 --port 5050",
        "",
        "# הפעלת לקוח CLI",
        "python -m secure_chess.client --host 127.0.0.1 --port 5050 --cli",
        "",
        "# הוכחת הצפנה",
        "Get-Content .\\data\\users.json",
        "",
        "# הרצת כל הבדיקות",
        "pytest -v",
    ])

    heading(doc, "מילים-מפתח לתוך כל תשובה", level=2)
    for word in [
        "\"השרת הוא ה-source of truth\"",
        "\"thread per connection\"",
        "\"thread-safe תחת lock\"",
        "\"bcrypt עם salt מובנה\"",
        "\"JSON-Lines נותן framing\"",
        "\"alpha-beta pruning חוסך זמן\"",
        "\"duck typing\" (כשמדברים על AISession)",
        "\"השרת לא שומר סיסמאות בטקסט גלוי\"",
        "\"אין הצפנת תקשורת — לפי דרישת המטלה\"",
    ]:
        he_paragraph(doc, f"• {word}", size=11, space_after=2)

    heading(doc, "אם נתקעת — שלוש תשובות בטחון", level=2)
    he_paragraph(doc,
                 "1. \"אני לא בטוח, אבל לפי הקוד נראה לי ש‑X. תן לי להראות לך.\"")
    he_paragraph(doc,
                 "2. \"זה תרחיש שלא בדקתי, אבל הייתי בודק אותו על-ידי X.\"")
    he_paragraph(doc,
                 "3. \"שאלה טובה — בנוי כך כי X, אבל בהחלט אפשר היה גם Y.\"")



if __name__ == "__main__":
    here = Path(__file__).parent
    build_document(here / "secure-chess-speaker-guide.docx")
