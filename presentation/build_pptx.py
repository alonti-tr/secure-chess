from __future__ import annotations

import copy
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt
from lxml import etree


SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

MARGIN_X = Inches(0.5)
MARGIN_TOP = Inches(0.35)
CONTENT_TOP = Inches(1.15)

HEBREW_FONT = "Arial"
CODE_FONT = "Consolas"

COLOR_DARK = RGBColor(0x1F, 0x23, 0x28)
COLOR_ACCENT = RGBColor(0x1F, 0x6F, 0xEB)
COLOR_ACCENT_SOFT = RGBColor(0xEA, 0xF2, 0xFD)
COLOR_MUTED = RGBColor(0x57, 0x60, 0x6A)
COLOR_CODE_BG = RGBColor(0x27, 0x28, 0x22)
COLOR_CODE_FG = RGBColor(0xF8, 0xF8, 0xF2)
COLOR_CODE_KEYWORD = RGBColor(0xF9, 0x26, 0x72)
COLOR_CODE_STRING = RGBColor(0xE6, 0xDB, 0x74)
COLOR_CODE_COMMENT = RGBColor(0x75, 0x71, 0x5E)
COLOR_CALLOUT_BG = RGBColor(0xEA, 0xF2, 0xFD)
COLOR_TIP_BG = RGBColor(0xFF, 0xF8, 0xE1)
COLOR_TIP_BORDER = RGBColor(0xF5, 0xA6, 0x23)
COLOR_PILL_BG = RGBColor(0xE7, 0xF5, 0xEC)
COLOR_PILL_FG = RGBColor(0x1F, 0x7A, 0x3E)
COLOR_TABLE_HEAD = RGBColor(0xEA, 0xF2, 0xFD)


def set_rtl(paragraph) -> None:
    pPr = paragraph._pPr
    if pPr is None:
        pPr = paragraph._p.get_or_add_pPr()
    pPr.set("rtl", "1")


def set_ltr(paragraph) -> None:
    pPr = paragraph._pPr
    if pPr is None:
        pPr = paragraph._p.get_or_add_pPr()
    pPr.set("rtl", "0")


def add_textbox(slide, left, top, width, height, *, fill=None, line=None):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.04)
    tf.margin_bottom = Inches(0.04)
    if fill is not None:
        box.fill.solid()
        box.fill.fore_color.rgb = fill
    else:
        box.fill.background()
    if line is None:
        box.line.fill.background()
    return box


def add_paragraph(text_frame, *, first=False):
    if first and text_frame.paragraphs and not text_frame.paragraphs[0].runs:
        return text_frame.paragraphs[0]
    return text_frame.add_paragraph()


XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def add_run(paragraph, text: str, *, font=HEBREW_FONT, size=18, bold=False,
            color=COLOR_DARK, italic=False):
    run = paragraph.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    t_elem = run._r.find(qn("a:t"))
    if t_elem is not None:
        t_elem.set(XML_SPACE, "preserve")
    return run


def add_title(slide, text: str, *, color=COLOR_ACCENT, size=30):
    title_box = add_textbox(slide, MARGIN_X, MARGIN_TOP,
                            SLIDE_W - 2 * MARGIN_X, Inches(0.7))
    title_box.text_frame.margin_left = Inches(0.05)
    p = title_box.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    set_rtl(p)
    add_run(p, text, font=HEBREW_FONT, size=size, bold=True, color=color)
    return title_box


NBSP = "\u00A0"
EM_SPACE = "\u2003"


def numbered_segment(n: int):
    return (f"{n}.{EM_SPACE}", "num")


def _normalize_segments(segments):
    normalized = []
    for seg in segments:
        if isinstance(seg, str):
            text, kind = seg, "he"
        else:
            text, kind = seg
        normalized.append([text, kind])

    for i, (text, kind) in enumerate(normalized):
        prev_kind = normalized[i - 1][1] if i > 0 else None
        next_kind = normalized[i + 1][1] if i + 1 < len(normalized) else None
        if prev_kind is not None and prev_kind != kind and text.startswith(" "):
            text = NBSP + text.lstrip(" ")
        if next_kind is not None and next_kind != kind and text.endswith(" "):
            text = text.rstrip(" ") + NBSP
        normalized[i][0] = text
    return [tuple(x) for x in normalized]


def add_he_paragraph(text_frame, segments, *, size=16, bold=False,
                     bullet=False, alignment=PP_ALIGN.RIGHT, first=False,
                     space_after=4):
    p = add_paragraph(text_frame, first=first)
    p.alignment = alignment
    set_rtl(p)
    p.space_after = Pt(space_after)
    if bullet:
        _set_bullet(p)
    for text, kind in _normalize_segments(segments):
        if kind == "en":
            run = add_run(p, text, font=CODE_FONT, size=size - 1,
                          color=RGBColor(0x03, 0x2F, 0x62), bold=False)
            run.font.highlight = None
        elif kind == "bold":
            add_run(p, text, font=HEBREW_FONT, size=size, bold=True)
        elif kind == "muted":
            add_run(p, text, font=HEBREW_FONT, size=size,
                    color=COLOR_MUTED)
        elif kind == "accent":
            add_run(p, text, font=HEBREW_FONT, size=size, bold=True,
                    color=COLOR_ACCENT)
        elif kind == "num":
            add_run(p, text, font=HEBREW_FONT, size=size, bold=bold)
        else:
            add_run(p, text, font=HEBREW_FONT, size=size, bold=bold)
    return p


def _set_bullet(paragraph) -> None:
    pPr = paragraph._p.get_or_add_pPr()
    for child in pPr.findall(qn("a:buChar")):
        pPr.remove(child)
    for child in pPr.findall(qn("a:buNone")):
        pPr.remove(child)
    buChar = etree.SubElement(pPr, qn("a:buChar"))
    buChar.set("char", "•")
    pPr.set("indent", str(Emu(-Inches(0.22))))
    pPr.set("marL", str(Emu(Inches(0.30))))


def add_code_block(slide, left, top, width, height, lines,
                   *, font_size=11):
    bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top,
                                width, height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = COLOR_CODE_BG
    bg.line.fill.background()
    bg.shadow.inherit = False
    bg.adjustments[0] = 0.04
    bg.text_frame.text = ""
    bg.text_frame.margin_left = Inches(0.18)
    bg.text_frame.margin_right = Inches(0.15)
    bg.text_frame.margin_top = Inches(0.12)
    bg.text_frame.margin_bottom = Inches(0.12)

    tf = bg.text_frame
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        set_ltr(p)
        p.space_after = Pt(0)
        _highlight_line(p, line, font_size)
    return bg


PY_KEYWORDS = {
    "def", "class", "if", "elif", "else", "for", "while", "try", "except",
    "finally", "raise", "return", "with", "as", "import", "from", "in",
    "is", "not", "and", "or", "lambda", "yield", "self", "None", "True",
    "False", "pass", "break", "continue", "assert", "global",
}


def _highlight_line(paragraph, text: str, font_size: int) -> None:
    if not text:
        run = paragraph.add_run()
        run.text = " "
        run.font.name = CODE_FONT
        run.font.size = Pt(font_size)
        run.font.color.rgb = COLOR_CODE_FG
        return

    leading = len(text) - len(text.lstrip(" "))
    if leading:
        _add_code_run(paragraph, " " * leading, font_size, COLOR_CODE_FG)
        text = text[leading:]

    comment_pos = _find_comment(text)
    if comment_pos == 0:
        _add_code_run(paragraph, text, font_size, COLOR_CODE_COMMENT)
        return

    code_part = text if comment_pos < 0 else text[:comment_pos]
    comment_part = "" if comment_pos < 0 else text[comment_pos:]

    _emit_tokens(paragraph, code_part, font_size)

    if comment_part:
        _add_code_run(paragraph, comment_part, font_size, COLOR_CODE_COMMENT)


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


def _emit_tokens(paragraph, code: str, font_size: int) -> None:
    i = 0
    buf = ""

    def flush_buf():
        nonlocal buf
        if buf:
            _add_code_run(paragraph, buf, font_size, COLOR_CODE_FG)
            buf = ""

    while i < len(code):
        ch = code[i]
        if ch in ('"', "'"):
            flush_buf()
            quote = ch
            j = i + 1
            while j < len(code) and code[j] != quote:
                if code[j] == "\\":
                    j += 2
                    continue
                j += 1
            j = min(j + 1, len(code))
            _add_code_run(paragraph, code[i:j], font_size, COLOR_CODE_STRING)
            i = j
            continue
        if ch.isalpha() or ch == "_":
            j = i
            while j < len(code) and (code[j].isalnum() or code[j] == "_"):
                j += 1
            word = code[i:j]
            if word in PY_KEYWORDS:
                flush_buf()
                _add_code_run(paragraph, word, font_size, COLOR_CODE_KEYWORD)
            else:
                buf += word
            i = j
            continue
        buf += ch
        i += 1
    flush_buf()


def _add_code_run(paragraph, text: str, font_size: int, color: RGBColor):
    run = paragraph.add_run()
    run.text = text
    run.font.name = CODE_FONT
    run.font.size = Pt(font_size)
    run.font.color.rgb = color
    t_elem = run._r.find(qn("a:t"))
    if t_elem is not None:
        t_elem.set(XML_SPACE, "preserve")


def add_callout(slide, left, top, width, height, segments,
                *, bg=COLOR_CALLOUT_BG, border=COLOR_ACCENT,
                font_size=15):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top,
                                   width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = bg
    shape.line.color.rgb = border
    shape.line.width = Pt(0)
    shape.adjustments[0] = 0.10
    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.18)
    tf.margin_right = Inches(0.18)
    tf.margin_top = Inches(0.10)
    tf.margin_bottom = Inches(0.10)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    first = True
    if isinstance(segments[0], list):
        for paragraph_segs in segments:
            add_he_paragraph(tf, paragraph_segs, size=font_size, first=first,
                             space_after=2)
            first = False
    else:
        add_he_paragraph(tf, segments, size=font_size, first=True)
    return shape


def add_speaker_tip(slide, left, top, width, height, body):
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top,
                                 width, height)
    box.fill.solid()
    box.fill.fore_color.rgb = COLOR_TIP_BG
    box.line.color.rgb = COLOR_TIP_BORDER
    box.line.width = Pt(0)
    box.adjustments[0] = 0.10
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.18)
    tf.margin_right = Inches(0.18)
    tf.margin_top = Inches(0.08)
    tf.margin_bottom = Inches(0.08)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    segments = [("טיפ למצגת: ", "tip-label")]
    if isinstance(body, str):
        segments.append((body, "he"))
    else:
        segments.extend(body)

    p = add_paragraph(tf, first=True)
    p.alignment = PP_ALIGN.RIGHT
    set_rtl(p)
    for text, kind in _normalize_segments(segments):
        if kind == "tip-label":
            add_run(p, text, font=HEBREW_FONT, size=13, bold=True,
                    color=RGBColor(0xB2, 0x5A, 0x00))
        elif kind == "en":
            add_run(p, text, font=CODE_FONT, size=12,
                    color=RGBColor(0x03, 0x2F, 0x62))
        elif kind == "bold":
            add_run(p, text, font=HEBREW_FONT, size=13, bold=True)
        else:
            add_run(p, text, font=HEBREW_FONT, size=13)
    return box


def add_slide_number(slide, index: int, total: int) -> None:
    box = add_textbox(slide, SLIDE_W - Inches(1.2), SLIDE_H - Inches(0.4),
                      Inches(1.0), Inches(0.3))
    p = box.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    set_ltr(p)
    add_run(p, f"{index} / {total}", font=CODE_FONT, size=10,
            color=COLOR_MUTED)


def blank_slide(prs) -> "Slide":
    return prs.slides.add_slide(prs.slide_layouts[6])


def build_presentation(out_path: Path) -> None:
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    builders = [
        slide_01_title, slide_02_agenda, slide_03_what, slide_04_coverage,
        slide_05_arch, slide_06_layout, slide_07_classes, slide_08_protocol,
        slide_09_protocol_code, slide_10_bcrypt, slide_11_user_store,
        slide_12_proof, slide_13_server_accept, slide_14_state_machine,
        slide_15_dispatch, slide_16_lobby, slide_17_registry, slide_18_move_flow,
        slide_19_rules, slide_20_ai_alphabeta, slide_21_eval, slide_22_ai_integ,
        slide_23_parallel, slide_24_gui, slide_25_cli, slide_26_errors,
        slide_27_tests, slide_28_demo, slide_29_qa, slide_30_closing,
    ]
    total = len(builders)
    for i, builder in enumerate(builders, start=1):
        slide = blank_slide(prs)
        builder(slide)
        add_slide_number(slide, i, total)

    prs.save(str(out_path))
    print(f"wrote {out_path}  ({total} slides)")



def _center_title(slide, text: str, size: int = 44):
    box = add_textbox(slide, MARGIN_X, Inches(1.2),
                      SLIDE_W - 2 * MARGIN_X, Inches(1.0))
    p = box.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    add_run(p, text, font=HEBREW_FONT, size=size, bold=True)
    return box


def slide_01_title(slide):
    _center_title(slide, "Secure Two-Player Chess", size=44)

    sub = add_textbox(slide, MARGIN_X, Inches(2.2),
                      SLIDE_W - 2 * MARGIN_X, Inches(0.7))
    p = sub.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    set_rtl(p)
    add_run(p, "משחק שחמט מאובטח בין שני שחקנים — מערכת לקוח/שרת ב‑",
            font=HEBREW_FONT, size=22, color=COLOR_MUTED)
    add_run(p, "Python", font=CODE_FONT, size=20, color=COLOR_MUTED)

    badge_w = Inches(4.5)
    badge = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        (SLIDE_W - badge_w) / 2, Inches(3.2),
        badge_w, Inches(0.55))
    badge.fill.solid()
    badge.fill.fore_color.rgb = COLOR_ACCENT_SOFT
    badge.line.fill.background()
    badge.adjustments[0] = 0.5
    bp = badge.text_frame.paragraphs[0]
    bp.alignment = PP_ALIGN.CENTER
    set_rtl(bp)
    add_run(bp, "מדריך לתלמיד — מצגת לבוחן",
            font=HEBREW_FONT, size=18, bold=True, color=COLOR_ACCENT)

    desc = add_textbox(slide, Inches(1.5), Inches(4.4),
                       SLIDE_W - Inches(3.0), Inches(1.5))
    add_he_paragraph(
        desc.text_frame,
        [
            ("הגרסה כוללת: אימות עם סיסמאות מוצפנות ", "he"),
            ("(bcrypt)", "en"),
            (", פרוטוקול ", "he"),
            ("JSON-Lines", "en"),
            (" מעל ", "he"),
            ("TCP", "en"),
            (", ריבוי משחקים במקביל, ויריב ", "he"),
            ("AI (alpha-beta minimax)", "en"),
            (".", "he"),
        ],
        size=15, alignment=PP_ALIGN.CENTER, first=True,
    )


def slide_02_agenda(slide):
    add_title(slide, "סדר הצגה לבוחן")
    box = add_textbox(slide, MARGIN_X, CONTENT_TOP,
                      SLIDE_W - 2 * MARGIN_X, Inches(5.6))
    items = [
        [("מה הפרויקט עושה — סקירה ב‑30 שניות", "he")],
        [("איך הוא עומד בדרישות המטלה (טבלת כיסוי)", "he")],
        [("ארכיטקטורה — לקוח, שרת, וזרימה כוללת", "he")],
        [("מחלקות ועצמים (≥ 2 — יש 9!)", "he")],
        [("פרוטוקול התקשורת מעל ה‑", "he"), ("TCP", "en")],
        [("הצפנת סיסמאות — ", "he"), ("bcrypt", "en"), (" בלבד", "he")],
        [("זרימת התחברות והרשמה (", "he"), ("register", "en"),
         (" / ", "he"), ("login", "en"), (")", "he")],
        [("לובי, התאמה (", "he"), ("matchmaking", "en"),
         (") ומשחק חי", "he")],
        [("בונוס: ריבוי משחקים במקביל (", "he"), ("threading", "en"),
         (")", "he")],
        [("בונוס: ", "he"), ("AI", "en"), (" בשיטת ", "he"),
         ("alpha-beta minimax", "en")],
        [("ממשק המשתמש: ", "he"), ("Tkinter GUI", "en"), (" + ", "he"),
         ("CLI", "en")],
        [("בדיקות אוטומטיות — ", "he"), ("pytest", "en")],
        [("הדגמה חיה", "he")],
    ]
    first = True
    for segs in items:
        add_he_paragraph(box.text_frame, segs, size=17, bullet=True,
                         first=first, space_after=4)
        first = False

    add_speaker_tip(slide, MARGIN_X, SLIDE_H - Inches(0.95),
                    SLIDE_W - 2 * MARGIN_X, Inches(0.5),
                    "להגיד \"עכשיו אעבור איתך שלב אחרי שלב — בכל שלב אראה גם את הקוד\".")


def slide_03_what(slide):
    add_title(slide, "מה הפרויקט עושה?")
    box = add_textbox(slide, MARGIN_X, CONTENT_TOP,
                      SLIDE_W - 2 * MARGIN_X, Inches(5.5))
    bullets = [
        [("שני משתמשים נרשמים בשרת מרכזי עם ", "he"),
         ("שם משתמש + סיסמה", "bold"), (".", "he")],
        [("הסיסמאות נשמרות על הדיסק ", "he"),
         ("בלבד", "bold"), (" כ‑", "he"), ("hash", "en"),
         (" מאובטח (", "he"), ("bcrypt", "en"),
         (") — לא טקסט גלוי.", "he")],
        [("לאחר התחברות הם נכנסים ל‑", "he"),
         ("לובי", "bold"), ("; השרת מצמיד שני שחקנים זמינים ופותח משחק שחמט.", "he")],
        [("המסירות עוברות במסר ", "he"), ("JSON", "en"),
         (" אחד בכל שורה מעל חיבור ", "he"), ("TCP", "en"), (".", "he")],
        [("השרת מנהל את הלוח, בודק חוקיות מהלכים, מזהה שחמט / פט / תיקו ושולח עדכוני מצב.", "he")],
        [("בונוס: אפשר לשחק נגד ", "he"), ("AI", "en"),
         (" מקומי במקום נגד יריב אנושי.", "he")],
        [("בונוס: השרת תומך בכמה זוגות שחקנים ", "he"),
         ("במקביל", "bold"), (".", "he")],
    ]
    first = True
    for segs in bullets:
        add_he_paragraph(box.text_frame, segs, size=17, bullet=True,
                         first=first, space_after=5)
        first = False


def _table_row(table, row_idx, left_text_segs, right_text):
    cell_left = table.cell(row_idx, 0)
    cell_left.text = ""
    p = cell_left.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    set_ltr(p)
    add_run(p, left_text_segs, font=CODE_FONT, size=12,
            color=RGBColor(0x03, 0x2F, 0x62))

    cell_right = table.cell(row_idx, 1)
    cell_right.text = ""
    p2 = cell_right.text_frame.paragraphs[0]
    p2.alignment = PP_ALIGN.RIGHT
    set_rtl(p2)
    add_run(p2, right_text, font=HEBREW_FONT, size=13)


def slide_04_coverage(slide):
    add_title(slide, "כיסוי דרישות המטלה")
    rows = [
        ("דרישה", "איפה בקוד"),
        ("שתי מחלקות ומעלה עם עצמים ופעולות",
         "Piece · Board · Move · Game · Account · UserStore · Session · Lobby · GameRegistry"),
        ("מערכת לקוח‑שרת",
         "src/secure_chess/server/ + src/secure_chess/client/"),
        ("ריבוי לקוחות במקביל  (בונוס)",
         "thread per connection + thread-safe Lobby/GameRegistry"),
        ("שימוש ב‑AI  (בונוס)",
         "common/ai.py — alpha-beta minimax"),
        ("הצפנת סיסמאות בלבד (התקשורת לא מוצפנת)",
         "common/crypto.py (bcrypt) → data/users.json"),
        ("ממשק משתמש",
         "client/gui.py (Tkinter) + client/cli.py"),
    ]
    table_shape = slide.shapes.add_table(
        rows=len(rows), cols=2,
        left=MARGIN_X, top=CONTENT_TOP,
        width=SLIDE_W - 2 * MARGIN_X, height=Inches(5.0))
    table = table_shape.table
    table.columns[0].width = Inches(7.5)
    table.columns[1].width = Inches(SLIDE_W.inches - 2 * MARGIN_X.inches - 7.5)

    for col, header in enumerate(["איפה בקוד", "דרישה"]):
        cell = table.cell(0, col)
        cell.text = ""
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_TABLE_HEAD
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.RIGHT
        set_rtl(p)
        add_run(p, header, font=HEBREW_FONT, size=14, bold=True,
                color=COLOR_ACCENT)

    for r, (he, en) in enumerate(rows[1:], start=1):
        _table_row(table, r, en, he)

    add_speaker_tip(
        slide, MARGIN_X, SLIDE_H - Inches(0.75),
        SLIDE_W - 2 * MARGIN_X, Inches(0.45),
        [("לפתוח לבוחן את ", "he"),
         ("README.md", "en"),
         (" ולהראות בדיוק את הטבלה הזו ב‑", "he"),
         ("repo", "en"),
         (".", "he")],
    )


def slide_05_arch(slide):
    add_title(slide, "ארכיטקטורה — מבט על")
    diagram = [
        "   +-----------------+         TCP / JSON-Lines         +-----------------+",
        "   |   Client A      | <------------------------------> |                 |",
        "   |  (GUI or CLI)   |                                  |                 |",
        "   +-----------------+                                  |                 |",
        "                                                        |   ChessServer   |",
        "   +-----------------+         TCP / JSON-Lines         |  (accept loop)  |",
        "   |   Client B      | <------------------------------> |                 |",
        "   |  (GUI or CLI)   |                                  |                 |",
        "   +-----------------+                                  +--------+--------+",
        "                                                                 |",
        "                                                  +--------------+--------------+",
        "                                                  |              |              |",
        "                                              UserStore       Lobby        GameRegistry",
        "                                            (users.json,    (queue of    (active Games,",
        "                                              bcrypt)       waiting     one per match)",
        "                                                           sessions)",
    ]
    add_code_block(slide, MARGIN_X, CONTENT_TOP,
                   SLIDE_W - 2 * MARGIN_X, Inches(3.6),
                   diagram, font_size=10)

    bullets_box = add_textbox(slide, MARGIN_X, Inches(5.0),
                              SLIDE_W - 2 * MARGIN_X, Inches(2.0))
    bullets = [
        [("כל חיבור = ", "he"), ("Session", "en"), (" אחד בתוך ", "he"),
         ("thread", "en"), (" דדיקטיבי.", "he")],
        [("שלוש \"מאגרים\" משותפים: משתמשים (דיסק), לובי (תור בזיכרון), משחקים פעילים.", "he")],
        [("השרת לעולם ", "he"), ("לא", "bold"),
         (" שולח את הסיסמאות חזרה ללקוח, ולא שומר אותן בטקסט.", "he")],
    ]
    first = True
    for segs in bullets:
        add_he_paragraph(bullets_box.text_frame, segs, size=14, bullet=True,
                         first=first)
        first = False


def slide_06_layout(slide):
    add_title(slide, "מבנה התיקיות")
    lines = [
        "secure-chess/",
        "|-- pyproject.toml          # dependencies (bcrypt, pytest)",
        "|-- README.md",
        "|-- src/secure_chess/",
        "|   |-- common/             # pure logic - no network, no I/O",
        "|   |   |-- pieces.py       # Color, PieceType, Piece + pseudo-legal moves",
        "|   |   |-- board.py        # Board + full legality, mate / stale / draws",
        "|   |   |-- move.py         # Square + Move + UCI parsing",
        "|   |   |-- game.py         # Game + Result",
        "|   |   |-- crypto.py       # bcrypt wrappers",
        "|   |   |-- user_store.py   # Account + JSON-backed UserStore",
        "|   |   |-- protocol.py     # JSON-Lines framing",
        "|   |   |-- ai.py           # alpha-beta minimax            BONUS",
        "|   |   `-- errors.py       # exception hierarchy",
        "|   |-- server/             # network layer",
        "|   |   |-- server.py       # ChessServer + accept loop",
        "|   |   |-- session.py      # Session, SessionState, AISession",
        "|   |   `-- lobby.py        # Lobby + GameRegistry",
        "|   `-- client/",
        "|       |-- gui.py          # Tkinter GUI front-end",
        "|       `-- cli.py          # text-mode REPL",
        "`-- tests/",
        "    |-- unit/               # unit tests",
        "    `-- integration/        # end-to-end tests over real TCP",
    ]
    add_code_block(slide, MARGIN_X, CONTENT_TOP,
                   SLIDE_W - 2 * MARGIN_X, Inches(5.4),
                   lines, font_size=11)

    add_speaker_tip(
        slide, MARGIN_X, SLIDE_H - Inches(0.75),
        SLIDE_W - 2 * MARGIN_X, Inches(0.45),
        [("להגיד: \"הקפדתי על הפרדה — ", "he"),
         ("common", "en"),
         (" זה לוגיקה בלי ", "he"),
         ("I/O", "en"),
         (", אז קל לבדוק אותה\".", "he")],
    )


def slide_07_classes(slide):
    add_title(slide, "9 מחלקות עיקריות (הדרישה: ≥ 2)")
    rows = [
        ("Piece", "כלי שחמט (צבע + סוג) + הולכת תנועות פסאודו-חוקיות"),
        ("Board", "הלוח 8x8 + חוקיות מלאה + זיהוי שחמט/פט/תיקו"),
        ("Move", "מהלך (ממ-עד-הכתרה) + פירוק UCI"),
        ("Game", "משחק חי בין שני שחקנים — מחזיק לוח + תוצאה"),
        ("Account", "שורת משתמש: שם + password_hash + תאריך"),
        ("UserStore", "אחסון משתמשים ב‑JSON עם נעילת thread"),
        ("Session", "חיבור לקוח אחד בשרת + מכונת מצבים"),
        ("Lobby", "תור המתנה — מצמיד שני שחקנים"),
        ("GameRegistry", "מאגר משחקים פעילים + שידור מצב"),
    ]
    table_shape = slide.shapes.add_table(
        rows=len(rows) + 1, cols=2,
        left=MARGIN_X, top=CONTENT_TOP,
        width=SLIDE_W - 2 * MARGIN_X, height=Inches(5.5))
    table = table_shape.table
    table.columns[0].width = Inches(3.0)
    table.columns[1].width = Inches(SLIDE_W.inches - 2 * MARGIN_X.inches - 3.0)
    for col, header in enumerate(["תפקיד", "מחלקה"]):
        cell = table.cell(0, col)
        cell.text = ""
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_TABLE_HEAD
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.RIGHT
        set_rtl(p)
        add_run(p, header, font=HEBREW_FONT, size=14, bold=True,
                color=COLOR_ACCENT)
    for r, (cls, role) in enumerate(rows, start=1):
        _table_row(table, r, cls, role)


def slide_08_protocol(slide):
    add_title(slide, "פרוטוקול ה‑JSON-Lines")
    intro = add_textbox(slide, MARGIN_X, CONTENT_TOP,
                        SLIDE_W - 2 * MARGIN_X, Inches(0.55))
    add_he_paragraph(
        intro.text_frame,
        [("אובייקט ", "he"), ("JSON", "en"),
         (" אחד בכל שורה, מסתיים ב‑", "he"), ("\\n", "en"),
         (", מעל ", "he"), ("TCP", "en"),
         (" רגיל (ללא הצפנה — לפי דרישת המטלה).", "he")],
        size=15, first=True,
    )

    lines = [
        'C -> S  {"type":"register","username":"alice","password":"hunter2!"}',
        'S -> C  {"type":"ok"}',
        'C -> S  {"type":"play_human"}',
        'S -> C  {"type":"ok"}',
        'S -> C  {"type":"game_started","you_are":"white","opponent":"bob", ...}',
        'C -> S  {"type":"move","uci":"e2e4"}',
        'S -> C  {"type":"ok"}',
        'S -> C  {"type":"game_state", ...}',
        '...',
        'S -> C  {"type":"game_ended","result":"white_wins","reason":"checkmate"}',
    ]
    add_code_block(slide, MARGIN_X, Inches(2.0),
                   SLIDE_W - 2 * MARGIN_X, Inches(3.0),
                   lines, font_size=12)

    add_callout(
        slide, MARGIN_X, Inches(5.3),
        SLIDE_W - 2 * MARGIN_X, Inches(1.5),
        [
            [("למה ", "he"), ("JSON-Lines", "en"), ("?", "bold"),
             (" כי ", "he"), ("TCP", "en"),
             (" הוא זרם בייטים בלי גבולות הודעה.", "he")],
            [("סיום שורה ב‑", "he"), ("\\n", "en"),
             (" נותן לנו ", "he"),
             ("framing", "en"),
             (" פשוט שאפשר לקרוא ולפענח בלי שום ספרייה מיוחדת.", "he")],
        ],
    )


def slide_09_protocol_code(slide):
    add_title(slide, "קוד הפרוטוקול — protocol.py")
    lines = [
        "def send_message(sock: socket.socket, msg: Dict[str, Any]) -> None:",
        "    line = (json.dumps(msg, separators=(\",\", \":\")) + \"\\n\").encode(\"utf-8\")",
        "    try:",
        "        sock.sendall(line)",
        "    except OSError as exc:",
        "        raise ConnectionClosed(str(exc)) from exc",
        "",
        "",
        "class _LineReader:",
        "    \"\"\"Reads from the socket and buffers until a \\n is seen.",
        "    Yields exactly one JSON line per read_line() call.\"\"\"",
        "    def __init__(self, sock: socket.socket) -> None:",
        "        self._sock = sock",
        "        self._buf = bytearray()",
        "",
        "    def read_line(self) -> bytes:",
        "        while b\"\\n\" not in self._buf:",
        "            chunk = self._sock.recv(4096)",
        "            if not chunk:",
        "                raise ConnectionClosed(\"peer closed the connection\")",
        "            self._buf.extend(chunk)",
        "        idx = self._buf.index(b\"\\n\")",
        "        line = bytes(self._buf[:idx])",
        "        del self._buf[: idx + 1]",
        "        return line",
    ]
    add_code_block(slide, MARGIN_X, CONTENT_TOP,
                   SLIDE_W - 2 * MARGIN_X, Inches(4.8),
                   lines, font_size=11)

    bottom = add_textbox(slide, MARGIN_X, Inches(6.05),
                         SLIDE_W - 2 * MARGIN_X, Inches(1.0))
    add_he_paragraph(
        bottom.text_frame,
        [("מה להגיד לבוחן: \"השרת והלקוח שניהם משתמשים ב‑", "he"),
         ("_LineReader", "en"),
         (", אז אותה לוגיקת ", "he"),
         ("framing", "en"),
         (" רצה בשני הצדדים — אין סיכוי לתקרים של חצי-הודעה\".", "he")],
        size=13, first=True,
    )


def slide_10_bcrypt(slide):
    add_title(slide, "הצפנת סיסמאות — bcrypt")
    intro = add_textbox(slide, MARGIN_X, CONTENT_TOP,
                        SLIDE_W - 2 * MARGIN_X, Inches(0.55))
    add_he_paragraph(
        intro.text_frame,
        [("הדרישה הייתה: ", "he"),
         ("רק הסיסמאות מוצפנות", "bold"),
         (", לא התקשורת. השתמשתי ב‑", "he"),
         ("bcrypt", "en"),
         (" — סטנדרט תעשייתי שמכיל ", "he"),
         ("salt", "en"),
         (" אוטומטי בתוך ה‑", "he"),
         ("hash", "en"),
         (".", "he")],
        size=15, first=True,
    )

    lines = [
        "# common/crypto.py",
        "import bcrypt",
        "",
        "",
        "def hash_password(plain: str) -> str:",
        "    if not isinstance(plain, str):",
        "        raise TypeError(\"password must be a string\")",
        "    return bcrypt.hashpw(plain.encode(\"utf-8\"), bcrypt.gensalt()).decode(\"ascii\")",
        "",
        "",
        "def verify_password(plain: str, hashed: str) -> bool:",
        "    if not isinstance(plain, str) or not isinstance(hashed, str):",
        "        return False",
        "    try:",
        "        return bcrypt.checkpw(plain.encode(\"utf-8\"), hashed.encode(\"ascii\"))",
        "    except (ValueError, TypeError):",
        "        return False",
    ]
    add_code_block(slide, MARGIN_X, Inches(2.0),
                   SLIDE_W - 2 * MARGIN_X, Inches(3.6),
                   lines, font_size=11)

    add_callout(
        slide, MARGIN_X, Inches(5.85),
        SLIDE_W - 2 * MARGIN_X, Inches(1.15),
        [
            [("למה ", "he"), ("bcrypt", "en"), (" ולא ", "he"),
             ("SHA-256", "en"), ("?", "bold"),
             (" כי ", "he"), ("bcrypt", "en"),
             (" איטי בכוונה (", "he"), ("cost factor", "en"),
             (") — תקיפת ", "he"), ("brute-force", "en"),
             (" הופכת לבלתי מעשית, גם אם הקובץ דלף.", "he")],
        ],
        font_size=13,
    )


def slide_11_user_store(slide):
    add_title(slide, "שמירה לדיסק — UserStore")
    lines = [
        "class UserStore:",
        "    def __init__(self, path: str | Path) -> None:",
        "        self.path = Path(path)",
        "        self._lock = threading.Lock()",
        "        self._accounts: Dict[str, Account] = {}",
        "        self._load()",
        "",
        "    def register(self, username: str, password: str) -> Account:",
        "        _validate_username(username)        # 3-32 chars [A-Za-z0-9_-]",
        "        _validate_password(password)        # min 8 characters",
        "        hashed = hash_password(password)    # bcrypt one-way hash",
        "        with self._lock:",
        "            if username in self._accounts:",
        "                raise DuplicateUserError(...)",
        "            account = Account(username, hashed, _now_utc())",
        "            self._accounts[username] = account",
        "            self._persist_locked()          # write-tmp + os.replace = atomic",
        "            return account",
        "",
        "    def authenticate(self, username, password) -> Optional[Account]:",
        "        with self._lock:",
        "            account = self._accounts.get(username)",
        "        if account is None:",
        "            return None",
        "        return account if verify_password(password, account.password_hash) else None",
    ]
    add_code_block(slide, MARGIN_X, CONTENT_TOP,
                   SLIDE_W - 2 * MARGIN_X, Inches(5.4),
                   lines, font_size=11)

    bottom = add_textbox(slide, MARGIN_X, SLIDE_H - Inches(0.7),
                         SLIDE_W - 2 * MARGIN_X, Inches(0.4))
    add_he_paragraph(
        bottom.text_frame,
        [("שלוש נקודות חשובות: ולידציה, ", "he"),
         ("thread-safe lock", "en"),
         (", וכתיבה אטומית לקובץ.", "he")],
        size=13, first=True,
    )


def slide_12_proof(slide):
    add_title(slide, "הוכחה שאין סיסמה גלויה בדיסק")
    intro = add_textbox(slide, MARGIN_X, CONTENT_TOP,
                        SLIDE_W - 2 * MARGIN_X, Inches(0.5))
    add_he_paragraph(
        intro.text_frame,
        [("אחרי שמשתמש נרשם — אפשר לפתוח את ", "he"),
         ("data/users.json", "en"),
         (" ולוודא שאין שם טקסט גלוי:", "he")],
        size=15, first=True,
    )

    json_lines = [
        "{",
        "  \"alice\": {",
        "    \"password_hash\": \"$2b$12$KIXr5w9.2T0e1m9k3rN6qO...\",",
        "    \"created_at\": \"2026-05-14T13:01:22Z\"",
        "  }",
        "}",
    ]
    add_code_block(slide, MARGIN_X, Inches(1.85),
                   SLIDE_W - 2 * MARGIN_X, Inches(2.1),
                   json_lines, font_size=12)

    ps_lines = [
        "# shows only hashes - no plaintext passwords",
        "Get-Content .\\data\\users.json",
        "",
        "# searching for the raw password should return nothing",
        "Select-String -Path .\\data\\users.json -Pattern \"hunter2!|s3cretpw\"",
        "# expected: no matches",
    ]
    add_code_block(slide, MARGIN_X, Inches(4.15),
                   SLIDE_W - 2 * MARGIN_X, Inches(2.1),
                   ps_lines, font_size=12)

    add_speaker_tip(slide, MARGIN_X, SLIDE_H - Inches(0.75),
                    SLIDE_W - 2 * MARGIN_X, Inches(0.45),
                    "להריץ את הפקודות האלה מול הבוחן — הוכחה ויזואלית מנצחת.")


def slide_13_server_accept(slide):
    add_title(slide, "השרת — לולאת קבלת חיבורים")
    lines = [
        "class ChessServer:",
        "    def start(self) -> tuple[str, int]:",
        "        self._sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)",
        "        self._sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)",
        "        self._sock.bind((self.host, self.port))",
        "        self._sock.listen(self.max_clients)",
        "        # background accept thread so start() returns immediately",
        "        self._accept_thread = threading.Thread(target=self._accept_loop, daemon=True)",
        "        self._accept_thread.start()",
        "        return self.bound_address",
        "",
        "    def _accept_loop(self) -> None:",
        "        while not self._stop_event.is_set():",
        "            client_sock, peer = self._sock.accept()",
        "            # one Session per client, each on its own thread",
        "            session = Session(client_sock, peer, self.user_store,",
        "                              self.lobby, self.registry, ...)",
        "            with self._sessions_lock:",
        "                self._sessions.append(session)",
        "            threading.Thread(target=session.run, daemon=True,",
        "                             name=f\"sess-{session.id[:6]}\").start()",
    ]
    add_code_block(slide, MARGIN_X, CONTENT_TOP,
                   SLIDE_W - 2 * MARGIN_X, Inches(4.6),
                   lines, font_size=11)

    add_callout(
        slide, MARGIN_X, Inches(5.95),
        SLIDE_W - 2 * MARGIN_X, Inches(1.15),
        [
            [("הבונוס של ריבוי לקוחות:", "bold"),
             (" כל חיבור הוא ", "he"), ("thread", "en"),
             (" נפרד. המאגרים המשותפים ", "he"),
             ("(UserStore, Lobby, GameRegistry)", "en"),
             (" כולם מוגנים ב‑", "he"),
             ("threading.Lock", "en"), (".", "he")],
        ],
        font_size=13,
    )


def slide_14_state_machine(slide):
    add_title(slide, "Session — מכונת מצבים")
    lines = [
        "   ┌──────────────┐  register/login   ┌────────────────┐",
        "   │  ANONYMOUS   │  ───────────────► │ AUTHENTICATED  │",
        "   └──────────────┘                   └───────┬────────┘",
        "                                              │ play_human",
        "                                              ▼",
        "                                       ┌────────────┐",
        "                                       │  IN_LOBBY  │ ── 2 שחקנים",
        "                                       └─────┬──────┘    ─────────┐",
        "                                             │ cancel_lobby      │",
        "                                             ▼                   ▼",
        "                                      AUTHENTICATED           IN_GAME",
        "                                                          ┌──────────┐",
        "                                                          │ submit   │",
        "                                                          │ moves... │",
        "                                                          └─────┬────┘",
        "                                                                │ checkmate /",
        "                                                                │ resign / quit",
        "                                                                ▼",
        "                                                         AUTHENTICATED",
    ]
    add_code_block(slide, MARGIN_X, CONTENT_TOP,
                   SLIDE_W - 2 * MARGIN_X, Inches(5.0),
                   lines, font_size=10)

    bottom = add_textbox(slide, MARGIN_X, Inches(6.3),
                         SLIDE_W - 2 * MARGIN_X, Inches(0.7))
    add_he_paragraph(
        bottom.text_frame,
        [("לכל הודעה נכנסת השרת בודק ", "he"),
         ("גם סוג", "bold"),
         (" וגם ", "he"), ("מצב", "bold"),
         (" — אם לא תקין, נזרק ", "he"),
         ("BadStateError", "en"),
         (" והלקוח מקבל הודעת ", "he"),
         ("error", "en"), (" מסודרת.", "he")],
        size=13, first=True,
    )


def slide_15_dispatch(slide):
    add_title(slide, "לולאת השליחה של Session")
    lines = [
        "def run(self) -> None:",
        "    log.info(\"[%s] connected from %s:%d\", self.id[:8], *self.peer)",
        "    try:",
        "        while not self._closed:",
        "            try:",
        "                msg = protocol.recv_message(self._reader)",
        "            except ConnectionClosed:",
        "                break",
        "            except ProtocolError as exc:",
        "                self.send(protocol.error(exc.code, str(exc)))",
        "                continue",
        "            try:",
        "                self._dispatch(msg)",
        "            except SecureChessError as exc:",
        "                # any domain error -> structured wire-protocol 'error' message",
        "                self.send(protocol.error(exc.code, str(exc)))",
        "    finally:",
        "        self._on_disconnect()",
        "        self.close()",
        "",
        "def _dispatch(self, msg):",
        "    kind = msg[\"type\"]                        # 'register' / 'login' / 'move' ...",
        "    handler = getattr(self, f\"handle_{kind}\", None)",
        "    if handler is None:",
        "        raise BadStateError(f\"command {kind!r} not legal\")",
        "    handler(msg)",
    ]
    add_code_block(slide, MARGIN_X, CONTENT_TOP,
                   SLIDE_W - 2 * MARGIN_X, Inches(5.4),
                   lines, font_size=11)

    bottom = add_textbox(slide, MARGIN_X, SLIDE_H - Inches(0.7),
                         SLIDE_W - 2 * MARGIN_X, Inches(0.4))
    add_he_paragraph(
        bottom.text_frame,
        [("דפוס: ", "he"),
         ("הודעה → handler → ok או error", "bold"),
         (". אחיד, נקי, בדיק.", "he")],
        size=13, first=True,
    )


def slide_16_lobby(slide):
    add_title(slide, "הלובי — איך מצמידים שני שחקנים")
    lines = [
        "class Lobby:",
        "    def __init__(self) -> None:",
        "        self._queue: List[\"Session\"] = []",
        "        self._lock = threading.Lock()",
        "",
        "    def enqueue(self, session: \"Session\") -> Optional[Tuple[\"Session\", \"Session\"]]:",
        "        with self._lock:",
        "            self._queue.append(session)",
        "            if len(self._queue) >= 2:",
        "                white = self._queue.pop(0)   # FIFO - longest-waiting first",
        "                black = self._queue.pop(0)",
        "                return white, black",
        "        return None",
    ]
    add_code_block(slide, MARGIN_X, CONTENT_TOP,
                   SLIDE_W - 2 * MARGIN_X, Inches(3.2),
                   lines, font_size=12)

    box = add_textbox(slide, MARGIN_X, Inches(4.55),
                      SLIDE_W - 2 * MARGIN_X, Inches(2.5))
    bullets = [
        [("תור פשוט בזיכרון, מוגן ב‑", "he"), ("Lock", "en"),
         (" כי הוא נגיש מ‑", "he"), ("threads", "en"),
         (" מרובים.", "he")],
        [("שני שחקנים = השרת קורא ל‑", "he"),
         ("GameRegistry.create(white, black)", "en"),
         (" שמשגר ", "he"), ("game_started", "en"),
         (" לשניהם.", "he")],
        [("שחקן שמבטל המתנה (", "he"), ("cancel_lobby", "en"),
         (") — מוסרים אותו דרך ", "he"),
         ("Lobby.remove", "en"), (".", "he")],
    ]
    first = True
    for segs in bullets:
        add_he_paragraph(box.text_frame, segs, size=14, bullet=True,
                         first=first)
        first = False


def slide_17_registry(slide):
    add_title(slide, "GameRegistry — שידור מצב המשחק")
    lines = [
        "def broadcast_state(self, game: Game, last_move: str) -> None:",
        "    msg = {",
        "        \"type\": \"game_state\",",
        "        \"game_id\": game.id,",
        "        \"last_move\": last_move,",
        "        \"board\": game.board.to_fen_short(),",
        "        \"to_move\": \"white\" if game.board.side_to_move is Color.WHITE else \"black\",",
        "        \"in_check\": game.board.is_in_check(game.board.side_to_move),",
        "    }",
        "    game.white.send(msg)",
        "    game.black.send(msg)",
    ]
    add_code_block(slide, MARGIN_X, CONTENT_TOP,
                   SLIDE_W - 2 * MARGIN_X, Inches(2.9),
                   lines, font_size=12)

    intro = add_textbox(slide, MARGIN_X, Inches(4.25),
                        SLIDE_W - 2 * MARGIN_X, Inches(0.6))
    add_he_paragraph(
        intro.text_frame,
        [("אחרי כל מהלך מוצלח, ", "he"), ("שני", "bold"),
         (" השחקנים מקבלים אותה הודעת ", "he"),
         ("game_state", "en"), (" עם מצב הלוח ב‑", "he"),
         ("FEN", "en"), (".", "he")],
        size=15, first=True,
    )

    add_callout(
        slide, MARGIN_X, Inches(5.0),
        SLIDE_W - 2 * MARGIN_X, Inches(1.6),
        [
            [("למה ", "he"), ("FEN", "en"), ("?", "bold"),
             (" כי זה ייצוג סטנדרטי, קומפקטי, של מצב הלוח —", "he")],
            [("הלקוח ", "he"), ("(GUI / CLI)", "en"),
             (" מצייר ממנו לוח גרפי או טקסטואלי.", "he")],
        ],
    )


def slide_18_move_flow(slide):
    add_title(slide, "זרימת מהלך — מקצה לקצה")
    lines = [
        "# Session.handle_move  (server-side)",
        "def handle_move(self, msg):",
        "    if self.state is not SessionState.IN_GAME:",
        "        raise BadStateError(\"no active game\")",
        "    move = Move.parse(msg[\"uci\"])            # \"e2e4\" -> Move object",
        "    game = self.current_game",
        "    game.submit_move(self, move)             # 1) Game: check it is your turn",
        "                                             # 2) Board: full legality check",
        "                                             # 3) Board: apply + detect mate/stale/draw",
        "    self.send(protocol.ok())",
        "    if game.result is None:",
        "        self.registry.broadcast_state(game, last_move=msg[\"uci\"])",
        "        next_session = game.current_session()",
        "        if isinstance(next_session, AISession):",
        "            self.registry.drive_ai(next_session)   # AI's turn? make it play now",
        "    else:",
        "        self.registry.end(game)              # send 'game_ended' to both sides",
    ]
    add_code_block(slide, MARGIN_X, CONTENT_TOP,
                   SLIDE_W - 2 * MARGIN_X, Inches(4.6),
                   lines, font_size=11)

    bottom = add_textbox(slide, MARGIN_X, Inches(6.0),
                         SLIDE_W - 2 * MARGIN_X, Inches(0.9))
    add_he_paragraph(
        bottom.text_frame,
        [("שימו לב: ", "he"),
         ("הלקוח לא מאמת חוקיות בעצמו", "bold"),
         (". השרת הוא ה‑", "he"),
         ("source of truth", "en"),
         (". כך לקוח זדוני לא יכול \"לרמות\".", "he")],
        size=13, first=True,
    )


def slide_19_rules(slide):
    add_title(slide, "חוקיות מהלכים — Game.submit_move")
    lines = [
        "def submit_move(self, session, move: Move) -> Move:",
        "    if self.result is not None:",
        "        raise IllegalMoveError(\"game has already ended\")",
        "    mover_color = self.color_of(session)",
        "    if mover_color is None:",
        "        raise NotYourTurnError(\"session is not a participant\")",
        "    if mover_color is not self.board.side_to_move:",
        "        raise NotYourTurnError(\"it is not your turn\")",
        "",
        "    resolved = self.board.apply(move)   # single source of truth for legality",
        "    self.last_move_uci = resolved.to_uci()",
        "",
        "    if self.board.is_checkmate():",
        "        self.result = (Result.white_wins(\"checkmate\") if mover_color is Color.WHITE",
        "                       else Result.black_wins(\"checkmate\"))",
        "    elif self.board.is_stalemate():",
        "        self.result = Result.draw(\"stalemate\")",
        "    elif self.board.is_threefold_repetition():",
        "        self.result = Result.draw(\"threefold\")",
        "    elif self.board.is_fifty_move_draw():",
        "        self.result = Result.draw(\"fifty-move\")",
        "    return resolved",
    ]
    add_code_block(slide, MARGIN_X, CONTENT_TOP,
                   SLIDE_W - 2 * MARGIN_X, Inches(5.4),
                   lines, font_size=11)

    bottom = add_textbox(slide, MARGIN_X, SLIDE_H - Inches(0.7),
                         SLIDE_W - 2 * MARGIN_X, Inches(0.4))
    add_he_paragraph(
        bottom.text_frame,
        [("חוקי השחמט עצמם נמצאים במחלקה ", "he"), ("Board", "en"),
         (" (", "he"), ("~315 lines", "en"),
         (") — תנועות פסאודו, סינון לפי שח, רוקדה, ", "he"),
         ("en-passant", "en"), (", הכתרה.", "he")],
        size=13, first=True,
    )


def slide_20_ai_alphabeta(slide):
    add_title(slide, "בונוס: יריב AI — alpha-beta minimax")
    lines = [
        "class AIPlayer:",
        "    def choose_move(self, board: Board, depth: int = 3) -> Move:",
        "        side = board.side_to_move",
        "        legal = board.legal_moves(side)",
        "        best_score, best_moves = None, []",
        "        alpha, beta = -math.inf, math.inf",
        "        for move in legal:",
        "            snap = board._snapshot()",
        "            try:",
        "                board._apply_unchecked(move)",
        "                score = -self._negamax(board, depth - 1, -beta, -alpha,",
        "                                       perspective=side.opposite())",
        "            finally:",
        "                board._restore(snap)",
        "            if best_score is None or score > best_score:",
        "                best_score, best_moves = score, [move]",
        "            elif score == best_score:",
        "                best_moves.append(move)",
        "            if score > alpha:",
        "                alpha = score",
        "        return random.choice(best_moves)   # random tie-break between equal moves",
    ]
    add_code_block(slide, MARGIN_X, CONTENT_TOP,
                   SLIDE_W - 2 * MARGIN_X, Inches(4.8),
                   lines, font_size=11)

    add_callout(
        slide, MARGIN_X, Inches(6.1),
        SLIDE_W - 2 * MARGIN_X, Inches(1.05),
        [
            [("negamax + alpha-beta:", "bold"),
             (" אותה פונקציה מנקודת מבט של הצד הנוכחי, עם גזירה ", "he"),
             ("(pruning)", "en"),
             (" כשמוצאים תוצאה שמשתחררת. עומק ברירת מחדל = 3.", "he")],
        ],
        font_size=13,
    )


def slide_21_eval(slide):
    add_title(slide, "הערכת מצב — material balance")
    lines = [
        "PIECE_VALUE = {",
        "    PieceType.PAWN:   100,",
        "    PieceType.KNIGHT: 300,",
        "    PieceType.BISHOP: 300,",
        "    PieceType.ROOK:   500,",
        "    PieceType.QUEEN:  900,",
        "    PieceType.KING:   0,        # no value - mate is handled via CHECKMATE_SCORE",
        "}",
        "CHECKMATE_SCORE = 1_000_000",
        "",
        "def evaluate(board: Board) -> int:",
        "    \"\"\"Material balance from White's perspective, in centipawns.\"\"\"",
        "    score = 0",
        "    for f in range(8):",
        "        for r in range(8):",
        "            p = board.squares[f][r]",
        "            if p is None:",
        "                continue",
        "            val = PIECE_VALUE[p.kind]",
        "            score += val if p.color is Color.WHITE else -val",
        "    return score",
    ]
    add_code_block(slide, MARGIN_X, CONTENT_TOP,
                   SLIDE_W - 2 * MARGIN_X, Inches(5.0),
                   lines, font_size=11)

    bottom = add_textbox(slide, MARGIN_X, Inches(6.3),
                         SLIDE_W - 2 * MARGIN_X, Inches(0.7))
    add_he_paragraph(
        bottom.text_frame,
        [("פונקציית הערכה פשוטה אבל אפקטיבית — מספיק כדי לנצח שחקן מתחיל. הדרישה היא ", "he"),
         ("שימוש", "bold"), (" ב‑", "he"), ("AI", "en"),
         (", לא מנוע ברמת ", "he"), ("Stockfish", "en"), (".", "he")],
        size=13, first=True,
    )


def slide_22_ai_integ(slide):
    add_title(slide, "איך ה‑AI משתלב בפרוטוקול?")
    lines = [
        "# Session.handle_play_ai",
        "def handle_play_ai(self, msg):",
        "    if self.state is not SessionState.AUTHENTICATED:",
        "        raise BadStateError(...)",
        "    requested_color = msg.get(\"color\", \"white\")",
        "    depth = max(1, min(4, int(msg.get(\"depth\", self.ai_depth))))",
        "",
        "    ai_session = AISession(depth=depth, registry=self.registry)",
        "    if requested_color == \"white\":",
        "        white, black = self, ai_session",
        "    else:",
        "        white, black = ai_session, self",
        "    self.registry.create(white, black)",
        "    if isinstance(white, AISession):",
        "        self.registry.drive_ai(white)        # AI opens if it plays White",
    ]
    add_code_block(slide, MARGIN_X, CONTENT_TOP,
                   SLIDE_W - 2 * MARGIN_X, Inches(3.8),
                   lines, font_size=11)

    desc = add_textbox(slide, MARGIN_X, Inches(5.15),
                       SLIDE_W - 2 * MARGIN_X, Inches(1.0))
    add_he_paragraph(
        desc.text_frame,
        [("טריק יפה:", "bold"),
         (" ", "he"), ("AISession", "en"),
         (" מממש את אותו ", "he"),
         ("חוזה", "bold"),
         (" כמו ", "he"), ("Session", "en"),
         (" אמיתי (", "he"),
         (".send, .account, .current_game", "en"), (").", "he")],
        size=14, first=True,
    )
    add_he_paragraph(
        desc.text_frame,
        [("לכן ", "he"), ("Game", "en"),
         (" ו‑", "he"), ("GameRegistry", "en"),
         (" בכלל לא יודעים שזה לא בן-אדם.", "he")],
        size=14,
    )

    add_callout(
        slide, MARGIN_X, Inches(6.55),
        SLIDE_W - 2 * MARGIN_X, Inches(0.6),
        [
            [("שימוש בעיקרון Liskov:", "bold"),
             (" בני אדם ו‑", "he"), ("AI", "en"),
             (" משחקים דרך אותו ממשק.", "he")],
        ],
        font_size=13,
    )


def slide_23_parallel(slide):
    add_title(slide, "בונוס: ריבוי משחקים במקביל")
    box = add_textbox(slide, MARGIN_X, CONTENT_TOP,
                      SLIDE_W - 2 * MARGIN_X, Inches(2.0))
    bullets = [
        [("כל ", "he"), ("Session", "en"), (" רץ ב‑", "he"),
         ("thread", "en"), (" משלו → אין חסימה בין לקוחות.", "he")],
        [("UserStore, Lobby, GameRegistry", "en"),
         (" — כולם נעולים ב‑", "he"),
         ("threading.Lock", "en"), (".", "he")],
        [("הלובי מצמיד שני שחקנים זמינים → כל זוג מקבל ", "he"),
         ("Game", "en"), (" משל עצמו ב‑", "he"),
         ("GameRegistry", "en"), (".", "he")],
        [("שחקנים זוגות אחרים מקבלים ", "he"),
         ("game_state", "en"), (" רק על המשחק שלהם.", "he")],
    ]
    first = True
    for segs in bullets:
        add_he_paragraph(box.text_frame, segs, size=14, bullet=True,
                         first=first)
        first = False

    lines = [
        "# lobby.py - locked under multi-thread access",
        "class GameRegistry:",
        "    def __init__(self) -> None:",
        "        self._games: Dict[str, Game] = {}",
        "        self._lock = threading.Lock()",
        "",
        "    def create(self, white, black) -> Game:",
        "        game = Game(white=white, black=black)",
        "        with self._lock:",
        "            self._games[game.id] = game   # atomic wrt other threads",
        "        ...",
    ]
    add_code_block(slide, MARGIN_X, Inches(3.4),
                   SLIDE_W - 2 * MARGIN_X, Inches(2.9),
                   lines, font_size=12)

    bottom = add_textbox(slide, MARGIN_X, Inches(6.5),
                         SLIDE_W - 2 * MARGIN_X, Inches(0.5))
    add_he_paragraph(
        bottom.text_frame,
        [("קיים ", "he"),
         ("integration test", "en"), (" מלא: ", "he"),
         ("tests/integration/test_parallel_games.py", "en"),
         (".", "he")],
        size=13, first=True,
    )


def slide_24_gui(slide):
    add_title(slide, "ממשק המשתמש — Tkinter GUI")
    box = add_textbox(slide, MARGIN_X, CONTENT_TOP,
                      SLIDE_W - 2 * MARGIN_X, Inches(2.8))
    bullets = [
        [("שלושה מסכים בתוך אותו חלון: ", "he"),
         ("Login → Lobby → Game", "bold"), (".", "he")],
        [("לוח 8×8 מצויר על ", "he"), ("Canvas", "en"),
         (" עם סמלי ", "he"), ("Unicode", "en"),
         (" של כלי שחמט.", "he")],
        [("קליק על הכלי שלי + קליק על משבצת יעד → שולח ", "he"),
         ('{"type":"move","uci":"e2e4"}', "en"), (".", "he")],
        [("Threading", "en"), (" בלקוח: ", "he"),
         ("thread רקע", "bold"),
         (" שקורא מהשרת ודוחף ל‑", "he"),
         ("queue.Queue", "en"),
         ("; ה‑", "he"), ("main loop", "en"),
         (" של ", "he"), ("Tk", "en"),
         (" דוגם אותה כל ", "he"), ("50ms", "en"), (".", "he")],
    ]
    first = True
    for segs in bullets:
        add_he_paragraph(box.text_frame, segs, size=14, bullet=True,
                         first=first)
        first = False

    lines = [
        "# gui.py - screen-switching pattern",
        "def _show_screen(self, name: str) -> None:",
        "    for widget in self._content.winfo_children():",
        "        widget.destroy()",
        "    builder = {",
        "        \"login\": self._build_login,",
        "        \"lobby\": self._build_lobby,",
        "        \"game\":  self._build_game,",
        "    }[name]",
        "    builder(self._content)",
    ]
    add_code_block(slide, MARGIN_X, Inches(4.2),
                   SLIDE_W - 2 * MARGIN_X, Inches(2.5),
                   lines, font_size=12)

    bottom = add_textbox(slide, MARGIN_X, Inches(6.9),
                         SLIDE_W - 2 * MARGIN_X, Inches(0.4))
    add_he_paragraph(
        bottom.text_frame,
        [("Tkinter", "en"),
         (" נכלל ב‑", "he"), ("Python", "en"),
         (" סטנדרטי, אז אין תלות חיצונית.", "he")],
        size=13, first=True,
    )


def slide_25_cli(slide):
    add_title(slide, "ממשק חלופי — CLI (טקסט)")
    lines = [
        "$ python -m secure_chess.client --host 127.0.0.1 --port 5050 --cli",
        "connected to 127.0.0.1:5050",
        "> register alice hunter2!",
        "ok",
        "> play_human",
        "ok",
        "",
        "== game started ==",
        "you are: white   opponent: bob   to move: white",
        "8 r n b q k b n r",
        "7 p p p p p p p p",
        "6 . . . . . . . .",
        "5 . . . . . . . .",
        "4 . . . . . . . .",
        "3 . . . . . . . .",
        "2 P P P P P P P P",
        "1 R N B Q K B N R",
        "  a b c d e f g h",
        "> move e2e4",
        "ok",
    ]
    add_code_block(slide, MARGIN_X, CONTENT_TOP,
                   SLIDE_W - 2 * MARGIN_X, Inches(5.0),
                   lines, font_size=11)

    box = add_textbox(slide, MARGIN_X, Inches(6.3),
                      SLIDE_W - 2 * MARGIN_X, Inches(0.9))
    bullets = [
        [("אותו פרוטוקול בדיוק כמו ה‑", "he"),
         ("GUI", "en"), (" → ", "he"),
         ("GUI", "en"), (" יכול לשחק נגד ", "he"),
         ("CLI", "en"), (".", "he")],
        [("thread", "en"), (" נפרד קורא הודעות שרת בזמן שהמשתמש מקליד.", "he")],
    ]
    first = True
    for segs in bullets:
        add_he_paragraph(box.text_frame, segs, size=13, bullet=True,
                         first=first)
        first = False


def slide_26_errors(slide):
    add_title(slide, "היררכיית שגיאות — errors.py")
    lines = [
        "class SecureChessError(Exception):",
        "    code: str = \"internal_error\"",
        "",
        "class MoveParseError(SecureChessError):        code = \"move_parse_error\"",
        "class IllegalMoveError(SecureChessError):      code = \"illegal_move\"",
        "class NotYourTurnError(SecureChessError):      code = \"not_your_turn\"",
        "class ProtocolError(SecureChessError):         code = \"bad_request\"",
        "class DuplicateUserError(SecureChessError):    code = \"duplicate_user\"",
        "class WeakPasswordError(SecureChessError):     code = \"weak_password\"",
        "class AuthenticationError(SecureChessError):   code = \"auth_failed\"",
        "class AlreadyLoggedInError(SecureChessError):  code = \"already_logged_in\"",
        "class ValidationError(SecureChessError):       code = \"validation\"",
        "class BadStateError(SecureChessError):         code = \"bad_state\"",
        "class ServerFullError(SecureChessError):       code = \"server_full\"",
    ]
    add_code_block(slide, MARGIN_X, CONTENT_TOP,
                   SLIDE_W - 2 * MARGIN_X, Inches(4.0),
                   lines, font_size=11)

    add_callout(
        slide, MARGIN_X, Inches(5.3),
        SLIDE_W - 2 * MARGIN_X, Inches(1.7),
        [
            [("כל שגיאה יודעת מה ה‑", "he"), ("code", "en"),
             (" שלה. ה‑", "he"), ("dispatcher", "en"),
             (" תופס ", "he"), ("SecureChessError", "en"),
             (" אחד וממיר אותה אוטומטית להודעת ", "he"),
             ("error", "en"), (" בפרוטוקול:", "he")],
            [('{"type":"error","code":"auth_failed","message":"..."}', "en")],
        ],
        font_size=13,
    )


def slide_27_tests(slide):
    add_title(slide, "בדיקות — pytest")

    col_w = (SLIDE_W - 3 * MARGIN_X) / 2
    left_box = add_textbox(slide, MARGIN_X, CONTENT_TOP,
                           col_w, Inches(3.6))
    add_he_paragraph(
        left_box.text_frame,
        [("בדיקות יחידה", "bold")],
        size=16, first=True,
    )
    for segs in [
        [("test_crypto.py", "en"), (" — ", "he"),
         ("bcrypt hash/verify", "en")],
        [("test_user_store.py", "en"), (" — רישום / כפילות / סיסמה חלשה", "he")],
        [("test_pieces.py", "en"), (" — תנועות לכל סוג כלי", "he")],
        [("test_board_rules.py", "en"),
         (" — שחמט / פט / רוקדה", "he")],
        [("test_ai.py", "en"), (" — בחירת מהלך, גזירה", "he")],
        [("test_protocol_auth.py / test_protocol_game.py", "en")],
        [("test_client_gui.py", "en"), (" — לוגיקת ", "he"),
         ("FEN", "en"), (" ב‑", "he"), ("GUI", "en")],
    ]:
        add_he_paragraph(left_box.text_frame, segs, size=13, bullet=True)

    right_box = add_textbox(slide, MARGIN_X * 2 + col_w, CONTENT_TOP,
                            col_w, Inches(3.6))
    add_he_paragraph(
        right_box.text_frame,
        [("בדיקות אינטגרציה", "bold")],
        size=16, first=True,
    )
    for segs in [
        [("test_register_login.py", "en")],
        [("test_full_game.py", "en"),
         (" — שני לקוחות, מהלכים, סיום", "he")],
        [("test_parallel_games.py", "en"), (" — בונוס", "he")],
        [("test_play_ai.py", "en"), (" — בונוס", "he")],
        [("test_cancel_lobby.py", "en")],
    ]:
        add_he_paragraph(right_box.text_frame, segs, size=13, bullet=True)

    add_code_block(slide, MARGIN_X, Inches(5.6),
                   SLIDE_W - 2 * MARGIN_X, Inches(1.0),
                   ["pytest -v",
                    "# expected: 60+ tests pass in < 60 seconds"],
                   font_size=13)


def slide_28_demo(slide):
    add_title(slide, "הדגמה חיה — סדר הצעדים")

    box = add_textbox(slide, MARGIN_X, CONTENT_TOP,
                      SLIDE_W - 2 * MARGIN_X, Inches(0.5))
    add_he_paragraph(
        box.text_frame,
        [("שלב 1 — שרת:", "bold")],
        size=14, first=True, bullet=True,
    )
    add_code_block(
        slide, MARGIN_X + Inches(0.4), Inches(1.55),
        SLIDE_W - 2 * MARGIN_X - Inches(0.4), Inches(1.0),
        [
            "cd secure-chess",
            ".\\.venv\\Scripts\\Activate.ps1",
            "python -m secure_chess.server --host 127.0.0.1 --port 5050 --data-dir ./data",
        ],
        font_size=11,
    )

    rest_box = add_textbox(slide, MARGIN_X, Inches(2.7),
                           SLIDE_W - 2 * MARGIN_X, Inches(4.2))
    items = [
        [("שלב 2 — לקוח ", "bold"),
         ("Alice", "en"), (":", "bold"),
         (" ", "he"), ("python -m secure_chess.client", "en"),
         (" → רישום → ", "he"), ("Join Lobby", "en"), (".", "he")],
        [("שלב 3 — לקוח ", "bold"),
         ("Bob", "en"), (":", "bold"),
         (" ", "he"), ("python -m secure_chess.client", "en"),
         (" → רישום → ", "he"), ("Join Lobby", "en"),
         (". השרת מצמיד והמשחק מתחיל.", "he")],
        [("שלב 4:", "bold"),
         (" לבצע 2–3 מהלכים, להראות שעדכון מצב מגיע ", "he"),
         ("לשני", "bold"), (" הלקוחות.", "he")],
        [("שלב 5:", "bold"),
         (" לפתוח את ", "he"), ("data/users.json", "en"),
         (" ולהראות ", "he"), ("hashes", "en"), (" בלבד.", "he")],
        [("שלב 6 — בונוס ", "bold"), ("AI", "en"), (":", "bold"),
         (" ", "he"), ("python -m secure_chess.client --cli", "en"),
         (" → ", "he"), ("play_ai", "en"), (".", "he")],
        [("שלב 7 — בונוס מקבילי:", "bold"),
         (" לפתוח עוד 2 לקוחות → שני משחקים בו זמנית.", "he")],
        [("שלב 8:", "bold"),
         (" להריץ ", "he"), ("pytest -v", "en"),
         (" ולהראות 60+ בדיקות ירוקות.", "he")],
    ]
    first = True
    for segs in items:
        add_he_paragraph(rest_box.text_frame, segs, size=13, bullet=True,
                         first=first, space_after=3)
        first = False

    add_speaker_tip(slide, MARGIN_X, SLIDE_H - Inches(0.75),
                    SLIDE_W - 2 * MARGIN_X, Inches(0.45),
                    "להכין מראש את ארבעת הטרמינלים פתוחים — לא מאבדים זמן על הקלדה מול הבוחן.")


def slide_29_qa(slide):
    add_title(slide, "שאלות צפויות מהבוחן — תשובות מוכנות")
    box = add_textbox(slide, MARGIN_X, CONTENT_TOP,
                      SLIDE_W - 2 * MARGIN_X, Inches(5.8))
    qa = [
        [("\"למה ", "he"), ("bcrypt", "en"), (" ולא ", "he"),
         ("SHA-256", "en"), ("?\" ", "bold"),
         ("— ", "he"), ("bcrypt", "en"),
         (" איטי בכוונה ומכיל ", "he"), ("salt", "en"),
         (" פנימי, ולכן עמיד ל‑", "he"), ("brute-force", "en"),
         (" וטבלאות ", "he"), ("rainbow", "en"), (".", "he")],
        [("\"איך אתה מבטיח שהתקשורת בין שני המהלכים לא תתערבב?\" ", "bold"),
         ("— ", "he"), ("JSON-Lines", "en"),
         (" נותן ", "he"),
         ("framing", "en"),
         (" ברור (", "he"), ("\\n", "en"),
         ("), ו‑", "he"), ("_LineReader", "en"),
         (" צובר עד שמגיע סוף שורה.", "he")],
        [("\"מה קורה אם לקוח מתנתק באמצע משחק?\" ", "bold"),
         ("— ", "he"), ("_on_disconnect", "en"),
         (" ב‑", "he"), ("Session", "en"),
         (" קורא ל‑", "he"), ("Game.handle_disconnect", "en"),
         (" שמסיים את המשחק בניצחון לצד השני (", "he"),
         ("reason=\"disconnect\"", "en"), (").", "he")],
        [("\"איך השרת מטפל בכמה שחקנים בו זמנית?\" ", "bold"),
         ("— ", "he"), ("thread per connection", "en"),
         (" + מבני נתונים משותפים תחת ", "he"),
         ("threading.Lock", "en"), (".", "he")],
        [("\"איפה החוקיות נאכפת?\" ", "bold"),
         ("— רק ב‑", "he"), ("Board.apply", "en"),
         (" בשרת. הלקוח לא מאמת — אחרת לקוח זדוני יוכל לרמות.", "he")],
        [("\"איך ה‑", "he"), ("AI", "en"),
         (" משתלב מבלי לכתוב מסלול מיוחד?\" ", "bold"),
         ("— ", "he"), ("AISession", "en"),
         (" מחקה את הממשק של ", "he"), ("Session", "en"),
         (" (", "he"), ("duck typing", "en"),
         (") ולכן ", "he"), ("Game", "en"),
         (" לא מבדיל ביניהם.", "he")],
    ]
    first = True
    for segs in qa:
        add_he_paragraph(box.text_frame, segs, size=13, bullet=True,
                         first=first, space_after=5)
        first = False


def slide_30_closing(slide):
    _center_title(slide, "סיכום", size=44)
    desc = add_textbox(slide, Inches(1.5), Inches(2.5),
                       SLIDE_W - Inches(3.0), Inches(2.0))
    add_he_paragraph(
        desc.text_frame,
        [("ארכיטקטורה נקייה, פרוטוקול ", "he"),
         ("JSON-Lines", "en"),
         (" פשוט אך עמיד, סיסמאות מוגנות ב‑", "he"),
         ("bcrypt", "en"),
         (", ריבוי משחקים תחת ", "he"),
         ("threading", "en"),
         (", ויריב ", "he"), ("AI", "en"),
         (" מבוסס ", "he"),
         ("alpha-beta minimax", "en"),
         (" — הכל מכוסה ב‑60+ בדיקות ", "he"),
         ("pytest", "en"), (".", "he")],
        size=18, alignment=PP_ALIGN.CENTER, first=True,
    )

    badge_w = Inches(4.0)
    badge = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        (SLIDE_W - badge_w) / 2, Inches(5.2),
        badge_w, Inches(0.6))
    badge.fill.solid()
    badge.fill.fore_color.rgb = COLOR_ACCENT_SOFT
    badge.line.fill.background()
    badge.adjustments[0] = 0.5
    bp = badge.text_frame.paragraphs[0]
    bp.alignment = PP_ALIGN.CENTER
    set_rtl(bp)
    add_run(bp, "תודה — נשמח לשאלות",
            font=HEBREW_FONT, size=20, bold=True, color=COLOR_ACCENT)


if __name__ == "__main__":
    here = Path(__file__).parent
    build_presentation(here / "secure-chess.pptx")
