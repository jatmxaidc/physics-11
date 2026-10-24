"""Motion 01 — frame of reference, rest/motion, distance and displacement (PW 01)."""
from __future__ import annotations

import math
import re
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt

BASE = Path(__file__).resolve().parent
OUTPUT = BASE / "lecture slides" / "Lecture_Motion_01_Frame_Distance_Displacement.pptx"

NAVY = RGBColor(8, 30, 63)
BLUE = RGBColor(64, 156, 255)
TEAL = RGBColor(40, 180, 190)
GREEN = RGBColor(90, 214, 128)
AMBER = RGBColor(255, 204, 72)
ORANGE = RGBColor(255, 158, 64)
RED = RGBColor(255, 96, 96)
TEXT = RGBColor(255, 255, 255)
MUTED = RGBColor(176, 182, 192)
STROKE = RGBColor(72, 78, 90)
WHITE = RGBColor(255, 255, 255)
SOFT = RGBColor(22, 24, 28)
BLACK = RGBColor(0, 0, 0)
GOLD = RGBColor(255, 214, 70)
NCERT_FILL = RGBColor(255, 248, 225)
NCERT_LINE = RGBColor(218, 168, 64)
QCOL = ORANGE
WORK = RGBColor(90, 200, 255)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

XL, CW, Y0, YF = 0.65, 12.05, 1.18, 7.12
LH = 0.42
NG = 0.18
PT_SCALE = 0.90

_SUP_MAP = str.maketrans({
    "0": "⁰", "1": "¹", "2": "²", "3": "³", "4": "⁴",
    "5": "⁵", "6": "⁶", "7": "⁷", "8": "⁸", "9": "⁹",
    "+": "⁺", "-": "⁻", "=": "⁼", "(": "⁽", ")": "⁾",
    "a": "ᵃ", "b": "ᵇ", "c": "ᶜ", "d": "ᵈ", "e": "ᵉ",
    "f": "ᶠ", "g": "ᵍ", "h": "ʰ", "i": "ⁱ", "j": "ʲ",
    "k": "ᵏ", "l": "ˡ", "m": "ᵐ", "n": "ⁿ", "o": "ᵒ",
    "p": "ᵖ", "r": "ʳ", "s": "ˢ", "t": "ᵗ", "u": "ᵘ",
    "v": "ᵛ", "w": "ʷ", "x": "ˣ", "y": "ʸ", "z": "ᶻ",
    "A": "ᴬ", "B": "ᴮ", "D": "ᴰ", "E": "ᴱ", "G": "ᴳ",
    "H": "ᴴ", "I": "ᴵ", "J": "ᴶ", "K": "ᴷ", "L": "ᴸ",
    "M": "ᴹ", "N": "ᴺ", "O": "ᴼ", "P": "ᴾ", "R": "ᴿ",
    "T": "ᵀ", "U": "ᵁ", "V": "ⱽ", "W": "ᵂ",
    "/": "ᐟ",
})


def _to_sup(s: str) -> str:
    return s.translate(_SUP_MAP)


_GREEK_SUP = {"α": "ᵅ", "β": "ᵝ", "γ": "ᵞ", "δ": "ᵟ", "ε": "ᵋ", "θ": "ᶿ"}


_SUB_CHARS = {
    "0": "₀", "1": "₁", "2": "₂", "3": "₃", "4": "₄",
    "5": "₅", "6": "₆", "7": "₇", "8": "₈", "9": "₉",
    "a": "ₐ", "e": "ₑ", "h": "ₕ", "i": "ᵢ", "j": "ⱼ",
    "k": "ₖ", "l": "ₗ", "m": "ₘ", "n": "ₙ", "o": "ₒ",
    "p": "ₚ", "r": "ᵣ", "s": "ₛ", "t": "ₜ", "u": "ᵤ",
    "v": "ᵥ", "x": "ₓ", "y": "ᵧ",
    # no reliable Unicode subscript f — use PPT native subscript run instead
}


def _format_powers(text: str) -> str:
    """Turn caret powers into Unicode superscripts: 10^(-3) → 10⁻³, L^2 → L²."""
    text = text.replace("^(-1/2)", "⁻½").replace("^(-3/2)", "⁻³⁄²")
    text = text.replace("^(-5/2)", "⁻⁵⁄²").replace("^(-7/2)", "⁻⁷⁄²")
    text = text.replace("^(1/2)", "½").replace("^(3/2)", "³⁄²")
    text = text.replace("^(5/2)", "⁵⁄²").replace("^(7/2)", "⁷⁄²")
    text = re.sub(r"\^\(([^)]+)\)", lambda m: _to_sup(m.group(1)), text)
    text = re.sub(r"\^(-?\d+)", lambda m: _to_sup(m.group(1)), text)
    text = re.sub(r"\^([A-Za-z])", lambda m: _to_sup(m.group(1)), text)
    text = re.sub(r"\^([αβγδεθ])", lambda m: _GREEK_SUP[m.group(1)], text)
    return text


def _format_subs_unicode(text: str) -> str:
    """Convert _i, _{fi}, _x into Unicode subscripts; unknown letters → PPT subscript markers."""
    def brace(m):
        out = []
        for ch in m.group(1):
            if ch in _SUB_CHARS:
                out.append(_SUB_CHARS[ch])
            elif ch.isalnum():
                out.append(f"\0{ch}\0")
            else:
                out.append(ch)
        return "".join(out)

    text = re.sub(r"_\{([^}]+)\}", brace, text)
    text = re.sub(
        r"_([A-Za-z0-9])",
        lambda m: _SUB_CHARS[m.group(1)] if m.group(1) in _SUB_CHARS else f"\0{m.group(1)}\0",
        text,
    )
    return text


def _bg(slide, color=BLACK):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color


def _solid(slide, x, y, w, h, fill, no_line=True):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if no_line:
        s.line.fill.background()
    else:
        s.line.color.rgb = fill
    return s


def _card_bg(slide, x, y, w, h, fill, border, line_pt=1.0, rounded=True):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    s = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.color.rgb = border
    s.line.width = Pt(line_pt)
    return s


def tx(slide, text, x, y, w, h, size=15, color=TEXT, bold=False,
       align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP, spa=1, lsp=1.1,
       font="Aptos", scaled=True):
    text = _format_powers(text)
    for old, new in (
        ("<<", "≪"),
        (">>", "≫"),
        ("sqrt(2500)", "√2500"),
        ("sqrt(200)", "√200"),
        ("sqrt(125)", "√125"),
        ("sqrt(100 + 100)", "√(100 + 100)"),
        ("sqrt(50)", "√50"),
        ("sqrt(48)", "√48"),
        ("sqrt(37)", "√37"),
        ("sqrt(25)", "√25"),
        ("sqrt(13)", "√13"),
        ("sqrt(8)", "√8"),
        ("sqrt(5)", "√5"),
        ("sqrt(3)", "√3"),
        ("sqrt(2)", "√2"),
        ("sqrt", "√"),
        ("cos 90", "cos 90°"),
        ("cos 60", "cos 60°"),
        ("cos 53", "cos 53°"),
        ("cos 45", "cos 45°"),
        ("cos 37", "cos 37°"),
        ("sin(60)", "sin 60°"),
        ("cos(60)", "cos 60°"),
        ("sin 60", "sin 60°"),
        ("sin 53", "sin 53°"),
        ("sin 45", "sin 45°"),
        ("sin 37", "sin 37°"),
        ("i-hat", "î"),
        ("j-hat", "ĵ"),
        ("k-hat", "k̂"),
        ("A-hat", "Â"),
        (" dot ", " · "),
        ("· product", "dot product"),
        (" times ", " × "),
        ("*", "×"),
        (" degrees", "°"),
        (" degree", "°"),
        ("°°", "°"),
        ("theta", "θ"),
        ("alpha", "α"),
        ("beta", "β"),
        ("cos inverse", "cos⁻¹"),
        ("tan inverse", "tan⁻¹"),
        ("F1", "F₁"),
        ("F2", "F₂"),
        ("Fx", "Fₓ"),
        ("Fy", "Fᵧ"),
        ("Fz", "Fz"),
        ("omega", "ω"),
        ("lambda", "λ"),
        (" proportional to ", " ∝ "),
        (">=", "≥"),
        ("<=", "≤"),
        ("->", "→"),
    ):
        text = text.replace(old, new)
    text = re.sub(r"\s*=\s*", " = ", text)
    text = re.sub(r"\s*([<>])\s*", r" \1 ", text)
    text = re.sub(r",\s*", ", ", text)
    # Subscripts: prefer Unicode; letters without a Unicode form use a PPT subscript run.
    text = _format_subs_unicode(text)
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(0)
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]
    p.alignment = align
    p.space_before = Pt(0)
    p.space_after = Pt(spa)
    p.line_spacing = lsp
    pt_size = Pt(size if not scaled else max(7, round(size * PT_SCALE, 1)))

    def _add_run(chunk: str, as_sub: bool = False):
        if not chunk:
            return
        run = p.add_run()
        run.text = chunk
        run.font.name = font
        run.font.size = pt_size
        run.font.bold = bold
        run.font.color.rgb = color
        if as_sub:
            # python-pptx Font.subscript is unreliable — set OOXML baseline directly
            rPr = run._r.get_or_add_rPr()
            for old in rPr.findall(qn("a:baseline")):
                rPr.remove(old)
            bl = OxmlElement("a:baseline")
            bl.set("val", "-25000")
            rPr.append(bl)

    # \0X\0 markers = letters that need PPT native subscript (e.g. f)
    parts = re.split(r"\0([A-Za-z0-9])\0", text)
    if len(parts) == 1:
        _add_run(text)
    else:
        for i, part in enumerate(parts):
            _add_run(part, as_sub=(i % 2 == 1))
    return box


def _bar(slide, label, color=BLUE):
    _solid(slide, 0, 0, 13.333, 0.29, color)
    if label:
        tx(slide, label.upper(), XL, 0.05, 9, 0.20, 8, WHITE, True, spa=0)


def _title(slide, text):
    tx(slide, text, XL, 0.34, CW, 0.72, 26, GOLD, True, spa=0)


def _sep(slide, y=1.10):
    _solid(slide, XL, y, CW, 0.018, STROKE)


def _footer(slide, n):
    tx(slide, f"{n}", 12.25, YF, 0.8, 0.25, 11, MUTED, True, PP_ALIGN.RIGHT, spa=0)


def add_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


_CONTENT = 0
_ROUGH_DUE = False
_ANSWER_STEPS = {}


def add_rough_work():
    slide = prs.slides.add_slide(BLANK)
    _bg(slide)
    _bar(slide, "Rough work", MUTED)
    tx(slide, "Rough work  /  extra work", XL, 0.34, CW, 0.50, 22, GOLD, True, spa=0)
    _sep(slide)
    _footer(slide, len(prs.slides))
    add_notes(slide, "Empty board for extra working.")


def after_content_slide(defer_rough=False):
    global _CONTENT, _ROUGH_DUE
    _CONTENT += 1
    if defer_rough and _CONTENT % 5 == 0:
        _ROUGH_DUE = True
    elif not defer_rough and (_ROUGH_DUE or _CONTENT % 5 == 0):
        add_rough_work()
        _ROUGH_DUE = False


def new_slide(section, title, bar_color=BLUE, defer_rough=False):
    slide = prs.slides.add_slide(BLANK)
    _bg(slide)
    _bar(slide, section, bar_color)
    _title(slide, title)
    _sep(slide)
    _footer(slide, len(prs.slides))
    after_content_slide(defer_rough)
    return slide


_LAST_QUESTION_TITLE = ""


def question_slide(title):
    global _LAST_QUESTION_TITLE
    _LAST_QUESTION_TITLE = title
    return new_slide("Question", title, ORANGE, defer_rough=True)


def answer_slide(title=None):
    title = title or f"Solution: {_LAST_QUESTION_TITLE}"
    slide = new_slide("Answer", title, GREEN)
    _ANSWER_STEPS[id(slide)] = 0
    return slide


def ncert(slide, formula, x, y, w, h, size=20):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = NCERT_FILL
    s.line.color.rgb = NCERT_LINE
    s.line.width = Pt(1.5)
    tx(slide, formula, x + 0.18, y + 0.16, w - 0.36, h - 0.32,
       size, NAVY, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE, spa=0, lsp=1.15)


def fml(slide, text, x, y, w, h=0.44, size=22):
    if id(slide) in _ANSWER_STEPS:
        _ANSWER_STEPS[id(slide)] += 1
        text = f"Step {_ANSWER_STEPS[id(slide)]}:  {text}"
    return tx(slide, text, x, y, w, h, size, WORK, True, spa=0)


def body(slide, text, x, y, w, h=0.44, size=17):
    return tx(slide, text, x, y, w, h, size, TEXT, False, spa=0, scaled=False)


def para(slide, text, x, y, w, h=0.95, size=16):
    """Long theory paragraph — wraps; use taller h and spill to next slide when full."""
    return tx(slide, text, x, y, w, h, size, TEXT, False, spa=0, lsp=1.28, scaled=False)


def bullet(slide, text, x, y, w, h=0.42, size=16):
    return tx(slide, "•  " + text, x, y, w, h, size, TEXT, False, spa=0, scaled=False)


def qtx(slide, text, x, y, w, h=0.44, size=18):
    return tx(slide, text, x, y, w, h, max(size, 19), QCOL, True, spa=0, scaled=False)


def result_card(slide, text, x, y, w, h=0.36, tsize=18):
    return tx(slide, text, x, y, w, h, max(tsize, 18), GREEN, True, spa=0)


def trap_card(slide, text, x, y, w, h=0.40):
    return tx(slide, text, x, y, w, h, 16, ORANGE, False, spa=0, scaled=False)


def choices(slide, items, x, y, w=None):
    w = w or CW
    step = 0.52
    for i, item in enumerate(items):
        cy = y + i * step
        circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(cy + 0.06), Inches(0.32), Inches(0.32))
        circ.fill.solid()
        circ.fill.fore_color.rgb = TEAL
        circ.line.fill.background()
        tx(slide, str(i + 1), x, cy + 0.08, 0.32, 0.28, 14, WHITE, True, PP_ALIGN.CENTER, spa=0, scaled=False)
        tx(slide, item, x + 0.46, cy + 0.04, w - 0.46, 0.42, 17, TEXT, False, spa=0, scaled=False)


def grid_table(slide, headers, rows, x, y, col_w, row_h=0.42, hsize=14, csize=15):
    col_x = [x]
    for w in col_w[:-1]:
        col_x.append(col_x[-1] + w)
    for j, ht in enumerate(headers):
        _solid(slide, col_x[j], y, col_w[j], row_h, RGBColor(28, 36, 52), False)
        tx(slide, ht, col_x[j] + 0.04, y, col_w[j] - 0.08, row_h,
           hsize, WHITE, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE, spa=0, scaled=False)
    for i, row in enumerate(rows):
        ry = y + row_h + i * row_h
        fill = RGBColor(16, 16, 20) if i % 2 == 0 else SOFT
        for j, val in enumerate(row):
            s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(col_x[j]), Inches(ry), Inches(col_w[j]), Inches(row_h))
            s.fill.solid()
            s.fill.fore_color.rgb = fill
            s.line.color.rgb = STROKE
            s.line.width = Pt(0.5)
            tx(slide, val, col_x[j] + 0.04, ry, col_w[j] - 0.08, row_h,
               csize, WORK if j == 0 else TEXT, j == 0,
               PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE, spa=0, scaled=False)


def tx_3d(slide, text, x, y, w, h, size, face, edge, font="Arial Black"):
    tx(slide, text, x + 0.045, y + 0.045, w, h, size, edge, True, PP_ALIGN.CENTER, spa=0, scaled=False, font=font)
    tx(slide, text, x, y, w, h, size, face, True, PP_ALIGN.CENTER, spa=0, scaled=False, font=font)


def draw_line(slide, x1, y1, x2, y2, color=WHITE, pt=1.75):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = color
    c.line.width = Pt(pt)
    return c


def arrow_line(slide, x1, y1, x2, y2, color=GOLD, pt=2.4):
    c = draw_line(slide, x1, y1, x2, y2, color, pt)
    head = OxmlElement("a:tailEnd")
    head.set("type", "triangle")
    c.line._ln.append(head)
    return c


def vec(slide, x1, y1, x2, y2, label, color=GOLD, pt=2.8):
    arrow_line(slide, x1, y1, x2, y2, color, pt)
    tx(slide, label, (x1 + x2) / 2 + 0.08, (y1 + y2) / 2 - 0.28, 1.2, 0.30,
       16, color, True, spa=0, scaled=False)


def axis_diagram(slide, x=8.0, y=1.55, w=4.0, h=3.6, coord=True):
    ox, oy = x + w * 0.40, y + h * 0.58
    arrow_line(slide, x, oy, x + w, oy, MUTED, 1.5)
    arrow_line(slide, ox, y + h, ox, y, MUTED, 1.5)
    if coord:
        tx(slide, "+x", x + w - 0.05, oy + 0.06, 0.45, 0.24, 12, MUTED, True, spa=0, scaled=False)
        tx(slide, "+y", ox + 0.08, y, 0.45, 0.24, 12, MUTED, True, spa=0, scaled=False)
    return ox, oy


def small_dot(slide, x, y, color=WHITE):
    d = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x - 0.045), Inches(y - 0.045), Inches(0.09), Inches(0.09))
    d.fill.solid()
    d.fill.fore_color.rgb = color
    d.line.fill.background()
    return d


def angle_mark(slide, cx, cy, r, a1, a2, label=None, color=ORANGE):
    pts = []
    steps = 18
    for i in range(steps + 1):
        a = math.radians(a1 + (a2 - a1) * i / steps)
        pts.append((cx + r * math.cos(a), cy - r * math.sin(a)))
    for (x1, y1), (x2, y2) in zip(pts, pts[1:]):
        draw_line(slide, x1, y1, x2, y2, color, 1.4)
    if label:
        mid = math.radians((a1 + a2) / 2)
        tx(slide, label, cx + (r + 0.08) * math.cos(mid), cy - (r + 0.08) * math.sin(mid),
           0.75, 0.25, 13, color, True, spa=0, scaled=False)


def mini_card(slide, title, detail, x, y, w=3.6, h=1.15, color=TEAL):
    _card_bg(slide, x, y, w, h, RGBColor(12, 14, 18), color, 1.0, True)
    tx(slide, title, x + 0.15, y + 0.12, w - 0.3, 0.30, 15, color, True, spa=0, scaled=False)
    tx(slide, detail, x + 0.15, y + 0.48, w - 0.3, h - 0.58, 12, TEXT, False, spa=0, scaled=False)


def axes(slide, x=7.5, y=1.55, w=4.5, h=4.5, xlab="x", ylab="y"):
    ox, oy = x + w * 0.40, y + h * 0.58
    draw_line(slide, x, oy, x + w, oy, MUTED, 1.5)
    draw_line(slide, ox, y, ox, y + h, MUTED, 1.5)
    tx(slide, xlab, x + w - 0.15, oy + 0.06, 0.3, 0.22, 13, MUTED, True, spa=0)
    tx(slide, ylab, ox + 0.08, y, 0.3, 0.22, 13, MUTED, True, spa=0)
    return ox, oy, w, h


def line_on_axes(slide, ox, oy, dx1, dy1, dx2, dy2, color=GOLD, pt=2.4):
    return draw_line(slide, ox + dx1, oy - dy1, ox + dx2, oy - dy2, color, pt)


def curve(slide, points, color=GOLD, pt=2.4):
    for (x1, y1), (x2, y2) in zip(points, points[1:]):
        draw_line(slide, x1, y1, x2, y2, color, pt)


def graph_axes(slide, x, y, w, h, xmin, xmax, ymin, ymax, xlab="x", ylab="y"):
    def sx(value):
        return x + (value - xmin) * w / (xmax - xmin)

    def sy(value):
        return y + h - (value - ymin) * h / (ymax - ymin)

    if ymin <= 0 <= ymax:
        draw_line(slide, x, sy(0), x + w, sy(0), MUTED, 1.4)
    if xmin <= 0 <= xmax:
        draw_line(slide, sx(0), y, sx(0), y + h, MUTED, 1.4)
    tx(slide, xlab, x + w - 0.18, sy(0) + 0.05 if ymin <= 0 <= ymax else y + h - 0.22,
       0.35, 0.24, 13, MUTED, True, spa=0, scaled=False)
    tx(slide, ylab, sx(0) + 0.07 if xmin <= 0 <= xmax else x + 0.05, y,
       0.45, 0.24, 13, MUTED, True, spa=0, scaled=False)
    return sx, sy


def axis_ticks(slide, sx, sy, x_values, y_values=None, color=MUTED):
    """Mark readable values on one or both axes."""
    y_values = x_values if y_values is None else y_values
    for v in x_values:
        if v == 0:
            continue
        draw_line(slide, sx(v), sy(0) - 0.06, sx(v), sy(0) + 0.06, color, 1.1)
        tx(slide, str(v), sx(v) - 0.18, sy(0) + 0.08, 0.36, 0.22, 12, color, True,
           PP_ALIGN.CENTER, spa=0, scaled=False)
    for v in y_values:
        if v == 0:
            continue
        draw_line(slide, sx(0) - 0.06, sy(v), sx(0) + 0.06, sy(v), color, 1.1)
        tx(slide, str(v), sx(0) - 0.46, sy(v) - 0.11, 0.38, 0.22, 12, color, True,
           PP_ALIGN.RIGHT, spa=0, scaled=False)
    tx(slide, "0", sx(0) - 0.32, sy(0) + 0.08, 0.26, 0.22, 12, color, True,
       PP_ALIGN.RIGHT, spa=0, scaled=False)


def plot_circle(slide, sx, sy, h, k, r, color=GOLD, pt=2.4):
    """True geometric circle via oval (equal axis units required for round look)."""
    left, top = sx(h - r), sy(k + r)
    width = sx(h + r) - sx(h - r)
    height = sy(k - r) - sy(k + r)
    oval = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(left), Inches(top), Inches(width), Inches(height))
    oval.fill.background()
    oval.line.color.rgb = color
    oval.line.width = Pt(pt)
    return oval


def plot_ellipse(slide, sx, sy, h, k, a, b, color=GOLD, pt=2.4):
    left, top = sx(h - a), sy(k + b)
    width = sx(h + a) - sx(h - a)
    height = sy(k - b) - sy(k + b)
    oval = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(left), Inches(top), Inches(width), Inches(height))
    oval.fill.background()
    oval.line.color.rgb = color
    oval.line.width = Pt(pt)
    return oval


def mark_point(slide, sx, sy, x, y, label, color=GREEN, dx=0.10, dy=-0.26):
    dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(sx(x) - 0.055), Inches(sy(y) - 0.055),
                                 Inches(0.11), Inches(0.11))
    dot.fill.solid()
    dot.fill.fore_color.rgb = color
    dot.line.fill.background()
    tx(slide, label, sx(x) + dx, sy(y) + dy, 0.90, 0.24, 12, color, True, spa=0, scaled=False)


def plot_function(slide, fn, sx, sy, xmin, xmax, ymin, ymax, color=GOLD, samples=160, pt=2.2):
    previous = None
    for i in range(samples + 1):
        xv = xmin + (xmax - xmin) * i / samples
        try:
            yv = fn(xv)
        except (ValueError, ZeroDivisionError, OverflowError):
            previous = None
            continue
        if not math.isfinite(yv) or yv < ymin or yv > ymax:
            previous = None
            continue
        current = (sx(xv), sy(yv))
        if previous is not None:
            draw_line(slide, previous[0], previous[1], current[0], current[1], color, pt)
        previous = current


def plot_parametric(slide, fx, fy, sx, sy, t0, t1, color=GOLD, samples=180, pt=2.2):
    points = []
    for i in range(samples + 1):
        t = t0 + (t1 - t0) * i / samples
        points.append((sx(fx(t)), sy(fy(t))))
    curve(slide, points, color, pt)









# ===== Figure helpers (Motion 01) =====

def stick(slide, x, y, color=WHITE, h=0.55, label=None, lab_color=None):
    """Simple standing person: head + body."""
    r = 0.09
    head = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x - r), Inches(y - h), Inches(2 * r), Inches(2 * r))
    head.fill.solid()
    head.fill.fore_color.rgb = color
    head.line.fill.background()
    draw_line(slide, x, y - h + 2 * r, x, y, color, 2.0)
    draw_line(slide, x, y - h * 0.45, x - 0.16, y - h * 0.15, color, 1.6)
    draw_line(slide, x, y - h * 0.45, x + 0.16, y - h * 0.15, color, 1.6)
    draw_line(slide, x, y, x - 0.12, y + 0.22, color, 1.6)
    draw_line(slide, x, y, x + 0.12, y + 0.22, color, 1.6)
    if label:
        tx(slide, label, x - 0.35, y + 0.26, 0.70, 0.28, 13, lab_color or color, True,
           PP_ALIGN.CENTER, spa=0, scaled=False)


def ground_line(slide, x1, x2, y, color=MUTED, label="ground"):
    draw_line(slide, x1, y, x2, y, color, 2.0)
    for i in range(8):
        gx = x1 + 0.25 + i * ((x2 - x1 - 0.4) / 7)
        draw_line(slide, gx, y, gx - 0.12, y + 0.14, color, 1.2)
    tx(slide, label, (x1 + x2) / 2 - 0.5, y + 0.16, 1.0, 0.26, 12, color, True,
       PP_ALIGN.CENTER, spa=0, scaled=False)


def rounded_box(slide, x, y, w, h, fill, border, pt=1.5):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.color.rgb = border
    s.line.width = Pt(pt)
    return s


def number_line(slide, x, y, w, xmin, xmax, ticks=None):
    """Horizontal number line; returns function pos -> screen x."""
    draw_line(slide, x, y, x + w, y, MUTED, 2.0)
    arrow_line(slide, x + w - 0.15, y, x + w + 0.05, y, MUTED, 1.6)

    def sx(v):
        return x + (v - xmin) * w / (xmax - xmin)

    ticks = ticks if ticks is not None else list(range(xmin, xmax + 1))
    for v in ticks:
        px = sx(v)
        draw_line(slide, px, y - 0.08, px, y + 0.08, MUTED, 1.2)
        tx(slide, str(v), px - 0.18, y + 0.10, 0.36, 0.22, 11, MUTED, True,
           PP_ALIGN.CENTER, spa=0, scaled=False)
    return sx


def mark_point(slide, px, y, label, color=GOLD, above=True):
    small_dot(slide, px, y, color)
    dy = -0.32 if above else 0.18
    tx(slide, label, px - 0.35, y + dy, 0.70, 0.28, 13, color, True,
       PP_ALIGN.CENTER, spa=0, scaled=False)


def path_polyline(slide, points, color=GOLD, pt=2.6):
    for (x1, y1), (x2, y2) in zip(points, points[1:]):
        arrow_line(slide, x1, y1, x2, y2, color, pt)


# ===== COVER =====
slide = prs.slides.add_slide(BLANK)
_bg(slide)
tx_3d(slide, "EXCELLENCE", XL, 0.52, CW, 1.10, 62, GOLD, RGBColor(210, 210, 220))
tx_3d(slide, "CLASS 11   -   2026-27", XL, 1.64, CW, 0.58, 30, WHITE, RGBColor(90, 90, 100), font="Aptos")
_solid(slide, 4.15, 2.28, 5.00, 0.028, WHITE)
bx, by, bw = 2.05, 2.70, 9.20
_card_bg(slide, bx, by, bw, 2.35, WHITE, WHITE, 0)
_card_bg(slide, bx, by + 0.72, bw, 1.63, BLACK, RGBColor(220, 220, 220), 1.25)
_solid(slide, bx + bw / 2, by + 0.16, 0.018, 0.42, RGBColor(40, 40, 40))
tx(slide, "Physics", bx, by + 0.14, bw / 2 - 0.10, 0.50, 25, BLACK, True, PP_ALIGN.CENTER, spa=0, scaled=False)
tx(slide, "Motion 01", bx + bw / 2, by + 0.14, bw / 2, 0.50, 25, BLACK, True, PP_ALIGN.CENTER, spa=0, scaled=False)
tx(slide, "Motion in a Straight Line", bx, by + 1.02, bw, 0.55, 28, WHITE, True, PP_ALIGN.CENTER, spa=0, scaled=False, font="Georgia")
tx(slide, "Frame of Reference · Rest & Motion · Distance & Displacement", bx, by + 1.62, bw, 0.40, 14, GOLD, True, PP_ALIGN.CENTER, spa=0, scaled=False)
tx(slide, "By - Excellence Programme", XL, 5.28, CW, 0.50, 25, WHITE, True, PP_ALIGN.CENTER, spa=0, scaled=False)
tx(slide, "20 October 2026", XL, 5.80, CW, 0.32, 18, MUTED, True, PP_ALIGN.CENTER, spa=0, scaled=False)
_footer(slide, 1)
add_notes(slide, "From PW Motion in a Straight Line Lecture 01. Full figures + explanations.")
after_content_slide()

slide = prs.slides.add_slide(BLANK)
_bg(slide)
tx(slide, "What we learn today", XL, 0.35, CW, 0.64, 32, GOLD, True, spa=0, scaled=False)
_solid(slide, XL, 1.00, 3.40, 0.035, GOLD)
topics = [
    "Frame of reference and the observer",
    "Position → rest and motion (with x–t graphs)",
    "Distance vs displacement — definitions and figures",
    "Rules: distance ≥ |displacement|",
    "Number-line, plane-path and 3-D examples",
    "Zero cases, square park, board MCQs",
]
for i, t in enumerate(topics):
    y = 1.20 + i * 0.72
    circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(XL), Inches(y), Inches(0.42), Inches(0.42))
    circ.fill.solid()
    circ.fill.fore_color.rgb = TEAL
    circ.line.fill.background()
    tx(slide, f"{i + 1:02d}", XL, y + 0.05, 0.42, 0.32, 14, WHITE, True, PP_ALIGN.CENTER, spa=0, scaled=False)
    tx(slide, t, XL + 0.60, y + 0.05, 10.5, 0.35, 19, TEXT, False, spa=0, scaled=False)
_footer(slide, 2)
after_content_slide()

slide = new_slide("Class note", "How to study this lecture", TEAL)
para(slide, "Every idea comes with a picture. First look at the figure, then read the boxed rule, then try the green solution steps without peeking.", XL, Y0, CW, 1.20, 16)
para(slide, "We start Motion in a Straight Line after Vectors. Today is only foundations: frame, rest–motion, distance and displacement. Speed and acceleration come later.", XL, Y0 + 1.40, CW, 1.35, 16)
ncert(slide, "Class notes come after teaching. These slides are for the board.", XL, Y0 + 3.00, CW, 0.95, 15)

# ===== FRAME =====
slide = new_slide("Idea", "Why “distance of Delhi?” fails", ORANGE)
para(slide, "Someone asks: “What is the distance of Delhi railway station?” Without saying from where we measure, the question has no meaning.", XL, Y0, 6.4, 1.35, 16)
ncert(slide, "Every length needs a starting place — a frame of reference", XL, Y0 + 1.55, 6.4, 1.05, 14)
bullet(slide, "Wrong question → reference not given", XL, Y0 + 2.85, 6.4, 0.50, 15)
bullet(slide, "Better: distance of Delhi from Patna along a named path", XL, Y0 + 3.40, 6.4, 0.55, 15)
# mini map figure
rounded_box(slide, 8.0, 1.45, 4.4, 4.4, RGBColor(14, 16, 22), TEAL, 1.4)
tx(slide, "Need a start point", 8.2, 1.65, 4.0, 0.35, 14, GOLD, True, PP_ALIGN.CENTER, spa=0, scaled=False)
small_dot(slide, 9.2, 4.6, GREEN)
tx(slide, "You (here)", 8.5, 4.85, 1.5, 0.28, 12, GREEN, True, spa=0, scaled=False)
small_dot(slide, 11.4, 2.4, ORANGE)
tx(slide, "Delhi", 11.0, 2.05, 1.2, 0.28, 12, ORANGE, True, spa=0, scaled=False)
arrow_line(slide, 9.35, 4.45, 11.2, 2.6, GOLD, 2.2)
tx(slide, "path?", 9.9, 3.2, 1.0, 0.28, 13, MUTED, True, spa=0, scaled=False)

slide = new_slide("Idea", "Frame of reference — the stand we stand on", BLUE)
para(slide, "A frame of reference is the place (and the axes) from which we take the observation. The observer is the person who takes that observation.", XL, Y0, CW, 1.10, 16)
# lift figure
rounded_box(slide, 1.2, 2.55, 4.6, 3.5, RGBColor(14, 16, 22), GREEN, 2.0)
tx(slide, "Lift = frame", 1.4, 2.70, 4.2, 0.35, 15, GREEN, True, PP_ALIGN.CENTER, spa=0, scaled=False)
stick(slide, 3.5, 5.0, GOLD, 0.70, "observer", GOLD)
tx(slide, "Observer rides with the lift", 1.4, 5.55, 4.2, 0.30, 12, MUTED, True, PP_ALIGN.CENTER, spa=0, scaled=False)
# ground frame
rounded_box(slide, 7.3, 2.55, 5.0, 3.5, RGBColor(14, 16, 22), ORANGE, 2.0)
tx(slide, "Ground = another frame", 7.5, 2.70, 4.6, 0.35, 15, ORANGE, True, PP_ALIGN.CENTER, spa=0, scaled=False)
ground_line(slide, 7.6, 11.9, 5.35)
stick(slide, 9.6, 5.15, WHITE, 0.65, "on ground", WHITE)

slide = new_slide("Figure", "Bus at 25 m/s — who is at rest?", TEAL)
para(slide, "The bus moves to the right at 25 m/s relative to the ground. Passengers A, B, C sit inside. Person D stands on the ground.", XL, Y0, CW, 0.85, 15)
# bus body
rounded_box(slide, 3.2, 2.30, 7.6, 2.55, RGBColor(18, 22, 30), BLUE, 2.2)
tx(slide, "BUS", 3.4, 2.40, 1.2, 0.30, 14, BLUE, True, spa=0, scaled=False)
arrow_line(slide, 10.5, 3.40, 12.2, 3.40, GOLD, 2.8)
tx(slide, "v = 25 m/s", 10.6, 2.95, 2.0, 0.30, 14, GOLD, True, spa=0, scaled=False)
stick(slide, 4.6, 4.15, RED, 0.55, "A", RED)
stick(slide, 6.6, 4.15, RED, 0.55, "B", RED)
stick(slide, 8.6, 4.15, GREEN, 0.55, "C", GREEN)
ground_line(slide, 1.0, 12.5, 5.55)
stick(slide, 1.8, 5.35, WHITE, 0.55, "D (ground)", WHITE)
ncert(slide, "A is at rest with respect to A. Observer always assumes he is at rest w.r.t. himself.", XL, Y0 + 4.55, CW, 0.85, 13)

slide = new_slide("Rule", "Rest and motion are relative", GREEN)
para(slide, "Relative to passenger A, B may be at rest (both sit still in the bus). Relative to ground observer D, both A and B move at 25 m/s with the bus.", XL, Y0, CW, 1.25, 16)
bullet(slide, "Same object can be “at rest” in one frame and “in motion” in another", XL, Y0 + 1.45, CW, 0.55, 16)
bullet(slide, "Never say rest or motion without naming the frame", XL, Y0 + 2.10, CW, 0.55, 16)
ncert(slide, "Nothing is absolute for rest and motion in this universe", XL, Y0 + 2.90, CW, 1.00, 15)
trap_card(slide, "“When will Delhi come?” — from the bus frame, Delhi approaches you. Physics allows that language.", XL, Y0 + 4.15, CW, 0.70)

# ===== POSITION / REST / MOTION =====
slide = new_slide("Idea", "Position", BLUE)
para(slide, "Position means the location of an object with respect to a chosen frame of reference. On a number line we write it as a coordinate (with a unit).", XL, Y0, 6.5, 1.35, 16)
ncert(slide, "Position = location of object w.r.t. the frame", XL, Y0 + 1.55, 6.5, 0.95, 15)
sx = number_line(slide, 1.2, 4.6, 11.0, -2, 6, list(range(-2, 7)))
mark_point(slide, sx(0), 4.6, "O", WHITE)
mark_point(slide, sx(3), 4.6, "object at x = 3 m", GOLD)
tx(slide, "Frame axis (+x →)", 9.2, 5.35, 3.0, 0.30, 13, MUTED, True, spa=0, scaled=False)

slide = new_slide("Idea", "Rest and motion from position", ORANGE)
para(slide, "Watch the position as time passes in the same frame.", XL, Y0, CW, 0.70, 16)
mini_card(slide, "Rest", "Position does not change with time.\nx stays the same.", XL, Y0 + 0.90, 5.7, 1.55, GREEN)
mini_card(slide, "Motion", "Position changes with time.\nx moves to a new value.", XL + 6.2, Y0 + 0.90, 5.7, 1.55, ORANGE)
ncert(slide, "Rest & motion are relative to the frame of reference", XL, Y0 + 2.80, CW, 1.00, 15)

slide = new_slide("Graph", "x–t picture: rest vs motion", GREEN)
para(slide, "On a position–time graph, rest is a horizontal line. Motion is a rising or falling line (or curve).", XL, Y0, CW, 0.75, 15)
sx, sy = graph_axes(slide, 1.0, 2.15, 5.0, 3.6, 0, 4, 0, 12, "t", "x")
axis_ticks(slide, sx, sy, [1, 2, 3], [5, 10])
draw_line(slide, sx(0), sy(10), sx(3.6), sy(10), GREEN, 2.6)
tx(slide, "Rest: x = 10 m", 1.15, 2.25, 2.6, 0.28, 13, GREEN, True, spa=0, scaled=False)
sx2, sy2 = graph_axes(slide, 7.3, 2.15, 5.0, 3.6, 0, 4, 0, 12, "t", "x")
axis_ticks(slide, sx2, sy2, [1, 2, 3], [5, 10])
draw_line(slide, sx2(0), sy2(2), sx2(3.6), sy2(11), GOLD, 2.6)
tx(slide, "Motion: x changes", 7.45, 2.25, 2.8, 0.28, 13, GOLD, True, spa=0, scaled=False)

# ===== DISTANCE VS DISPLACEMENT =====
slide = new_slide("Idea", "Distance — the path you really walked", ORANGE)
para(slide, "Distance is the length of the actual path covered by the object. A vehicle odometer reading is distance.", XL, Y0, 6.3, 1.15, 16)
bullet(slide, "Scalar — only size, no direction", XL, Y0 + 1.35, 6.3, 0.45, 15)
bullet(slide, "Unit: metre (m)", XL, Y0 + 1.85, 6.3, 0.40, 15)
bullet(slide, "Cannot be negative", XL, Y0 + 2.35, 6.3, 0.40, 15)
bullet(slide, "Keeps increasing while the object moves", XL, Y0 + 2.85, 6.3, 0.45, 15)
bullet(slide, "Depends on the path between two points", XL, Y0 + 3.35, 6.3, 0.45, 15)
trap_card(slide, "If the path is unknown, you cannot calculate distance.", XL, Y0 + 4.05, 6.3, 0.60)
# odometer-style card
rounded_box(slide, 8.0, 1.50, 4.4, 4.6, RGBColor(14, 16, 22), ORANGE, 1.5)
tx(slide, "Odometer idea", 8.2, 1.70, 4.0, 0.35, 15, ORANGE, True, PP_ALIGN.CENTER, spa=0, scaled=False)
tx(slide, "Adds every metre of path", 8.2, 2.20, 4.0, 0.40, 14, TEXT, False, PP_ALIGN.CENTER, spa=0, scaled=False)
tx(slide, "never goes backwards", 8.2, 2.70, 4.0, 0.35, 14, MUTED, True, PP_ALIGN.CENTER, spa=0, scaled=False)
# winding path sketch
pts = [(8.6, 5.3), (9.3, 4.4), (10.2, 5.0), (11.0, 3.9), (11.6, 4.8)]
path_polyline(slide, pts, GOLD, 2.2)
tx(slide, "long path = large distance", 8.3, 5.55, 3.8, 0.30, 12, GOLD, True, PP_ALIGN.CENTER, spa=0, scaled=False)

slide = new_slide("Idea", "Displacement — the net shift", TEAL)
para(slide, "Displacement is how much the position shifted, and in which direction. It is the shortest directed jump from start to finish.", XL, Y0, CW, 1.05, 16)
ncert(slide, "Δr = r_f − r_i", XL, Y0 + 1.25, CW, 0.90, 20)
bullet(slide, "Vector — size and direction; unit metre", XL, Y0 + 2.40, CW, 0.45, 15)
bullet(slide, "Does not depend on the path taken", XL, Y0 + 2.90, CW, 0.45, 15)
bullet(slide, "Can be positive or negative on a line (sign = direction)", XL, Y0 + 3.40, CW, 0.50, 15)
bullet(slide, "Can decrease even while the object is still moving", XL, Y0 + 3.90, CW, 0.50, 15)
trap_card(slide, "If initial and final positions are the same → displacement = 0", XL, Y0 + 4.55, CW, 0.55)

slide = new_slide("Figure", "Path (distance) vs straight jump (displacement)", GREEN)
para(slide, "Same start A and end B. The winding route is distance. The straight arrow is displacement.", XL, Y0, CW, 0.75, 15)
# draw A, B
ax, ay, bx_, by_ = 2.5, 5.4, 10.5, 2.6
small_dot(slide, ax, ay, WHITE)
tx(slide, "A  (start)", ax - 0.6, ay + 0.25, 1.4, 0.28, 13, WHITE, True, spa=0, scaled=False)
small_dot(slide, bx_, by_, WHITE)
tx(slide, "B  (end)", bx_ - 0.3, by_ - 0.40, 1.4, 0.28, 13, WHITE, True, spa=0, scaled=False)
# winding path
wpath = [(ax, ay), (3.8, 3.2), (5.5, 5.0), (7.2, 2.9), (8.8, 4.6), (bx_, by_)]
path_polyline(slide, wpath, ORANGE, 2.4)
tx(slide, "distance = length of orange path", 3.5, 5.85, 5.5, 0.30, 14, ORANGE, True, spa=0, scaled=False)
# displacement
arrow_line(slide, ax, ay, bx_, by_, GREEN, 3.0)
tx(slide, "displacement", 5.8, 3.6, 2.2, 0.30, 14, GREEN, True, spa=0, scaled=False)

slide = new_slide("Rule", "Three inequalities you must own", BLUE)
para(slide, "In one dimension, if the object never turns back, the size of displacement equals the distance. In every case, distance cannot be smaller than |displacement|.", XL, Y0, CW, 1.15, 15)
fml(slide, "No change of direction (1-D):   |displacement| = distance", XL, Y0 + 1.35, CW, LH, 16)
fml(slide, "Always (any path):   distance ≥ |displacement|", XL, Y0 + 1.35 + LH + NG, CW, LH, 17)
fml(slide, "distance / |displacement|  ≥  1", XL, Y0 + 1.35 + 2 * (LH + NG), CW, LH, 17)
trap_card(slide, "|displacement| > distance is never possible.", XL, Y0 + 3.55, CW, 0.55)

# ===== MCQs inequalities =====
slide = question_slide("MCQ — motion in a straight line")
qtx(slide, "Which option is correct for motion in a straight line (1-D)?", XL, Y0, CW, 0.55, 17)
choices(slide, [
    "distance = |displacement|",
    "distance > |displacement|",
    "distance ≥ |displacement|",
    "distance < |displacement|",
], XL, Y0 + 0.80)

slide = answer_slide()
body(slide, "Idea: without a turn, equal; with a turn, distance is larger. Both sit under ≥.", XL, Y0, CW, 0.55)
fml(slide, "Correct: distance ≥ |displacement|", XL, Y0 + 0.80, CW, LH, 17)
result_card(slide, "Answer: option 3", XL, Y0 + 1.60, CW, 0.40, 18)

slide = question_slide("MCQ — numerical ratio")
qtx(slide, "The numerical ratio of distance to |displacement| is always", XL, Y0, CW, 0.65, 17)
choices(slide, [
    "less than one",
    "equal to one",
    "equal to or less than one",
    "equal to or greater than one",
], XL, Y0 + 0.90)

slide = answer_slide()
fml(slide, "distance / |displacement|  ≥  1", XL, Y0, CW, LH, 18)
result_card(slide, "Answer: equal to or greater than one", XL, Y0 + 0.95, CW, 0.45, 17)

slide = question_slide("MCQ — 1-D without change of direction")
qtx(slide, "For 1-D motion without change in direction, which is correct?", XL, Y0, CW, 0.60, 17)
choices(slide, [
    "distance = |displacement|",
    "distance > |displacement|",
    "distance ≥ |displacement|",
    "distance < |displacement|",
], XL, Y0 + 0.85)

slide = answer_slide()
body(slide, "Idea: the object never turns, so path length equals net shift size.", XL, Y0, CW, 0.55)
result_card(slide, "Answer: distance = |displacement|", XL, Y0 + 0.85, CW, 0.45, 17)

slide = question_slide("MCQ — 1-D with change of direction")
qtx(slide, "For 1-D motion with a change of direction, which is correct?", XL, Y0, CW, 0.60, 17)
# small turn figure
sx = number_line(slide, 7.5, 2.0, 5.0, 0, 6, [0, 2, 4, 6])
mark_point(slide, sx(1), 2.0, "A", GREEN)
mark_point(slide, sx(5), 2.0, "turn", ORANGE, above=False)
mark_point(slide, sx(2), 2.0, "B", GOLD)
arrow_line(slide, sx(1), 1.75, sx(5), 1.75, ORANGE, 2.0)
arrow_line(slide, sx(5), 1.55, sx(2), 1.55, ORANGE, 2.0)
choices(slide, [
    "distance = |displacement|",
    "distance > |displacement|",
    "distance ≥ |displacement|",
    "distance < |displacement|",
], XL, Y0 + 0.85, w=6.2)

slide = answer_slide()
result_card(slide, "Answer: distance > |displacement|", XL, Y0, CW, 0.45, 18)
trap_card(slide, "Because it turned, the path is longer than the straight jump.", XL, Y0 + 0.75, CW, 0.55)

# ===== NUMBER LINE EXAMPLES =====
slide = new_slide("Method", "Number-line method", TEAL)
para(slide, "Mark start and end on a line. Distance = sum of every segment you walk. Displacement = final coordinate − initial coordinate (as a vector along î).", XL, Y0, CW, 1.15, 15)
ncert(slide, "displacement = r_f − r_i    →    on a line: (x_f − x_i) î", XL, Y0 + 1.35, CW, 1.00, 15)  # subscripts via _format
# example 1 figure
sx = number_line(slide, 1.5, 4.3, 10.5, 0, 6, list(range(0, 7)))
mark_point(slide, sx(1), 4.3, "A", GREEN)
mark_point(slide, sx(5), 4.3, "B", GOLD)
arrow_line(slide, sx(1), 3.95, sx(5), 3.95, GOLD, 2.4)
fml(slide, "① A at +1 → B at +5:   distance = 4 m ,   displacement = +4 î", XL, Y0 + 4.40, CW, LH, 15)

slide = new_slide("Example", "With a turn on the line", GREEN)
para(slide, "Start at +1 m, go to +4 m, then reverse and stop at −3 m.", XL, Y0, CW, 0.70, 15)
sx = number_line(slide, 1.2, 2.55, 11.0, -4, 5, list(range(-4, 6)))
mark_point(slide, sx(1), 2.55, "start", GREEN)
mark_point(slide, sx(4), 2.55, "turn", ORANGE, above=False)
mark_point(slide, sx(-3), 2.55, "end", GOLD)
arrow_line(slide, sx(1), 2.20, sx(4), 2.20, ORANGE, 2.2)
arrow_line(slide, sx(4), 1.95, sx(-3), 1.95, ORANGE, 2.2)
arrow_line(slide, sx(1), 3.05, sx(-3), 3.05, GREEN, 2.6)
tx(slide, "displacement", 4.5, 3.15, 2.0, 0.28, 13, GREEN, True, spa=0, scaled=False)
fml(slide, "distance = 3 + 7 = 10 m", XL, Y0 + 3.35, CW, LH, 16)
fml(slide, "displacement = (−3)î − (1)î = −4 î", XL, Y0 + 3.35 + LH + NG, CW, LH, 16)

slide = new_slide("Example", "Two more line drills", BLUE)
fml(slide, "③ Start +2 → go to −4 → end −1", XL, Y0, CW, LH, 15)
fml(slide, "distance = 6 + 3 = 9 m ,   displacement = (−1 − 2)î = −3 î", XL, Y0 + LH + NG, CW, LH, 14)
sx = number_line(slide, 1.5, 3.15, 10.5, -5, 3, list(range(-5, 4)))
mark_point(slide, sx(2), 3.15, "A", GREEN)
mark_point(slide, sx(-4), 3.15, "", ORANGE, above=False)
mark_point(slide, sx(-1), 3.15, "B", GOLD)
fml(slide, "④ Start −3 → go to +4 → end −4", XL, Y0 + 2.55, CW, LH, 15)
fml(slide, "distance = 7 + 8 = 15 m ,   displacement = (−4 − (−3))î = −1 î", XL, Y0 + 2.55 + LH + NG, CW, LH, 14)

slide = new_slide("Example", "3-D shift of position", ORANGE)
qtx(slide, "Initial (2, 4, 7), final (5, 8, 12). Find displacement.", XL, Y0, CW, 0.55, 16)
# 3D-ish arrow card
rounded_box(slide, 8.2, 1.9, 4.3, 3.2, RGBColor(14, 16, 22), GOLD, 1.5)
small_dot(slide, 9.0, 4.4, GREEN)
tx(slide, "r_i", 8.6, 4.55, 0.8, 0.25, 12, GREEN, True, spa=0, scaled=False)
small_dot(slide, 11.4, 2.5, ORANGE)
tx(slide, "r_f", 11.2, 2.15, 0.8, 0.25, 12, ORANGE, True, spa=0, scaled=False)
arrow_line(slide, 9.1, 4.25, 11.25, 2.65, GOLD, 2.8)
tx(slide, "Δr", 9.8, 3.2, 0.8, 0.28, 14, GOLD, True, spa=0, scaled=False)
fml(slide, "Δr = (5−2)î + (8−4)ĵ + (12−7)k̂ = 3î + 4ĵ + 5k̂", XL, Y0 + 0.80, 7.2, LH, 14)
fml(slide, "|Δr| = √(9 + 16 + 25) = √50 = 5√2 m", XL, Y0 + 0.80 + LH + NG, 7.2, LH, 15)
result_card(slide, "Displacement 3î + 4ĵ + 5k̂ ; magnitude 5√2 m", XL, Y0 + 2.40, 7.2, 0.55, 14)

slide = new_slide("Figure", "Broken path A → B on a plane", TEAL)
para(slide, "Walk: 4 m right, 5 m down, 1 m left, 2 m up. Find distance and displacement.", XL, Y0, CW, 0.70, 15)
# path drawing
ox, oy = 3.5, 2.4
# scale ~0.45 per metre
s = 0.42
pts = [
    (ox, oy),
    (ox + 4 * s, oy),
    (ox + 4 * s, oy + 5 * s),
    (ox + 3 * s, oy + 5 * s),
    (ox + 3 * s, oy + 3 * s),
]
path_polyline(slide, pts, ORANGE, 2.6)
small_dot(slide, pts[0][0], pts[0][1], GREEN)
tx(slide, "A", pts[0][0] - 0.45, pts[0][1] - 0.15, 0.4, 0.28, 14, GREEN, True, spa=0, scaled=False)
small_dot(slide, pts[-1][0], pts[-1][1], GOLD)
tx(slide, "B", pts[-1][0] + 0.15, pts[-1][1] - 0.15, 0.4, 0.28, 14, GOLD, True, spa=0, scaled=False)
arrow_line(slide, pts[0][0], pts[0][1], pts[-1][0], pts[-1][1], GREEN, 2.8)
tx(slide, "4 m", ox + 1.5 * s, oy - 0.35, 0.8, 0.25, 12, MUTED, True, spa=0, scaled=False)
tx(slide, "5 m", ox + 4 * s + 0.1, oy + 2.2 * s, 0.7, 0.25, 12, MUTED, True, spa=0, scaled=False)
fml(slide, "distance = 4+5+1+2 = 12 m", 7.5, 2.3, 5.0, LH, 15)
fml(slide, "net: 3 m right, 3 m down", 7.5, 2.3 + LH + NG, 5.0, LH, 14)
fml(slide, "Δr = 3î − 3ĵ", 7.5, 2.3 + 2 * (LH + NG), 5.0, LH, 15)
fml(slide, "|Δr| = 3√2 m", 7.5, 2.3 + 3 * (LH + NG), 5.0, LH, 15)

slide = new_slide("Figure", "Spiral path (segments 1 to 8)", GREEN)
para(slide, "Path lengths 1, 2, 3, 4, 5, 6, 7, 8 m in turn (right, up, left, down, …).", XL, Y0, CW, 0.70, 15)
# simplified spiral sketch
ox, oy = 4.0, 4.6
s = 0.22
dirs = [(1, 0), (0, -1), (-1, 0), (0, 1), (1, 0), (0, -1), (-1, 0), (0, 1)]
lens = [1, 2, 3, 4, 5, 6, 7, 8]
x, y = ox, oy
pts = [(x, y)]
for L, (dx, dy) in zip(lens, dirs):
    x += dx * L * s
    y += dy * L * s
    pts.append((x, y))
path_polyline(slide, pts, ORANGE, 2.2)
small_dot(slide, pts[0][0], pts[0][1], GREEN)
tx(slide, "A", pts[0][0] - 0.4, pts[0][1] + 0.1, 0.4, 0.25, 13, GREEN, True, spa=0, scaled=False)
small_dot(slide, pts[-1][0], pts[-1][1], GOLD)
tx(slide, "B", pts[-1][0] + 0.15, pts[-1][1], 0.4, 0.25, 13, GOLD, True, spa=0, scaled=False)
arrow_line(slide, pts[0][0], pts[0][1], pts[-1][0], pts[-1][1], GREEN, 2.5)
fml(slide, "distance = 1+2+…+8 = 36 m", 8.0, 2.2, 4.6, LH, 15)
fml(slide, "net forms 4–4 right triangle", 8.0, 2.2 + LH + NG, 4.6, LH, 14)
fml(slide, "|Δr| = √(16+16) = 4√2 m", 8.0, 2.2 + 2 * (LH + NG), 4.6, LH, 15)

# ===== ZERO LOGIC =====
slide = new_slide("Rule", "Zero distance / zero displacement", ORANGE)
grid_table(
    slide,
    ["If …", "Then …"],
    [
        ["distance = 0", "displacement must be 0 (did not move)"],
        ["distance ≠ 0", "displacement may be 0 or not"],
        ["displacement = 0", "distance may be 0 or not"],
        ["displacement ≠ 0", "distance must be > 0"],
    ],
    XL, Y0, [4.0, 7.8], row_h=0.55, hsize=14, csize=14,
)
# round trip mini figure
rounded_box(slide, 8.8, 4.0, 3.7, 2.2, RGBColor(14, 16, 22), TEAL, 1.2)
small_dot(slide, 10.4, 5.4, GOLD)
tx(slide, "A = B", 9.9, 5.55, 1.2, 0.25, 12, GOLD, True, PP_ALIGN.CENTER, spa=0, scaled=False)
# loop
draw_line(slide, 10.4, 5.4, 11.5, 4.5, ORANGE, 2.0)
draw_line(slide, 11.5, 4.5, 10.4, 4.3, ORANGE, 2.0)
draw_line(slide, 10.4, 4.3, 9.3, 4.5, ORANGE, 2.0)
draw_line(slide, 9.3, 4.5, 10.4, 5.4, ORANGE, 2.0)
tx(slide, "disp = 0, distance > 0", 8.95, 5.85, 3.4, 0.25, 12, ORANGE, True, PP_ALIGN.CENTER, spa=0, scaled=False)

slide = new_slide("Question", "True / False — lock these", ORANGE)
choices(slide, [
    "Distance must equal |displacement| in every 1-D motion.  → False",
    "If displacement is zero, distance must be zero.  → False",
    "If distance is zero, displacement must be zero.  → True",
    "|displacement| is always equal to distance.  → False",
], XL, Y0)
para(slide, "Only when direction never changes in 1-D can we say |displacement| equals distance.", XL, Y0 + 2.70, CW, 1.05, 15)

# ===== SQUARE + MORE =====
slide = new_slide("Figure", "Square park of side ℓ", BLUE)
para(slide, "Start at corner A of square ABCD. Walk on the boundary. Side length = ℓ.", XL, Y0, CW, 0.65, 15)
# square
sq = 2.6
ax, ay = 2.2, 5.5
corners = [(ax, ay), (ax + sq, ay), (ax + sq, ay - sq), (ax, ay - sq), (ax, ay)]
for (x1, y1), (x2, y2) in zip(corners, corners[1:]):
    draw_line(slide, x1, y1, x2, y2, WHITE, 2.0)
arrow_line(slide, ax, ay, ax + sq, ay - sq, GOLD, 2.4)
tx(slide, "A", ax - 0.35, ay + 0.05, 0.35, 0.25, 13, GREEN, True, spa=0, scaled=False)
tx(slide, "B", ax + sq + 0.08, ay + 0.05, 0.35, 0.25, 13, WHITE, True, spa=0, scaled=False)
tx(slide, "C", ax + sq + 0.08, ay - sq - 0.30, 0.35, 0.25, 13, WHITE, True, spa=0, scaled=False)
tx(slide, "D", ax - 0.35, ay - sq - 0.30, 0.35, 0.25, 13, WHITE, True, spa=0, scaled=False)
tx(slide, "diagonal AC", ax + 0.7, ay - 1.5, 1.5, 0.25, 12, GOLD, True, spa=0, scaled=False)
grid_table(
    slide,
    ["Path", "distance", "|disp|"],
    [
        ["A → B", "ℓ", "ℓ"],
        ["A → C (via B)", "2ℓ", "ℓ√2"],
        ["A → D (via B,C)", "3ℓ", "ℓ"],
        ["A → A (full lap)", "4ℓ", "0"],
    ],
    6.3, Y0 + 0.85, [3.3, 2.4, 2.4], row_h=0.48, hsize=13, csize=14,
)

slide = new_slide("Example", "When distance cannot be found", TEAL)
para(slide, "Particle goes from (0, 0) to (2, −4) on a curved flight path. Without the exact path length, distance is unknown — displacement is still easy.", XL, Y0, CW, 1.05, 15)
# cliff sketch
draw_line(slide, 1.5, 4.8, 5.5, 4.8, MUTED, 1.5)
draw_line(slide, 3.2, 4.8, 3.2, 2.6, MUTED, 1.5)
small_dot(slide, 2.0, 4.8, GREEN)
tx(slide, "A (0,0)", 1.5, 5.0, 1.5, 0.25, 12, GREEN, True, spa=0, scaled=False)
small_dot(slide, 4.6, 3.2, GOLD)
tx(slide, "B (2,−4)", 4.3, 2.85, 1.6, 0.25, 12, GOLD, True, spa=0, scaled=False)
# curve approx
curve(slide, [(2.0, 4.8), (2.6, 3.9), (3.4, 3.3), (4.6, 3.2)], ORANGE, 2.2)
arrow_line(slide, 2.0, 4.8, 4.6, 3.2, GREEN, 2.4)
fml(slide, "Δr = 2î − 4ĵ ,   |Δr| = 2√5", 6.5, 2.4, 6.0, LH, 15)
para(slide, "Straight drop (0,0) → (0,−18): distance = 18 m and displacement = −18 ĵ.", 6.5, 3.3, 6.0, 1.10, 14)

slide = new_slide("Figure", "L-shaped path on axes", GREEN)
para(slide, "A → B = 4 m right, B → C = 6 m up, C → D = 4 m right.", XL, Y0, CW, 0.65, 15)
ox, oy = 2.0, 5.6
s = 0.45
pts = [(ox, oy), (ox + 4 * s, oy), (ox + 4 * s, oy - 6 * s), (ox + 8 * s, oy - 6 * s)]
path_polyline(slide, pts, ORANGE, 2.6)
labs = ["A", "B", "C", "D"]
for p, lab in zip(pts, labs):
    small_dot(slide, p[0], p[1], WHITE)
    tx(slide, lab, p[0] - 0.15, p[1] + 0.15, 0.4, 0.25, 13, WHITE, True, spa=0, scaled=False)
arrow_line(slide, pts[0][0], pts[0][1], pts[-1][0], pts[-1][1], GREEN, 2.6)
tx(slide, "4 m", ox + 1.5 * s, oy + 0.15, 0.8, 0.25, 12, MUTED, True, spa=0, scaled=False)
tx(slide, "6 m", ox + 4 * s + 0.12, oy - 3 * s, 0.7, 0.25, 12, MUTED, True, spa=0, scaled=False)
tx(slide, "4 m", ox + 6 * s, oy - 6 * s - 0.35, 0.8, 0.25, 12, MUTED, True, spa=0, scaled=False)
fml(slide, "distance = 14 m", 8.0, 2.5, 4.5, LH, 16)
fml(slide, "net: 8 m right, 6 m up", 8.0, 2.5 + LH + NG, 4.5, LH, 15)
fml(slide, "|Δr| = √(64+36) = 10 m", 8.0, 2.5 + 2 * (LH + NG), 4.5, LH, 16)
result_card(slide, "distance 14 m , |displacement| 10 m", 8.0, 4.6, 4.5, 0.50, 14)

slide = question_slide("MCQ — displacement P → Q")
qtx(slide, "Particle moves from P(2, 3, 5) to Q(3, 4, 5). Displacement is", XL, Y0, CW, 0.65, 16)
choices(slide, [
    "î + ĵ + 10 k̂",
    "î + ĵ + 5 k̂",
    "î + ĵ",
    "2î + 4ĵ + 6k̂",
], XL, Y0 + 0.90)

slide = answer_slide()
fml(slide, "Δr = (3−2)î + (4−3)ĵ + (5−5)k̂ = î + ĵ", XL, Y0, CW, LH, 16)
result_card(slide, "Answer: î + ĵ", XL, Y0 + 0.95, CW, 0.40, 18)

slide = question_slide("MCQ — correct statement")
qtx(slide, "Correct statement among the following is:", XL, Y0, CW, 0.50, 17)
choices(slide, [
    "When displacement is zero, distance travelled is not zero.",
    "When displacement is zero, distance travelled is also zero.",
    "When distance is zero, displacement is not zero.",
    "Distance travelled and displacement are always equal.",
], XL, Y0 + 0.75)
trap_card(slide, "Option 1 can happen (round trip) but is not always true. Option 3 is impossible.", XL, Y0 + 3.30, CW, 0.70)

slide = answer_slide()
body(slide, "Idea: none of 2–4 are always true. The useful lock line is: displacement zero does not force distance zero.", XL, Y0, CW, 0.85)
para(slide, "Best teaching answer from the logic table: if displacement = 0, distance may or may not be zero. Statement 1 describes a possible case (return trip).", XL, Y0 + 1.05, CW, 1.20, 15)
result_card(slide, "Focus: zero displacement ⇏ zero distance", XL, Y0 + 2.50, CW, 0.50, 16)

slide = new_slide("Example", "East–South–North-West walk", BLUE)
para(slide, "Walk 10 m east, then 20 m south, then 20√2 m north-west. Find displacement.", XL, Y0, CW, 0.85, 15)
# vector tips from origin
ox, oy = 3.2, 4.8
# east 10
arrow_line(slide, ox, oy, ox + 2.0, oy, ORANGE, 2.4)
tx(slide, "10 E", ox + 0.6, oy + 0.15, 1.0, 0.25, 12, ORANGE, True, spa=0, scaled=False)
# south 20
arrow_line(slide, ox + 2.0, oy, ox + 2.0, oy + 2.0, ORANGE, 2.4)
tx(slide, "20 S", ox + 2.15, oy + 0.8, 1.0, 0.25, 12, ORANGE, True, spa=0, scaled=False)
# NW 20√2 ~ equal west and north components visually
arrow_line(slide, ox + 2.0, oy + 2.0, ox + 2.0 - 2.0, oy + 2.0 - 2.0, GOLD, 2.4)
tx(slide, "20√2 NW", ox + 0.3, oy + 1.5, 1.6, 0.25, 12, GOLD, True, spa=0, scaled=False)
# resultant west
arrow_line(slide, ox, oy, ox - 1.5, oy, GREEN, 2.8)
tx(slide, "10 W", ox - 1.5, oy - 0.40, 1.2, 0.25, 13, GREEN, True, spa=0, scaled=False)
fml(slide, "NW step = 20√2 × (−î + ĵ)/√2 = −20 î + 20 ĵ", 7.0, 2.2, 5.5, LH, 13)
fml(slide, "Net = 10î − 20ĵ − 20î + 20ĵ = −10 î", 7.0, 2.2 + LH + NG, 5.5, LH, 14)
result_card(slide, "Displacement = 10 m due west", 7.0, 4.0, 5.5, 0.50, 15)

slide = new_slide("Example", "Curved path on a slope (2,4) → (8,10)", TEAL)
para(slide, "Object moves from (2, 4) to (8, 10) along a curved path on a line. Path length may be unknown; displacement is not.", XL, Y0, CW, 0.95, 15)
sx, sy = graph_axes(slide, 1.2, 2.4, 5.5, 3.5, 0, 10, 0, 12, "x", "y")
small_dot(slide, sx(2), sy(4), GREEN)
tx(slide, "(2,4)", sx(2) - 0.5, sy(4) + 0.15, 1.0, 0.25, 11, GREEN, True, spa=0, scaled=False)
small_dot(slide, sx(8), sy(10), GOLD)
tx(slide, "(8,10)", sx(8) - 0.3, sy(10) - 0.40, 1.1, 0.25, 11, GOLD, True, spa=0, scaled=False)
curve(slide, [(sx(2), sy(4)), (sx(4), sy(7)), (sx(6), sy(8.5)), (sx(8), sy(10))], ORANGE, 2.2)
arrow_line(slide, sx(2), sy(4), sx(8), sy(10), GREEN, 2.4)
fml(slide, "Δr = (8î+10ĵ) − (2î+4ĵ) = 6î + 6ĵ", 7.3, 2.6, 5.3, LH, 14)
fml(slide, "|Δr| = 6√2", 7.3, 2.6 + LH + NG, 5.3, LH, 15)

# ===== LOCK =====
slide = new_slide("Lock lines", "Say these aloud", GREEN)
bullet(slide, "Frame first — every observation needs a reference.", XL, Y0, CW, 0.48, 16)
bullet(slide, "Observer assumes he is at rest with respect to himself.", XL, Y0 + 0.52, CW, 0.48, 16)
bullet(slide, "Position changes with time → motion; else rest (in that frame).", XL, Y0 + 1.04, CW, 0.48, 16)
bullet(slide, "Distance = path length (scalar). Displacement = r_f − r_i (vector).", XL, Y0 + 1.56, CW, 0.52, 16)
bullet(slide, "Always: distance ≥ |displacement|. Never the reverse.", XL, Y0 + 2.10, CW, 0.48, 16)
bullet(slide, "Same start and end → displacement = 0 (distance may still be large).", XL, Y0 + 2.62, CW, 0.52, 16)
bullet(slide, "No turn in 1-D → |displacement| = distance.", XL, Y0 + 3.16, CW, 0.48, 16)
ncert(slide, "Next: deeper distance–displacement practice (PW 02).", XL, Y0 + 3.85, CW, 0.90, 14)

slide = prs.slides.add_slide(BLANK)
_bg(slide)
_bar(slide, "End", GOLD)
tx_3d(slide, "Thank you", XL, 2.4, CW, 1.0, 48, GOLD, RGBColor(120, 120, 130))
tx(slide, "Excellence  ·  Motion 01  ·  Frame · Rest–Motion · Distance · Displacement", XL, 3.7, CW, 0.45, 15, MUTED, True, PP_ALIGN.CENTER, spa=0, scaled=False)
_footer(slide, len(prs.slides))
after_content_slide()

OUTPUT.parent.mkdir(parents=True, exist_ok=True)
prs.save(OUTPUT)
print(OUTPUT)
print("slides", len(prs.slides))
