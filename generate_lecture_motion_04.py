"""Motion 04 — velocity as a vector, v-t graphs, new average cases (PW 04)."""
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
OUTPUT = BASE / "lecture slides" / "Lecture_Motion_04_Velocity.pptx"

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
tx(slide, "Motion 04", bx + bw / 2, by + 0.14, bw / 2, 0.50, 25, BLACK, True, PP_ALIGN.CENTER, spa=0, scaled=False)
tx(slide, "Motion in a Straight Line", bx, by + 1.02, bw, 0.55, 28, WHITE, True, PP_ALIGN.CENTER, spa=0, scaled=False, font="Georgia")
tx(slide, "Velocity as a vector · v–t graphs · when averages fail", bx, by + 1.62, bw, 0.40, 14, GOLD, True, PP_ALIGN.CENTER, spa=0, scaled=False)
tx(slide, "By - Excellence Programme", XL, 5.28, CW, 0.50, 25, WHITE, True, PP_ALIGN.CENTER, spa=0, scaled=False)
tx(slide, "22 October 2026", XL, 5.80, CW, 0.32, 18, MUTED, True, PP_ALIGN.CENTER, spa=0, scaled=False)
_footer(slide, 1)
add_notes(slide, "PW Motion 04. Skips speed recipes already taught in Motion 03.")
after_content_slide()

slide = prs.slides.add_slide(BLANK)
_bg(slide)
tx(slide, "What is new today", XL, 0.35, CW, 0.64, 32, GOLD, True, spa=0, scaled=False)
_solid(slide, XL, 1.00, 3.40, 0.035, GOLD)
topics = [
    "Instantaneous velocity, and how it differs from speed",
    "What a graph’s slope and area actually give",
    "v–t: a kink is not a U-turn",
    "v = 0 with no reversal",
    "Average velocity along the chord in 2-D",
    "One nested average, and one impossible average",
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

slide = new_slide("Class note", "Left in Motion 03", TEAL)
para(slide, "Average speed, average velocity, the time-mix formula, the distance-mix formula, 4v/3, the 40 km/h round trip, and the three-thirds 16.4 m/s problem stay in lecture 03. We will not calculate them again.", XL, Y0, CW, 1.40, 16)
para(slide, "Today the new job is direction: when velocity changes, what a v–t graph is allowed to say, and when an average cannot exist.", XL, Y0 + 1.60, CW, 1.15, 16)

# ===== INSTANT =====
slide = new_slide("Idea", "Instantaneous velocity", ORANGE)
para(slide, "Average velocity looks at the whole trip. Instantaneous velocity is the velocity at one moment. In one dimension it is the slope of the x–t graph at that instant.", XL, Y0, CW, 1.20, 16)
ncert(slide, "v = dx/dt     ·     instantaneous speed = |v|", XL, Y0 + 1.40, CW, 0.95, 18)
bullet(slide, "Differentiate x(t) when you want the instant, not the average.", XL, Y0 + 2.55, CW, 0.55, 16)
bullet(slide, "If the instantaneous velocity stays constant, the average velocity equals that same vector.", XL, Y0 + 3.20, CW, 0.70, 16)

slide = new_slide("Rule", "Speed fixed does not fix velocity", BLUE)
para(slide, "Speed is only the size. Velocity also carries direction. A particle can keep the same speed and still change its velocity by turning.", XL, Y0, CW, 1.15, 16)
fml(slide, "constant velocity  →  constant speed", XL, Y0 + 1.40, CW, LH, 17)
fml(slide, "constant speed  →  velocity may still change", XL, Y0 + 1.40 + LH + NG, CW, LH, 17)
fml(slide, "variable speed  →  velocity is variable", XL, Y0 + 1.40 + 2 * (LH + NG), CW, LH, 17)
trap_card(slide, "Uniform circular motion is the clean picture: speed stays v, velocity keeps turning.", XL, Y0 + 3.40, CW, 0.65)

slide = new_slide("Rule", "|average velocity| cannot beat average speed", GREEN)
para(slide, "Distance is at least the straight shift, so the speed average is at least the size of the velocity average.", XL, Y0, CW, 1.05, 16)
ncert(slide, "|average velocity|  ≤  average speed", XL, Y0 + 1.25, CW, 0.95, 18)
para(slide, "They are equal only in one dimension when the object never reverses. Any turn, or any detour, makes the speed average strictly larger.", XL, Y0 + 2.45, CW, 1.15, 16)

slide = new_slide("Rule", "Zero links — speed and velocity", TEAL)
grid_table(
    slide,
    ["If this is zero", "Then the other"],
    [
        ["average speed = 0", "average velocity must be 0"],
        ["average velocity = 0", "average speed may still be > 0"],
        ["average velocity ≠ 0", "average speed must be > 0"],
    ],
    XL, Y0, [5.2, 6.6], row_h=0.70, hsize=15, csize=15,
)
trap_card(slide, "A closed loop has average velocity zero and average speed greater than zero.", XL, Y0 + 3.00, CW, 0.60)

# ===== GRAPHS =====
slide = new_slide("Method", "Four readings from one graph", ORANGE)
para(slide, "Before a new graph, name the axes. Then only four things are available.", XL, Y0, CW, 0.85, 16)
bullet(slide, "The y-reading is the quantity on the vertical axis at that instant.", XL, Y0 + 1.05, CW, 0.55, 16)
bullet(slide, "The x-reading is the quantity on the horizontal axis.", XL, Y0 + 1.65, CW, 0.50, 16)
bullet(slide, "Slope = Δy / Δx. On an x–t graph the slope is velocity.", XL, Y0 + 2.25, CW, 0.55, 16)
bullet(slide, "Area under a v–t graph equals displacement (watch the sign below the axis).", XL, Y0 + 2.90, CW, 0.65, 16)

slide = new_slide("Figure", "A kink on v–t is not a U-turn", BLUE)
para(slide, "Direction of motion changes only when velocity crosses zero. A sharp peak or a sharp valley, while v stays positive, is only a change of acceleration.", XL, Y0, CW, 1.10, 15)
sx, sy = graph_axes(slide, 1.4, 2.55, 10.2, 3.4, 0, 10, 0, 8, "t", "v")
pts = [(sx(0), sy(2)), (sx(4), sy(7)), (sx(7), sy(1.5)), (sx(9.5), sy(4))]
path_polyline(slide, pts, GOLD, 2.4)
tx(slide, "A", sx(0) - 0.3, sy(2) - 0.35, 0.4, 0.25, 13, WHITE, True, spa=0, scaled=False)
tx(slide, "B", sx(4) - 0.15, sy(7) - 0.40, 0.4, 0.25, 13, ORANGE, True, spa=0, scaled=False)
tx(slide, "C", sx(7) - 0.15, sy(1.5) + 0.15, 0.4, 0.25, 13, ORANGE, True, spa=0, scaled=False)
tx(slide, "D", sx(9.5) + 0.08, sy(4), 0.4, 0.25, 13, WHITE, True, spa=0, scaled=False)
result_card(slide, "v never crosses 0  →  direction of motion does not change", XL, Y0 + 4.35, CW, 0.45, 14)

slide = new_slide("Rule", "Where a real U-turn sits", GREEN)
para(slide, "On a v–t graph, mark only the crossings of the time axis. Those are the instants when velocity changes sign. A maximum or a minimum of speed, away from the axis, is not a turn-around.", XL, Y0, CW, 1.25, 16)
sx, sy = graph_axes(slide, 1.5, 2.7, 10.0, 3.3, 0, 10, -4, 4, "t", "v")
pts = [(sx(0), sy(3)), (sx(2.5), sy(0)), (sx(5), sy(-3)), (sx(7.5), sy(0)), (sx(9.5), sy(2.5))]
curve(slide, pts, GOLD, 2.4)
tx(slide, "cross", sx(2.5) - 0.35, sy(0) + 0.15, 0.9, 0.25, 12, GREEN, True, spa=0, scaled=False)
tx(slide, "cross", sx(7.5) - 0.35, sy(0) - 0.40, 0.9, 0.25, 12, GREEN, True, spa=0, scaled=False)
tx(slide, "not a turn", sx(5) - 0.5, sy(-3) - 0.35, 1.3, 0.25, 12, ORANGE, True, spa=0, scaled=False)

slide = new_slide("Example", "v = 0, and still no reversal", ORANGE)
para(slide, "x = t^3/3 − 2t^2 + 4t. Find distance and displacement in the first 3 s.", XL, Y0, CW, 0.80, 16)
fml(slide, "v = t^2 − 4t + 4 = (t − 2)^2", XL, Y0 + 1.00, CW, LH, 16)
fml(slide, "v = 0 at t = 2 s, but (t − 2)^2 never becomes negative.", XL, Y0 + 1.00 + LH + NG, CW, LH, 15)
fml(slide, "x(0) = 0 ,   x(3) = 3. No split of the path.", XL, Y0 + 1.00 + 2 * (LH + NG), CW, LH, 16)
result_card(slide, "distance = 3 m    ·    displacement = 3 m", XL, Y0 + 3.15, CW, 0.55, 16)
trap_card(slide, "A root of v = 0 is a U-turn only when v changes sign.", XL, Y0 + 3.90, CW, 0.50)

slide = new_slide("Example", "Read the start from x = αt^2 + βt − γ", TEAL)
fml(slide, "x(0) = −γ     →     the object starts at x = −γ", XL, Y0, CW, LH, 16)
fml(slide, "v = 2αt + β     →     v(0) = β", XL, Y0 + LH + NG, CW, LH, 16)
fml(slide, "v(2 s) = 4α + β", XL, Y0 + 2 * (LH + NG), CW, LH, 17)
para(slide, "α, β and γ are constants given in the question. Do not treat them as variables you must eliminate.", XL, Y0 + 2.40, CW, 1.00, 15)

slide = new_slide("Example", "Average speed from v = 3t^2", BLUE)
para(slide, "Speed at time t is 3t^2 (already positive). Average speed over the first 2 s is the area under the speed–time curve, divided by 2 s.", XL, Y0, CW, 1.10, 15)
fml(slide, "area = ∫ from 0 to 2 of 3t^2 dt = [t^3] from 0 to 2 = 8", XL, Y0 + 1.30, CW, LH, 16)
fml(slide, "average speed = 8 / 2 = 4 m/s", XL, Y0 + 1.30 + LH + NG, CW, LH, 17)
trap_card(slide, "Do not average  v(0) and v(2)  by adding and dividing by 2. The speed is not linear.", XL, Y0 + 2.90, CW, 0.70)

# ===== 2D VELOCITY =====
slide = new_slide("Figure", "Chord versus tangent", GREEN)
para(slide, "On a curved path, average velocity points along the straight displacement from the start to the end. Instantaneous velocity points along the tangent at that moment.", XL, Y0, CW, 1.10, 15)
# curve A to B
pts = [(2.2, 5.2), (3.6, 3.4), (5.4, 3.0), (7.2, 4.2)]
curve(slide, pts, ORANGE, 2.6)
arrow_line(slide, 2.2, 5.2, 7.2, 4.2, GREEN, 2.6)
# tangent at end
arrow_line(slide, 7.2, 4.2, 8.6, 5.3, GOLD, 2.4)
small_dot(slide, 2.2, 5.2, WHITE)
small_dot(slide, 7.2, 4.2, WHITE)
tx(slide, "A", 1.8, 5.35, 0.4, 0.25, 14, WHITE, True, spa=0, scaled=False)
tx(slide, "B", 7.3, 3.85, 0.4, 0.25, 14, WHITE, True, spa=0, scaled=False)
tx(slide, "v_avg", 4.2, 4.9, 1.2, 0.28, 14, GREEN, True, spa=0, scaled=False)
tx(slide, "v at B", 8.5, 5.35, 1.2, 0.28, 14, GOLD, True, spa=0, scaled=False)

slide = new_slide("Rule", "Same end-points, same average velocity", TEAL)
para(slide, "Average velocity uses only the displacement and the time taken. Two different x–t curves that start together and finish together have the same average velocity, even if one is straight and one bends.", XL, Y0, CW, 1.30, 15)
sx, sy = graph_axes(slide, 1.6, 2.8, 10.0, 3.2, 0, 8, 0, 8, "t", "x")
draw_line(slide, sx(0), sy(0), sx(6), sy(6), GOLD, 2.2)
curve(slide, [(sx(0), sy(0)), (sx(2), sy(4)), (sx(4), sy(3.2)), (sx(6), sy(6))], ORANGE, 2.2)
curve(slide, [(sx(0), sy(0)), (sx(2), sy(1.2)), (sx(4), sy(4.5)), (sx(6), sy(6))], GREEN, 2.2)
result_card(slide, "Same Δx and same Δt  →  same average velocity", XL, Y0 + 4.20, CW, 0.45, 15)

slide = new_slide("Example", "Average velocity in a plane", ORANGE)
para(slide, "At t = 0 the position is (2, 3) m. At t = 5 s the position is (13, 14) m. A point recorded at t = 2 s is not needed.", XL, Y0, CW, 1.05, 15)
fml(slide, "Δr = (13 − 2)î + (14 − 3)ĵ = 11î + 11ĵ", XL, Y0 + 1.25, CW, LH, 16)
fml(slide, "v_avg = Δr / 5 = (11/5)î + (11/5)ĵ", XL, Y0 + 1.25 + LH + NG, CW, LH, 16)
result_card(slide, "v_avg = (11/5)(î + ĵ) m/s", XL, Y0 + 2.80, CW, 0.55, 17)

slide = new_slide("Example", "Slope ratio on x–t", BLUE)
para(slide, "Two straight x–t lines leave the origin. Line A makes 45° with the time axis. Line B makes 30° with the time axis. Velocity is the slope.", XL, Y0, CW, 1.05, 15)
fml(slide, "v_A / v_B = tan 45° / tan 30°", XL, Y0 + 1.25, CW, LH, 17)
fml(slide, "= 1 / (1/√3) = √3", XL, Y0 + 1.25 + LH + NG, CW, LH, 17)
# two rays
sx, sy = graph_axes(slide, 7.6, 3.3, 4.8, 2.6, 0, 6, 0, 6, "t", "x")
draw_line(slide, sx(0), sy(0), sx(3.2), sy(3.2), GOLD, 2.2)
draw_line(slide, sx(0), sy(0), sx(4.2), sy(4.2 / 1.732), ORANGE, 2.2)
tx(slide, "45°", sx(2.2), sy(2.6), 0.7, 0.25, 12, GOLD, True, spa=0, scaled=False)
tx(slide, "30°", sx(3.2), sy(1.2), 0.7, 0.25, 12, ORANGE, True, spa=0, scaled=False)

slide = question_slide("MCQ — half a circle at steady speed")
qtx(slide, "Uniform speed v on a circle. After half a circle, the change in velocity and the change in speed are", XL, Y0, CW, 0.80, 16)
choices(slide, [
    "zero , zero",
    "2v , zero",
    "2v , 2v",
    "zero , 2v",
], XL, Y0 + 1.05)
# semicircle hint
cx, cy, R = 10.2, 4.6, 1.1
arc = [(cx + R * math.cos(math.radians(a)), cy - R * math.sin(math.radians(a))) for a in range(0, 181, 12)]
curve(slide, arc, ORANGE, 2.0)
arrow_line(slide, cx + R, cy, cx + R, cy - 0.7, GOLD, 2.0)
arrow_line(slide, cx - R, cy, cx - R, cy + 0.7, GREEN, 2.0)

slide = answer_slide()
body(slide, "Idea: the two velocity arrows have the same size and opposite direction, so their difference has size 2v. Speed never changed.", XL, Y0, CW, 0.85)
result_card(slide, "Answer: 2v , zero", XL, Y0 + 1.10, CW, 0.45, 18)

# ===== NEW AVERAGES =====
slide = new_slide("Example", "A half that is itself an average", GREEN)
para(slide, "First half of the distance at 5 m/s. The second half is covered in two equal times, at 15 m/s and at 25 m/s.", XL, Y0, CW, 1.05, 16)
fml(slide, "On the second half, times are equal:  (15 + 25) / 2 = 20 m/s", XL, Y0 + 1.25, CW, LH, 15)
fml(slide, "Whole trip: equal distances at 5 m/s and 20 m/s", XL, Y0 + 1.25 + LH + NG, CW, LH, 15)
fml(slide, "average = 2 × 5 × 20 / (5 + 20) = 8 m/s", XL, Y0 + 1.25 + 2 * (LH + NG), CW, LH, 16)
result_card(slide, "8 m/s", XL, Y0 + 3.30, CW, 0.45, 18)

slide = new_slide("Example", "An average that cannot exist", ORANGE)
para(slide, "Equal distances. First half at 100 m/s. Someone claims the average for the whole trip is 200 m/s. Find the speed on the second half.", XL, Y0, CW, 1.10, 16)
fml(slide, "2 × 100 × v / (100 + v) = 200", XL, Y0 + 1.30, CW, LH, 16)
fml(slide, "100 v = 100(100 + v)", XL, Y0 + 1.30 + LH + NG, CW, LH, 16)
fml(slide, "0 = 100     →     no real speed v", XL, Y0 + 1.30 + 2 * (LH + NG), CW, LH, 16)
trap_card(slide, "Average speed on equal distances always lies between the two speeds. 200 cannot sit above both if one of them is 100.", XL, Y0 + 3.30, CW, 0.80)

slide = new_slide("Lock lines", "Only today’s lines", GREEN)
bullet(slide, "v = dx/dt. Speed at that instant is |v|.", XL, Y0, CW, 0.48, 16)
bullet(slide, "|average velocity| ≤ average speed.", XL, Y0 + 0.55, CW, 0.48, 16)
bullet(slide, "Constant speed can still have a changing velocity.", XL, Y0 + 1.10, CW, 0.48, 16)
bullet(slide, "On v–t, direction changes only where the graph crosses zero.", XL, Y0 + 1.65, CW, 0.52, 16)
bullet(slide, "v = 0 with no sign change is not a U-turn.", XL, Y0 + 2.20, CW, 0.48, 16)
bullet(slide, "Average velocity follows the displacement chord, not the bend of the path.", XL, Y0 + 2.75, CW, 0.55, 16)
bullet(slide, "Area under a speed–time curve, divided by time, is average speed.", XL, Y0 + 3.35, CW, 0.55, 16)

slide = prs.slides.add_slide(BLANK)
_bg(slide)
_bar(slide, "End", GOLD)
tx_3d(slide, "Thank you", XL, 2.4, CW, 1.0, 48, GOLD, RGBColor(120, 120, 130))
tx(slide, "Excellence  ·  Motion 04  ·  Velocity · v–t · averages that are new", XL, 3.7, CW, 0.45, 15, MUTED, True, PP_ALIGN.CENTER, spa=0, scaled=False)
_footer(slide, len(prs.slides))
after_content_slide()

OUTPUT.parent.mkdir(parents=True, exist_ok=True)
prs.save(OUTPUT)
print(OUTPUT)
print("slides", len(prs.slides))
