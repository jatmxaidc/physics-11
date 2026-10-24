"""Motion 02 — deeper distance & displacement, arcs, x-t graphs (PW 02)."""
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
OUTPUT = BASE / "lecture slides" / "Lecture_Motion_02_Distance_Displacement_Deep.pptx"

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
tx(slide, "Motion 02", bx + bw / 2, by + 0.14, bw / 2, 0.50, 25, BLACK, True, PP_ALIGN.CENTER, spa=0, scaled=False)
tx(slide, "Motion in a Straight Line", bx, by + 1.02, bw, 0.55, 28, WHITE, True, PP_ALIGN.CENTER, spa=0, scaled=False, font="Georgia")
tx(slide, "Distance & Displacement Deep Dive · Arcs · x–t Graphs", bx, by + 1.62, bw, 0.40, 14, GOLD, True, PP_ALIGN.CENTER, spa=0, scaled=False)
tx(slide, "By - Excellence Programme", XL, 5.28, CW, 0.50, 25, WHITE, True, PP_ALIGN.CENTER, spa=0, scaled=False)
tx(slide, "20 October 2026", XL, 5.80, CW, 0.32, 18, MUTED, True, PP_ALIGN.CENTER, spa=0, scaled=False)
_footer(slide, 1)
add_notes(slide, "From PW Motion in a Straight Line Lecture 02.")
after_content_slide()

slide = prs.slides.add_slide(BLANK)
_bg(slide)
tx(slide, "What we learn today", XL, 0.35, CW, 0.64, 32, GOLD, True, spa=0, scaled=False)
_solid(slide, XL, 1.00, 3.40, 0.035, GOLD)
topics = [
    "Lock the may / must logic for distance & displacement",
    "More 1-D, 2-D and 3-D path calculations",
    "Square, hexagon and triangle geometry shortcuts",
    "Motion on a circular arc — distance and chord",
    "Reading x–t graphs (graph is not the path)",
    "Board MCQs from today’s ideas",
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

slide = new_slide("Class note", "Link from Motion 01", TEAL)
para(slide, "Last class we met frame, rest–motion, and the basic meaning of distance and displacement. Today we go deeper: harder paths, polygon shortcuts, circular arcs, and graphs.", XL, Y0, CW, 1.35, 16)
ncert(slide, "Δr = r_f − r_i    |    distance = length of the actual path", XL, Y0 + 1.60, CW, 1.05, 15)
trap_card(slide, "Same unit (metre) does not mean same quantity. Distance is scalar. Displacement is vector.", XL, Y0 + 2.95, CW, 0.70)

# ===== LOGIC =====
slide = new_slide("Rule", "May / must logic — lock this table", ORANGE)
grid_table(
    slide,
    ["Statement", "Truth"],
    [
        ["Distance = 0  ⇒  displacement = 0", "MUST (object did not move)"],
        ["Displacement = 0  ⇒  distance = 0", "MAY / MAY NOT (round trip)"],
        ["Displacement ≠ 0  ⇒  distance = 0", "MUST NOT"],
        ["Distance ≠ 0  ⇒  displacement = 0", "MAY / MAY NOT"],
        ["In every 1-D motion, distance = |disp|", "FALSE (turns break equality)"],
    ],
    XL, Y0, [6.8, 5.0], row_h=0.52, hsize=13, csize=13,
)
# round trip figure
rounded_box(slide, 9.0, 4.55, 3.5, 1.85, RGBColor(14, 16, 22), TEAL, 1.2)
small_dot(slide, 10.6, 5.6, GOLD)
path_polyline(slide, [(10.6, 5.6), (11.6, 5.0), (10.6, 4.9), (9.6, 5.0), (10.6, 5.6)], ORANGE, 2.0)
tx(slide, "disp = 0, dist > 0", 9.2, 5.95, 3.1, 0.28, 12, ORANGE, True, PP_ALIGN.CENTER, spa=0, scaled=False)

slide = question_slide("MCQ — two statements")
qtx(slide, "Stmt-1: Distance and displacement are different quantities.", XL, Y0, CW, 0.45, 16)
qtx(slide, "Stmt-2: They have the same unit and same dimension.", XL, Y0 + 0.55, CW, 0.45, 16)
choices(slide, [
    "Only Stmt-1 is correct",
    "Both are correct",
    "Only Stmt-2 is correct",
    "Both are incorrect",
], XL, Y0 + 1.20)

slide = answer_slide()
body(slide, "Idea: different nature (scalar vs vector), but both measured in metre with [L].", XL, Y0, CW, 0.55)
result_card(slide, "Answer: both are correct", XL, Y0 + 0.80, CW, 0.45, 17)

# ===== 1D / 2D CORE =====
slide = new_slide("Example", "1-D: out and past the origin", GREEN)
para(slide, "On the x-axis: start at origin, go to +20 m, then reverse and stop at −20 m.", XL, Y0, CW, 0.70, 15)
sx = number_line(slide, 1.2, 2.70, 11.0, -25, 25, [-20, 0, 20])
mark_point(slide, sx(0), 2.70, "O", WHITE)
mark_point(slide, sx(20), 2.70, "+20", ORANGE, above=False)
mark_point(slide, sx(-20), 2.70, "end", GOLD)
arrow_line(slide, sx(0), 2.35, sx(20), 2.35, ORANGE, 2.2)
arrow_line(slide, sx(20), 2.10, sx(-20), 2.10, ORANGE, 2.2)
arrow_line(slide, sx(0), 3.15, sx(-20), 3.15, GREEN, 2.6)
tx(slide, "displacement", 3.5, 3.25, 2.0, 0.28, 13, GREEN, True, spa=0, scaled=False)
fml(slide, "distance = 20 + 40 = 60 m", XL, Y0 + 3.40, CW, LH, 16)
fml(slide, "displacement = (−20)î − 0 = −20 î m", XL, Y0 + 3.40 + LH + NG, CW, LH, 16)

slide = new_slide("Figure", "L-path: 4 m + 6 m + 4 m", TEAL)
para(slide, "A → B = 4 m right, B → C = 6 m up, C → D = 4 m right.", XL, Y0, CW, 0.60, 15)
ox, oy, s = 2.0, 5.5, 0.42
pts = [(ox, oy), (ox + 4 * s, oy), (ox + 4 * s, oy - 6 * s), (ox + 8 * s, oy - 6 * s)]
path_polyline(slide, pts, ORANGE, 2.6)
for p, lab in zip(pts, ["A", "B", "C", "D"]):
    small_dot(slide, p[0], p[1], WHITE)
    tx(slide, lab, p[0] - 0.15, p[1] + 0.12, 0.4, 0.25, 13, WHITE, True, spa=0, scaled=False)
arrow_line(slide, pts[0][0], pts[0][1], pts[-1][0], pts[-1][1], GREEN, 2.6)
# right triangle helper
draw_line(slide, pts[0][0], pts[0][1], pts[-1][0], pts[0][1], MUTED, 1.2)
draw_line(slide, pts[-1][0], pts[0][1], pts[-1][0], pts[-1][1], MUTED, 1.2)
tx(slide, "8 m", ox + 3.5 * s, oy + 0.15, 0.8, 0.25, 12, MUTED, True, spa=0, scaled=False)
tx(slide, "6 m", pts[-1][0] + 0.15, oy - 3 * s, 0.7, 0.25, 12, MUTED, True, spa=0, scaled=False)
fml(slide, "distance = 14 m", 8.0, 2.4, 4.5, LH, 16)
fml(slide, "|Δr| = √(8^2 + 6^2) = 10 m", 8.0, 2.4 + LH + NG, 4.5, LH, 15)
result_card(slide, "14 m , 10 m", 8.0, 4.0, 4.5, 0.45, 16)

slide = question_slide("MCQ — L-path numbers")
qtx(slide, "For the path above, distance and |displacement| are", XL, Y0, CW, 0.55, 17)
choices(slide, [
    "14 m, 8 m",
    "14 m, 6 m",
    "14 m, 10 m",
    "None",
], XL, Y0 + 0.80)

slide = answer_slide()
fml(slide, "distance = 14 m ,  |Δr| = 10 m", XL, Y0, CW, LH, 17)
result_card(slide, "Answer: option 3", XL, Y0 + 0.90, CW, 0.40, 18)

slide = new_slide("Example", "Endpoints only — distance may be unknown", ORANGE)
para(slide, "Object goes from (0, 0) to (2, −4) on a curved path. Without the path, distance cannot be found. Displacement still can.", XL, Y0, CW, 1.00, 15)
draw_line(slide, 1.5, 4.6, 5.2, 4.6, MUTED, 1.4)
small_dot(slide, 2.0, 4.6, GREEN)
tx(slide, "A", 1.6, 4.8, 0.5, 0.25, 12, GREEN, True, spa=0, scaled=False)
small_dot(slide, 4.4, 3.2, GOLD)
tx(slide, "B", 4.2, 2.85, 0.5, 0.25, 12, GOLD, True, spa=0, scaled=False)
curve(slide, [(2.0, 4.6), (2.7, 3.7), (3.5, 3.2), (4.4, 3.2)], ORANGE, 2.2)
arrow_line(slide, 2.0, 4.6, 4.4, 3.2, GREEN, 2.4)
fml(slide, "Δr = 2î − 4ĵ ,  |Δr| = 2√5 m", 6.5, 2.5, 6.0, LH, 15)
trap_card(slide, "Endpoints alone → displacement yes, distance maybe no.", 6.5, 3.5, 6.0, 0.60)

slide = new_slide("Example", "Projectile landing & 3–4–5", BLUE)
para(slide, "Ends at (8, 6) after a curved flight from the origin. Path length unknown; net shift is a 3–4–5 style triangle ×2.", XL, Y0, CW, 0.95, 15)
sx, sy = graph_axes(slide, 1.2, 2.35, 5.2, 3.5, 0, 10, 0, 8, "x", "y")
small_dot(slide, sx(0), sy(0), GREEN)
small_dot(slide, sx(8), sy(6), GOLD)
curve(slide, [(sx(0), sy(0)), (sx(3), sy(5)), (sx(6), sy(6.5)), (sx(8), sy(6))], ORANGE, 2.2)
arrow_line(slide, sx(0), sy(0), sx(8), sy(6), GREEN, 2.4)
fml(slide, "Δr = 8î + 6ĵ", 7.2, 2.5, 5.3, LH, 16)
fml(slide, "|Δr| = √(64+36) = 10 m", 7.2, 2.5 + LH + NG, 5.3, LH, 16)

slide = new_slide("Example", "3-D walk: North, East, Up", TEAL)
para(slide, "Walk 10 m North, 10 m East, 10 m vertically up.", XL, Y0, CW, 0.70, 15)
fml(slide, "distance = 10 + 10 + 10 = 30 m", XL, Y0 + 0.90, CW, LH, 16)
fml(slide, "Δr = 10î + 10ĵ + 10k̂", XL, Y0 + 0.90 + LH + NG, CW, LH, 16)
fml(slide, "|Δr| = 10√3 m", XL, Y0 + 0.90 + 2 * (LH + NG), CW, LH, 16)
# tiny 3-axis sketch
ox, oy = 9.5, 5.2
arrow_line(slide, ox, oy, ox + 1.6, oy, ORANGE, 2.2)
tx(slide, "E î", ox + 1.7, oy - 0.15, 0.8, 0.25, 12, ORANGE, True, spa=0, scaled=False)
arrow_line(slide, ox, oy, ox, oy - 1.4, GREEN, 2.2)
tx(slide, "N ĵ", ox + 0.1, oy - 1.6, 0.8, 0.25, 12, GREEN, True, spa=0, scaled=False)
arrow_line(slide, ox, oy, ox - 0.9, oy - 0.9, GOLD, 2.2)
tx(slide, "up k̂", ox - 1.6, oy - 1.1, 1.0, 0.25, 12, GOLD, True, spa=0, scaled=False)

slide = question_slide("MCQ — P(2,3,5) → Q(3,4,5)")
qtx(slide, "Displacement vector is", XL, Y0, CW, 0.50, 17)
choices(slide, [
    "î + ĵ + 10 k̂",
    "î + ĵ + 5 k̂",
    "î + ĵ",
    "2î + 4ĵ + 6k̂",
], XL, Y0 + 0.75)

slide = answer_slide()
fml(slide, "Δr = (3−2)î + (4−3)ĵ + (5−5)k̂ = î + ĵ", XL, Y0, CW, LH, 16)
result_card(slide, "Answer: î + ĵ", XL, Y0 + 0.95, CW, 0.40, 18)

slide = new_slide("Example", "IIT-style walk: N, E, then SW", ORANGE)
para(slide, "30 m North, 20 m East, then 30√2 m South-West. Find displacement.", XL, Y0, CW, 0.85, 15)
ox, oy = 2.8, 5.2
arrow_line(slide, ox, oy, ox, oy - 1.5, GREEN, 2.4)
tx(slide, "30 N", ox + 0.15, oy - 0.9, 1.0, 0.25, 12, GREEN, True, spa=0, scaled=False)
arrow_line(slide, ox, oy - 1.5, ox + 1.2, oy - 1.5, ORANGE, 2.4)
tx(slide, "20 E", ox + 0.3, oy - 1.85, 1.0, 0.25, 12, ORANGE, True, spa=0, scaled=False)
arrow_line(slide, ox + 1.2, oy - 1.5, ox + 1.2 - 1.5, oy - 1.5 + 1.5, GOLD, 2.4)
tx(slide, "30√2 SW", ox - 0.4, oy - 0.5, 1.6, 0.25, 12, GOLD, True, spa=0, scaled=False)
arrow_line(slide, ox, oy, ox - 0.9, oy, TEAL, 2.6)
tx(slide, "10 W", ox - 1.5, oy - 0.35, 1.2, 0.25, 13, TEAL, True, spa=0, scaled=False)
fml(slide, "SW = 30√2 × (−î − ĵ)/√2 = −30î − 30ĵ", 7.0, 2.3, 5.5, LH, 13)
fml(slide, "Net = 20î + 30ĵ − 30î − 30ĵ = −10î", 7.0, 2.3 + LH + NG, 5.5, LH, 14)
result_card(slide, "Displacement = 10 m West", 7.0, 4.0, 5.5, 0.50, 15)

slide = question_slide("MCQ — IIT walk answer")
qtx(slide, "Net displacement in the walk above is", XL, Y0, CW, 0.55, 17)
choices(slide, [
    "10 m East",
    "10 m North",
    "10 m West",
    "10√2 m South-West",
], XL, Y0 + 0.80)

slide = answer_slide()
result_card(slide, "Answer: 10 m West", XL, Y0, CW, 0.45, 18)

slide = new_slide("Example", "Spiral segments 1…8 m", GREEN)
para(slide, "Alternate right, up, left, down with lengths 1, 2, 3, 4, 5, 6, 7, 8 m.", XL, Y0, CW, 0.65, 15)
ox, oy, sc = 3.5, 4.8, 0.20
dirs = [(1, 0), (0, -1), (-1, 0), (0, 1), (1, 0), (0, -1), (-1, 0), (0, 1)]
x, y = ox, oy
pts = [(x, y)]
for L, (dx, dy) in zip([1, 2, 3, 4, 5, 6, 7, 8], dirs):
    x += dx * L * sc
    y += dy * L * sc
    pts.append((x, y))
path_polyline(slide, pts, ORANGE, 2.0)
small_dot(slide, pts[0][0], pts[0][1], GREEN)
small_dot(slide, pts[-1][0], pts[-1][1], GOLD)
arrow_line(slide, pts[0][0], pts[0][1], pts[-1][0], pts[-1][1], GREEN, 2.4)
fml(slide, "distance = 36 m", 8.0, 2.4, 4.5, LH, 15)
fml(slide, "Δr = −4î − 4ĵ", 8.0, 2.4 + LH + NG, 4.5, LH, 15)
fml(slide, "|Δr| = 4√2 m", 8.0, 2.4 + 2 * (LH + NG), 4.5, LH, 15)

slide = new_slide("Example", "Staircase path → 7–24–25", BLUE)
para(slide, "Horizontal pieces add to 24 m. Vertical pieces add to 7 m. Net is a right triangle.", XL, Y0, CW, 0.90, 15)
fml(slide, "|Δr| = √(24^2 + 7^2) = √(576+49) = √625 = 25 m", XL, Y0 + 1.15, CW, LH, 16)
# triangle sketch
draw_line(slide, 3.0, 5.5, 8.5, 5.5, ORANGE, 2.2)
draw_line(slide, 8.5, 5.5, 8.5, 3.2, ORANGE, 2.2)
arrow_line(slide, 3.0, 5.5, 8.5, 3.2, GREEN, 2.6)
tx(slide, "24 m", 5.2, 5.65, 1.2, 0.25, 13, ORANGE, True, spa=0, scaled=False)
tx(slide, "7 m", 8.7, 4.2, 1.0, 0.25, 13, ORANGE, True, spa=0, scaled=False)
tx(slide, "25 m", 5.0, 4.0, 1.4, 0.28, 14, GREEN, True, spa=0, scaled=False)

# ===== POLYGON GEOMETRY =====
slide = new_slide("Rule", "Square park shortcuts", TEAL)
para(slide, "Square ABCD of side ℓ. Start at A. Walk on the boundary.", XL, Y0, CW, 0.65, 15)
sq = 2.4
ax, ay = 1.8, 5.4
corners = [(ax, ay), (ax + sq, ay), (ax + sq, ay - sq), (ax, ay - sq)]
for (x1, y1), (x2, y2) in zip(corners, corners[1:] + [corners[0]]):
    draw_line(slide, x1, y1, x2, y2, WHITE, 2.0)
arrow_line(slide, ax, ay, ax + sq, ay - sq, GOLD, 2.4)
tx(slide, "A", ax - 0.35, ay + 0.05, 0.35, 0.25, 13, GREEN, True, spa=0, scaled=False)
tx(slide, "C", ax + sq + 0.08, ay - sq - 0.30, 0.35, 0.25, 13, GOLD, True, spa=0, scaled=False)
tx(slide, "√2 ℓ", ax + 0.7, ay - 1.3, 1.2, 0.25, 13, GOLD, True, spa=0, scaled=False)
grid_table(
    slide,
    ["Path", "distance", "|disp|"],
    [
        ["A → B", "ℓ", "ℓ"],
        ["A → C via B", "2ℓ", "ℓ√2"],
        ["A → D via B,C", "3ℓ", "ℓ"],
        ["A → A lap", "4ℓ", "0"],
    ],
    5.8, Y0 + 0.80, [3.2, 2.4, 2.4], row_h=0.48, hsize=13, csize=14,
)

slide = new_slide("Rule", "Regular hexagon — distance vs chord", GREEN)
para(slide, "Regular hexagon side ℓ. From vertex A, walking around the boundary.", XL, Y0, CW, 0.70, 15)
# hexagon approx
cx, cy, R = 3.2, 4.2, 1.5
hex_pts = []
for k in range(6):
    ang = math.radians(-90 + 60 * k)
    hex_pts.append((cx + R * math.cos(ang), cy + R * math.sin(ang)))
for a, b in zip(hex_pts, hex_pts[1:] + [hex_pts[0]]):
    draw_line(slide, a[0], a[1], b[0], b[1], WHITE, 1.8)
arrow_line(slide, hex_pts[0][0], hex_pts[0][1], hex_pts[2][0], hex_pts[2][1], GOLD, 2.2)
tx(slide, "A", hex_pts[0][0] - 0.35, hex_pts[0][1] - 0.15, 0.35, 0.25, 12, GREEN, True, spa=0, scaled=False)
grid_table(
    slide,
    ["Sides walked", "distance", "|disp|"],
    [
        ["1", "ℓ", "ℓ"],
        ["2", "2ℓ", "ℓ√3"],
        ["3", "3ℓ", "2ℓ"],
        ["6 (full)", "6ℓ", "0"],
    ],
    6.0, Y0 + 0.90, [3.0, 2.4, 2.4], row_h=0.48, hsize=13, csize=14,
)

slide = new_slide("Rule", "Equilateral triangle + centroid", BLUE)
para(slide, "Equilateral triangle side ℓ. Centroid O. Useful lengths: vertex→centroid = ℓ/√3 ; centroid→side midpoint = ℓ/(2√3).", XL, Y0, CW, 1.10, 15)
# triangle
t = [(3.5, 5.4), (5.8, 5.4), (4.65, 3.2)]
draw_line(slide, t[0][0], t[0][1], t[1][0], t[1][1], WHITE, 1.8)
draw_line(slide, t[1][0], t[1][1], t[2][0], t[2][1], WHITE, 1.8)
draw_line(slide, t[2][0], t[2][1], t[0][0], t[0][1], WHITE, 1.8)
ox = (t[0][0] + t[1][0] + t[2][0]) / 3
oy = (t[0][1] + t[1][1] + t[2][1]) / 3
small_dot(slide, ox, oy, GOLD)
tx(slide, "O", ox + 0.12, oy - 0.05, 0.4, 0.25, 12, GOLD, True, spa=0, scaled=False)
tx(slide, "A", t[0][0] - 0.35, t[0][1] + 0.05, 0.35, 0.25, 12, GREEN, True, spa=0, scaled=False)
tx(slide, "B", t[1][0] + 0.08, t[1][1] + 0.05, 0.35, 0.25, 12, GREEN, True, spa=0, scaled=False)
fml(slide, "Example A → B → O:  distance = ℓ + ℓ/√3 ,  |disp| = ℓ/√3", 6.5, 2.5, 6.0, LH, 13)
fml(slide, "Example O → midpoint → B: use ℓ/(2√3) + ℓ/2 for distance", 6.5, 2.5 + LH + NG, 6.0, LH, 13)

# ===== CIRCULAR ARCS =====
slide = new_slide("Idea", "Circular arc — the two lengths", ORANGE)
para(slide, "Object moves on a circle of radius R through angle θ (in radian). Distance follows the arc. Displacement is the chord.", XL, Y0, CW, 1.05, 15)
# arc figure
cx, cy, R = 3.5, 4.5, 1.8
# radii
a1, a2 = math.radians(20), math.radians(110)
draw_line(slide, cx, cy, cx + R * math.cos(a1), cy - R * math.sin(a1), MUTED, 1.5)
draw_line(slide, cx, cy, cx + R * math.cos(a2), cy - R * math.sin(a2), MUTED, 1.5)
# approximate arc
arc_pts = []
for i in range(16):
    a = a1 + (a2 - a1) * i / 15
    arc_pts.append((cx + R * math.cos(a), cy - R * math.sin(a)))
curve(slide, arc_pts, ORANGE, 2.6)
# chord
arrow_line(slide, arc_pts[0][0], arc_pts[0][1], arc_pts[-1][0], arc_pts[-1][1], GREEN, 2.6)
small_dot(slide, cx, cy, WHITE)
tx(slide, "O", cx - 0.35, cy - 0.1, 0.35, 0.25, 12, WHITE, True, spa=0, scaled=False)
tx(slide, "θ", cx + 0.35, cy - 0.55, 0.4, 0.25, 14, GOLD, True, spa=0, scaled=False)
tx(slide, "arc", 4.8, 3.0, 1.0, 0.25, 13, ORANGE, True, spa=0, scaled=False)
tx(slide, "chord", 4.2, 5.5, 1.2, 0.25, 13, GREEN, True, spa=0, scaled=False)
ncert(slide, "distance = R θ     |     |displacement| = 2 R sin(θ/2)", 6.8, 2.4, 5.7, 1.20, 14)

slide = new_slide("Rule", "Why the chord formula works", TEAL)
para(slide, "Join O to the ends A and B. Triangle OAB is isosceles with OA = OB = R and angle AOB = θ. Drop a perpendicular from O to chord AB — it bisects θ and the chord.", XL, Y0, CW, 1.25, 15)
fml(slide, "Half chord = R sin(θ/2)", XL, Y0 + 1.45, CW, LH, 17)
fml(slide, "Full chord = |Δr| = 2 R sin(θ/2)", XL, Y0 + 1.45 + LH + NG, CW, LH, 17)
trap_card(slide, "Put θ in radian for distance = Rθ. For 60° use θ = π/3, not 60.", XL, Y0 + 2.90, CW, 0.65)

slide = new_slide("Example", "Arc angles you must own", GREEN)
grid_table(
    slide,
    ["θ", "distance Rθ", "|disp| = 2R sin(θ/2)"],
    [
        ["60° = π/3", "R π/3", "R"],
        ["90° = π/2", "R π/2", "R√2"],
        ["120° = 2π/3", "2πR/3", "R√3"],
        ["180° = π", "π R", "2R"],
        ["270° = 3π/2", "3πR/2", "R√2"],
    ],
    XL, Y0, [3.2, 3.8, 4.8], row_h=0.52, hsize=13, csize=14,
)
trap_card(slide, "Small θ: sin(θ/2) ≈ θ/2 ⇒ |disp| ≈ Rθ = distance (almost equal).", XL, Y0 + 3.40, CW, 0.60)

slide = new_slide("Example", "Radial then semicircle", BLUE)
para(slide, "From centre O go out along a radius to A, then along the upper semicircle to B (opposite end of diameter).", XL, Y0, CW, 0.95, 15)
cx, cy, R = 3.5, 4.6, 1.7
draw_line(slide, cx - R, cy, cx + R, cy, MUTED, 1.4)
# semicircle approx
arc = [(cx + R * math.cos(math.radians(a)), cy - R * math.sin(math.radians(a))) for a in range(0, 181, 10)]
curve(slide, arc, ORANGE, 2.4)
arrow_line(slide, cx, cy, cx - R, cy, GOLD, 2.2)
small_dot(slide, cx, cy, WHITE)
tx(slide, "O", cx - 0.15, cy + 0.15, 0.4, 0.25, 12, WHITE, True, spa=0, scaled=False)
tx(slide, "A", cx - R - 0.4, cy - 0.15, 0.4, 0.25, 12, GREEN, True, spa=0, scaled=False)
tx(slide, "B", cx + R + 0.1, cy - 0.15, 0.4, 0.25, 12, GOLD, True, spa=0, scaled=False)
fml(slide, "distance = R + π R", 7.0, 2.5, 5.5, LH, 16)
fml(slide, "displacement = R (from O to B along diameter)", 7.0, 2.5 + LH + NG, 5.5, LH, 14)
result_card(slide, "dist = R(1+π) , |disp| = R", 7.0, 4.2, 5.5, 0.50, 15)

slide = new_slide("Example", "3.5 revolutions on a circle", ORANGE)
para(slide, "Period for one lap = 40 s. Radius R. After 140 s the object has completed 140/40 = 3.5 revolutions.", XL, Y0, CW, 1.00, 15)
fml(slide, "distance = 3.5 × 2πR = 7π R", XL, Y0 + 1.20, CW, LH, 16)
fml(slide, "After half-integer turns, ends at the opposite point on diameter", XL, Y0 + 1.20 + LH + NG, CW, LH, 14)
fml(slide, "|displacement| = 2R", XL, Y0 + 1.20 + 2 * (LH + NG), CW, LH, 16)
fml(slide, "ratio distance / |disp| = 7π / 2", XL, Y0 + 1.20 + 3 * (LH + NG), CW, LH, 16)

slide = question_slide("MCQ — circle (x−5)^2 + y^2 = 25")
qtx(slide, "Particle goes along the upper semicircle from (0,0) to (10,0). Distance is", XL, Y0, CW, 0.70, 16)
# mini circle
cx, cy, R = 9.5, 4.0, 1.3
arc = [(cx + R * math.cos(math.radians(a)), cy - R * math.sin(math.radians(a))) for a in range(0, 181, 12)]
curve(slide, arc, ORANGE, 2.2)
draw_line(slide, cx - R, cy, cx + R, cy, MUTED, 1.3)
small_dot(slide, cx - R, cy, GREEN)
small_dot(slide, cx + R, cy, GOLD)
choices(slide, [
    "10",
    "5π / 2",
    "5π",
    "10π",
], XL, Y0 + 0.95, w=6.5)

slide = answer_slide()
body(slide, "Idea: centre (5,0), R = 5. Upper semicircle arc = πR = 5π.", XL, Y0, CW, 0.55)
result_card(slide, "Answer: 5π", XL, Y0 + 0.80, CW, 0.40, 18)

# ===== X-T GRAPHS =====
slide = new_slide("Idea", "x–t graph is not the path", TEAL)
para(slide, "A position–time graph plots x against t. The curve on paper is not the road the object walked. Read change in x for displacement. Add |Δx| of each piece for distance when it turns.", XL, Y0, CW, 1.35, 15)
ncert(slide, "displacement = x_f − x_i     |     distance = sum of |Δx| on each leg", XL, Y0 + 1.60, CW, 1.15, 14)
trap_card(slide, "A sloping line on x–t means uniform velocity — not a sloping road.", XL, Y0 + 3.10, CW, 0.65)

slide = new_slide("Figure", "Two travellers, same Δx", GREEN)
para(slide, "A reaches x_0 at time t_1. B reaches the same x_0 at later time t_2. Both start from x = 0.", XL, Y0, CW, 0.75, 15)
sx, sy = graph_axes(slide, 1.5, 2.2, 10.0, 3.8, 0, 10, 0, 8, "t", "x")
draw_line(slide, sx(0), sy(0), sx(4), sy(6), ORANGE, 2.4)
draw_line(slide, sx(0), sy(0), sx(8), sy(6), GREEN, 2.4)
tx(slide, "A", sx(4) + 0.1, sy(6) - 0.35, 0.5, 0.25, 13, ORANGE, True, spa=0, scaled=False)
tx(slide, "B", sx(8) + 0.1, sy(6) - 0.35, 0.5, 0.25, 13, GREEN, True, spa=0, scaled=False)
draw_line(slide, sx(0), sy(6), sx(9), sy(6), MUTED, 1.0)
tx(slide, "x_0", sx(0) - 0.55, sy(6) - 0.12, 0.5, 0.25, 12, MUTED, True, spa=0, scaled=False)

slide = question_slide("MCQ — same final x")
qtx(slide, "For A and B above, which is correct?", XL, Y0, CW, 0.55, 17)
choices(slide, [
    "displacement of A = displacement of B",
    "distance of A = distance of B always",
    "A and B have equal average speeds",
    "A’s distance is always greater",
], XL, Y0 + 0.80)

slide = answer_slide()
body(slide, "Idea: both go 0 → x_0 with no turn, so Δx is the same. Speeds differ because times differ.", XL, Y0, CW, 0.70)
result_card(slide, "Answer: displacement of A = displacement of B", XL, Y0 + 0.95, CW, 0.50, 15)

slide = new_slide("Example", "x–t with rest and return", BLUE)
para(slide, "Read carefully: go forward, rest (horizontal), then return.", XL, Y0, CW, 0.65, 15)
sx, sy = graph_axes(slide, 1.2, 2.15, 6.0, 3.6, 0, 14, 0, 12, "t", "x")
# 0->10, rest, back to 0
draw_line(slide, sx(0), sy(0), sx(3), sy(10), ORANGE, 2.4)
draw_line(slide, sx(3), sy(10), sx(7), sy(10), GOLD, 2.4)
draw_line(slide, sx(7), sy(10), sx(12), sy(0), ORANGE, 2.4)
fml(slide, "displacement = 0 − 0 = 0", 7.8, 2.4, 4.8, LH, 15)
fml(slide, "distance = 10 + 0 + 10 = 20 m", 7.8, 2.4 + LH + NG, 4.8, LH, 15)
result_card(slide, "Back to start → disp 0, distance 20 m", 7.8, 4.2, 4.8, 0.55, 13)

slide = new_slide("Example", "x–t: 10 → 20 → −20", ORANGE)
sx, sy = graph_axes(slide, 1.2, 2.0, 6.2, 4.0, 0, 10, -25, 25, "t", "x")
draw_line(slide, sx(0), sy(10), sx(3), sy(20), GREEN, 2.4)
draw_line(slide, sx(3), sy(20), sx(9), sy(-20), ORANGE, 2.4)
fml(slide, "displacement = (−20) − 10 = −30 m", 7.9, 2.3, 4.7, LH, 14)
fml(slide, "distance = 10 + 40 = 50 m", 7.9, 2.3 + LH + NG, 4.7, LH, 15)
trap_card(slide, "Distance adds the absolute climbs. Displacement only cares about start and end x.", 7.9, 4.0, 4.7, 0.85)

slide = new_slide("Rule", "Valid distance–time graph", GREEN)
para(slide, "Distance never decreases with time for a real moving clock. A distance–time graph cannot go down, cannot be negative, and cannot show two distances at one time.", XL, Y0, CW, 1.20, 15)
# four mini panels conceptually as text cards
mini_card(slide, "Reject", "Decreasing d–t\n(distance fell)", XL, Y0 + 1.50, 2.8, 1.50, RED)
mini_card(slide, "Reject", "Negative distance", XL + 3.1, Y0 + 1.50, 2.8, 1.50, RED)
mini_card(slide, "Reject", "Two values at\none time", XL + 6.2, Y0 + 1.50, 2.8, 1.50, RED)
mini_card(slide, "Accept", "Non-decreasing\nd ≥ 0", XL + 9.3, Y0 + 1.50, 2.8, 1.50, GREEN)

slide = question_slide("MCQ — which d–t is valid?")
qtx(slide, "Which distance–time graph can be correct?", XL, Y0, CW, 0.55, 17)
choices(slide, [
    "A graph that decreases after some time",
    "A graph that goes below d = 0",
    "A vertical line (two d at one t)",
    "A non-decreasing graph with d ≥ 0",
], XL, Y0 + 0.80)

slide = answer_slide()
result_card(slide, "Answer: non-decreasing graph with d ≥ 0", XL, Y0, CW, 0.50, 16)

slide = new_slide("Rule", "Displacement–time can go down", TEAL)
para(slide, "Displacement (or position) can increase or decrease. An s–t or x–t graph may slope up or down. That is allowed — the object can reverse.", XL, Y0, CW, 1.20, 16)
bullet(slide, "Distance–time: never decreases", XL, Y0 + 1.45, CW, 0.50, 16)
bullet(slide, "Displacement–time: may decrease (sign / direction change)", XL, Y0 + 2.05, CW, 0.55, 16)
ncert(slide, "Do not mix the two graph rules.", XL, Y0 + 2.85, CW, 0.90, 16)

# ===== MORE PRACTICE =====
slide = new_slide("Example", "Square from centre O", BLUE)
para(slide, "From centre O of a square side ℓ, go to a corner, then walk three sides to the opposite corner B.", XL, Y0, CW, 1.00, 15)
fml(slide, "distance = ℓ/√2 + 3ℓ", XL, Y0 + 1.20, CW, LH, 16)
fml(slide, "|displacement| = ℓ/√2   (O to opposite corner is also half-diagonal)", XL, Y0 + 1.20 + LH + NG, CW, LH, 14)
# square with center
sq = 2.6
ax, ay = 8.2, 5.6
corners = [(ax, ay), (ax + sq, ay), (ax + sq, ay - sq), (ax, ay - sq)]
for a, b in zip(corners, corners[1:] + [corners[0]]):
    draw_line(slide, a[0], a[1], b[0], b[1], WHITE, 1.6)
ox = ax + sq / 2
oy = ay - sq / 2
small_dot(slide, ox, oy, GOLD)
tx(slide, "O", ox + 0.1, oy - 0.05, 0.35, 0.25, 12, GOLD, True, spa=0, scaled=False)
arrow_line(slide, ox, oy, corners[0][0], corners[0][1], ORANGE, 2.0)

slide = new_slide("Example", "Pendulum release to lowest point", ORANGE)
para(slide, "Bob released from angle θ with vertical length ℓ. Ends at the bottom (ℓ, 0) if we set coordinates from the support carefully — use start (ℓ cosθ, ℓ sinθ) → (ℓ, 0) in the usual school sketch.", XL, Y0, CW, 1.20, 14)
fml(slide, "Δr = ℓ(1 − cosθ) î − ℓ sinθ ĵ", XL, Y0 + 1.40, CW, LH, 15)
# pendulum sketch
sx0, sy0 = 9.5, 2.6
draw_line(slide, sx0, sy0, sx0, sy0 + 2.2, MUTED, 1.4)
draw_line(slide, sx0, sy0, sx0 + 1.5, sy0 + 1.8, ORANGE, 2.2)
small_dot(slide, sx0 + 1.5, sy0 + 1.8, GOLD)
small_dot(slide, sx0, sy0 + 2.2, GREEN)
tx(slide, "θ", sx0 + 0.25, sy0 + 0.55, 0.4, 0.25, 13, GOLD, True, spa=0, scaled=False)

slide = new_slide("Lock lines", "Say these aloud", GREEN)
bullet(slide, "Distance = path sum. Displacement = r_f − r_i.", XL, Y0, CW, 0.48, 16)
bullet(slide, "Same unit ≠ same quantity (scalar vs vector).", XL, Y0 + 0.52, CW, 0.48, 16)
bullet(slide, "Circle: distance = Rθ , |disp| = 2R sin(θ/2).", XL, Y0 + 1.04, CW, 0.50, 16)
bullet(slide, "Hexagon / square chords beat long perimeter walks.", XL, Y0 + 1.56, CW, 0.50, 16)
bullet(slide, "x–t is not the road — read Δx carefully.", XL, Y0 + 2.08, CW, 0.48, 16)
bullet(slide, "d–t never falls; s–t / x–t may fall.", XL, Y0 + 2.60, CW, 0.48, 16)
ncert(slide, "Next: speed and velocity (PW 03 / 04).", XL, Y0 + 3.35, CW, 0.90, 14)

slide = prs.slides.add_slide(BLANK)
_bg(slide)
_bar(slide, "End", GOLD)
tx_3d(slide, "Thank you", XL, 2.4, CW, 1.0, 48, GOLD, RGBColor(120, 120, 130))
tx(slide, "Excellence  ·  Motion 02  ·  Distance · Displacement · Arcs · Graphs", XL, 3.7, CW, 0.45, 15, MUTED, True, PP_ALIGN.CENTER, spa=0, scaled=False)
_footer(slide, len(prs.slides))
after_content_slide()

OUTPUT.parent.mkdir(parents=True, exist_ok=True)
prs.save(OUTPUT)
print(OUTPUT)
print("slides", len(prs.slides))
