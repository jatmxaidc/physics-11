"""Units 05 — dimensional questions from Units 04 slide 49, then error combination (PW 07)."""
from __future__ import annotations

import math
import re
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt

BASE = Path(__file__).resolve().parent
OUTPUT = BASE / "lecture slides" / "Lecture_Units_05_Error_Combination.pptx"

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
    r = p.add_run()
    r.text = text
    r.font.name = font
    r.font.size = Pt(size if not scaled else max(7, round(size * PT_SCALE, 1)))
    r.font.bold = bold
    r.font.color.rgb = color
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









# COVER
slide=prs.slides.add_slide(BLANK); _bg(slide)
tx_3d(slide,"EXCELLENCE",XL,0.52,CW,1.10,62,GOLD,RGBColor(210,210,220))
tx_3d(slide,"CLASS 11   -   2026-27",XL,1.64,CW,0.58,30,WHITE,RGBColor(90,90,100),font="Aptos")
_solid(slide,4.15,2.28,5.00,0.028,WHITE); bx,by,bw=2.05,2.70,9.20
_card_bg(slide,bx,by,bw,2.35,WHITE,WHITE,0); _card_bg(slide,bx,by+0.72,bw,1.63,BLACK,RGBColor(220,220,220),1.25)
_solid(slide,bx+bw/2,by+0.16,0.018,0.42,RGBColor(40,40,40))
tx(slide,"Physics",bx,by+0.14,bw/2-0.10,0.50,25,BLACK,True,PP_ALIGN.CENTER,spa=0,scaled=False)
tx(slide,"Units 05",bx+bw/2,by+0.14,bw/2,0.50,25,BLACK,True,PP_ALIGN.CENTER,spa=0,scaled=False)
tx(slide,"Units and Measurement",bx,by+1.02,bw,0.62,32,WHITE,True,PP_ALIGN.CENTER,spa=0,scaled=False,font="Georgia")
tx(slide,"Dimensional Questions and Error Combination",bx,by+1.62,bw,0.40,15,GOLD,True,PP_ALIGN.CENTER,spa=0,scaled=False)
tx(slide,"By - Excellence Programme",XL,5.28,CW,0.50,25,WHITE,True,PP_ALIGN.CENTER,spa=0,scaled=False)
tx(slide,"16 October 2026",XL,5.80,CW,0.32,18,MUTED,True,PP_ALIGN.CENTER,spa=0,scaled=False)
_footer(slide,1); add_notes(slide,"Starts at Units 04 slide 49. Then PW Units and Measurement Lecture 07."); after_content_slide()

slide=prs.slides.add_slide(BLANK); _bg(slide)
tx(slide,"Topics to be covered",XL,0.35,CW,0.64,32,GOLD,True,spa=0,scaled=False); _solid(slide,XL,1.00,3.40,0.035,GOLD)
topics=[
"Dimensional questions continued from Units 04",
"Error in a sum and a difference",
"Error in a product, a quotient and a power",
"Percentage error in energy, density and g",
"Error in a sine, a lens and an exponential",
]
for i,t in enumerate(topics):
    y=1.22+i*0.72
    circ=slide.shapes.add_shape(MSO_SHAPE.OVAL,Inches(XL),Inches(y),Inches(0.42),Inches(0.42))
    circ.fill.solid(); circ.fill.fore_color.rgb=TEAL; circ.line.fill.background()
    tx(slide,f"{i+1:02d}",XL,y+0.05,0.42,0.32,14,WHITE,True,PP_ALIGN.CENTER,spa=0,scaled=False)
    tx(slide,t,XL+0.60,y+0.05,10.5,0.35,20,TEXT,False,spa=0,scaled=False)
_footer(slide,2); after_content_slide()

slide=new_slide("Class note","Where we left Units 04",TEAL)
para(slide,"Last class we finished significant figures, accuracy, precision, and the mean error of a pendulum. The next board in that deck is slide 49: matching powers of M, L and T.",XL,Y0,CW,1.35,16)
para(slide,"We finish those dimensional questions first. After that we learn how an error grows when two measured quantities are added, multiplied, divided, or raised to a power.",XL,Y0+1.50,CW,1.30,16)
ncert(slide,"Write the formula first. The formula tells you whether to add absolute errors or relative errors.",XL,Y0+3.00,CW,0.90,16)
trap_card(slide,"Class notes will be saved after teaching. This deck is for the board.",XL,Y0+4.15,CW,0.45)

# CONTINUED FROM UNITS 04 SLIDE 49

slide=new_slide("Recall","Matching powers still works",BLUE)
para(slide,"Write the unknown as a product of powers. Put dimensions on both sides. Equate the powers of M, of L and of T separately. A pure number such as 2π or 6π never comes out of this method.",XL,Y0,CW,1.40,16)
para(slide,"If two quantities in the list have the same dimensions, they cannot both be used as separate base quantities. If a quantity is an angle, it is a pure number and cannot build a force or a time by itself.",XL,Y0+1.55,CW,1.30,16)
ncert(slide,"Same dimensions, or a pure angle, means the method stops.",XL,Y0+3.05,CW,0.80,16)

slide=question_slide("Electric field and magnetic field")
qtx(slide,"The relation between [E] and [B] is",XL,Y0,CW,0.45,18)
choices(slide,[
"[E] = [B] [L] [T]",
"[E] = [B] [L]^(-1) [T]",
"[E] = [B] [L] [T]^(-1)",
"[E] = [B] [L]^(-1) [T]^(-1)",
], XL, Y0+0.65, 10.8)
slide=answer_slide()
body(slide,"In an electromagnetic wave, E = c B. Speed c carries the extra dimensions.",XL,Y0,CW,0.50)
fml(slide,"[c] = L T^(-1), so [E] = [B] × L T^(-1).",XL,Y0+0.60,CW,LH,16)
fml(slide,"That is [E] = [B] [L] [T]^(-1).",XL,Y0+0.60+LH+NG,CW,LH,17)
result_card(slide,"Final answer: option 3",XL,Y0+0.60+2*LH+2*NG,CW,0.48)

slide=question_slide("Permittivity and permeability")
qtx(slide,"The relation between [ε_0] and [μ_0] is",XL,Y0,CW,0.45,18)
choices(slide,[
"[μ_0] = [ε_0] [L]^2 [T]^(-2)",
"[μ_0] = [ε_0] [L]^(-2) [T]^2",
"[μ_0] = [ε_0]^(-1) [L]^2 [T]^(-2)",
"[μ_0] = [ε_0]^(-1) [L]^(-2) [T]^2",
], XL, Y0+0.60, 11.0)
slide=answer_slide()
body(slide,"Use c = 1 / √(μ_0 ε_0). Square both sides.",XL,Y0,CW,0.45)
fml(slide,"μ_0 ε_0 = 1 / c^2, so [μ_0] [ε_0] = L^(-2) T^2.",XL,Y0+0.55,CW,LH,16)
fml(slide,"[μ_0] = [ε_0]^(-1) [L]^(-2) [T]^2.",XL,Y0+0.55+LH+NG,CW,LH,16)
result_card(slide,"Final answer: option 4",XL,Y0+0.55+2*LH+2*NG,CW,0.48)

slide=question_slide("Same dimensions as S")
qtx(slide,"S = (E × B) / μ_0. Its dimensions match which of these?",XL,Y0,CW,0.55,17)
choices(slide,[
"Energy / (charge × current)",
"Force / (length × time)",
"Energy / volume",
"Power / area",
], XL, Y0+0.75, 10.8)
slide=answer_slide()
body(slide,"Build [S] from the known formulas, then compare.",XL,Y0,CW,0.42)
fml(slide,"[E] = M L T^(-3) I^(-1),  [B] = M T^(-2) I^(-1),  [μ_0] = M L T^(-2) I^(-2).",XL,Y0+0.50,CW,0.65,14)
fml(slide,"[E B / μ_0] = M T^(-3).",XL,Y0+1.25,CW,LH,17)
fml(slide,"Force / (length × time) = M T^(-3).  Power / area = M T^(-3).",XL,Y0+1.25+LH+NG,CW,0.60,15)
result_card(slide,"Final answer: options 2 and 4",XL,Y0+2.50,CW,0.48)

slide=question_slide("Mass and angular momentum are pure numbers")
qtx(slide,"In a new system, mass and angular momentum are dimensionless, and length has dimension L. Which statements are correct?",XL,Y0,CW,0.85,16)
choices(slide,[
"Dimension of force is L^(-3)",
"Dimension of energy is L^(-2)",
"Dimension of power is L^(-5)",
"Dimension of linear momentum is L^(-1)",
], XL, Y0+1.05, 10.8)
slide=answer_slide()
body(slide,"Angular momentum m v r is a pure number, and m is a pure number, so velocity must cancel the length.",XL,Y0,CW,0.60)
fml(slide,"[v] = L^(-1).  Then [time] = length / speed = L^2.",XL,Y0+0.70,CW,LH,15)
fml(slide,"[force] = [mass × acceleration] = L^(-3).  [energy] = force × length = L^(-2).",XL,Y0+0.70+LH+NG,CW,0.60,14)
fml(slide,"[power] = energy / time = L^(-4), not L^(-5).  [momentum] = m v = L^(-1).",XL,Y0+1.45+LH+NG,CW,0.60,14)
result_card(slide,"Correct: options 1, 2 and 4.  Option 3 is wrong.",XL,Y0+2.70,CW,0.48)

slide=question_slide("Young’s modulus in F, A and V")
qtx(slide,"Speed V, area A and force F are the base quantities. [Young’s modulus] is",XL,Y0,CW,0.70,17)
choices(slide,["F A^(-1) V^0","F A^2 V^(-1)","F A^2 V^(-2)","F A^2 V^(-3)"], XL, Y0+0.90, 10.5)
slide=answer_slide()
body(slide,"Young’s modulus is stress / strain, and strain is a pure number.",XL,Y0,CW,0.48)
fml(slide,"Y = F / A.",XL,Y0+0.58,CW,LH,18)
fml(slide,"Speed is not needed, so its power is 0.",XL,Y0+0.58+LH+NG,CW,LH,16)
result_card(slide,"Final answer: F A^(-1) V^0   (option 1)",XL,Y0+0.58+2*LH+2*NG,CW,0.48)

slide=question_slide("Density from force, velocity and time")
qtx(slide,"F, V and T are base quantities. The dimensional formula of density is",XL,Y0,CW,0.65,17)
choices(slide,["F V^4 T^(-6)","F V^(-4) T^(-2)","F^2 V^(-2) T^6","F V^(-2) T^2"], XL, Y0+0.85, 10.5)
slide=answer_slide()
body(slide,"Assume density = F^a V^b T^c and match M, L, T.",XL,Y0,CW,0.45)
fml(slide,"Mass: a = 1.  Length: a + b = -3, so b = -4.",XL,Y0+0.55,CW,LH,16)
fml(slide,"Time: -2a - b + c = 0 → -2 - (-4) + c = 0 → c = -2.",XL,Y0+0.55+LH+NG,CW,LH,15)
result_card(slide,"Final answer: F V^(-4) T^(-2)   (option 2)",XL,Y0+0.55+2*LH+2*NG,CW,0.48)

slide=question_slide("Viscosity from momentum, area and time")
qtx(slide,"P, A and T are base quantities. The dimensional formula of viscosity is",XL,Y0,CW,0.65,17)
choices(slide,["[P A^(-1)]","[P A T^(-1)]","[P A^(-1) T]","[P A^(-1) T^(-1)]"], XL, Y0+0.85, 10.5)
slide=answer_slide()
body(slide,"From F = 6π η r v, the dimensions of η are M L^(-1) T^(-1).",XL,Y0,CW,0.50)
fml(slide,"η = P^a A^b T^c.  Mass gives a = 1.",XL,Y0+0.60,CW,LH,16)
fml(slide,"Length: 1 + 2b = -1, so b = -1.  Time: -1 + c = -1, so c = 0.",XL,Y0+0.60+LH+NG,CW,LH,15)
result_card(slide,"Final answer: [P A^(-1)]   (option 1)",XL,Y0+0.60+2*LH+2*NG,CW,0.48)

slide=question_slide("Mass from time, velocity and angular momentum")
qtx(slide,"Express mass using time t, velocity v and angular momentum J.",XL,Y0,CW,0.60,17)
slide=answer_slide()
body(slide,"[J] = M L^2 T^(-1). Assume M = J^a v^b t^c.",XL,Y0,CW,0.48)
fml(slide,"Mass: a = 1.  Length: 2 + b = 0, so b = -2.",XL,Y0+0.58,CW,LH,16)
fml(slide,"Time: -1 - (-2) + c = 0, so c = -1.",XL,Y0+0.58+LH+NG,CW,LH,16)
result_card(slide,"Mass = J / (v^2 t)",XL,Y0+0.58+2*LH+2*NG,CW,0.50)

slide=question_slide("Young’s modulus from c, h and G")
qtx(slide,"Young’s modulus is a product of powers of c, h and G. Which powers are correct?",XL,Y0,CW,0.55,17)
choices(slide,[
"α = 7, β = -1, γ = -2",
"α = -7, β = -1, γ = -2",
"α = 7, β = -1, γ = 2",
"α = -7, β = 1, γ = -2",
], XL, Y0+0.70, 10.5)
slide=answer_slide("Solution: equations for Y")
body(slide,"[Y] = M L^(-1) T^(-2),  [c] = L T^(-1),  [h] = M L^2 T^(-1),  [G] = M^(-1) L^3 T^(-2).",XL,Y0,CW,0.70)
fml(slide,"Mass: β - γ = 1.",XL,Y0+0.85,CW,LH,17)
fml(slide,"Length: α + 2β + 3γ = -1.",XL,Y0+0.85+LH+NG,CW,LH,17)
fml(slide,"Time: -α - β - 2γ = -2, so α + β + 2γ = 2.",XL,Y0+0.85+2*LH+2*NG,CW,LH,15)
slide=answer_slide("Solution: the three powers")
body(slide,"From mass, β = 1 + γ. Put that into the time equation.",XL,Y0,CW,0.48)
fml(slide,"α + (1 + γ) + 2γ = 2  →  α + 3γ = 1.",XL,Y0+0.58,CW,LH,16)
fml(slide,"Length then gives α + 5γ = -3. Subtract: 2γ = -4, so γ = -2.",XL,Y0+0.58+LH+NG,CW,0.60,15)
fml(slide,"α = 1 - 3(-2) = 7.  β = 1 + (-2) = -1.",XL,Y0+1.30+LH+NG,CW,LH,16)
result_card(slide,"α = 7, β = -1, γ = -2   (option 1)",XL,Y0+2.40,CW,0.48)

slide=question_slide("A leaking balloon")
qtx(slide,"Gas speed depends on the radius through a power a. The emptying time T depends on surface tension S, outlet area A, density ρ and starting radius R through powers α, β, γ and δ.",XL,Y0,CW,1.05,16)
choices(slide,[
"a = 1/2, α = 1/2, β = -1, γ = 1, δ = 3/2",
"a = -1/2, α = -1/2, β = -1, γ = -1/2, δ = 5/2",
"a = -1/2, α = -1/2, β = -1, γ = 1/2, δ = 7/2",
"a = 1/2, α = 1/2, β = -1/2, γ = 1/2, δ = 7/2",
], XL, Y0+1.20, 12.0)
slide=answer_slide("Solution: speed of the gas")
body(slide,"The speed of the leaving gas is fixed by S, ρ and the present radius r.",XL,Y0,CW,0.50)
fml(slide,"v ∝ S^(1/2) ρ^(-1/2) r^(-1/2).",XL,Y0+0.60,CW,LH,16)
fml(slide,"Mass cancels. What remains is a speed, and the power of r is -1/2. So a = -1/2.",XL,Y0+0.60+LH+NG,CW,0.70,15)
slide=answer_slide("Solution: time to empty")
body(slide,"Volume leaves through area A, so the radius falls as dr/dt ∝ A v / r^2.",XL,Y0,CW,0.55)
fml(slide,"With v ∝ r^(-1/2), we get dr/dt ∝ A S^(1/2) ρ^(-1/2) r^(-5/2).",XL,Y0+0.65,CW,0.65,14)
fml(slide,"Time to fall from R to 0 is T ∝ R^(7/2) / (A S^(1/2) ρ^(-1/2)).",XL,Y0+1.40,CW,0.60,15)
fml(slide,"So α = -1/2, β = -1, γ = 1/2, δ = 7/2.",XL,Y0+2.10,CW,LH,16)
result_card(slide,"Final answer: option 3",XL,Y0+2.65,CW,0.45)

slide=new_slide("Stop","Three cases where dimensions refuse",ORANGE)
bullet(slide,"Torque, energy and work have the same dimensions. Using all three as base quantities cannot produce a new formula for time.",XL,Y0,CW,0.85,16)
bullet(slide,"If g is said to depend on force, energy, mass and velocity, that is four quantities and only three equations (M, L, T). The powers are not unique.",XL,Y0+1.05,CW,0.90,16)
bullet(slide,"If force is said to depend on an angle, a velocity and an energy, the angle is a pure number. It cannot supply the dimensions of force.",XL,Y0+2.15,CW,0.90,16)
trap_card(slide,"Write “not possible” and name the reason. That is a full answer.",XL,Y0+3.20,CW,0.48)

slide=question_slide("Which equation cannot be fully derived")
qtx(slide,"Which equation cannot be obtained from dimensions alone?",XL,Y0,CW,0.50,17)
choices(slide,[
"F = 6π η r v",
"θ = ω t",
"P = ρ g h",
"v = d / t",
], XL, Y0+0.70, 10.5)
slide=answer_slide()
body(slide,"Dimensions can give the powers. They cannot give a pure number such as 6π.",XL,Y0,CW,0.55)
fml(slide,"F, η, r and v can be matched, but 6π stays unknown.",XL,Y0+0.65,CW,LH,16)
fml(slide,"θ = ω t, P = ρ g h and v = d / t are products of powers. Their number in front is 1.",XL,Y0+0.65+LH+NG,CW,0.70,15)
result_card(slide,"Final answer: F = 6π η r v   (option 1)",XL,Y0+2.15,CW,0.50)


# ERROR COMBINATION
slide=new_slide("Idea","Which error to write first",GREEN)
para(slide,"An absolute error has the same unit as the quantity. A relative error is the absolute error divided by the measured value, so it has no unit. Percentage error is the relative error times 100.",XL,Y0,CW,1.35,16)
para(slide,"For a sum or a difference, write the absolute error directly. For a product, a quotient, or a power, write the relative error directly, then multiply by the answer if the absolute error is asked.",XL,Y0+1.50,CW,1.30,16)
ncert(slide,"Report the result as (value ± absolute error) with the unit.",XL,Y0+3.00,CW,0.80,16)

slide=new_slide("Rule","Sum and difference",TEAL)
para(slide,"If Y = A + B, or Y = A − B, the maximum absolute error is the sum of the two absolute errors. The errors are added even when the quantities are subtracted.",XL,Y0,CW,1.25,16)
fml(slide,"ΔY = ΔA + ΔB",XL,Y0+1.40,CW,LH,20)
fml(slide,"Relative error of a sum = (ΔA + ΔB) / (A + B)",XL,Y0+1.40+LH+NG,CW,LH,16)
fml(slide,"Relative error of a difference = (ΔA + ΔB) / (A − B)",XL,Y0+1.40+2*LH+2*NG,CW,LH,16)
trap_card(slide,"Two close numbers subtracted leave a small answer and a large percentage error.",XL,Y0+3.55,CW,0.50)

slide=new_slide("Rule","Product, quotient and power",BLUE)
para(slide,"If Y = A B, or Y = A / B, the relative errors are added. The same rule covers a power: each relative error is multiplied by the size of its power.",XL,Y0,CW,1.20,16)
fml(slide,"ΔY / Y = ΔA / A + ΔB / B",XL,Y0+1.35,CW,LH,18)
fml(slide,"Y = k A^m B^n    →    ΔY / Y = |m| ΔA/A + |n| ΔB/B",XL,Y0+1.35+LH+NG,CW,LH,16)
ncert(slide,"A fixed number such as 1/2, 2 or π has no measurement error. Its power still multiplies the relative error of the measured quantity.",XL,Y0+2.85,CW,1.05,16)

slide=new_slide("Method","From the derivative",ORANGE)
para(slide,"A small change follows the derivative. If y = x^n, then Δy / y = |n| Δx / x. That is the power rule written in one line.",XL,Y0,CW,1.15,16)
fml(slide,"y = sin θ    →    Δy = |cos θ| Δθ, with Δθ in radians",XL,Y0+1.30,CW,LH,16)
fml(slide,"y = 1/x    →    Δy / y = Δx / x",XL,Y0+1.30+LH+NG,CW,LH,16)
ncert(slide,"Take every power as a positive number when you want the maximum error.",XL,Y0+2.90,CW,0.80,16)

slide=question_slide("Percentage error in X")
qtx(slide,"X = M^a L^b T^(-c). Percentage errors in M, L and T are α, β and γ, and they are small. The maximum percentage error in X is",XL,Y0,CW,0.90,16)
choices(slide,[
"(aα − bβ + cγ) %",
"(aα + bβ + cγ) %",
"(aα + bβ − cγ) %",
"(α + β + γ) %",
], XL, Y0+1.10, 10.8)
slide=answer_slide()
body(slide,"The power of T is −c. For the maximum error we use the size of that power, which is c.",XL,Y0,CW,0.65)
fml(slide,"ΔX / X × 100 = |a|α + |b|β + |−c|γ.",XL,Y0+0.80,CW,LH,16)
fml(slide,"With positive a, b and c this is (aα + bβ + cγ) %.",XL,Y0+0.80+LH+NG,CW,LH,16)
result_card(slide,"Final answer: option 2",XL,Y0+0.80+2*LH+2*NG,CW,0.48)

slide=question_slide("A product of powers")
qtx(slide,"P = a^2 b^3 / (c √d). Percentage errors in a, b, c and d are 1%, 2%, 3% and 4%. The percentage error in P is",XL,Y0,CW,0.90,16)
choices(slide,["13 %","16 %","12 %","10 %"], XL, Y0+1.10, 10.5)
slide=answer_slide()
body(slide,"√d means d raised to 1/2. Each percentage is multiplied by the size of its power.",XL,Y0,CW,0.60)
fml(slide,"2×1 + 3×2 + 1×3 + (1/2)×4",XL,Y0+0.75,CW,LH,18)
fml(slide,"= 2 + 6 + 3 + 2 = 13.",XL,Y0+0.75+LH+NG,CW,LH,18)
result_card(slide,"Final answer: 13 %   (option 1)",XL,Y0+0.75+2*LH+2*NG,CW,0.48)

slide=question_slide("Area of a rectangular plate")
qtx(slide,"A plate is (2.00 ± 0.02) cm by (1.00 ± 0.01) cm. The percentage error in its area is",XL,Y0,CW,0.70,17)
choices(slide,["1 %","2 %","3 %","4 %"], XL, Y0+0.90, 10.5)
slide=answer_slide()
body(slide,"Area = length × breadth, so the relative errors are added.",XL,Y0,CW,0.50)
fml(slide,"ΔA / A = 0.02/2.00 + 0.01/1.00 = 0.01 + 0.01.",XL,Y0+0.65,CW,LH,16)
fml(slide,"Percentage error = 0.02 × 100 = 2 %.",XL,Y0+0.65+LH+NG,CW,LH,16)
result_card(slide,"Final answer: 2 %   (option 2)",XL,Y0+0.65+2*LH+2*NG,CW,0.48)

slide=question_slide("Kinetic energy")
qtx(slide,"m = (5.0 ± 0.5) kg and v = (20.0 ± 0.4) m/s. The kinetic energy with its error is",XL,Y0,CW,0.70,17)
choices(slide,[
"(1000 ± 140) J",
"(1000 ± 0.14) J",
"(500 ± 140) J",
"(1000 ± 70) J",
], XL, Y0+0.90, 10.8)
slide=answer_slide()
body(slide,"K = (1/2) m v^2. The 1/2 is exact, so it adds no error term. The power of v is 2.",XL,Y0,CW,0.65)
fml(slide,"K = (1/2)×5×400 = 1000 J.",XL,Y0+0.80,CW,LH,17)
fml(slide,"ΔK / K = 0.5/5 + 2×(0.4/20) = 0.10 + 0.04 = 0.14.",XL,Y0+0.80+LH+NG,CW,LH,15)
fml(slide,"ΔK = 0.14 × 1000 = 140 J.",XL,Y0+0.80+2*LH+2*NG,CW,LH,17)
result_card(slide,"Final answer: (1000 ± 140) J   (option 1)",XL,Y0+2.70,CW,0.48)

slide=question_slide("Side of a square park")
qtx(slide,"The area of a square park is (100 ± 0.2) m^2. The side of the park is",XL,Y0,CW,0.65,17)
choices(slide,[
"(10.00 ± 0.01) m",
"(10.00 ± 0.02) m",
"(10.0 ± 0.1) m",
"(10.00 ± 0.20) m",
], XL, Y0+0.85, 10.8)
slide=answer_slide()
body(slide,"Find the side first. Area = side^2, so the relative error in the side is half the relative error in the area.",XL,Y0,CW,0.70)
fml(slide,"Side = √100 = 10 m.",XL,Y0+0.85,CW,LH,17)
fml(slide,"Δl / l = (1/2) × (0.2 / 100) = 0.001.",XL,Y0+0.85+LH+NG,CW,LH,16)
fml(slide,"Δl = 0.001 × 10 = 0.01 m.",XL,Y0+0.85+2*LH+2*NG,CW,LH,17)
result_card(slide,"Final answer: (10.00 ± 0.01) m   (option 1)",XL,Y0+2.75,CW,0.48)

slide=question_slide("Pressure")
qtx(slide,"P = F / L^2. The percentage error in F is 4 % and the percentage error in L is 2 %. The percentage error in P is",XL,Y0,CW,0.80,16)
choices(slide,["2 %","4 %","6 %","8 %"], XL, Y0+1.00, 10.5)
slide=answer_slide()
body(slide,"L is squared, so its percentage error is doubled and then added to the percentage error in F.",XL,Y0,CW,0.65)
fml(slide,"ΔP / P × 100 = 4 + 2×2 = 8.",XL,Y0+0.80,CW,LH,18)
result_card(slide,"Final answer: 8 %   (option 4)",XL,Y0+0.80+LH+NG,CW,0.48)

slide=question_slide("Density of a wire")
qtx(slide,"ρ = m / (π r^2 L). m = (0.4 ± 0.01) g, L = (8 ± 0.04) cm, r = (6 ± 0.03) mm. The percentage error in density is",XL,Y0,CW,0.85,16)
choices(slide,["1 %","2 %","4 %","5 %"], XL, Y0+1.05, 10.5)
slide=answer_slide()
body(slide,"π is exact. Each relative error uses its own unit, so millimetres and centimetres can stay as they are.",XL,Y0,CW,0.70)
fml(slide,"(0.01/0.4)×100 + 2×(0.03/6)×100 + (0.04/8)×100",XL,Y0+0.85,CW,LH,15)
fml(slide,"= 2.5 + 1.0 + 0.5 = 4.",XL,Y0+0.85+LH+NG,CW,LH,18)
result_card(slide,"Final answer: 4 %   (option 3)",XL,Y0+0.85+2*LH+2*NG,CW,0.48)

slide=question_slide("Relative error in Z")
qtx(slide,"Z = A^2 B^3 / C^4. The maximum relative error in Z is",XL,Y0,CW,0.60,17)
choices(slide,[
"ΔA/A + ΔB/B + ΔC/C",
"2ΔA/A + 3ΔB/B − 4ΔC/C",
"2ΔA/A + 3ΔB/B + 4ΔC/C",
"2ΔA/A + 3ΔB/B + ΔC/C",
], XL, Y0+0.80, 11.2)
slide=answer_slide()
body(slide,"The power of C is −4. For the maximum error that power is taken as 4, and the term is added.",XL,Y0,CW,0.65)
fml(slide,"ΔZ / Z = 2 ΔA/A + 3 ΔB/B + 4 ΔC/C.",XL,Y0+0.80,CW,LH,16)
result_card(slide,"Final answer: option 3",XL,Y0+0.80+LH+NG,CW,0.48)

slide=question_slide("Torque")
qtx(slide,"Torque has dimensions M L^2 T^(-2). If M, L and T are each measured to 5 % accuracy, the percentage error in torque is",XL,Y0,CW,0.80,16)
choices(slide,["5 %","25 %","15 %","10 %"], XL, Y0+1.00, 10.5)
slide=answer_slide()
body(slide,"The powers are 1 for mass, 2 for length and 2 for time.",XL,Y0,CW,0.50)
fml(slide,"1×5 + 2×5 + 2×5 = 5 + 10 + 10 = 25.",XL,Y0+0.65,CW,LH,17)
result_card(slide,"Final answer: 25 %   (option 2)",XL,Y0+0.65+LH+NG,CW,0.48)

slide=question_slide("g from a pendulum")
qtx(slide,"In T = 2π √(l/g), the relative errors in l and T are α and β. The percentage error in g is",XL,Y0,CW,0.75,16)
choices(slide,[
"(α + β) × 100",
"(α + β/2) × 100",
"(α/2 + β) × 100",
"(α + 2β) × 100",
], XL, Y0+0.95, 10.8)
slide=answer_slide()
body(slide,"Square the period formula and solve for g. The 4π^2 is exact.",XL,Y0,CW,0.55)
fml(slide,"g = 4π^2 l / T^2.",XL,Y0+0.70,CW,LH,18)
fml(slide,"Δg / g = Δl/l + 2 ΔT/T = α + 2β.",XL,Y0+0.70+LH+NG,CW,LH,16)
fml(slide,"Percentage error = (α + 2β) × 100.",XL,Y0+0.70+2*LH+2*NG,CW,LH,16)
result_card(slide,"Final answer: option 4",XL,Y0+2.85,CW,0.48)

slide=question_slide("A measured pendulum")
qtx(slide,"l = 10 cm, known to 1 mm. The time for 200 oscillations is 100 s, read on a clock of 1 s resolution. The percentage error in g is",XL,Y0,CW,0.90,16)
choices(slide,["3 %","5 %","4 %","2 %"], XL, Y0+1.10, 10.5)
slide=answer_slide()
body(slide,"The number of oscillations is exact, so the relative error in the period equals the relative error in the total time.",XL,Y0,CW,0.70)
fml(slide,"Δl / l = 0.1 cm / 10 cm = 0.01.",XL,Y0+0.85,CW,LH,16)
fml(slide,"ΔT / T = 1 s / 100 s = 0.01.",XL,Y0+0.85+LH+NG,CW,LH,16)
fml(slide,"Δg / g = 0.01 + 2×0.01 = 0.03, so the percentage error is 3 %.",XL,Y0+0.85+2*LH+2*NG,CW,0.65,15)
result_card(slide,"Final answer: 3 %   (option 1). The fractional error is 0.03.",XL,Y0+3.05,CW,0.55)

slide=question_slide("Error in sin θ")
qtx(slide,"y = sin θ. At θ = π/6 the percentage error in θ is 2 %. Find the percentage error in y.",XL,Y0,CW,0.75,17)
slide=answer_slide()
body(slide,"Δy = |cos θ| Δθ, and Δθ must be in radians. A 2 % error means Δθ = 0.02 × θ.",XL,Y0,CW,0.65)
fml(slide,"Δy / y = |cot θ| × Δθ = √3 × 0.02 × (π/6).",XL,Y0+0.80,CW,LH,15)
fml(slide,"Percentage error = √3 × 0.02 × (π/6) × 100 = π/√3 %.",XL,Y0+0.80+LH+NG,CW,0.65,15)
result_card(slide,"Final answer: π/√3 %  (about 1.81 %)",XL,Y0+2.35,CW,0.50)

slide=question_slide("Where the percentage error is smallest")
qtx(slide,"y = m tan θ, and m is exact. The absolute error in θ is the same at every angle. The percentage error in y is smallest at",XL,Y0,CW,0.85,16)
choices(slide,["θ = 30°","θ = 45°","θ = 60°","θ = 90°"], XL, Y0+1.05, 10.5)
slide=answer_slide()
body(slide,"Differentiate y = m tan θ. The constant m cancels in the relative error.",XL,Y0,CW,0.55)
fml(slide,"Δy / y = 2 Δθ / sin 2θ.",XL,Y0+0.70,CW,LH,18)
fml(slide,"With Δθ fixed, this is smallest when sin 2θ = 1.",XL,Y0+0.70+LH+NG,CW,LH,16)
fml(slide,"2θ = 90°, so θ = 45°.",XL,Y0+0.70+2*LH+2*NG,CW,LH,17)
result_card(slide,"Final answer: θ = 45°   (option 2)",XL,Y0+2.85,CW,0.48)

slide=new_slide("Worked","Series, product and difference",TEAL)
body(slide,"R_1 = (10 ± 1) Ω and R_2 = (20 ± 1) Ω.",XL,Y0,CW,0.40)
fml(slide,"Series: R = 30 Ω,  ΔR = 2 Ω,  percentage = (2/30)×100 = 6.67 %.",XL,Y0+0.55,CW,0.70,15)
fml(slide,"Product: R_1 R_2 = 200,  relative error = 0.10 + 0.05 = 0.15.",XL,Y0+1.35,CW,0.60,15)
fml(slide,"Absolute error in the product = 200 × 0.15 = 30.",XL,Y0+2.05,CW,LH,16)
fml(slide,"Difference: R = 10 Ω,  ΔR = 2 Ω,  percentage = 20 %.",XL,Y0+2.05+LH+NG,CW,LH,16)
ncert(slide,"Quotient R_2 / R_1 = 2 also has relative error 0.15, so its absolute error is 0.30 Ω.",XL,Y0+3.40,CW,0.80,15)

slide=new_slide("Rule","Parallel resistors and a thin lens",BLUE)
para(slide,"For two resistors in parallel, 1/R = 1/R_1 + 1/R_2. Differentiating gives the absolute error in R from the squares of the two resistances.",XL,Y0,CW,1.15,16)
fml(slide,"ΔR = R^2 ( ΔR_1 / R_1^2 + ΔR_2 / R_2^2 )",XL,Y0+1.30,CW,LH,16)
para(slide,"A thin lens with a real image uses the same shape: 1/f = 1/u + 1/v, so Δf = f^2 (Δu/u^2 + Δv/v^2).",XL,Y0+2.15,CW,1.05,16)
ncert(slide,"Series capacitors, and 1/f = 1/u + 1/v, follow this same squared form.",XL,Y0+3.35,CW,0.75,16)

slide=question_slide("Two resistors in parallel")
qtx(slide,"R_1 = (20 ± 2) Ω and R_2 = (5 ± 1) Ω are in parallel. Find the percentage error and the absolute error in the equivalent resistance.",XL,Y0,CW,0.90,16)
slide=answer_slide()
body(slide,"First find the equivalent resistance, then use the parallel error formula.",XL,Y0,CW,0.50)
fml(slide,"R = (20×5)/(20+5) = 4 Ω.",XL,Y0+0.65,CW,LH,17)
fml(slide,"ΔR / R = 4 (2/20^2 + 1/5^2) = 4 (0.005 + 0.04) = 0.18.",XL,Y0+0.65+LH+NG,CW,0.65,15)
fml(slide,"Percentage error = 18 %.  Absolute error = 0.18 × 4 = 0.72 Ω.",XL,Y0+1.45+LH+NG,CW,0.65,15)
result_card(slide,"Final answer: 18 % and 0.72 Ω",XL,Y0+2.80,CW,0.48)

slide=question_slide("Focal length of a convex lens")
qtx(slide,"Object distance u = (10 ± 0.1) cm and image distance v = (20 ± 0.2) cm. The error in the focal length is n %. Find n.",XL,Y0,CW,0.85,16)
slide=answer_slide()
body(slide,"Use 1/f = 1/u + 1/v with both distances taken positive.",XL,Y0,CW,0.48)
fml(slide,"1/f = 1/10 + 1/20 = 3/20, so f = 20/3 cm.",XL,Y0+0.60,CW,LH,16)
fml(slide,"Δf / f = f (0.1/10^2 + 0.2/20^2) = (20/3)(0.001 + 0.0005).",XL,Y0+0.60+LH+NG,CW,0.65,14)
fml(slide,"(20/3)×0.0015 = 0.01, so the percentage error is 1 %.",XL,Y0+1.40+LH+NG,CW,0.60,15)
result_card(slide,"Final answer: n = 1",XL,Y0+2.70,CW,0.48)

slide=question_slide("Error in d from 2d sin θ = λ")
qtx(slide,"λ is known exactly, and the error in θ is the same at every angle. As θ increases from 0° to 90°, which statement is true?",XL,Y0,CW,0.80,16)
choices(slide,[
"The absolute error in d stays constant",
"The absolute error in d increases",
"The fractional error in d stays constant",
"The fractional error in d decreases",
], XL, Y0+1.00, 11.5)
slide=answer_slide()
body(slide,"d = λ / (2 sin θ). Differentiating with respect to θ gives the fractional error.",XL,Y0,CW,0.60)
fml(slide,"Δd / d = |cot θ| Δθ.",XL,Y0+0.75,CW,LH,18)
fml(slide,"As θ rises from 0° toward 90°, cot θ falls, so the fractional error falls.",XL,Y0+0.75+LH+NG,CW,0.65,15)
result_card(slide,"Final answer: option 4",XL,Y0+2.20,CW,0.48)

slide=question_slide("Error in the ratio r")
qtx(slide,"r = (1 − a)/(1 + a), and the error in the dimensionless quantity a is Δa. Then Δr equals",XL,Y0,CW,0.70,16)
choices(slide,[
"Δa / (1 + a)^2",
"−2 Δa / (1 + a)^2",
"2 Δa / (1 − a)^2",
"2 a Δa / (1 − a^2)",
], XL, Y0+0.90, 10.8)
slide=answer_slide()
body(slide,"Differentiate r with respect to a. The two minus signs in the quotient rule add.",XL,Y0,CW,0.60)
fml(slide,"dr/da = −2 / (1 + a)^2.",XL,Y0+0.75,CW,LH,18)
fml(slide,"Δr = −2 Δa / (1 + a)^2.",XL,Y0+0.75+LH+NG,CW,LH,18)
result_card(slide,"Final answer: option 2. The size of the error is 2 Δa / (1 + a)^2.",XL,Y0+2.30,CW,0.55)

slide=question_slide("Energy that falls with time")
qtx(slide,"E = A^2 exp(−α t), with α = 0.2 s^(-1). The error in A is 1.25 % and the error in t is 1.50 %. At t = 5 s, the percentage error in E is",XL,Y0,CW,0.95,16)
slide=answer_slide()
body(slide,"Take the logarithm. The power of A contributes 2. The exponential contributes α times the absolute error in t.",XL,Y0,CW,0.70)
fml(slide,"Δt = 0.015 × 5 = 0.075 s, so α Δt = 0.2 × 0.075 = 0.015.",XL,Y0+0.85,CW,0.65,15)
fml(slide,"Percentage from A: 2 × 1.25 = 2.5. From time: 1.5.",XL,Y0+1.60,CW,LH,16)
fml(slide,"Total percentage error = 2.5 + 1.5 = 4.",XL,Y0+1.60+LH+NG,CW,LH,17)
result_card(slide,"Final answer: 4 %",XL,Y0+3.05,CW,0.48)

slide=question_slide("Current in a diode")
qtx(slide,"I = (exp(1000 V/T) − 1) mA. At 300 K the current is 5 mA, and the voltage error is 0.01 V. The error in the current is",XL,Y0,CW,0.85,16)
choices(slide,["0.2 mA","0.02 mA","0.5 mA","0.05 mA"], XL, Y0+1.05, 10.5)
slide=answer_slide()
body(slide,"The measured current fixes the exponential. Then differentiate with respect to V.",XL,Y0,CW,0.55)
fml(slide,"exp(1000 V/T) = I + 1 = 6.",XL,Y0+0.70,CW,LH,17)
fml(slide,"ΔI = 6 × (1000/300) × 0.01 = 6 × (10/3) × 0.01.",XL,Y0+0.70+LH+NG,CW,0.65,15)
fml(slide,"ΔI = 0.2 mA.",XL,Y0+1.50+LH+NG,CW,LH,18)
result_card(slide,"Final answer: 0.2 mA   (option 1)",XL,Y0+2.85,CW,0.48)

slide=question_slide("g from a periodic motion")
qtx(slide,"T = 2π √[7(R − r)/(5g)]. R = (60 ± 1) mm, r = (10 ± 1) mm. Five periods are 0.52, 0.56, 0.57, 0.54 and 0.59 s. The watch reads to 0.01 s. Which statements are true?",XL,Y0,CW,1.15,15)
choices(slide,[
"The error in r is 10 %",
"The error in T is 3.57 %",
"The error in T is 2 %",
"The error in g is 11 %",
], XL, Y0+1.30, 10.8)
slide=answer_slide("Solution: the mean period")
body(slide,"Mean of the five times: (0.52 + 0.56 + 0.57 + 0.54 + 0.59)/5 = 0.556 s, taken as 0.56 s.",XL,Y0,CW,0.70)
fml(slide,"Mean absolute error ≈ 0.02 s, larger than the least count 0.01 s.",XL,Y0+0.85,CW,0.60,15)
fml(slide,"ΔT / T × 100 = (0.02 / 0.56) × 100 = 3.57 %.",XL,Y0+1.55,CW,LH,16)
fml(slide,"Error in r: (1/10)×100 = 10 %. The 2 % statement is false.",XL,Y0+1.55+LH+NG,CW,LH,15)
result_card(slide,"Statements 1 and 2 are true.",XL,Y0+3.05,CW,0.48)
slide=answer_slide("Solution: the error in g")
body(slide,"From the period formula, g depends on (R − r) and on T^2.",XL,Y0,CW,0.50)
fml(slide,"R − r = 50 mm,  Δ(R − r) = 1 + 1 = 2 mm, so 4 %.",XL,Y0+0.65,CW,LH,16)
fml(slide,"The T^2 term contributes 2 × 3.57 % = 7.14 %.",XL,Y0+0.65+LH+NG,CW,LH,16)
fml(slide,"Percentage error in g = 4 + 7.14 ≈ 11 %.",XL,Y0+0.65+2*LH+2*NG,CW,LH,16)
result_card(slide,"True statements: 1, 2 and 4.",XL,Y0+2.70,CW,0.48)

slide=new_slide("Rule","A reading written as 2.5 g",ORANGE)
para(slide,"When a mass is written 2.5 g and no separate instrument error is given, the last written digit is the uncertain one. The absolute error is one unit in that place.",XL,Y0,CW,1.25,16)
fml(slide,"Absolute error = 0.1 g.",XL,Y0+1.45,CW,LH,18)
fml(slide,"Percentage error = (0.1 / 2.5) × 100 = 4 %.",XL,Y0+1.45+LH+NG,CW,LH,17)
ncert(slide,"If the instrument’s least count is given, use that least count as the absolute error.",XL,Y0+3.00,CW,0.85,16)

# CLOSE
slide=new_slide("Class summary","Units 05 lock lines",TEAL)
fml(slide,"Sum or difference: add the absolute errors.",XL,Y0,CW,LH,17)
fml(slide,"Product or quotient: add the relative errors.",XL,Y0+LH+NG,CW,LH,17)
fml(slide,"A power multiplies its own relative error. Use the positive size of the power.",XL,Y0+2*LH+2*NG,CW,0.70,16)
fml(slide,"Parallel resistors and a lens use ΔR = R^2 (ΔR_1/R_1^2 + ΔR_2/R_2^2).",XL,Y0+2.55,CW,0.70,15)
ncert(slide,"Dimensions still cannot produce a pure number such as 6π. That limit from the first half of class stays.",XL,Y0+3.40,CW,0.85,15)

slide=new_slide("For the exam","Lines worth writing in full",GREEN)
body(slide,"Maximum absolute error in a sum or a difference is ΔA + ΔB.",XL,Y0,CW,0.55)
body(slide,"Maximum relative error in a product or a quotient is ΔA/A + ΔB/B.",XL,Y0+0.70,CW,0.60)
body(slide,"For Y = A^m B^n, the maximum relative error is |m|ΔA/A + |n|ΔB/B.",XL,Y0+1.45,CW,0.60)
body(slide,"A result is written (measured value ± absolute error) unit.",XL,Y0+2.20,CW,0.55)
ncert(slide,"These four lines cover the error chapter in the board paper.",XL,Y0+3.00,CW,0.70,16)

slide=new_slide("Practice","Before the next class",ORANGE)
body(slide,"1. A = (2.0 ± 0.1) and B = (3.0 ± 0.2). Find A B with its absolute error.",XL,Y0,CW,0.60)
body(slide,"2. A sphere’s mass has 2 % error and its diameter has 1 % error. Find the percentage error in density.",XL,Y0+0.75,CW,0.70)
body(slide,"3. R_1 = (100 ± 3) Ω and R_2 = (200 ± 4) Ω in series. Find the percentage error.",XL,Y0+1.60,CW,0.65)
body(slide,"4. For a pendulum, l has 1 % error and T has 2 % error. Find the percentage error in g.",XL,Y0+2.40,CW,0.65)
trap_card(slide,"Next class: metre scale, vernier callipers and screw gauge.",XL,Y0+3.25,CW,0.48)

OUTPUT.parent.mkdir(parents=True, exist_ok=True)
prs.save(OUTPUT)
print(f"Created {len(prs.slides)} slides")
print(OUTPUT)
