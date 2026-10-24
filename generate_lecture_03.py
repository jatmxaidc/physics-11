"""
Lecture 03 — Feel of trigonometry (2-hour class)
Follows PW Basic Maths class-02 teaching order, rewritten on our board.
Leftover L01/L02 questions are recap or last-class homework only.
compound A±B → double/half → allied → graphs → even/odd → 37/53 splits
→ sin(45±30) → small angle → tonight.
Class 11 Physics  •  CBSE + NEET Foundation
"""
from __future__ import annotations

import math
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

BASE = Path(__file__).resolve().parent
OUTPUT = BASE / "Lecture_03_Trigonometry_Feel_Compound_Small_Angle.pptx"

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
TRI_X = XL + 0.55
LH = 0.32
NG = 0.18
PT_SCALE = 0.80


def _bg(slide, color=BLACK):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color


def _solid(slide, x, y, w, h, fill, no_line=True):
    s = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h)
    )
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


def tx(
    slide, text, x, y, w, h,
    size=15, color=TEXT, bold=False,
    align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP,
    spa=1, lsp=1.1, font="Aptos", scaled=True,
):
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
    pt = size if not scaled else max(7, round(size * PT_SCALE, 1))
    r.font.size = Pt(pt)
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


def new_slide(section, title, bar_color=BLUE):
    slide = prs.slides.add_slide(BLANK)
    _bg(slide)
    _bar(slide, section, bar_color)
    _title(slide, title)
    _sep(slide)
    _footer(slide, len(prs.slides))
    return slide


def add_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def ncert(slide, formula, x, y, w, h, size=20):
    s = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h)
    )
    s.fill.solid()
    s.fill.fore_color.rgb = NCERT_FILL
    s.line.color.rgb = NCERT_LINE
    s.line.width = Pt(1.5)
    tx(
        slide, formula, x + 0.18, y + 0.16, w - 0.36, h - 0.32,
        size, NAVY, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE, spa=0, lsp=1.15,
    )


def fml(slide, text, x, y, w, h=0.44, size=22):
    return tx(slide, text, x, y, w, h, size, WORK, True, spa=0)


def body(slide, text, x, y, w, h=0.40, size=16):
    return tx(slide, text, x, y, w, h, size, TEXT, False, spa=0, scaled=False)


def qtx(slide, text, x, y, w, h=0.44, size=18):
    return tx(slide, text, x, y, w, h, size, QCOL, True, spa=0, scaled=False)


def choices(slide, items, x, y, w=None):
    if w is None:
        w = CW
    row = 0.40
    for i, item in enumerate(items):
        cy = y + i * row
        circ = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(x), Inches(cy + 0.04), Inches(0.30), Inches(0.30)
        )
        circ.fill.solid()
        circ.fill.fore_color.rgb = TEAL
        circ.line.fill.background()
        tx(slide, str(i + 1), x, cy + 0.06, 0.30, 0.26, 14, WHITE, True, PP_ALIGN.CENTER, spa=0, scaled=False)
        tx(slide, item, x + 0.42, cy + 0.06, w - 0.42, 0.28, 18, TEXT, False, spa=0, scaled=False)
    return y + len(items) * row


def formula_card(slide, formula, note, x, y, w, h, fsize=18):
    fml(slide, formula, x, y, w, LH, fsize)
    if note:
        body(slide, note, x, y + LH, w, 0.28)
    return y + (LH + 0.28 if note else LH)


def trap_card(slide, text, x, y, w, h):
    tx(slide, text, x, y, w, h, 16, ORANGE, False, spa=0, scaled=False)


def result_card(slide, text, x, y, w, h, tsize=14):
    return tx(slide, text, x, y, w, 0.36, max(tsize, 18), GREEN, True, spa=0)


def warn_row(slide, wrong, correct, x, y, w):
    tx(slide, f"Wrong:  {wrong}", x, y, w, 0.30, 16, RED, False, spa=0, scaled=False)
    tx(slide, f"Right:  {correct}", x, y + LH, w, 0.30, 16, GREEN, False, spa=0, scaled=False)
    return y + LH + LH


def right_triangle(slide, x, y, w, h, opp="opp", adj="adj", hyp="hyp"):
    tri = slide.shapes.add_shape(MSO_SHAPE.RIGHT_TRIANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    tri.fill.solid()
    tri.fill.fore_color.rgb = RGBColor(42, 112, 178)
    tri.line.color.rgb = RGBColor(170, 220, 255)
    tri.line.width = Pt(2.25)
    ms, pad = 0.13, 0.012
    mark = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(x + pad), Inches(y + h - ms - pad), Inches(ms), Inches(ms)
    )
    mark.fill.background()
    mark.line.color.rgb = WHITE
    mark.line.width = Pt(1.25)
    tx(slide, adj, x + w * 0.28, y + h + 0.04, w * 0.5, 0.24, 14, BLUE, True, PP_ALIGN.CENTER, spa=0)
    tx(slide, opp, x - 0.55, y + h * 0.32, 0.52, 0.24, 14, ORANGE, True, PP_ALIGN.RIGHT, spa=0)
    tx(slide, hyp, x + w * 0.36, y + h * 0.12, 0.90, 0.24, 14, GREEN, True, spa=0)
    tx(slide, "90°", x + pad + 0.16, y + h - ms - pad - 0.18, 0.40, 0.16, 11, WHITE, True, spa=0)
    tw, th = 0.24, 0.18
    from_right = max(tw + 0.08, (th + 0.08) * w / h)
    tx(slide, "θ", x + w - from_right, y + h - th - 0.02, tw, th, 16, GOLD, True, spa=0)


def grid_table(slide, headers, rows, x, y, col_w, row_h=0.42, hsize=11, csize=12, scaled=True):
    col_x = [x]
    for w in col_w[:-1]:
        col_x.append(col_x[-1] + w)
    for j, ht in enumerate(headers):
        _solid(slide, col_x[j], y, col_w[j], row_h, RGBColor(28, 36, 52), False)
        tx(
            slide, ht, col_x[j] + 0.04, y, col_w[j] - 0.08, row_h,
            hsize, WHITE, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE, spa=0, scaled=scaled,
        )
    for i, row in enumerate(rows):
        ry = y + row_h + i * row_h
        fill = RGBColor(16, 16, 20) if i % 2 == 0 else SOFT
        for j, val in enumerate(row):
            s = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(col_x[j]), Inches(ry), Inches(col_w[j]), Inches(row_h)
            )
            s.fill.solid()
            s.fill.fore_color.rgb = fill
            s.line.color.rgb = STROKE
            s.line.width = Pt(0.5)
            tx(
                slide, val, col_x[j] + 0.04, ry, col_w[j] - 0.08, row_h,
                csize, WORK if j == 0 else TEXT, j == 0, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE,
                spa=0, scaled=scaled,
            )


def tx_3d(slide, text, x, y, w, h, size, face, edge, font="Arial Black"):
    tx(slide, text, x + 0.045, y + 0.045, w, h, size, edge, True, PP_ALIGN.CENTER, spa=0, scaled=False, font=font)
    tx(slide, text, x, y, w, h, size, face, True, PP_ALIGN.CENTER, spa=0, scaled=False, font=font)


def draw_line(slide, x1, y1, x2, y2, color=WHITE, pt=1.75, dash=None):
    c = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2)
    )
    c.line.color.rgb = color
    c.line.width = Pt(pt)
    if dash is not None:
        c.line.dash_style = dash
    return c


def axes(slide, x, y, w, h, xlabel="θ", ylabel="", y_top=""):
    zero = y + h / 2
    draw_line(slide, x, y, x, y + h, MUTED, 1.75)
    draw_line(slide, x, zero, x + w, zero, MUTED, 1.75)
    for i in (1, 2, 3, 4):
        xv = x + i * w / 4
        draw_line(slide, xv, y, xv, y + h, STROKE, 1.15, MSO_LINE.DASH)
    tx(slide, "0", x - 0.45, zero - 0.11, 0.40, 0.22, 12, MUTED, True, spa=0)
    if y_top:
        tx(slide, y_top, x - 0.45, y, 0.40, 0.22, 12, MUTED, True, spa=0)
        tx(slide, f"−{y_top.lstrip('+')}", x - 0.55, y + h - 0.22, 0.52, 0.22, 12, MUTED, True, spa=0)
    tx(slide, xlabel, x + w - 0.25, y + h + 0.04, 0.40, 0.22, 12, MUTED, True, spa=0)
    if ylabel:
        tx(slide, ylabel, x - 0.70, y - 0.02, 0.68, 0.22, 12, MUTED, True, PP_ALIGN.RIGHT, spa=0)
    for i, lab in enumerate(["0°", "90°", "180°", "270°", "360°"]):
        tx(slide, lab, x + i * w / 4 - 0.28, y + h + 0.22, 0.60, 0.22, 12, MUTED, True, PP_ALIGN.CENTER, spa=0)


def trig_val(kind, ang):
    s, c = math.sin(ang), math.cos(ang)
    if kind == "sin":
        return s
    if kind == "cos":
        return c
    if kind == "tan":
        if abs(c) < 1e-12:
            return None
        return s / c
    if kind == "cot":
        if abs(s) < 1e-12:
            return None
        return c / s
    if kind == "sec":
        if abs(c) < 1e-12:
            return None
        return 1 / c
    if kind == "cosec":
        if abs(s) < 1e-12:
            return None
        return 1 / s
    return None


def _wave_intervals(kind):
    if kind in ("tan", "sec"):
        return [(0.0, 90.0), (90.0, 270.0), (270.0, 360.0)]
    if kind in ("cot", "cosec"):
        return [(0.0, 180.0), (180.0, 360.0)]
    return [(0.0, 360.0)]


def wave(slide, x0, y0, w, h, kind="sin", color=GREEN, y_span=1.0):
    amp = h / 2 - 0.08
    eps = 0.6
    clip = kind not in ("sin", "cos")

    def ypix(val):
        v = max(-y_span, min(y_span, val))
        return y0 + h / 2 - v * (amp / y_span)

    def xpix(deg):
        return x0 + (deg / 360.0) * w

    for a, b in _wave_intervals(kind):
        n = max(120, int((b - a) * 3))
        raw = []
        for i in range(n + 1):
            deg = a + eps + (b - a - 2 * eps) * i / n
            val = trig_val(kind, math.radians(deg))
            if val is not None:
                raw.append((deg, val))
        if len(raw) < 2:
            continue
        if not clip:
            pts = [(xpix(d), ypix(v)) for d, v in raw]
        else:
            pts = []
            i = 0
            while i < len(raw):
                d, v = raw[i]
                if abs(v) <= y_span * 0.72:
                    pts.append((xpix(d), ypix(v)))
                    i += 1
                    continue
                j = i
                while j < len(raw) and abs(raw[j][1]) > y_span * 0.72:
                    j += 1
                run = raw[i:j]
                keep = min(run, key=lambda p: min(abs(p[0] - a), abs(p[0] - b)))
                pts.append((xpix(keep[0]), ypix(keep[1])))
                i = j
        for p, q in zip(pts, pts[1:]):
            draw_line(slide, p[0], p[1], q[0], q[1], color, 2.0)


def graph_board(slide, kind, ylabel, lock, color, y_span=1.0, y_top="+1"):
    fml(slide, lock, XL, Y0, CW, LH, 16)
    ax, ay, aw, ah = XL + 0.80, 1.85, 10.4, 3.6
    axes(slide, ax, ay, aw, ah, "θ", ylabel, y_top)
    wave(slide, ax, ay, aw, ah, kind, color, y_span)


# ── 01 COVER ──────────────────────────────────────────────────────────────────
slide = prs.slides.add_slide(BLANK)
_bg(slide)
tx_3d(slide, "EXCELLENCE", XL, 0.55, CW, 1.05, 62, GOLD, RGBColor(210, 210, 220))
tx_3d(slide, "CLASS 11   ·   2026–27", XL, 1.62, CW, 0.55, 30, WHITE, RGBColor(90, 90, 100), font="Aptos")
_solid(slide, 4.15, 2.28, 5.00, 0.028, WHITE)

bx, by, bw = 2.05, 2.70, 9.20
_card_bg(slide, bx, by, bw, 2.35, WHITE, WHITE, 0)
_card_bg(slide, bx, by + 0.72, bw, 1.63, BLACK, RGBColor(220, 220, 220), 1.25)
_solid(slide, bx + bw / 2, by + 0.16, 0.018, 0.42, RGBColor(40, 40, 40))
tx(slide, "Physics", bx, by + 0.16, bw / 2 - 0.10, 0.44, 25, BLACK, True, PP_ALIGN.CENTER, spa=0, scaled=False)
tx(slide, "Lecture - 03", bx + bw / 2, by + 0.16, bw / 2, 0.44, 25, BLACK, True, PP_ALIGN.CENTER, spa=0, scaled=False)
tx(slide, "Basic Maths and Calculus", bx, by + 1.05, bw, 0.55, 32, WHITE, True, PP_ALIGN.CENTER, spa=0, scaled=False, font="Georgia")
tx(slide, "Leftover radian, then feel of trigonometry", bx, by + 1.62, bw, 0.40, 21, GOLD, True, PP_ALIGN.CENTER, spa=0, scaled=False)

tx(slide, "By — Excellence Programme", XL, 5.30, CW, 0.42, 25, WHITE, True, PP_ALIGN.CENTER, spa=0, scaled=False)
tx(slide, "3 September 2026", XL, 5.80, CW, 0.32, 18, MUTED, True, PP_ALIGN.CENTER, spa=0, scaled=False)
_footer(slide, 1)
add_notes(slide, "Cover. 2-hour class. Hour 1: leftover L02 from slide 29. Hour 2: PW class-02 feel.")

# ── 02 TOPICS ─────────────────────────────────────────────────────────────────
slide = prs.slides.add_slide(BLANK)
_bg(slide)
tx(slide, "Topics to be covered", XL, 0.38, CW, 0.55, 32, GOLD, True, spa=0, scaled=False)
_solid(slide, XL, 0.98, 3.40, 0.035, GOLD)

topics = [
    "Inequalities, identities, brother angle, then the 0°–180° table for all six ratios",
    "Leftover from last class, from slide 29  —  radian,  conversions,  leftover PYQs",
    "Feel of trigonometry  —  θ outside a ratio must be in radian",
    "Compound angles  A ± B,  then double angle and half angle",
    "Allied angles,  even / odd,  graphs of all six ratios  0° to 360°",
    "Small-angle approximation  —  sin θ ≈ θ,  tan θ ≈ θ,  cos θ ≈ 1",
]
for i, line in enumerate(topics):
    y = 1.18 + i * 0.78
    tx(slide, f"{i + 1:02d}", XL, y, 0.70, 0.50, 23, GOLD, True, spa=0, scaled=False)
    tx(slide, line, XL + 0.90, y, CW - 0.90, 0.70, 18, WHITE, False, spa=0, scaled=False)

_footer(slide, 2)
add_notes(slide, "Slide 2. Inequalities and full 0–180 table first. Then leftover radian. Then feel.")

# ── INEQUALITIES FROM L01 AND L02 ─────────────────────────────────────────────
slide = new_slide("Copy", "Inequalities from Lectures 01 and 02", BLUE)
add_notes(slide, "They say each line. H is longest. sin=2 is impossible. tan 90 not defined.")
body(slide, "H is the longest side, so P ≤ H and B ≤ H. That is the whole range.", XL, Y0, CW)
y = Y0 + LH + NG
fml(slide, "−1  ≤  sin θ  ≤  1", XL, y, 7.0, LH, 26)
fml(slide, "−1  ≤  cos θ  ≤  1", XL, y + LH, 7.0, LH, 26)
ncert(slide, "sin θ = P/H     H > P     so     |sin θ| ≤ 1", XL + 7.15, y, 4.90, 1.10, 16)
y = y + 2 * LH + NG
fml(slide, "tan θ  has no finite range", XL, y, CW, LH, 22)
fml(slide, "tan 90°  is not defined", XL, y + LH, CW, LH, 22)
trap_card(slide, "If a question writes sin θ = 2, stop. No real angle. Do not open the identity.", XL, y + 2 * LH + NG, CW, 0.40)

slide = new_slide("Copy", "Reciprocals, and 0° to 90°", BLUE)
add_notes(slide, "cosec and sec cannot sit between −1 and 1. 0 to 90: sin up, cos down, tan up.")
fml(slide, "|cosec θ|  ≥  1          |sec θ|  ≥  1", XL, Y0, CW, LH, 24)
body(slide, "When they are defined. Cosec is 1/sin, sec is 1/cos. If |sin| ≤ 1, then |cosec| cannot be 1/2.", XL, Y0 + LH, CW)
y = Y0 + 2 * LH + NG
fml(slide, "cosec 0° and cosec 180°  are not defined          cot 0° and cot 180°  are not defined", XL, y, CW, LH, 18)
fml(slide, "sec 90°  is not defined          tan 90°  is not defined", XL, y + LH, CW, LH, 18)
y = y + 2 * LH + NG
result_card(slide, "0° → 90°:     sin increases     cos decreases     tan increases", XL, y, CW, 0.36, 18)
trap_card(slide, "cot and tan can be any real number, except at the angles where they are not defined.", XL, y + LH + NG, CW, 0.40)

slide = new_slide("Copy", "The three identities — one Pythagoras", BLUE)
add_notes(slide, "Read the three divisions. No writing today. Keep sin²+cos²=1 if memory is short.")
body(slide, "Start from  P² + B² = H².  Divide by H², by B², by P².", XL, Y0, CW)
y = Y0 + LH + NG
fml(slide, "÷ H²:    (P/H)² + (B/H)² = 1          →     sin²θ + cos²θ = 1", XL, y, CW, LH, 20)
fml(slide, "÷ B²:    (P/B)² + 1 = (H/B)²          →     1 + tan²θ = sec²θ", XL, y + LH, CW, LH, 20)
fml(slide, "÷ P²:    1 + (B/P)² = (H/P)²          →     1 + cot²θ = cosec²θ", XL, y + 2 * LH, CW, LH, 20)
ncert(slide, "If you keep only one line:  sin²θ + cos²θ = 1", XL, y + 3 * LH + NG, CW, 0.85, 18)

slide = new_slide("Copy", "Brother angle in the second quadrant", BLUE)
add_notes(slide, "sin(180−θ)=sin θ. cos(180−θ)=−cos θ. tan(180−θ)=−tan θ. 120° is the brother of 60°.")
body(slide, "An angle past 90° uses a brother angle in the first quadrant. ASTC gives the sign.", XL, Y0, CW)
y = Y0 + LH + NG
fml(slide, "sin (180° − θ)  =  sin θ", XL, y, 7.0, LH, 22)
fml(slide, "cos (180° − θ)  =  − cos θ", XL, y + LH, 7.0, LH, 22)
fml(slide, "tan (180° − θ)  =  − tan θ", XL, y + 2 * LH, 7.0, LH, 22)
ncert(slide, "120°  =  180° − 60°", XL + 7.15, y, 4.90, 1.35, 20)
body(slide, "Example:  sin 120° = sin 60° = √3/2.     cos 120° = − cos 60° = −1/2.", XL, y + 3 * LH + NG, CW)

# ── FULL 0–180 TABLE, ALL SIX RATIOS ──────────────────────────────────────────
TW = [1.70] + [2.07] * 5
RH = 0.62
slide = new_slide("Table", "All six ratios from 0° to 90°", TEAL)
add_notes(slide, "Copy the full-width table. Reciprocals of 0 and of 1. tan 90 and sec 90 undefined.")
headers90 = ["θ", "0°", "30°", "45°", "60°", "90°"]
rows90 = [
    ["sin", "0", "1/2", "1/√2", "√3/2", "1"],
    ["cos", "1", "√3/2", "1/√2", "1/2", "0"],
    ["tan", "0", "1/√3", "1", "√3", "undef."],
    ["cot", "undef.", "√3", "1", "1/√3", "0"],
    ["sec", "1", "2/√3", "√2", "2", "undef."],
    ["cosec", "undef.", "2", "√2", "2/√3", "1"],
]
grid_table(slide, headers90, rows90, XL, Y0, TW, row_h=RH, hsize=18, csize=18, scaled=False)
trap_card(slide, "tan 90° and sec 90° are not defined. cosec 0° and cot 0° are not defined.", XL, Y0 + 7 * RH + 0.12, CW, 0.36)

slide = new_slide("Table", "All six ratios from 90° to 180°", TEAL)
add_notes(slide, "Copy. Second quadrant: sin and cosec positive. tan 90 still undefined.")
headers180 = ["θ", "90°", "120°", "135°", "150°", "180°"]
rows180 = [
    ["sin", "1", "√3/2", "1/√2", "1/2", "0"],
    ["cos", "0", "−1/2", "−1/√2", "−√3/2", "−1"],
    ["tan", "undef.", "−√3", "−1", "−1/√3", "0"],
    ["cot", "0", "−1/√3", "−1", "−√3", "undef."],
    ["sec", "undef.", "−2", "−√2", "−2/√3", "−1"],
    ["cosec", "1", "2/√3", "√2", "2", "undef."],
]
grid_table(slide, headers180, rows180, XL, Y0, TW, row_h=RH, hsize=18, csize=18, scaled=False)
trap_card(slide, "Second quadrant: only sin and cosec are positive. cot 180° and cosec 180° are not defined.", XL, Y0 + 7 * RH + 0.12, CW, 0.36)

# ── 5-MINUTE RECAP ────────────────────────────────────────────────────────────
slide = new_slide("5-minute recap", "Last class — ASTC", BLUE)
add_notes(slide, "Five minutes. They say All Silver Tea Cups. Reciprocals take the same sign.")
ncert(slide, "A   S   T   C      →      All    sin    tan    cos", XL, Y0, CW, 0.90, 20)
body(slide, "First quadrant: all positive. Reciprocals take the same sign as the ratio.", XL, Y0 + 1.10, CW)
y = Y0 + 1.10 + LH + NG
fml(slide, "sin (180° − θ)  =  sin θ          cos (180° − θ)  =  − cos θ", XL, y, CW, LH, 20)
fml(slide, "tan 90° is not defined.     −1 ≤ sin θ ≤ 1     −1 ≤ cos θ ≤ 1", XL, y + LH, CW, LH, 18)

slide = new_slide("5-minute recap", "Last class — table and 37° / 53°", BLUE)
add_notes(slide, "They say the 3-4-5 lines. tan 90 not defined. We stopped at the break.")
fml(slide, "sin 37° = 3/5     cos 37° = 4/5     tan 37° = 3/4", XL, Y0, CW, LH, 18)
fml(slide, "sin 53° = 4/5     cos 53° = 3/5     tan 53° = 4/3", XL, Y0 + LH, CW, LH, 18)
body(slide, "0° → 90°:  sin increases,  cos decreases,  tan increases.", XL, Y0 + 2 * LH + NG, CW)
result_card(slide, "sin² 40° + cos² 40°  =  1", XL, Y0 + 3 * LH + 2 * NG, CW, 0.36, 18)
trap_card(slide, "Last class stopped at the break. Radian was not taught. That leftover is first today.", XL, Y0 + 4 * LH + 3 * NG, CW, 0.45)

slide = new_slide("5-minute recap", "Last class — height and distance", BLUE)
add_notes(slide, "Tangent is opposite over adjacent. Depression uses the difference of heights.")
fml(slide, "Elevation:     h  =  x tan θ", XL, Y0, CW, LH, 22)
fml(slide, "Depression:    D  =  (h₁ − h₂) / tan θ", XL, Y0 + LH, CW, LH, 20)
body(slide, "Sketch first. Opposite is the height, or the difference of heights, never the sum.", XL, Y0 + 2 * LH + NG, CW)
trap_card(slide, "We stopped at slide 29. First today: leftover radian, conversions, leftover questions.", XL, Y0 + 3 * LH + 2 * NG, CW, 0.55)

# ── LEFTOVER FROM L02 SLIDE 29 — RADIAN ───────────────────────────────────────
slide = new_slide("Leftover", "Last class stopped here — radian starts now", AMBER)
add_notes(slide, "Do not skip. Slide 29 was the break. Arc, 180=π, conversions, leftover PYQs.")
body(slide, "Last class ended at the break. The next board was radian. We do that leftover first.", XL, Y0, CW)
y = Y0 + LH + NG
fml(slide, "Arc length  →  180° = π rad  →  conversions  →  leftover questions", XL, y, CW, LH, 20)
body(slide, "Then, after the break today: feel of trigonometry.", XL, y + LH + NG, CW)

slide = new_slide("Angle and arc", "Arc length", TEAL)
add_notes(slide, "Arc = Rθ with θ in radian. Never put 90° into this formula.")
formula_card(slide, "Arc length  =  R θ", "θ must be in radian. Never put 90° into this formula.", XL, Y0, 7.0, 1.25, 26)
ncert(slide, "Arc  =  R θ", XL + 7.15, Y0, 4.90, 1.20, 24)
body(slide, "Practical unit of angle: degree.     SI unit of angle: radian.", XL, Y0 + LH + 0.28 + NG, CW)
body(slide, "+θ anti-clockwise     −θ clockwise", XL, Y0 + 2 * LH + 0.28 + NG, CW)

slide = new_slide("Angle and arc", "The one line:  180° = π rad", TEAL)
add_notes(slide, "Wherever you see π rad, put 180°.")
body(slide, "Lock this. Every conversion starts here.", XL, Y0, 7.0)
fml(slide, "180°  =  π rad", XL, Y0 + LH, 7.0, LH, 28)
ncert(slide, "Wherever you see π rad, put 180°.", XL + 7.15, Y0, 4.90, 1.20, 18)
trap_card(slide, "Never write π = 180°. Write π rad = 180°.", XL, Y0 + 2 * LH + NG, 7.0, 0.40)

slide = new_slide("Angle and arc", "Degree → radian", TEAL)
add_notes(slide, "1° = π/180 rad.")
y = formula_card(slide, "1°  =  π / 180  rad", "Divide the lock-line by 180.", XL, Y0, CW, 1.30, 26)
tx(slide, "So  θ°  in radian  =  θ × (π / 180).", XL, y + NG, CW, LH, 18, TEXT, spa=0)
result_card(slide, "Multiply the degree by π/180.", XL, y + NG + LH, CW, 0.36, 18)

slide = new_slide("Angle and arc", "Radian → degree", TEAL)
add_notes(slide, "1 rad = 180/π.")
y = formula_card(slide, "1 rad  =  180° / π", "Divide the lock-line by π.", XL, Y0, CW, 1.30, 26)
tx(slide, "So  θ rad  in degree  =  θ × (180 / π).", XL, y + NG, CW, LH, 18, TEXT, spa=0)
result_card(slide, "Multiply the radian by 180/π.", XL, y + NG + LH, CW, 0.36, 18)

slide = new_slide("Board", "Object turns 90° on a circle of radius R", GREEN)
add_notes(slide, "Read. Cross out R×90°. 90°=π/2. Distance=πR/2.")
qtx(slide, "An object moves on the circumference of a circle of radius R. If the object turns through 90°, find the total distance moved.", XL, Y0, CW, 0.55, 17)
y = warn_row(slide, "Distance = R × 90°", "Distance = R × (π/2 rad)  =  πR/2", XL, Y0 + 0.60, CW)
fml(slide, "distance = path = arc = Rθ          and          90° = π/2 rad", XL, y + NG, CW, LH, 18)
result_card(slide, "Distance  =  πR / 2", XL, y + NG + LH + NG, CW, 0.36, 22)

slide = new_slide("Board", "Convert π/3 radian into degree", GREEN)
add_notes(slide, "Read. Put π rad = 180°. 180/3 = 60°.")
qtx(slide, "Convert π/3 radian into degree.", XL, Y0, CW, 0.40, 20)
body(slide, "Wherever you see π rad, put 180°.", XL, Y0 + 0.50, CW)
fml(slide, "π / 3  rad  =  180° / 3  =  60°", XL, Y0 + 0.50 + LH + NG, 7.0, LH, 24)
ncert(slide, "π / 3 rad  =  60°", XL + 7.15, Y0 + 0.50 + LH + NG, 4.90, 1.10, 22)

slide = new_slide("Board", "Six conversions — read each line", GREEN)
add_notes(slide, "Read left to right. π rad = 180°.")
body(slide, "Put 180° wherever you see π rad. Multiply a degree by π/180 to get radian.", XL, Y0, CW)
live_a = [
    "π/2 rad  =  90°",
    "π/4 rad  =  45°",
    "π/6 rad  =  30°",
    "45°  =  π/4 rad",
    "90°  =  π/2 rad",
    "60°  =  π/3 rad",
]
for i, item in enumerate(live_a):
    r, c = divmod(i, 2)
    result_card(slide, item, XL + c * 6.1, Y0 + 0.50 + r * LH, 5.85, LH, 18)

slide = new_slide("Board", "Radian into degree — every line", GREEN)
add_notes(slide, "Read. Each one is  (number) × 180°.")
body(slide, "π rad = 180°. So  5π/6  =  5 × 180° / 6,  and so on.", XL, Y0, CW)
rad_a = [
    "5π/6  =  150°", "2π/3  =  120°", "4π/3  =  240°",
    "3π/5  =  108°", "5π  =  900°", "2π/5  =  72°",
    "7π/12  =  105°", "5π/2  =  450°", "3π  =  540°",
]
for i, item in enumerate(rad_a):
    r, c = divmod(i, 3)
    result_card(slide, item, XL + c * 4.0, Y0 + 0.48 + r * LH, 3.85, LH, 16)

slide = new_slide("Board", "Degree into radian — every line", GREEN)
add_notes(slide, "Read. π is already radian. π° would be π²/180.")
body(slide, "Multiply the degree by π/180.  (vi) is π, not π°, so it is already radian.", XL, Y0, CW)
deg_a = [
    "120°  =  2π/3 rad", "30°  =  π/6 rad", "150°  =  5π/6 rad",
    "270°  =  3π/2 rad", "135°  =  3π/4 rad", "π  is already radian",
]
for i, item in enumerate(deg_a):
    r, c = divmod(i, 3)
    result_card(slide, item, XL + c * 4.0, Y0 + 0.48 + r * LH, 3.85, LH, 16)
warn_row(slide, "π°  treated as π rad", "π  is already radian.  π°  would be  π × (π/180)  =  π²/180 rad", XL, Y0 + 0.48 + 2 * LH + NG, CW)

slide = new_slide("Board", "Find β — exterior angle", GREEN)
add_notes(slide, "Read. Interior third = 70°. Exterior β = 110°.")
qtx(slide, "A triangle has two interior angles 40° and 70°. The side next to the 70° angle is extended. Find the exterior angle β.", XL, Y0, CW, 0.70, 17)
fml(slide, "40°  +  70°  +  α  =  180°          α  =  70°", XL, Y0 + 0.80, CW, LH, 20)
fml(slide, "β  =  180°  −  α  =  110°", XL, Y0 + 0.80 + LH, CW, LH, 20)
result_card(slide, "β  =  110°", XL, Y0 + 0.80 + 2 * LH + NG, CW, 0.36, 22)
body(slide, "Shortcut: an exterior angle equals the sum of the two remote interiors.  40° + 70° = 110°.", XL, Y0 + 0.80 + 3 * LH + 2 * NG, CW)

slide = new_slide("Board", "sin 53° × cosec 53°", GREEN)
add_notes(slide, "Read. Same angle. Reciprocal. Value is 1.")
qtx(slide, "Find the value of  sin 53°  ×  cosec 53°.", XL, Y0, CW, 0.40, 20)
fml(slide, "sin θ  ×  cosec θ  =  1          for the same angle θ", XL, Y0 + 0.55, CW, LH, 20)
body(slide, "The angles are the same. Do not look up 53°. The product of a ratio and its reciprocal is 1.", XL, Y0 + 0.55 + LH + NG, CW)
result_card(slide, "sin 53°  ×  cosec 53°  =  1", XL, Y0 + 0.55 + 2 * LH + 2 * NG, CW, 0.36, 22)

slide = new_slide("Board", "sin 180° · cos 135° · tan 120°", GREEN)
add_notes(slide, "Read. sin180=0. Product is 0.")
qtx(slide, "Find the value of  sin 180°  ·  cos 135°  ·  tan 120°.", XL, Y0, CW, 0.40, 20)
fml(slide, "sin 180°  =  0", XL, Y0 + 0.55, CW, LH, 22)
body(slide, "Zero times anything is zero. You do not need cos 135° or tan 120° after that.", XL, Y0 + 0.55 + LH + NG, CW)
result_card(slide, "0", XL, Y0 + 0.55 + 2 * LH + 2 * NG, CW, 0.36, 24)

slide = new_slide("Board", "Shadow of a pole", GREEN)
add_notes(slide, "Read. tan θ = h/l. Answer is option 2.")
qtx(slide, "A vertical pole of height h casts a shadow of length l when the Sun is at an angle θ above the horizon. Which expression is l?", XL, Y0, CW, 0.55, 16)
choices(slide, [
    "l  =  h tan θ",
    "l  =  h / tan θ",
    "l  =  h sin θ",
    "l  =  h / sin θ",
], XL, Y0 + 0.60)
fml(slide, "tan θ  =  opposite / adjacent  =  h / l          so          l  =  h / tan θ  =  h cot θ", XL, Y0 + 2.30, CW, LH, 18)
result_card(slide, "Option 2     l  =  h / tan θ", XL, Y0 + 2.30 + LH + NG, CW, 0.36, 20)

slide = new_slide("Board", "Wire from the top of a pole", GREEN)
add_notes(slide, "Read. x = h(cosec−cot). None of the four options.")
qtx(slide, "A pole of height h has a wire from the top to the ground at angle θ. The wire breaks and falls without changing length. Find x from the base to the far end of the fallen wire.", XL, Y0, CW, 0.70, 16)
choices(slide, [
    "x  =  h cot θ",
    "x  =  h tan θ",
    "x  =  h sin θ",
    "x  =  h sec θ",
], XL, Y0 + 0.75)
fml(slide, "l  =  h / sin θ  =  h cosec θ          b  =  h / tan θ  =  h cot θ          x  =  l − b", XL, Y0 + 2.45, CW, LH, 16)
result_card(slide, "x  =  h (cosec θ  −  cot θ)      —  none of the four options", XL, Y0 + 2.45 + LH + NG, CW, 0.36, 18)
trap_card(slide, "Option 1 is only the original base b, not the far end after the wire falls.", XL, Y0 + 2.45 + 2 * LH + 2 * NG, CW, 0.36)

slide = new_slide("Board", "Shadow on sloping ground", GREEN)
add_notes(slide, "Read. a=l/2, b=l√3/2. tan60=(10+a)/b. l=10 m. Option 2.")
qtx(slide, "Pole h = 10 m on a slope α = 30°. Sun elevation θ = 60°. Find the shadow length l on the slope.", XL, Y0, CW, 0.55, 16)
choices(slide, [
    "5 m",
    "10 m",
    "10√3 m",
    "10 / (√3 − 1) m",
], XL, Y0 + 0.60)
fml(slide, "a  =  l sin 30°  =  l/2          b  =  l cos 30°  =  l √3 / 2", XL, Y0 + 2.30, CW, LH, 16)
fml(slide, "tan 60°  =  (10 + a)/b          √3  =  (10 + l/2) / (l √3 / 2)          3l/2  =  10 + l/2", XL, Y0 + 2.30 + LH, CW, LH, 16)
result_card(slide, "Option 2     l  =  10 m", XL, Y0 + 2.30 + 2 * LH + NG, CW, 0.36, 20)

# ── BREAK AFTER LEFTOVER ──────────────────────────────────────────────────────
slide = prs.slides.add_slide(BLANK)
_bg(slide)
_solid(slide, 0, 0, 13.333, 0.06, BLUE)
tx(slide, "BREAK", XL, 1.85, CW, 0.55, 34, GOLD, True, PP_ALIGN.CENTER, spa=4)
tx(slide, "Leftover radian is done. Next: feel of trigonometry", XL, 2.70, CW, 0.40, 16, BLUE, True, PP_ALIGN.CENTER, spa=0)
_footer(slide, len(prs.slides))
add_notes(slide, "Hour 1 done. Leftover from L02 slide 29. After break: PW class-02 feel.")

# ── LENGTH PREFIXES ───────────────────────────────────────────────────────────
slide = new_slide("Copy", "Length prefixes you will meet in physics", TEAL)
add_notes(slide, "Copy once. Å is 10^−10 m. fm is femtometre, nuclear size.")
body(slide, "Write this list once. You will use it from Units and Measurements onward.", XL, Y0, CW)
left = [
    "1 m   =  100 cm",
    "1 fm  =  10⁻¹⁵ m",
    "1 pm  =  10⁻¹² m",
    "1 nm  =  10⁻⁹ m",
    "1 μm  =  10⁻⁶ m",
]
right = [
    "1 mm  =  10⁻³ m",
    "1 cm  =  10⁻² m",
    "1 km  =  10³ m",
    "1 Mm  =  10⁶ m",
    "1 Å   =  10⁻¹⁰ m",
]
for i, line in enumerate(left):
    fml(slide, line, XL, Y0 + 0.50 + i * LH, 5.6, LH, 20)
for i, line in enumerate(right):
    fml(slide, line, XL + 6.2, Y0 + 0.50 + i * LH, 5.6, LH, 20)
trap_card(slide, "Å is ångström. Nuclear sizes are in fm. Light wavelengths are in nm or Å.", XL, Y0 + 0.50 + 5 * LH + NG, CW, 0.40)

# ── LINES AND ANGLES ──────────────────────────────────────────────────────────
slide = new_slide("Feel", "Lines — parallel, orthogonal, antiparallel", TEAL)
add_notes(slide, "Parallel: same direction, θ=0. Orthogonal: 90°. Antiparallel: opposite, 180°.")
body(slide, "These three words come back in vectors. Lock the picture today.", XL, Y0, CW)
# parallel
draw_line(slide, XL, 2.55, XL + 2.6, 2.55, GREEN, 2.25)
draw_line(slide, XL, 2.95, XL + 2.6, 2.95, GREEN, 2.25)
tx(slide, "Parallel", XL, 3.10, 2.6, 0.28, 16, GOLD, True, PP_ALIGN.CENTER, spa=0, scaled=False)
fml(slide, "θ  =  0°", XL, 3.40, 2.6, LH, 18)
# orthogonal
draw_line(slide, XL + 4.3, 3.15, XL + 6.5, 3.15, GREEN, 2.25)
draw_line(slide, XL + 5.4, 2.20, XL + 5.4, 3.15, GREEN, 2.25)
_card_bg(slide, XL + 5.4, 3.02, 0.14, 0.14, BLACK, WHITE, 1.0, False)
tx(slide, "Orthogonal  (perpendicular)", XL + 3.9, 3.28, 3.2, 0.28, 16, GOLD, True, PP_ALIGN.CENTER, spa=0, scaled=False)
fml(slide, "θ  =  90°", XL + 3.9, 3.58, 3.2, LH, 18)
# antiparallel
draw_line(slide, XL + 8.4, 2.70, XL + 11.4, 2.70, GREEN, 2.25)
tx(slide, "→", XL + 10.85, 2.50, 0.50, 0.28, 18, GREEN, True, spa=0, scaled=False)
tx(slide, "←", XL + 8.40, 2.50, 0.50, 0.28, 18, GREEN, True, spa=0, scaled=False)
tx(slide, "Antiparallel", XL + 8.4, 3.10, 3.0, 0.28, 16, GOLD, True, PP_ALIGN.CENTER, spa=0, scaled=False)
fml(slide, "θ  =  180°", XL + 8.4, 3.40, 3.0, LH, 18)

slide = new_slide("Feel", "Complementary, acute, obtuse, alternate", TEAL)
add_notes(slide, "Complementary sum to 90. Acute < 90. Obtuse > 90 and ≤ 180. Alternate equal when lines are parallel.")
fml(slide, "Complementary:     α + β  =  90°     so     α  =  90° − β", XL, Y0, CW, LH, 20)
body(slide, "In a right triangle the two acute angles are complementary. That is why sin(90° − θ) = cos θ.", XL, Y0 + LH, CW)
y = Y0 + 2 * LH + NG
fml(slide, "Acute:     0°  ≤  θ  <  90°", XL, y, 5.8, LH, 20)
fml(slide, "Obtuse:     90°  <  α  ≤  180°", XL + 6.2, y, 5.8, LH, 20)
body(slide, "Alternate angles: a transversal cuts two parallel lines. The inner opposite angles are equal.", XL, y + LH + NG, CW)

# ── y = sinθ/θ ────────────────────────────────────────────────────────────────
slide = new_slide("Board", "θ outside a ratio is not a degree", GREEN)
add_notes(slide, "Read. Do not write sin30/30. Convert 30° to π/6 rad. y = 3/π.")
qtx(slide, "If  y  =  sin θ / θ,  find the value of y when  θ  =  30°.", XL, Y0, CW, 0.45, 18)
y = warn_row(slide, "y  =  sin 30° / 30°", "y  =  sin 30°  /  (π/6)", XL, Y0 + 0.50, CW)
fml(slide, "30°  =  π/6 rad          sin 30°  =  1/2", XL, y + NG, CW, LH, 18)
fml(slide, "y  =  (1/2)  ÷  (π/6)  =  (1/2) × (6/π)  =  3/π", XL, y + NG + LH, CW, LH, 18)
ncert(slide, "If θ is not inside sin, cos or tan, convert it to radian first.", XL, y + NG + 2 * LH + NG, CW, 0.70, 16)
body(slide, "(a/b) ÷ (c/d)  =  ad / bc.          Answer:   y  =  3/π", XL, y + NG + 2 * LH + NG + 0.80, CW)

# ── COMPOUND ANGLES ───────────────────────────────────────────────────────────
slide = new_slide("Copy", "Compound angles — sine and cosine", TEAL)
add_notes(slide, "Copy. Write ten times on a rough page after class. A and B are angles.")
body(slide, "A and B are angles. Copy these four lines. Practice them ten times on a rough page.", XL, Y0, CW)
y = Y0 + LH + NG
fml(slide, "sin (A + B)  =  sin A cos B  +  cos A sin B", XL, y, CW, LH, 22)
fml(slide, "sin (A − B)  =  sin A cos B  −  cos A sin B", XL, y + LH, CW, LH, 22)
fml(slide, "cos (A + B)  =  cos A cos B  −  sin A sin B", XL, y + 2 * LH + NG, CW, LH, 22)
fml(slide, "cos (A − B)  =  cos A cos B  +  sin A sin B", XL, y + 3 * LH + NG, CW, LH, 22)
trap_card(slide, "Cosine of a sum has a minus. Sine of a sum has a plus. Do not swap those two signs.", XL, y + 4 * LH + 2 * NG, CW, 0.45)

slide = new_slide("Copy", "Compound angles — tangent", TEAL)
add_notes(slide, "Copy. Denominator 1 − tanA tanB for sum. Plus in the denominator for difference.")
fml(slide, "tan (A + B)  =  (tan A + tan B) / (1 − tan A tan B)", XL, Y0, CW, LH, 22)
fml(slide, "tan (A − B)  =  (tan A − tan B) / (1 + tan A tan B)", XL, Y0 + LH, CW, LH, 22)
body(slide, "The minus in the denominator belongs to the sum. The plus belongs to the difference.", XL, Y0 + 2 * LH + NG, CW)
ncert(slide, "If tan A tan B = 1, then tan(A + B) is not defined.", XL, Y0 + 3 * LH + 2 * NG, CW, 0.90, 18)

# ── DOUBLE / HALF SIN ─────────────────────────────────────────────────────────
slide = new_slide("Derivation", "Double angle from sin (A + B)", TEAL)
add_notes(slide, "Put A = B = θ. sin 2θ = 2 sinθ cosθ. Then 2θ = α gives the half-angle form.")
body(slide, "Put A = B = θ in the sine-of-a-sum line.", XL, Y0, CW)
fml(slide, "sin (θ + θ)  =  sin θ cos θ  +  cos θ sin θ", XL, Y0 + LH + NG, CW, LH, 20)
result_card(slide, "sin 2θ  =  2 sin θ cos θ", XL, Y0 + 2 * LH + 2 * NG, CW, 0.36, 22)
body(slide, "Now let  2θ = α,  so  θ = α/2.", XL, Y0 + 3 * LH + 3 * NG, CW)
ncert(slide, "sin α  =  2 sin(α/2) cos(α/2)", XL, Y0 + 4 * LH + 4 * NG, CW, 0.90, 20)

slide = new_slide("Derivation", "Double angle from cos (A + B)", TEAL)
add_notes(slide, "cos 2θ = cos²θ − sin²θ = 1 − 2sin²θ. Rearrange: sin θ = √[(1−cos2θ)/2].")
body(slide, "Put A = B = θ in the cosine-of-a-sum line.", XL, Y0, CW)
fml(slide, "cos 2θ  =  cos²θ  −  sin²θ", XL, Y0 + LH + NG, 7.0, LH, 22)
ncert(slide, "cos 2θ  =  1 − 2 sin²θ", XL + 7.15, Y0 + LH + NG, 4.90, 1.10, 18)
y = Y0 + 2 * LH + 2 * NG + 0.40
fml(slide, "2 sin²θ  =  1 − cos 2θ", XL, y, CW, LH, 20)
fml(slide, "sin θ  =  √[(1 − cos 2θ) / 2]", XL, y + LH, CW, LH, 20)
body(slide, "Half angle: let  2θ = α. Then  cos α  =  cos²(α/2) − sin²(α/2).", XL, y + 2 * LH + NG, CW)

# ── FIND VALUES ───────────────────────────────────────────────────────────────
slide = new_slide("Board", "sin 120°  and  cos 150°", GREEN)
add_notes(slide, "Read. sin120=√3/2. cos150=−√3/2.")
qtx(slide, "Find  (1)  sin 120°     and     (2)  cos 150°.", XL, Y0, CW, 0.40, 18)
body(slide, "Write each as a sum you already know.  120° = 90° + 30°.   150° = 90° + 60°.", XL, Y0 + 0.45, CW)
fml(slide, "sin 120°  =  sin (90° + 30°)  =  1 · (√3/2)  +  0  =  √3/2", XL, Y0 + 0.90, CW, LH, 18)
body(slide, "Same value from the double-angle line:  sin(2 × 60°)  =  2 sin 60° cos 60°  =  √3/2.", XL, Y0 + 0.90 + LH, CW)
result_card(slide, "(1)   sin 120°  =  √3/2", XL, Y0 + 0.90 + 2 * LH + NG, CW, 0.32, 18)
fml(slide, "cos 150°  =  cos (90° + 60°)  =  0  −  1 · (√3/2)  =  −√3/2", XL, Y0 + 0.90 + 3 * LH + 2 * NG, CW, LH, 18)
result_card(slide, "(2)   cos 150°  =  −√3/2", XL, Y0 + 0.90 + 4 * LH + 3 * NG, CW, 0.32, 18)

slide = new_slide("Board", "sin 240°  and  cos 120°", GREEN)
add_notes(slide, "Read. sin240=−√3/2. cos120=−1/2. Third quadrant, sine is negative.")
qtx(slide, "Find  (3)  sin 240°     and     (4)  cos 120°.", XL, Y0, CW, 0.40, 18)
fml(slide, "sin 240°  =  sin (180° + 60°)  =  0  +  (−1) · (√3/2)  =  −√3/2", XL, Y0 + 0.50, CW, LH, 18)
result_card(slide, "(3)   sin 240°  =  −√3/2", XL, Y0 + 0.50 + LH + NG, CW, 0.32, 18)
fml(slide, "cos 120°  =  cos (90° + 30°)  =  − sin 30°  =  −1/2", XL, Y0 + 0.50 + 2 * LH + 2 * NG, CW, LH, 18)
result_card(slide, "(4)   cos 120°  =  −1/2", XL, Y0 + 0.50 + 3 * LH + 3 * NG, CW, 0.32, 18)
trap_card(slide, "240° is in the third quadrant. Sine is negative there. The compound line already gives the sign.", XL, Y0 + 0.50 + 4 * LH + 4 * NG, CW, 0.40)

# ── ALLIED ANGLES ─────────────────────────────────────────────────────────────
slide = new_slide("Board", "sin (90° − θ)  and  sin (90° + θ)", GREEN)
add_notes(slide, "Read. Both equal cos θ. cos 90 = 0, sin 90 = 1.")
qtx(slide, "Simplify.   (5)  sin (90° − θ)     (6)  sin (90° + θ)", XL, Y0, CW, 0.40, 18)
fml(slide, "sin (90° − θ)  =  sin 90° cos θ  −  cos 90° sin θ  =  1 · cos θ  −  0  =  cos θ", XL, Y0 + 0.50, CW, LH, 16)
result_card(slide, "(5)   sin (90° − θ)  =  cos θ", XL, Y0 + 0.50 + LH + NG, CW, 0.32, 18)
fml(slide, "sin (90° + θ)  =  sin 90° cos θ  +  cos 90° sin θ  =  cos θ", XL, Y0 + 0.50 + 2 * LH + 2 * NG, CW, LH, 16)
result_card(slide, "(6)   sin (90° + θ)  =  cos θ", XL, Y0 + 0.50 + 3 * LH + 3 * NG, CW, 0.32, 18)
trap_card(slide, "Both become cosine. The plus or minus next to 90° does not change sine into minus cosine here.", XL, Y0 + 0.50 + 4 * LH + 4 * NG, CW, 0.40)

slide = new_slide("Board", "sin (180° − θ)  and  sin (270° + θ)", GREEN)
add_notes(slide, "Read. sin(180−θ)=sinθ. sin(270+θ)= −cosθ because sin270= −1, cos270=0.")
qtx(slide, "Simplify.   (7)  sin (180° − θ)     (8)  sin (270° + θ)", XL, Y0, CW, 0.40, 18)
fml(slide, "sin (180° − θ)  =  0 · cos θ  −  sin θ · (−1)  =  sin θ", XL, Y0 + 0.50, CW, LH, 16)
result_card(slide, "(7)   sin (180° − θ)  =  sin θ", XL, Y0 + 0.50 + LH + NG, CW, 0.32, 18)
fml(slide, "sin (270° + θ)  =  sin 270° cos θ  +  cos 270° sin θ  =  (−1) cos θ  +  0", XL, Y0 + 0.50 + 2 * LH + 2 * NG, CW, LH, 16)
result_card(slide, "(8)   sin (270° + θ)  =  − cos θ", XL, Y0 + 0.50 + 3 * LH + 3 * NG, CW, 0.32, 18)

# ── GRAPHS ────────────────────────────────────────────────────────────────────
slide = new_slide("Graph", "sin θ from 0° to 360°", TEAL)
add_notes(slide, "Zeros at 0, 180, 360. Peak +1 at 90. Trough −1 at 270.")
graph_board(
    slide, "sin", "sin θ",
    "sin 0° = sin 180° = sin 360° = 0          sin 90° = +1          sin 270° = −1",
    GREEN, 1.0, "+1",
)

slide = new_slide("Graph", "cos θ from 0° to 360°", TEAL)
add_notes(slide, "cos 0 = +1, cos 90 = 0, cos 180 = −1, cos 270 = 0, back to +1 at 360.")
graph_board(
    slide, "cos", "cos θ",
    "cos 0° = +1          cos 90° = 0          cos 180° = −1          cos 270° = 0",
    GOLD, 1.0, "+1",
)

slide = new_slide("Graph", "tan θ from 0° to 360°", TEAL)
add_notes(slide, "Asymptotes at 90° and 270°. Period 180°. Odd function.")
graph_board(
    slide, "tan", "tan θ",
    "tan 0° = tan 180° = tan 360° = 0          tan 90° and tan 270° are not defined",
    ORANGE, 3.0, "+3",
)
trap_card(slide, "Dotted lines at 90° and 270° are asymptotes. The graph never crosses them.", XL, 6.55, CW, 0.32)

slide = new_slide("Graph", "cot θ from 0° to 360°", TEAL)
add_notes(slide, "Asymptotes at 0°, 180°, 360°. cot 90° = 0.")
graph_board(
    slide, "cot", "cot θ",
    "cot 90° = cot 270° = 0          cot 0°, cot 180° and cot 360° are not defined",
    TEAL, 3.0, "+3",
)
trap_card(slide, "Dotted lines at 0°, 180° and 360° are asymptotes. Reciprocal of tan.", XL, 6.55, CW, 0.32)

slide = new_slide("Graph", "sec θ from 0° to 360°", TEAL)
add_notes(slide, "sec = 1/cos. |sec| ≥ 1. Asymptotes at 90° and 270°.")
graph_board(
    slide, "sec", "sec θ",
    "sec 0° = sec 360° = +1          sec 180° = −1          sec 90° and sec 270° are not defined",
    BLUE, 3.0, "+3",
)
trap_card(slide, "|sec θ| ≥ 1. The graph never sits between −1 and +1.", XL, 6.55, CW, 0.32)

slide = new_slide("Graph", "cosec θ from 0° to 360°", TEAL)
add_notes(slide, "cosec = 1/sin. |cosec| ≥ 1. Asymptotes at 0°, 180°, 360°.")
graph_board(
    slide, "cosec", "cosec θ",
    "cosec 90° = +1          cosec 270° = −1          cosec 0°, 180° and 360° are not defined",
    WORK, 3.0, "+3",
)
trap_card(slide, "|cosec θ| ≥ 1. Reciprocal of sine. Never write cosec 0°.", XL, 6.55, CW, 0.32)

# ── EVEN / ODD ────────────────────────────────────────────────────────────────
slide = new_slide("Even and odd", "Negative angles", TEAL)
add_notes(slide, "cos is even. sin and tan are odd. From 360 − θ and the table.")
body(slide, "cos 360° = 1 and sin 360° = 0. Put those into the compound lines.", XL, Y0, CW)
y = Y0 + LH + NG
fml(slide, "cos (−θ)  =  cos θ", XL, y, 7.0, LH, 22)
ncert(slide, "even function", XL + 7.15, y, 4.90, 0.70, 18)
fml(slide, "sin (−θ)  =  − sin θ", XL, y + LH + NG, 7.0, LH, 22)
fml(slide, "tan (−θ)  =  − tan θ", XL, y + 2 * LH + NG, 7.0, LH, 22)
ncert(slide, "odd functions", XL + 7.15, y + LH + NG, 4.90, 1.05, 18)
body(slide, "Example:  cos (−30°)  =  cos 30°  =  √3/2.     Also  sin (360° + θ)  =  sin θ.", XL, y + 3 * LH + 2 * NG, CW)

slide = new_slide("Board", "Negative angles — even and odd", GREEN)
add_notes(slide, "Read. sin odd, cos even. −1/2, +1/2, −√3/2, −1/2, −1/2.")
qtx(slide, "Find each value.  Sine is odd.  Cosine is even.", XL, Y0, CW, 0.36, 18)
fml(slide, "(1)   sin (−30°)   =  − sin 30°   =  −1/2", XL, Y0 + 0.42, CW, LH, 18)
fml(slide, "(2)   cos (−60°)   =  + cos 60°   =  +1/2", XL, Y0 + 0.42 + LH, CW, LH, 18)
fml(slide, "(3)   sin (−120°)  =  − sin 120°  =  −√3/2", XL, Y0 + 0.42 + 2 * LH, CW, LH, 18)
fml(slide, "(4)   cos (−120°)  =  + cos 120°  =  −1/2", XL, Y0 + 0.42 + 3 * LH, CW, LH, 18)
fml(slide, "(5)   sin (−150°)  =  − sin 150°  =  −1/2", XL, Y0 + 0.42 + 4 * LH, CW, LH, 18)
trap_card(slide, "Cosine keeps the sign of the positive angle. Sine flips. Do not flip cosine.", XL, Y0 + 0.42 + 5 * LH + NG, CW, 0.40)

# ── CONTINUE — 37 / 53 AND SMALL ANGLE ────────────────────────────────────────
slide = prs.slides.add_slide(BLANK)
_bg(slide)
_solid(slide, 0, 0, 13.333, 0.06, TEAL)
tx(slide, "NEXT", XL, 1.85, CW, 0.55, 34, GOLD, True, PP_ALIGN.CENTER, spa=4)
tx(slide, "37° / 53° splits  →  sin 15°, 75°, 105°  →  small angle  →  tonight", XL, 2.70, CW, 0.40, 16, BLUE, True, PP_ALIGN.CENTER, spa=0)
_footer(slide, len(prs.slides))
add_notes(slide, "Keep moving. No second break. 37/53 then small angle.")

# ── 37 / 53 SPLITS ────────────────────────────────────────────────────────────
slide = new_slide("Board", "127° and 143° — split, then 37° or 53°", GREEN)
add_notes(slide, "Read. sin127=cos37=4/5. cos143= −sin53= −4/5. tan143= −tan37= −3/4.")
qtx(slide, "Write each as 90° or 180° with 37° or 53°, then find the value.", XL, Y0, CW, 0.40, 16)
body(slide, "Last class:  sin 37° = 3/5,  cos 37° = 4/5,  tan 37° = 3/4.     sin 53° = 4/5,  cos 53° = 3/5,  tan 53° = 4/3.", XL, Y0 + 0.45, CW)
fml(slide, "(1)   sin 127°  =  sin (90° + 37°)  =  cos 37°  =  4/5", XL, Y0 + 0.90, CW, LH, 18)
fml(slide, "(2)   cos 143°  =  cos (90° + 53°)  =  − sin 53°  =  −4/5", XL, Y0 + 0.90 + LH, CW, LH, 18)
fml(slide, "(3)   tan 143°  =  tan (180° − 37°)  =  − tan 37°  =  −3/4", XL, Y0 + 0.90 + 2 * LH, CW, LH, 18)
result_card(slide, "sin 127°  =  4/5          cos 143°  =  −4/5          tan 143°  =  −3/4", XL, Y0 + 0.90 + 3 * LH + NG, CW, 0.36, 18)

# ── 15 / 75 / 105 ─────────────────────────────────────────────────────────────
slide = new_slide("Board", "sin 105°  =  sin (60° + 45°)", GREEN)
add_notes(slide, "Read. (√6+√2)/4.")
qtx(slide, "Find  sin 105°.   Use the sine-of-a-sum line. Keep √2 and √3 until the last step.", XL, Y0, CW, 0.45, 16)
fml(slide, "sin (60° + 45°)  =  sin 60° cos 45°  +  cos 60° sin 45°", XL, Y0 + 0.55, CW, LH, 18)
fml(slide, "                 =  (√3/2)(1/√2)  +  (1/2)(1/√2)", XL, Y0 + 0.55 + LH, CW, LH, 18)
fml(slide, "                 =  (√6 + √2) / 4", XL, Y0 + 0.55 + 2 * LH, CW, LH, 18)
result_card(slide, "sin 105°  =  (√6 + √2) / 4", XL, Y0 + 0.55 + 3 * LH + NG, CW, 0.36, 20)

slide = new_slide("Board", "sin 75°  and  sin 15°", GREEN)
add_notes(slide, "Read. sin75 = (√6+√2)/4 same as 105. sin15 = (√6−√2)/4.")
qtx(slide, "Find  (2)  sin 75°  =  sin (45° + 30°)     and     (3)  sin 15°  =  sin (45° − 30°).", XL, Y0, CW, 0.45, 16)
fml(slide, "sin 75°  =  sin (45° + 30°)  =  same expansion as 105°  =  (√6 + √2) / 4", XL, Y0 + 0.55, CW, LH, 16)
result_card(slide, "(2)   sin 75°  =  (√6 + √2) / 4", XL, Y0 + 0.55 + LH + NG, CW, 0.32, 18)
fml(slide, "sin 15°  =  sin 45° cos 30°  −  cos 45° sin 30°  =  (1/√2)(√3/2)  −  (1/√2)(1/2)", XL, Y0 + 0.55 + 2 * LH + 2 * NG, CW, LH, 16)
fml(slide, "         =  (√6 − √2) / 4", XL, Y0 + 0.55 + 3 * LH + 2 * NG, CW, LH, 18)
result_card(slide, "(3)   sin 15°  =  (√6 − √2) / 4", XL, Y0 + 0.55 + 4 * LH + 3 * NG, CW, 0.32, 18)

# ── SMALL ANGLE ───────────────────────────────────────────────────────────────
slide = new_slide("Small angle", "When θ is small, the arc looks like the opposite side", TEAL)
add_notes(slide, "θ ≤ 5°. sinθ≈θ, tanθ≈θ in radian. cosθ≈1. Never leave θ in degree.")
body(slide, "On a circle of radius R, a tiny angle makes opposite ≈ arc = Rθ, and adjacent ≈ R.", XL, Y0, CW)
y = Y0 + LH + NG
fml(slide, "sin θ  =  P/H  ≈  (Rθ)/R  =  θ", XL, y, 7.0, LH, 20)
fml(slide, "cos θ  =  B/H  ≈  R/R  =  1", XL, y + LH, 7.0, LH, 20)
fml(slide, "tan θ  =  sin θ / cos θ  ≈  θ / 1  =  θ", XL, y + 2 * LH, 7.0, LH, 20)
ncert(slide, "θ must be in radian.   Use only if θ ≤ 5°.", XL + 7.15, y, 4.90, 1.55, 18)
trap_card(slide, "sin 2° is not 2. Convert 2° to radian first, then the approximation is that radian number.", XL, y + 3 * LH + NG, CW, 0.50)

slide = new_slide("Board", "Small-angle values — convert first", GREEN)
add_notes(slide, "Read. π/90, π/60, 1.")
qtx(slide, "Using the small-angle approximation, find  sin 2°,  tan 3°  and  cos 3°.", XL, Y0, CW, 0.40, 16)
fml(slide, "(i)    sin 2°  ≈  2° in radian  =  2 × (π/180)  =  π/90", XL, Y0 + 0.50, CW, LH, 18)
fml(slide, "(ii)   tan 3°  ≈  3 × (π/180)  =  π/60", XL, Y0 + 0.50 + LH, CW, LH, 18)
fml(slide, "(iii)  cos 3°  ≈  1", XL, Y0 + 0.50 + 2 * LH, CW, LH, 18)
result_card(slide, "sin 2°  ≈  π/90          tan 3°  ≈  π/60          cos 3°  ≈  1", XL, Y0 + 0.50 + 3 * LH + NG, CW, 0.36, 18)

slide = new_slide("Board", "Which value of tan 2° is correct?", GREEN)
add_notes(slide, "Read. Option 2: π/90. Not 2°, not 2 rad.")
qtx(slide, "Which is correct for  tan 2°  under the small-angle approximation?", XL, Y0, CW, 0.40, 16)
choices(slide, [
    "2°",
    "π / 90",
    "π / 30",
    "2 rad",
], XL, Y0 + 0.50)
y = warn_row(slide, "tan 2°  =  2°", "tan 2°  ≈  2 × (π/180)  =  π/90", XL, Y0 + 2.20, CW)
result_card(slide, "Option 2     tan 2°  ≈  π/90", XL, y + NG, CW, 0.36, 20)
trap_card(slide, "The approximation replaces the ratio by the radian measure, never by the degree number.", XL, y + NG + 0.42, CW, 0.40)

# ── REMAINING ALLIED (not only sine) ──────────────────────────────────────────
slide = new_slide("Board", "Remaining allied — cosine and tangent", GREEN)
add_notes(slide, "Read. These were leftover after the sine allied board. Cosine at 90 changes to sine.")
qtx(slide, "Simplify.   Last board was sine only. These three are the remaining lines.", XL, Y0, CW, 0.40, 16)
fml(slide, "cos (90° − θ)  =  cos 90° cos θ  +  sin 90° sin θ  =  0  +  1 · sin θ  =  sin θ", XL, Y0 + 0.48, CW, LH, 16)
result_card(slide, "cos (90° − θ)  =  sin θ", XL, Y0 + 0.48 + LH + NG, CW, 0.32, 18)
fml(slide, "cos (90° + θ)  =  cos 90° cos θ  −  sin 90° sin θ  =  − sin θ", XL, Y0 + 0.48 + 2 * LH + 2 * NG, CW, LH, 16)
result_card(slide, "cos (90° + θ)  =  − sin θ", XL, Y0 + 0.48 + 3 * LH + 3 * NG, CW, 0.32, 18)
fml(slide, "tan (90° − θ)  =  cot θ", XL, Y0 + 0.48 + 4 * LH + 4 * NG, CW, LH, 18)
result_card(slide, "tan (90° − θ)  =  cot θ", XL, Y0 + 0.48 + 5 * LH + 5 * NG, CW, 0.32, 18)

slide = new_slide("Board", "tan 150°  —  leftover from even / odd", GREEN)
add_notes(slide, "Read. 150 = 180−30. tan negative in Q2. −1/√3.")
qtx(slide, "Find  tan 150°.", XL, Y0, CW, 0.36, 18)
fml(slide, "tan 150°  =  tan (180° − 30°)  =  − tan 30°  =  − 1/√3", XL, Y0 + 0.50, CW, LH, 20)
result_card(slide, "tan 150°  =  − 1/√3", XL, Y0 + 0.50 + LH + NG, CW, 0.36, 22)
body(slide, "Second quadrant: tangent is negative. The brother of 150° is 30°.", XL, Y0 + 0.50 + 2 * LH + 2 * NG, CW)

# ── TONIGHT — NOW WITH WORKING ────────────────────────────────────────────────
slide = new_slide("Board", "Tonight (13) and (14) — allied to finish", GREEN)
add_notes(slide, "Read. (13) sin θ. (14) − cos θ. Expand with compound. No pen needed.")
qtx(slide, "Simplify.", XL, Y0, CW, 0.32, 18)
qtx(slide, "(13)   sin (180° − θ)", XL, Y0 + 0.38, 6.0, 0.32, 18)
qtx(slide, "(14)   cos (180° + θ)", XL + 6.1, Y0 + 0.38, 5.9, 0.32, 18)
fml(slide, "sin (180° − θ)  =  sin 180° cos θ  −  cos 180° sin θ", XL, Y0 + 0.80, CW, LH, 16)
fml(slide, "                 =  0 · cos θ  −  (−1) sin θ  =  sin θ", XL, Y0 + 0.80 + LH, CW, LH, 16)
result_card(slide, "(13)   sin (180° − θ)  =  sin θ", XL, Y0 + 0.80 + 2 * LH + NG, CW, 0.32, 18)
fml(slide, "cos (180° + θ)  =  cos 180° cos θ  −  sin 180° sin θ  =  (−1) cos θ  −  0  =  − cos θ", XL, Y0 + 0.80 + 3 * LH + 2 * NG, CW, LH, 16)
result_card(slide, "(14)   cos (180° + θ)  =  − cos θ", XL, Y0 + 0.80 + 4 * LH + 3 * NG, CW, 0.32, 18)

slide = new_slide("Board", "Tonight — match the column, every row", GREEN)
add_notes(slide, "Read. A PT, B QT, C QT, D PS. Option 1. Modulus = magnitude.")
qtx(slide, "Match Column I to Column II.  Modulus means magnitude.", XL, Y0, CW, 0.32, 16)
grid_table(
    slide,
    ["Column I", "working", "match"],
    [
        ["(A)  sin 390°", "sin (360° + 30°)  =  +1/2", "P and T"],
        ["(B)  sin (−30°)", "− sin 30°  =  −1/2", "Q and T"],
        ["(C)  cos 120°", "− cos 60°  =  −1/2", "Q and T"],
        ["(D)  tan (−120°)", "− tan 120°  =  +√3", "P and S"],
    ],
    XL, Y0 + 0.38, [3.4, 5.3, 3.35], 0.40, 14, 14, scaled=False,
)
choices(slide, [
    "A → PT,   B → QT,   C → QT,   D → PS",
    "A → PT,   B → QS,   C → QT,   D → PS",
    "A → QT,   B → QS,   C → PT,   D → PS",
    "A → QS,   B → PT,   C → QT,   D → PS",
], XL, Y0 + 2.45, 7.4)
result_card(slide, "Option 1     A → PT,   B → QT,   C → QT,   D → PS", XL, Y0 + 4.12, CW, 0.36, 18)
trap_card(slide, "|√3| > 1 so D is P and S, not T.  |1/2| < 1 so A, B, C take T.", XL, Y0 + 4.52, CW, 0.32)

slide = new_slide("Tonight", "After class", AMBER)
add_notes(slide, "Daily test before the next lecture. Compound lines ten times on a rough page.")
body(slide, "Daily test before the next lecture.", XL, Y0, CW)
body(slide, "Write the six compound-angle lines ten times on a rough page.", XL, Y0 + LH, CW)
body(slide, "The two allied lines and the match-column are now on the board. Copy them once more at home.", XL, Y0 + 2 * LH + NG, CW)

# ── 5-MINUTE SUMMARY ──────────────────────────────────────────────────────────
slide = new_slide("5-minute summary", "Leftover radian — now locked", TEAL)
add_notes(slide, "This was last class from slide 29. They say 180=π. Arc=Rθ in radian.")
fml(slide, "180°  =  π rad", XL, Y0, 7.0, LH, 28)
ncert(slide, "Wherever you see π rad, put 180°.", XL + 7.15, Y0, 4.90, 1.10, 18)
y = Y0 + LH + NG
fml(slide, "degree → radian:     multiply by  π/180", XL, y + LH, CW, LH, 20)
fml(slide, "radian → degree:     multiply by  180/π", XL, y + 2 * LH, CW, LH, 20)
fml(slide, "Arc length  =  R θ     with θ in radian.     Never  R × 90°.", XL, y + 3 * LH, CW, LH, 20)
result_card(slide, "90° on a circle of radius R  →  distance  πR/2", XL, y + 4 * LH + NG, CW, 0.36, 18)

slide = new_slide("5-minute summary", "What we lock from today’s feel", TEAL)
add_notes(slide, "They say each boxed line. θ outside a ratio is radian. Compound signs.")
body(slide, "If θ is not inside sin, cos or tan, convert it to radian first.", XL, Y0, CW)
ncert(slide, "y = sin θ / θ   at   θ = 30°    →    3/π", XL, Y0 + LH + NG, CW, 0.85, 18)
y = Y0 + 2 * LH + NG + 0.70
fml(slide, "sin (A + B)  =  sin A cos B  +  cos A sin B", XL, y, CW, LH, 18)
fml(slide, "cos (A + B)  =  cos A cos B  −  sin A sin B", XL, y + LH, CW, LH, 18)
fml(slide, "sin 2θ  =  2 sin θ cos θ          cos 2θ  =  1 − 2 sin²θ", XL, y + 2 * LH, CW, LH, 18)

slide = new_slide("5-minute summary", "Even, odd, and allied", TEAL)
add_notes(slide, "cos even. sin and tan odd. sin(90−θ)=cosθ.")
fml(slide, "cos (−θ)  =  cos θ          sin (−θ)  =  − sin θ          tan (−θ)  =  − tan θ", XL, Y0, CW, LH, 18)
fml(slide, "sin (90° − θ)  =  cos θ          sin (180° − θ)  =  sin θ", XL, Y0 + LH + NG, CW, LH, 18)
result_card(slide, "sin 15°  =  (√6 − √2)/4          sin 75°  =  sin 105°  =  (√6 + √2)/4", XL, Y0 + 2 * LH + 2 * NG, CW, 0.36, 18)
result_card(slide, "sin (180° − θ)  =  sin θ          cos (180° + θ)  =  − cos θ", XL, Y0 + 3 * LH + 3 * NG, CW, 0.36, 16)

slide = new_slide("5-minute summary", "Small angle — the one lock-line", TEAL)
add_notes(slide, "sinθ≈θ, tanθ≈θ in radian. cosθ≈1. θ≤5°. tan2° is π/90 not 2°.")
ncert(slide, "sin θ  ≈  θ      tan θ  ≈  θ      cos θ  ≈  1      (θ in radian,  θ ≤ 5°)", XL, Y0, CW, 1.05, 18)
y = Y0 + 1.25
fml(slide, "sin 2°  ≈  π/90          tan 3°  ≈  π/60          cos 3°  ≈  1", XL, y, CW, LH, 18)
trap_card(slide, "Never leave the degree number in the approximation. Convert, then drop the word radian.", XL, y + LH + NG, CW, 0.45)
body(slide, "Daily test before the next lecture. Compound lines ten times tonight.", XL, y + 2 * LH + 2 * NG, CW)

prs.save(OUTPUT)
print(f"Created {len(prs.slides)} slides")
print(OUTPUT)
