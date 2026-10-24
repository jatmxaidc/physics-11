"""
Lecture 01 — Trigonometry
Follows the actual PW Basic Maths class-01 teaching order:
ratios → identities → find-x → range trap → algebra → similar triangles
→ height/distance → 0–180 table → 37°/53° → radian/arc → conversions.
Class 11 Physics  •  CBSE + NEET Foundation  •  2-hour session
"""
from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

BASE = Path(__file__).resolve().parent
OUTPUT = BASE / "Lecture_01_Trigonometry_Ratios_Identities.pptx"

NAVY = RGBColor(8, 30, 63)
BLUE = RGBColor(64, 156, 255)
BLUE_BG = RGBColor(18, 32, 52)
TEAL = RGBColor(40, 180, 190)
TEAL_BG = RGBColor(16, 40, 44)
GREEN = RGBColor(90, 214, 128)
GREEN_BG = RGBColor(18, 44, 28)
AMBER = RGBColor(255, 204, 72)
AMBER_BG = RGBColor(48, 36, 8)
ORANGE = RGBColor(255, 158, 64)
ORANGE_BG = RGBColor(48, 28, 8)
RED = RGBColor(255, 96, 96)
RED_BG = RGBColor(48, 16, 16)
TEXT = RGBColor(255, 255, 255)
MUTED = RGBColor(176, 182, 192)
STROKE = RGBColor(72, 78, 90)
WHITE = RGBColor(255, 255, 255)
SOFT = RGBColor(22, 24, 28)
BLACK = RGBColor(0, 0, 0)
DARK_PANEL = RGBColor(12, 14, 18)
GOLD = RGBColor(255, 214, 70)
NCERT_FILL = RGBColor(255, 248, 225)
NCERT_LINE = RGBColor(218, 168, 64)
# Same colour roles on every slide:
#   GOLD   — titles only
#   QCOL   — the question they must solve
#   WORK   — formulas and working steps
#   TEXT   — spoken explanation
#   GREEN  — final answer
#   RED    — wrong    ORANGE — trap / P    BLUE — B    GREEN — H on diagrams
QCOL = ORANGE
WORK = RGBColor(90, 200, 255)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

XL, CW, Y0, YN, YF = 0.65, 12.05, 1.18, 5.52, 7.12
# Room for the opposite-side label to the left of the triangle.
TRI_X = XL + 0.55
# Tight inside a group. Extra gap only between groups. Rest of board is for writing.
LH = 0.32
NG = 0.18
PT_SCALE = 0.70


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


def _accent(slide, x, y, h, color, w=0.065):
    return _solid(slide, x, y, w, h, color)


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


def tx_two_runs(slide, text1, color1, bold1, text2, color2, bold2, x, y, w, h, size=15):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(0)
    p = tf.paragraphs[0]
    p.space_before = Pt(0)
    p.space_after = Pt(0)
    p.line_spacing = 1.1
    for text, color, bold in ((text1, color1, bold1), (text2, color2, bold2)):
        r = p.add_run()
        r.text = text
        r.font.name = "Aptos"
        r.font.size = Pt(max(7, round(size * PT_SCALE, 1)))
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
    """One NCERT-style formula box. Use only for the result, usually on the right."""
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


def body(slide, text, x, y, w, h=0.40, size=14):
    return tx(slide, text, x, y, w, h, size, TEXT, False, spa=0, scaled=False)


def qtx(slide, text, x, y, w, h=0.44, size=16):
    return tx(slide, text, x, y, w, h, size, QCOL, True, spa=0, scaled=False)


def choices(slide, items, x, y, w=None):
    """NEET-style options: circled number + white text. Not question colour."""
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
        tx(slide, str(i + 1), x, cy + 0.06, 0.30, 0.26, 12, WHITE, True, PP_ALIGN.CENTER, spa=0, scaled=False)
        tx(slide, item, x + 0.42, cy + 0.06, w - 0.42, 0.28, 16, TEXT, False, spa=0, scaled=False)
    return y + len(items) * row


def formula_card(slide, formula, note, x, y, w, h, fsize=18):
    fml(slide, formula, x, y, w, LH, fsize)
    if note:
        body(slide, note, x, y + LH, w, 0.28)
    return y + (LH + 0.28 if note else LH)


def trap_card(slide, text, x, y, w, h):
    tx(slide, text, x, y, w, h, 14, ORANGE, False, spa=0, scaled=False)


def result_card(slide, text, x, y, w, h, tsize=14):
    return tx(slide, text, x, y, w, 0.36, max(tsize, 18), GREEN, True, spa=0)


def warn_row(slide, wrong, correct, x, y, w):
    tx(slide, f"Wrong:  {wrong}", x, y, w, 0.30, 14, RED, False, spa=0, scaled=False)
    tx(slide, f"Right:  {correct}", x, y + LH, w, 0.30, 14, GREEN, False, spa=0, scaled=False)
    return y + LH + LH


def right_triangle(slide, x, y, w, h, opp="opp", adj="adj", hyp="hyp"):
    """Right triangle: 90° at bottom-left, θ inside the bottom-right angle."""
    tri = slide.shapes.add_shape(MSO_SHAPE.RIGHT_TRIANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    tri.fill.solid()
    tri.fill.fore_color.rgb = RGBColor(42, 112, 178)
    tri.line.color.rgb = RGBColor(170, 220, 255)
    tri.line.width = Pt(2.25)
    # Square sits in the corner: left edge on the vertical side, bottom on the base.
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
    # θ inside the angle: on the base, left of the vertex, below the hypotenuse.
    tw, th = 0.24, 0.18
    from_right = max(tw + 0.08, (th + 0.08) * w / h)
    tx(slide, "θ", x + w - from_right, y + h - th - 0.02, tw, th, 16, GOLD, True, spa=0)


def grid_table(slide, headers, rows, x, y, col_w, row_h=0.42, hsize=11, csize=12):
    col_x = [x]
    for w in col_w[:-1]:
        col_x.append(col_x[-1] + w)
    for j, ht in enumerate(headers):
        _solid(slide, col_x[j], y, col_w[j], row_h, RGBColor(28, 36, 52), False)
        tx(slide, ht, col_x[j] + 0.04, y + 0.08, col_w[j] - 0.08, 0.26, hsize, WHITE, True, PP_ALIGN.CENTER, spa=0)
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
                slide, val, col_x[j] + 0.04, ry + 0.08, col_w[j] - 0.08, 0.26,
                csize, WORK if j == 0 else TEXT, j == 0, PP_ALIGN.CENTER, spa=0,
            )


def tx_3d(slide, text, x, y, w, h, size, face, edge, font="Arial Black"):
    tx(slide, text, x + 0.045, y + 0.045, w, h, size, edge, True, PP_ALIGN.CENTER, spa=0, scaled=False, font=font)
    tx(slide, text, x, y, w, h, size, face, True, PP_ALIGN.CENTER, spa=0, scaled=False, font=font)


# ── 01 COVER ──────────────────────────────────────────────────────────────────
slide = prs.slides.add_slide(BLANK)
_bg(slide)
tx_3d(slide, "EXCELLENCE", XL, 0.55, CW, 1.05, 54, GOLD, RGBColor(210, 210, 220))
tx_3d(slide, "CLASS 11   ·   2026–27", XL, 1.62, CW, 0.55, 26, WHITE, RGBColor(90, 90, 100), font="Aptos")
_solid(slide, 4.15, 2.28, 5.00, 0.028, WHITE)

bx, by, bw = 2.05, 2.70, 9.20
_card_bg(slide, bx, by, bw, 2.35, WHITE, WHITE, 0)
_card_bg(slide, bx, by + 0.72, bw, 1.63, BLACK, RGBColor(220, 220, 220), 1.25)
_solid(slide, bx + bw / 2, by + 0.16, 0.018, 0.42, RGBColor(40, 40, 40))
tx(slide, "Physics", bx, by + 0.16, bw / 2 - 0.10, 0.44, 22, BLACK, True, PP_ALIGN.CENTER, spa=0, scaled=False)
tx(slide, "Lecture - 01", bx + bw / 2, by + 0.16, bw / 2, 0.44, 22, BLACK, True, PP_ALIGN.CENTER, spa=0, scaled=False)
tx(slide, "Basic Maths and Calculus", bx, by + 1.05, bw, 0.55, 28, WHITE, True, PP_ALIGN.CENTER, spa=0, scaled=False, font="Georgia")
tx(slide, "Trigonometry", bx, by + 1.62, bw, 0.40, 18, GOLD, True, PP_ALIGN.CENTER, spa=0, scaled=False)

tx(slide, "By — Excellence Programme", XL, 5.30, CW, 0.42, 22, WHITE, True, PP_ALIGN.CENTER, spa=0, scaled=False)
tx(slide, "1 September 2026", XL, 5.80, CW, 0.32, 16, MUTED, True, PP_ALIGN.CENTER, spa=0, scaled=False)
_footer(slide, 1)
add_notes(slide, "Cover. Next slide is topics to be covered.")

# ── 01 TOPICS ─────────────────────────────────────────────────────────────────
slide = prs.slides.add_slide(BLANK)
_bg(slide)
tx(slide, "Topics to be covered", XL, 0.38, CW, 0.55, 28, GOLD, True, spa=0, scaled=False)
_solid(slide, XL, 0.98, 3.40, 0.035, GOLD)

topics = [
    "Trigonometric ratios  —  sin, cos, tan, cosec, sec, cot",
    "Trigonometric identities from  P² + B² = H²",
    "Finding unknown sides from a given ratio",
    "Range of sine and cosine",
    "Algebra with ratios, and similar triangles",
    "Height and distance",
    "Standard angles  0° to 180°,  and  37° / 53°",
    "Radian measure,  arc = Rθ,  degree ↔ radian",
]
for i, line in enumerate(topics):
    y = 1.22 + i * 0.62
    tx(slide, f"{i + 1:02d}", XL, y, 0.70, 0.42, 18, GOLD, True, spa=0, scaled=False)
    tx(slide, line, XL + 0.85, y, CW - 0.85, 0.42, 18, WHITE, False, spa=0, scaled=False)

_footer(slide, 2)
add_notes(slide, "Read the list once. Then name P, B, H.")

# ── 02 RATIOS — SIDES ─────────────────────────────────────────────────────────
slide = new_slide("Ratios", "Name the sides from θ", BLUE)
add_notes(slide, "P = perpendicular (opposite θ), B = base (adjacent), H = hypotenuse.")
right_triangle(slide, TRI_X, Y0, 3.30, 2.50, opp="P", adj="B", hyp="H")
tx_x = TRI_X + 3.80
tx(slide, "P  —  perpendicular  —  opposite θ", tx_x, Y0, 6.4, 0.26, 18, ORANGE, True, spa=0)
tx(slide, "B  —  base  —  adjacent to θ", tx_x, Y0 + 0.26, 6.4, 0.26, 18, BLUE, True, spa=0)
tx(slide, "H  —  hypotenuse  —  longest side", tx_x, Y0 + 0.52, 6.4, 0.26, 18, GREEN, True, spa=0)
trap_card(slide, "Not from how the figure sits.", tx_x, Y0 + 0.52 + 0.26 + NG, 6.4, 0.32)

# ── 02 RATIOS — SIN / COSEC ───────────────────────────────────────────────────
slide = new_slide("Ratios", "sin θ and cosec θ", BLUE)
add_notes(slide, "sin = P/H. cosec = H/P. Product is 1.")
y = Y0
fml(slide, "sin θ  =  P / H", XL, y, 6.8, LH, 24)
body(slide, "Opposite over hypotenuse.", XL, y + LH, 6.8)
y = y + LH + 0.28 + NG
fml(slide, "cosec θ  =  H / P", XL, y, 6.8, LH, 24)
body(slide, "Reciprocal of sin θ.   If sin θ = 3/5, then cosec θ = 5/3.", XL, y + LH, 6.8)
ncert(slide, "sin θ  ·  cosec θ  =  1", XL + 7.15, Y0, 4.90, 1.45, 20)
y = y + LH + 0.28 + NG
body(slide, "sin, cos, tan are dimensionless. Never write sin θ = 0.6 m.", XL, y, CW)
body(slide, "cosec θ = 1 / sin θ.", XL, y + LH, CW)

# ── 02 RATIOS — COS / SEC ─────────────────────────────────────────────────────
slide = new_slide("Ratios", "cos θ and sec θ", BLUE)
add_notes(slide, "cos = B/H. sec = H/B. Product is 1.")
y = Y0
fml(slide, "cos θ  =  B / H", XL, y, 6.8, LH, 24)
body(slide, "Adjacent over hypotenuse.", XL, y + LH, 6.8)
y = y + LH + 0.28 + NG
fml(slide, "sec θ  =  H / B", XL, y, 6.8, LH, 24)
body(slide, "Reciprocal of cos θ.   If cos θ = 4/5, then sec θ = 5/4.", XL, y + LH, 6.8)
ncert(slide, "cos θ  ·  sec θ  =  1", XL + 7.15, Y0, 4.90, 1.45, 20)
y = y + LH + 0.28 + NG
body(slide, "Never write cos θ = 0.8 m. A ratio has no unit.", XL, y, CW)
body(slide, "sec θ = 1 / cos θ.", XL, y + LH, CW)

# ── 02 RATIOS — TAN / COT ─────────────────────────────────────────────────────
slide = new_slide("Ratios", "tan θ and cot θ", BLUE)
add_notes(slide, "tan = P/B. cot = B/P. Product is 1.")
y = Y0
fml(slide, "tan θ  =  P / B", XL, y, 6.8, LH, 24)
body(slide, "Opposite over adjacent.", XL, y + LH, 6.8)
y = y + LH + 0.28 + NG
fml(slide, "cot θ  =  B / P", XL, y, 6.8, LH, 24)
body(slide, "Reciprocal of tan θ.   Also:  tan θ = sin θ / cos θ.", XL, y + LH, 6.8)
ncert(slide, "tan θ  ·  cot θ  =  1", XL + 7.15, Y0, 4.90, 1.45, 20)
y = y + LH + 0.28 + NG
body(slide, "sin · cosec = 1     cos · sec = 1     tan · cot = 1", XL, y, CW)
body(slide, "cot θ = 1 / tan θ.", XL, y + LH, CW)

# ── 03 IDENTITIES — ÷ H² ──────────────────────────────────────────────────────
slide = new_slide("Identities", "Divide Pythagoras by H²", BLUE)
add_notes(slide, "P²+B²=H². Divide by H². sin²+cos²=1. Rearrange for cos or sin.")
y = Y0
body(slide, "Start here. Divide both sides by H².", XL, y, 7.0)
y += LH
fml(slide, "P²  +  B²  =  H²", XL, y, 7.0, LH, 24)
fml(slide, "P² / H²  +  B² / H²  =  H² / H²", XL, y + LH, 7.0, LH, 20)
fml(slide, "(P / H)²  +  (B / H)²  =  1", XL, y + 2 * LH, 7.0, LH, 20)
ncert(slide, "sin²θ  +  cos²θ  =  1", XL + 7.15, Y0, 4.90, 1.70, 22)
body(slide, "Also:  cos θ = √(1 − sin²θ)     and     sin θ = √(1 − cos²θ)", XL, y + 3 * LH + NG, 7.0)

# ── 03 IDENTITIES — ÷ B² ──────────────────────────────────────────────────────
slide = new_slide("Identities", "Divide Pythagoras by B²", TEAL)
add_notes(slide, "Divide by B². 1 + tan² = sec². P/B = tan, H/B = sec.")
y = Y0
body(slide, "Same start. Now divide both sides by B².", XL, y, 7.0)
y += LH
fml(slide, "P²  +  B²  =  H²", XL, y, 7.0, LH, 24)
fml(slide, "P² / B²  +  B² / B²  =  H² / B²", XL, y + LH, 7.0, LH, 20)
fml(slide, "(P / B)²  +  1  =  (H / B)²", XL, y + 2 * LH, 7.0, LH, 20)
ncert(slide, "1  +  tan²θ  =  sec²θ", XL + 7.15, Y0, 4.90, 1.70, 22)
body(slide, "Because  P/B = tan θ     and     H/B = sec θ.", XL, y + 3 * LH + NG, 7.0)

# ── 03 IDENTITIES — ÷ P² ──────────────────────────────────────────────────────
slide = new_slide("Identities", "Divide Pythagoras by P²", AMBER)
add_notes(slide, "Divide by P². 1 + cot² = cosec². If they remember only one, sin²+cos²=1.")
y = Y0
body(slide, "Same start. Now divide both sides by P².", XL, y, 7.0)
y += LH
fml(slide, "P²  +  B²  =  H²", XL, y, 7.0, LH, 24)
fml(slide, "P² / P²  +  B² / P²  =  H² / P²", XL, y + LH, 7.0, LH, 20)
fml(slide, "1  +  (B / P)²  =  (H / P)²", XL, y + 2 * LH, 7.0, LH, 20)
ncert(slide, "1  +  cot²θ  =  cosec²θ", XL + 7.15, Y0, 4.90, 1.70, 22)
body(slide, "Because  B/P = cot θ     and     H/P = cosec θ.     If you keep only one line:  sin²θ + cos²θ = 1", XL, y + 3 * LH + NG, 7.0)

# ── 04 FIND X — ASK ───────────────────────────────────────────────────────────
slide = new_slide("Question", "Find x", GREEN)
add_notes(slide, "Wait. Do not click next until they try. Answer: H = √5, sin = 1/√5, cos = 2/√5, tan = 1/2.")
qtx(slide, "In the right triangle shown, the perpendicular is 1 and the base is 2. The hypotenuse is x. Find the value of x, sin θ, cos θ and tan θ.", XL, Y0, CW, 0.70, 16)
right_triangle(slide, TRI_X, Y0 + 0.80, 3.30, 1.75, opp="1", adj="2", hyp="x")

# ── 04 FIND X — SOLVE H ───────────────────────────────────────────────────────
slide = new_slide("Solution", "First find H", GREEN)
add_notes(slide, "P²+B²=H² → x = √5. Ratios on the next slide.")
body(slide, "Right triangle:  P = 1,  B = 2,  H = x.", XL, Y0, CW)
right_triangle(slide, TRI_X, Y0 + LH + NG, 3.30, 1.75, opp="1", adj="2", hyp="x")
rx = TRI_X + 3.80
fml(slide, "P² + B² = H²", rx, Y0 + LH + NG, 7.2, LH, 22)
body(slide, "1² + 2² = x²", rx, Y0 + LH + NG + LH, 7.2)
ncert(slide, "x  =  √5", rx, Y0 + LH + NG + 2 * LH + NG, 7.2, 0.95, 26)

# ── 04 FIND X — SOLVE RATIOS ──────────────────────────────────────────────────
slide = new_slide("Solution", "Now sin θ, cos θ, tan θ", GREEN)
add_notes(slide, "sin = 1/√5, cos = 2/√5, tan = 1/2.")
body(slide, "P = 1,  B = 2,  H = √5", XL, Y0, CW)
y = Y0 + LH + NG
fml(slide, "sin θ  =  P / H  =  1 / √5", XL, y, 7.0, LH, 22)
fml(slide, "cos θ  =  B / H  =  2 / √5", XL, y + LH, 7.0, LH, 22)
fml(slide, "tan θ  =  P / B  =  1 / 2", XL, y + 2 * LH, 7.0, LH, 22)
ncert(slide, "tan θ  =  sin θ / cos θ  =  1/2", XL + 7.15, y, 4.90, 1.20, 18)
body(slide, "Check:  (1/√5) / (2/√5)  =  1/2.  Same.", XL, y + 3 * LH + NG, CW)

# ── 05 TAN = 2 — ASK ──────────────────────────────────────────────────────────
slide = new_slide("Question", "If tan θ = 2", GREEN)
add_notes(slide, "Wait. Answer: B = k, P = 2k, H = √5 k, sin = 2/√5, cos = 1/√5.")
qtx(slide, "If tan θ = 2, then find the perpendicular, the base and the hypotenuse. Also find sin θ and cos θ. Take θ acute.", XL, Y0, CW, 0.70, 16)

# ── 05 TAN = 2 — SOLVE SIDES ──────────────────────────────────────────────────
slide = new_slide("Solution", "tan θ = 2  →  find P, B, H", GREEN)
add_notes(slide, "tan = 2/1. B = k, P = 2k, H = √5 k.")
tx(slide, "Given tan θ = 2  (θ acute).", XL, Y0, CW, LH, 18, WORK, True, spa=0)
y = Y0 + LH + NG
y = formula_card(slide, "tan θ  =  P / B  =  2  =  2 / 1", "Let B = k.  Then P = 2k.", XL, y, CW, 1.10, 22)
y = formula_card(slide, "H²  =  (2k)² + k²  =  5k²", "H = √5 k   (take + because θ is acute)", XL, y + NG, CW, 1.10, 22)
result_card(slide, "P = 2k     B = k     H = √5 k", XL, y + NG, CW, 0.36, 18)

# ── 05 TAN = 2 — SOLVE SIN COS ────────────────────────────────────────────────
slide = new_slide("Solution", "Now sin θ and cos θ", GREEN)
add_notes(slide, "k cancels. sin = 2/√5, cos = 1/√5.")
tx(slide, "P = 2k,   B = k,   H = √5 k", XL, Y0, CW, LH, 18, WORK, True, spa=0)
y = Y0 + LH + NG
y = formula_card(slide, "sin θ  =  P / H  =  2k / √5 k  =  2 / √5", "k cancels.", XL, y, CW, 1.15, 22)
y = formula_card(slide, "cos θ  =  B / H  =  k / √5 k  =  1 / √5", "k cancels.", XL, y + NG, CW, 1.15, 22)
body(slide, "A ratio only fixes the shape. The scale k always cancels in sin, cos, tan.", XL, y + NG, CW)

# ── 06 RANGE — ASK ────────────────────────────────────────────────────────────
slide = new_slide("Question", "Range of sin θ", GREEN)
add_notes(slide, "Wait. Do not hint it is impossible. Let them try. Then next slide.")
qtx(slide, "If sin θ = 2, then find the value of cos θ and tan θ.", XL, Y0, CW, 0.50, 18)

# ── 06 RANGE — SOLVE GEOMETRY ─────────────────────────────────────────────────
slide = new_slide("Solution", "Method 1 — geometry", ORANGE)
add_notes(slide, "H > P so sin ≤ 1. sin = 2 is impossible.")
body(slide, "If sin θ = 2, find cos θ and tan θ.", XL, Y0, CW)
y = warn_row(slide, "Start calculating cos θ from the identity without checking the range.", "First ask: can sin θ equal 2?", XL, Y0 + LH + NG, CW)
y = formula_card(slide, "sin θ  =  P / H", "H is the longest side, so H > P.  Then P/H ≤ 1.", XL, y + NG, CW, 1.20, 22)
result_card(slide, "sin θ = 2 is impossible.  There is no such triangle.", XL, y + NG, CW, 0.36, 18)

# ── 06 RANGE — SOLVE IDENTITY ─────────────────────────────────────────────────
slide = new_slide("Solution", "Method 2 — identity", ORANGE)
add_notes(slide, "cos² = −3. Not real. Then lock the range.")
body(slide, "Try the identity anyway. It also fails.", XL, Y0, CW)
y = Y0 + LH + NG
y = formula_card(slide, "sin²θ  +  cos²θ  =  1", "Put sin θ = 2.", XL, y, CW, 1.05, 22)
y = formula_card(slide, "(2)²  +  cos²θ  =  1", "cos²θ = 1 − 4 = −3.  Not a real value.", XL, y + NG, CW, 1.05, 22)
result_card(slide, "−1 ≤ sin θ ≤ 1     and     −1 ≤ cos θ ≤ 1     tan θ has no finite range.", XL, y + NG, CW, 0.36, 16)

# ── 07 ALGEBRA — ASK ──────────────────────────────────────────────────────────
slide = new_slide("Question", "Find tan θ from a sin–cos fraction", GREEN)
add_notes(slide, "Wait. Answer: tan θ = 8/3.")
qtx(slide, "If  (sin θ + cos θ) / (sin θ − cos θ)  =  11/5 ,  then find the value of tan θ.", XL, Y0, CW, 0.55, 18)

# ── 07 ALGEBRA — SOLVE EXPAND ─────────────────────────────────────────────────
slide = new_slide("Solution", "Cross multiply, then expand", GREEN)
add_notes(slide, "5(sin+cos)=11(sin−cos). Expand. Collect on next slide.")
tx(slide, "(sin θ + cos θ) / (sin θ − cos θ)  =  11/5", XL, Y0, CW, LH, 18, WORK, True, spa=0)
y = Y0 + LH + NG
y = formula_card(slide, "5 (sin θ + cos θ)  =  11 (sin θ − cos θ)", "Cross multiply.", XL, y, CW, 1.10, 20)
y = formula_card(slide, "5 sin θ + 5 cos θ  =  11 sin θ − 11 cos θ", "Open the brackets.", XL, y + NG, CW, 1.10, 20)
tx(slide, "Next: collect sin on one side, cos on the other.", XL, y + NG, CW, LH, 14, MUTED, spa=0, scaled=False)

# ── 07 ALGEBRA — SOLVE TAN ────────────────────────────────────────────────────
slide = new_slide("Solution", "Collect, then divide", GREEN)
add_notes(slide, "16 cos = 6 sin. tan = 8/3.")
tx(slide, "5 sin θ + 5 cos θ  =  11 sin θ − 11 cos θ", XL, Y0, CW, LH, 18, WORK, True, spa=0)
y = Y0 + LH + NG
fml(slide, "5 cos θ + 11 cos θ  =  11 sin θ − 5 sin θ", XL, y, CW, LH, 20)
fml(slide, "16 cos θ  =  6 sin θ", XL, y + LH, CW, LH, 20)
fml(slide, "sin θ / cos θ  =  16 / 6  =  8 / 3", XL, y + 2 * LH, 7.0, LH, 20)
ncert(slide, "tan θ  =  8 / 3", XL + 7.15, y, 4.90, 1.20, 24)

# ── 08 SIMILAR — ASK ──────────────────────────────────────────────────────────
slide = new_slide("Question", "Nested right triangles — same θ, same tan θ", GREEN)
add_notes(slide, "Wait. Answer: x = 8/3, y = 3/2.")
qtx(slide, "A large right triangle has opposite side 4 and adjacent side 3. Nested similar slices share the same angle θ. In the small slice the opposite is 2 and the adjacent is y. In the middle slice the opposite is x and the adjacent is 2. Find the values of x and y.", XL, Y0, CW, 1.05, 15)
choices(slide, [
    "x = 4/3 ,   y = 3/2",
    "x = 8/3 ,   y = 6/2",
    "x = 8/3 ,   y = 3/2",
    "x = 3/2 ,   y = 8/3",
], XL, Y0 + 1.15)

# ── 08 SIMILAR — SOLVE Y ──────────────────────────────────────────────────────
slide = new_slide("Solution", "Find y from tan θ = 4/3", GREEN)
add_notes(slide, "4/3 = 2/y → y = 3/2. x on the next slide.")
tx(slide, "Large triangle: opposite 4, adjacent 3.  Small slice: opposite 2, adjacent y.", XL, Y0, CW, LH, 14, TEXT, spa=0, scaled=False)
y = Y0 + LH + NG
y = formula_card(slide, "tan θ  =  4 / 3  =  2 / y", "Same θ, so same tan.", XL, y, CW, 1.15, 22)
y = formula_card(slide, "4 y  =  6", "Cross multiply.", XL, y + NG, CW, 1.00, 22)
result_card(slide, "y  =  3 / 2", XL, y + NG, CW, 0.36, 22)

# ── 08 SIMILAR — SOLVE X ──────────────────────────────────────────────────────
slide = new_slide("Solution", "Find x from tan θ = 4/3", GREEN)
add_notes(slide, "4/3 = x/2 → x = 8/3. Trap: do not mix triangles.")
tx(slide, "Middle slice: opposite x, adjacent 2.  Same θ as the large triangle.", XL, Y0, CW, LH, 14, TEXT, spa=0, scaled=False)
y = Y0 + LH + NG
y = formula_card(slide, "tan θ  =  4 / 3  =  x / 2", "Same θ, so same tan.", XL, y, CW, 1.10, 22)
y = formula_card(slide, "3 x  =  8", "Cross multiply.", XL, y + NG, CW, 0.85, 22)
result_card(slide, "x  =  8 / 3", XL, y + NG, CW, 0.36, 20)
trap_card(slide, "Do not mix a height from one triangle with a base from another. Each ratio must come from the SAME triangle.", XL, y + NG + LH + NG, CW, 0.50)

# ── 10 BREAK ──────────────────────────────────────────────────────────────────
slide = prs.slides.add_slide(BLANK)
_bg(slide)
_solid(slide, 0, 0, 13.333, 0.06, BLUE)
tx(slide, "BREAK", XL, 1.85, CW, 0.55, 34, GOLD, True, PP_ALIGN.CENTER, spa=4)
tx(slide, "Next: height and distance  →  0°–180° table  →  37°/53°  →  radian", XL, 2.70, CW, 0.40, 16, BLUE, True, PP_ALIGN.CENTER, spa=0)
_footer(slide, len(prs.slides))
add_notes(slide, "Break. Stretch. Then height and distance.")

# ── 11 ELEVATION — ASK ────────────────────────────────────────────────────────
slide = new_slide("Question", "Observer between two towers — elevation", TEAL)
add_notes(slide, "Wait. h1 = x tanθ, h2 = (d−x) tanφ. Difference = x tanθ − (d−x) tanφ.")
qtx(slide, "From a point on level ground between two vertical towers that are d metres apart, the angles of elevation to the tops of the nearer and farther towers are θ and φ respectively. The observer is x metres from the nearer tower. What is the difference in heights of the towers?", XL, Y0, CW, 1.05, 15)
choices(slide, [
    "d (tan θ − tan φ)",
    "x tan θ  −  (d − x) tan φ",
    "tan θ − tan φ · d",
    "(tan θ − tan φ) / d",
], XL, Y0 + 1.15)

# ── 11 ELEVATION — SOLVE H1 ───────────────────────────────────────────────────
slide = new_slide("Solution", "Nearer tower", TEAL)
add_notes(slide, "h1 = x tan θ.")
tx(slide, "Observer is x metres from the nearer tower. Elevation to that top is θ.", XL, Y0, CW, LH, 14, TEXT, spa=0, scaled=False)
y = formula_card(slide, "tan θ  =  h₁ / x", "Opposite = height.  Adjacent = x.", XL, Y0 + LH + NG, CW, 1.20, 24)
result_card(slide, "h₁  =  x tan θ", XL, y + NG, CW, 0.36, 24)
tx(slide, "Next: the farther tower.", XL, y + NG + LH + NG, CW, LH, 14, MUTED, spa=0, scaled=False)

# ── 11 ELEVATION — SOLVE H2 AND DIFF ──────────────────────────────────────────
slide = new_slide("Solution", "Farther tower, then the difference", TEAL)
add_notes(slide, "h2 = (d−x) tanφ. Difference = x tanθ − (d−x) tanφ.")
tx(slide, "Observer is (d − x) metres from the farther tower. Elevation to that top is φ.", XL, Y0, CW, LH, 14, TEXT, spa=0, scaled=False)
y = formula_card(slide, "h₂  =  (d − x) tan φ", "Same idea. Adjacent is now d − x.", XL, Y0 + LH + NG, CW, 1.05, 22)
result_card(slide, "h₁ − h₂  =  x tan θ  −  (d − x) tan φ", XL, y + NG, CW, 0.36, 20)
trap_card(slide, "If the question asks for the farther minus the nearer, flip the sign. Always sketch before you write tan.", XL, y + NG + LH + NG, CW, 0.55)

# ── 12 DEPRESSION — ASK ───────────────────────────────────────────────────────
slide = new_slide("Question", "From the top of one tower — depression", TEAL)
add_notes(slide, "Wait. D = (h1−h2)/tanθ.")
qtx(slide, "Two towers are separated by a horizontal distance D. From the top of the first tower (height h₁), the angle of depression to the top of the second tower (height h₂) is θ. Which relation correctly gives D in terms of h₁, h₂ and θ? Take h₁ > h₂.", XL, Y0, CW, 1.05, 15)
choices(slide, [
    "D  =  (h₁ + h₂) / tan θ",
    "D  =  (h₁ − h₂) / tan θ",
    "D  =  (h₂ − h₁) / tan θ",
    "D  =  tan θ · (h₁ + h₂)",
], XL, Y0 + 1.15)

# ── 12 DEPRESSION — SOLVE TAN ─────────────────────────────────────────────────
slide = new_slide("Solution", "Write tan θ first", TEAL)
add_notes(slide, "Opposite = h1−h2. Adjacent = D.")
tx(slide, "From the top of the first tower, look down to the top of the second. h₁ > h₂.", XL, Y0, CW, LH, 14, TEXT, spa=0, scaled=False)
y = formula_card(slide, "tan θ  =  (h₁ − h₂) / D", "Opposite = height drop.  Adjacent = D.", XL, Y0 + LH + NG, CW, 1.30, 24)
tx(slide, "Depression looks down, so the opposite side is the difference of heights, never the sum.", XL, y + NG, CW, LH, 14, MUTED, spa=0, scaled=False)
tx(slide, "Next: solve for D.", XL, y + NG + LH, CW, LH, 14, MUTED, spa=0, scaled=False)

# ── 12 DEPRESSION — SOLVE D ───────────────────────────────────────────────────
slide = new_slide("Solution", "Solve for D", TEAL)
add_notes(slide, "D = (h1−h2)/tanθ. Wrong options use sum or multiply by tan.")
tx(slide, "tan θ  =  (h₁ − h₂) / D", XL, Y0, CW, LH, 18, WORK, True, spa=0)
result_card(slide, "D  =  (h₁ − h₂) / tan θ", XL, Y0 + LH + NG, CW, 0.36, 24)
trap_card(slide, "Wrong options use (h₁ + h₂) or multiply by tan θ. Depression looks down, so the opposite side is the difference of heights, never the sum.", XL, Y0 + 2 * LH + 2 * NG, CW, 0.70)

# ── 13 TABLE 0–90 ─────────────────────────────────────────────────────────────
slide = new_slide("Table", "sin, cos, tan from 0° to 90°", TEAL)
add_notes(slide, "Fill column by column. tan 90° is not defined.")
headers = ["θ", "0°", "30°", "45°", "60°", "90°"]
rows = [
    ["sin", "0", "1/2", "1/√2", "√3/2", "1"],
    ["cos", "1", "√3/2", "1/√2", "1/2", "0"],
    ["tan", "0", "1/√3", "1", "√3", "undef."],
]
grid_table(slide, headers, rows, XL, Y0, [1.70] + [2.07] * 5, row_h=0.52, hsize=14, csize=16)
result_card(slide, "tan 90° is not defined.", XL, Y0 + 2.20 + NG, CW, 0.36, 18)

# ── 13 TABLE 90–180 ───────────────────────────────────────────────────────────
slide = new_slide("Table", "sin, cos, tan from 90° to 180°", TEAL)
add_notes(slide, "After 90° cosine is negative. Range lock on this slide.")
headers = ["θ", "90°", "120°", "135°", "150°", "180°"]
rows = [
    ["sin", "1", "√3/2", "1/√2", "1/2", "0"],
    ["cos", "0", "−1/2", "−1/√2", "−√3/2", "−1"],
    ["tan", "undef.", "−√3", "−1", "−1/√3", "0"],
]
grid_table(slide, headers, rows, XL, Y0, [1.70] + [2.07] * 5, row_h=0.52, hsize=14, csize=16)
result_card(slide, "Range: −1 ≤ sin θ ≤ 1     and     −1 ≤ cos θ ≤ 1", XL, Y0 + 2.20 + NG, CW, 0.36, 16)

# ── 13 TABLE — COPY VALUES (PW p23) ───────────────────────────────────────────
slide = new_slide("Question", "Read these from the table", GREEN)
add_notes(slide, "Wait. They copy from the 0–180 table. Answers on the next slide.")
qtx(slide, "From the 0° to 180° table, find the value of each of the following.", XL, Y0, CW, 0.40, 16)
vals_q = [
    "(i)     sin 120°",
    "(ii)    cos 150°",
    "(iii)   tan 135°",
    "(iv)    sin 150°",
    "(v)     cos 120°",
]
for i, line in enumerate(vals_q):
    tx(slide, line, XL, Y0 + 0.50 + i * LH, CW, LH, 16, TEXT, False, spa=0, scaled=False)

slide = new_slide("Solution", "From the 0°–180° table", GREEN)
add_notes(slide, "sin120=√3/2, cos150=−√3/2, tan135=−1, sin150=1/2, cos120=−1/2.")
vals_a = [
    "sin 120°  =  √3 / 2",
    "cos 150°  =  −√3 / 2",
    "tan 135°  =  −1",
    "sin 150°  =  1 / 2",
    "cos 120°  =  −1 / 2",
]
for i, line in enumerate(vals_a):
    result_card(slide, line, XL, Y0 + i * LH, CW, LH, 18)

# ── 14 37 53 — ASK RATIOS ─────────────────────────────────────────────────────
slide = new_slide("Question", "3–4–5 triangle — 37° and 53°", BLUE)
add_notes(slide, "Wait. Opposite 3 is 37°, opposite 4 is 53°.")
right_triangle(slide, TRI_X, Y0, 2.40, 2.70, opp="4", adj="3", hyp="5")
qx = TRI_X + 2.90
qtx(slide, "In a 3–4–5 right triangle, find sin, cos and tan of 37°, and sin, cos and tan of 53°.", qx, Y0, 7.0, 0.80, 16)

# ── 14 37 53 — SOLVE RATIOS ───────────────────────────────────────────────────
slide = new_slide("Solution", "3–4–5 triangle — 37° and 53°", BLUE)
add_notes(slide, "Opposite 3 is 37°, opposite 4 is 53°.")
right_triangle(slide, TRI_X, Y0, 2.40, 2.70, opp="4", adj="3", hyp="5")
rx = TRI_X + 2.90
fml(slide, "sin 37° = 3/5      cos 37° = 4/5      tan 37° = 3/4", rx, Y0, 7.5, LH, 18)
fml(slide, "sin 53° = 4/5      cos 53° = 3/5      tan 53° = 4/3", rx, Y0 + LH, 7.5, LH, 18)
body(slide, "tan 37° < tan 53°  because  3/4 < 4/3.", rx, Y0 + 2 * LH + NG, 7.5)
body(slide, "Also:  sin 53° × cosec 53° = 1.", rx, Y0 + 3 * LH + NG, 7.5)

# ── 14 0 TO 90 — ASK ──────────────────────────────────────────────────────────
slide = new_slide("Question", "From 0° to 90°", GREEN)
add_notes(slide, "Wait. sin increases, cos decreases.")
qtx(slide, "If the angle increases from 0° to 90°, then what happens to the value of sin θ and the value of cos θ?", XL, Y0, CW, 0.55, 16)
choices(slide, [
    "sin θ increases,  cos θ decreases",
    "sin θ decreases,  cos θ increases",
    "both decrease",
    "both increase",
], XL, Y0 + 0.70)

# ── 14 0 TO 90 — SOLVE ────────────────────────────────────────────────────────
slide = new_slide("Solution", "From 0° to 90°", GREEN)
add_notes(slide, "sin increases, cos decreases.")
formula_card(slide, "sin θ  increases", "0 → 1/2 → 1/√2 → √3/2 → 1", XL, Y0, 5.85, 1.40, 22)
formula_card(slide, "cos θ  decreases", "1 → √3/2 → 1/√2 → 1/2 → 0", XL + 6.20, Y0, 5.85, 1.40, 22)
result_card(slide, "0° → 90°:  sin ↑     cos ↓     tan ↑     (tan 90° not defined)", XL, Y0 + LH + 0.28 + NG, CW, 0.36, 18)

# ── 14 IDENTITY — ASK ─────────────────────────────────────────────────────────
slide = new_slide("Question", "Identity", GREEN)
add_notes(slide, "Wait. Same angle. Answer is 1. Do not compute 40°.")
qtx(slide, "Find the value of  sin² 40°  +  cos² 40°.", XL, Y0, CW, 0.50, 18)

# ── 14 IDENTITY — SOLVE ───────────────────────────────────────────────────────
slide = new_slide("Solution", "sin² 40° + cos² 40°", GREEN)
add_notes(slide, "Same angle. Identity. Do not compute 40°.")
body(slide, "True for every θ. Here θ = 40°.", XL, Y0, 7.0)
fml(slide, "sin²θ  +  cos²θ  =  1", XL, Y0 + LH, 7.0, LH, 24)
ncert(slide, "sin² 40° + cos² 40°  =  1", XL + 7.15, Y0, 4.90, 1.20, 20)
body(slide, "Do not compute 40°. Same angle, so the identity is enough.", XL, Y0 + 2 * LH + NG, CW)

# ── 15 RADIAN — ARC ───────────────────────────────────────────────────────────
slide = new_slide("Angle and arc", "Arc length", TEAL)
add_notes(slide, "Arc = Rθ with θ in radian. Never put 90° into this formula.")
formula_card(slide, "Arc length  =  R θ", "θ must be in radian. Never put 90° into this formula.", XL, Y0, 7.0, 1.25, 26)
ncert(slide, "Arc  =  R θ", XL + 7.15, Y0, 4.90, 1.20, 24)
body(slide, "Practical unit of angle: degree.     SI unit of angle: radian.", XL, Y0 + LH + 0.28 + NG, CW)
body(slide, "+θ anti-clockwise     −θ clockwise", XL, Y0 + 2 * LH + 0.28 + NG, CW)

# ── 15 RADIAN — 180 = π ───────────────────────────────────────────────────────
slide = new_slide("Angle and arc", "The one line:  180° = π rad", TEAL)
add_notes(slide, "Wherever you see π rad, put 180°.")
body(slide, "Lock this. Every conversion starts here.", XL, Y0, 7.0)
fml(slide, "180°  =  π rad", XL, Y0 + LH, 7.0, LH, 28)
ncert(slide, "Wherever you see π rad, put 180°.", XL + 7.15, Y0, 4.90, 1.20, 18)
body(slide, "Next: degree → radian, then radian → degree.", XL, Y0 + 2 * LH + NG, 7.0)

# ── 15 RADIAN — DEG TO RAD ────────────────────────────────────────────────────
slide = new_slide("Angle and arc", "Degree → radian", TEAL)
add_notes(slide, "1° = π/180 rad.")
y = formula_card(slide, "1°  =  π / 180  rad", "Divide the lock-line by 180.", XL, Y0, CW, 1.30, 26)
tx(slide, "So  θ°  in radian  =  θ × (π / 180).", XL, y + NG, CW, LH, 18, TEXT, spa=0)
result_card(slide, "Multiply the degree by π/180.", XL, y + NG + LH, CW, 0.36, 18)

# ── 15 RADIAN — RAD TO DEG ────────────────────────────────────────────────────
slide = new_slide("Angle and arc", "Radian → degree", TEAL)
add_notes(slide, "1 rad = 180/π.")
y = formula_card(slide, "1 rad  =  180° / π", "Divide the lock-line by π.", XL, Y0, CW, 1.30, 26)
tx(slide, "So  θ rad  in degree  =  θ × (180 / π).", XL, y + NG, CW, LH, 18, TEXT, spa=0)
result_card(slide, "Multiply the radian by 180/π.", XL, y + NG + LH, CW, 0.36, 18)

# ── 16 CIRCLE 90 — ASK ────────────────────────────────────────────────────────
slide = new_slide("Question", "Arc on a circle", GREEN)
add_notes(slide, "Wait. Distance = πR/2. Do not hint the degree trap.")
qtx(slide, "An object moves on the circumference of a circle of radius R. If the object turns through 90°, find the total distance moved by the object.", XL, Y0, CW, 0.70, 16)

# ── 16 CIRCLE 90 — SOLVE ──────────────────────────────────────────────────────
slide = new_slide("Solution", "Object turns 90° on a circle of radius R", ORANGE)
add_notes(slide, "Cross out R×90°. Convert 90° = π/2. Distance = πR/2.")
tx(slide, "An object moves on the circumference of a circle of radius R and turns through 90°. Find the distance.", XL, Y0, CW, LH, 14, TEXT, spa=0, scaled=False)
y = warn_row(slide, "Distance = R × 90°", "Distance = R × (π/2 rad)  =  πR/2", XL, Y0 + LH + NG, CW)
y = formula_card(slide, "distance = path = arc = Rθ", "Convert 90° first:  90° = π/2 rad", XL, y + NG, CW, 0.85, 18)
result_card(slide, "πR / 2", XL, y + NG, CW, 0.36)

# ── 16 π/3 COPY EXAMPLE (PW p29) ──────────────────────────────────────────────
slide = new_slide("Question", "First conversion", GREEN)
add_notes(slide, "Wait. π/3 rad = 60°.")
qtx(slide, "Convert π/3 radian into degree.", XL, Y0, CW, 0.45, 18)

slide = new_slide("Solution", "π / 3 rad", GREEN)
add_notes(slide, "Put π rad = 180°. 180/3 = 60°.")
body(slide, "Wherever you see π rad, put 180°.", XL, Y0, CW)
fml(slide, "π / 3  rad  =  180° / 3  =  60°", XL, Y0 + LH + NG, CW, LH, 24)
ncert(slide, "π / 3 rad  =  60°", XL + 7.15, Y0 + LH + NG, 4.90, 1.20, 22)

# ── 17 CONVERSIONS — ASK NOW ──────────────────────────────────────────────────
slide = new_slide("Question", "Convert these now", GREEN)
add_notes(slide, "Live board. Do not show answers. Tonight is the next list.")
qtx(slide, "Convert the following into degree or radian, as required.", XL, Y0, CW, 0.36, 16)
live = [
    "(i)     π/2 rad",
    "(ii)    π/4 rad",
    "(iii)   π/6 rad",
    "(iv)    45°",
    "(v)     90°",
    "(vi)    60°",
]
for i, line in enumerate(live):
    r, c = divmod(i, 2)
    x, y = XL + c * 6.1, Y0 + 0.50 + r * LH
    tx(slide, line, x, y, 5.85, LH, 16, TEXT, False, spa=0, scaled=False)

# ── 17 CONVERSIONS — SOLVE NOW ────────────────────────────────────────────────
slide = new_slide("Solution", "The six conversions from the board", GREEN)
add_notes(slide, "Tonight stays unanswered.")
live = [
    "π/2 rad = 90°",
    "π/4 rad = 45°",
    "π/6 rad = 30°",
    "45° = π/4 rad",
    "90° = π/2 rad",
    "60° = π/3 rad",
]
for i, line in enumerate(live):
    r, c = divmod(i, 2)
    x, y = XL + c * 6.1, Y0 + r * LH
    result_card(slide, line, x, y, 5.85, LH, 18)

# ── 17 TONIGHT — EXTERIOR β (PW p30) ──────────────────────────────────────────
slide = new_slide("Tonight", "Find β", AMBER)
add_notes(slide, "Do not reveal. Interior third angle = 70°. Exterior β = 110°.")
qtx(slide, "A triangle has two interior angles 40° and 70°. The side next to the 70° angle is extended. If β is the exterior angle formed at that vertex, find the value of β.", XL, Y0, CW, 0.90, 16)

# ── 17 CONVERSIONS — TONIGHT RAD TO DEG (PW p31) ──────────────────────────────
slide = new_slide("Tonight", "Convert radian into degree", AMBER)
add_notes(slide, "No answers. 5π/6=150°, 2π/3=120°, 4π/3=240°, 3π/5=108°, 5π=900°, 2π/5=72°, 7π/12=105°, 5π/2=450°, 3π=540°.")
qtx(slide, "Convert the following radian measures into degree.", XL, Y0, CW, 0.36, 16)
rad_hw = [
    "(i)  5π/6", "(ii)  2π/3", "(iii)  4π/3", "(iv)  3π/5",
    "(v)  5π", "(vi)  2π/5", "(vii)  7π/12", "(viii)  5π/2",
    "(ix)  3π",
]
for i, line in enumerate(rad_hw):
    r, c = divmod(i, 3)
    tx(slide, line, XL + c * 4.0, Y0 + 0.48 + r * LH, 3.8, LH, 15, TEXT, False, spa=0, scaled=False)

# ── 17 CONVERSIONS — TONIGHT DEG TO RAD (PW p32) ──────────────────────────────
slide = new_slide("Tonight", "Convert degree into radian", AMBER)
add_notes(slide, "No answers. 120°=2π/3, 30°=π/6, 150°=5π/6, 270°=3π/2, 135°=3π/4. π is already radian.")
qtx(slide, "Convert the following into radian.", XL, Y0, CW, 0.36, 16)
deg_hw = [
    "(i)  120°", "(ii)  30°", "(iii)  150°",
    "(iv)  270°", "(v)  135°", "(vi)  π",
]
for i, line in enumerate(deg_hw):
    r, c = divmod(i, 3)
    tx(slide, line, XL + c * 4.0, Y0 + 0.48 + r * LH, 3.8, LH, 15, TEXT, False, spa=0, scaled=False)

# ── 18 TONIGHT — sin 53 cosec 53 (PW p33) ─────────────────────────────────────
slide = new_slide("Tonight", "Product of a ratio and its reciprocal", AMBER)
add_notes(slide, "Do not reveal. sin 53 × cosec 53 = 1.")
qtx(slide, "Find the value of  sin 53°  ×  cosec 53°.", XL, Y0, CW, 0.45, 18)

# ── 18 TONIGHT — sin 180 cos 135 tan 120 (PW p34) ─────────────────────────────
slide = new_slide("Tonight", "Product of three values", AMBER)
add_notes(slide, "Do not reveal. sin180=0 so the product is 0.")
qtx(slide, "Find the value of  sin 180°  ·  cos 135°  ·  tan 120°.", XL, Y0, CW, 0.50, 18)

# ── 18 TONIGHT — POLE SHADOW (PW p35) ─────────────────────────────────────────
slide = new_slide("Tonight", "Shadow of a pole", AMBER)
add_notes(slide, "Do not reveal. tan θ = h/l so l = h / tan θ = h cot θ.")
qtx(slide, "A vertical pole of height h casts a shadow of length l when the Sun is at an angle θ above the horizon. Which of the following expresses the shadow length l in terms of h and θ?", XL, Y0, CW, 0.90, 15)
choices(slide, [
    "l  =  h tan θ",
    "l  =  h / tan θ",
    "l  =  h sin θ",
    "l  =  h / sin θ",
], XL, Y0 + 1.05)

# ── 18 TONIGHT — WIRE PYQ (PW p36) ────────────────────────────────────────────
slide = new_slide("Tonight", "Wire from the top of a pole", AMBER)
add_notes(slide, "Do not reveal. Wire length = h / sinθ. After it falls, x = h cot θ.")
qtx(slide, "A pole of height h has a wire attached from the top to the ground at an angle θ. If the wire breaks and falls to the ground without changing length, find the horizontal distance x from the base of the pole where the wire touches the ground.", XL, Y0, CW, 1.05, 15)
choices(slide, [
    "x  =  h cot θ",
    "x  =  h tan θ",
    "x  =  h sin θ",
    "x  =  h sec θ",
], XL, Y0 + 1.20)

# ── 18 TONIGHT — SLOPING GROUND PYQ (PW p37) ──────────────────────────────────
slide = new_slide("Tonight", "Shadow on sloping ground", AMBER)
add_notes(slide, "Do not reveal. Answer is 10/(√3−1) m. Draw: slope 30°, sun 60° to horizontal, pole 10 m.")
qtx(slide, "A vertical pole of height h = 10 m stands on ground that slopes upwards at a constant angle α = 30° with the horizontal. If the Sun’s angle of elevation above the horizontal is θ = 60°, what is the length of the shadow cast by the pole on the sloping ground?", XL, Y0, CW, 1.10, 15)
choices(slide, [
    "5 m",
    "10 m",
    "10√3 m",
    "10 / (√3 − 1) m",
], XL, Y0 + 1.25)

# ── 18 TONIGHT — DAILY TEST ───────────────────────────────────────────────────
slide = new_slide("Tonight", "After class", AMBER)
add_notes(slide, "Daily test before the next lecture.")
body(slide, "Daily test before the next lecture.", XL, Y0, CW)

# ── 19 SUMMARY ────────────────────────────────────────────────────────────────
take = [
    ("P, B, H", "sin = P/H     cos = B/H     tan = P/B     and the three reciprocals"),
    ("Identities", "sin²θ + cos²θ = 1     1 + tan²θ = sec²θ     1 + cot²θ = cosec²θ"),
    ("Range and table", "−1 ≤ sin, cos ≤ 1     0°–180° table     tan 90° not defined"),
    ("Radian", "Arc = Rθ with θ in radian     π rad → 180°"),
]
for head, line in take:
    slide = new_slide("Today", head, TEAL)
    add_notes(slide, line)
    ncert(slide, line, XL, Y0, CW, 1.20, 18)
slide = new_slide("Today", "After class", TEAL)
add_notes(slide, "Daily test tonight. All remaining conversions and PYQs tonight.")
body(slide, "Daily test tonight.", XL, Y0, CW)
body(slide, "Remaining conversions and PYQs tonight.", XL, Y0 + LH, CW)

prs.save(OUTPUT)
print(f"Created {len(prs.slides)} slides")
print(OUTPUT)
