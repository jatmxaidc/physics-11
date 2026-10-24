"""Weekly Test 02 — Paper discussion deck (question → next-slide solution).

Matches Excellence lecture style (L10 helpers) and class methods from L01–L10.
"""
from __future__ import annotations

import re
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

BASE = Path(__file__).resolve().parent
OUTPUT = BASE / "lecture slides" / "Weekly_Test_02_Paper_Discussion.pptx"

NAVY = RGBColor(8, 30, 63)
BLUE = RGBColor(64, 156, 255)
TEAL = RGBColor(40, 180, 190)
GREEN = RGBColor(90, 214, 128)
AMBER = RGBColor(255, 204, 72)
ORANGE = RGBColor(255, 158, 64)
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
LH = 0.40
NG = 0.14
PT_SCALE = 0.90


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
    for old, new in (
        ("^(n-1)", "ⁿ⁻¹"),
        ("^(-1)", "⁻¹"),
        ("^(-3x)", "⁻³ˣ"),
        ("^(-3)", "⁻³"),
        ("^n", "ⁿ"),
        ("^x", "ˣ"),
        ("^3", "³"),
        ("^2", "²"),
        ("^0", "⁰"),
        ("sqrt", "√"),
        ("omega", "ω"),
        (">=", "≥"),
        ("<=", "≤"),
        ("->", "→"),
    ):
        text = text.replace(old, new)
    text = re.sub(r"\s*=\s*", " = ", text)
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
    tx(slide, text, XL, 0.34, CW, 0.72, 24, GOLD, True, spa=0)


def _sep(slide, y=1.10):
    _solid(slide, XL, y, CW, 0.018, STROKE)


def _footer(slide, n):
    tx(slide, f"{n}", 12.25, YF, 0.8, 0.25, 11, MUTED, True, PP_ALIGN.RIGHT, spa=0)


def add_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


_CONTENT = 0
_ROUGH_DUE = False
_ANSWER_STEPS = {}
_LAST_QUESTION_TITLE = ""


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


def fml(slide, text, x, y, w, h=0.44, size=20):
    if id(slide) in _ANSWER_STEPS:
        _ANSWER_STEPS[id(slide)] += 1
        text = f"Step {_ANSWER_STEPS[id(slide)]}:  {text}"
    return tx(slide, text, x, y, w, h, size, WORK, True, spa=0)


def body(slide, text, x, y, w, h=0.44, size=17):
    return tx(slide, text, x, y, w, h, size, TEXT, False, spa=0, scaled=False)


def qtx(slide, text, x, y, w, h=0.44, size=18):
    return tx(slide, text, x, y, w, h, max(size, 18), QCOL, True, spa=0, scaled=False)


def result_card(slide, text, x, y, w, h=0.36, tsize=18):
    return tx(slide, text, x, y, w, h, max(tsize, 18), GREEN, True, spa=0)


def trap_card(slide, text, x, y, w, h=0.40):
    return tx(slide, text, x, y, w, h, 15, ORANGE, False, spa=0, scaled=False)


def choices(slide, items, x, y, w=None, correct=None):
    w = w or CW
    letters = "ABCD"
    for i, item in enumerate(items):
        cy = y + i * 0.42
        fill = GREEN if correct is not None and i == correct else TEAL
        circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(cy + 0.04), Inches(0.30), Inches(0.30))
        circ.fill.solid()
        circ.fill.fore_color.rgb = fill
        circ.line.fill.background()
        tx(slide, letters[i], x, cy + 0.06, 0.30, 0.26, 14, WHITE, True, PP_ALIGN.CENTER, spa=0, scaled=False)
        tx(slide, item, x + 0.42, cy + 0.06, w - 0.42, 0.28, 18, TEXT, False, spa=0, scaled=False)


def tx_3d(slide, text, x, y, w, h, size, face, edge, font="Arial Black"):
    tx(slide, text, x + 0.045, y + 0.045, w, h, size, edge, True, PP_ALIGN.CENTER, spa=0, scaled=False, font=font)
    tx(slide, text, x, y, w, h, size, face, True, PP_ALIGN.CENTER, spa=0, scaled=False, font=font)


def paper_q(n, topic, prompt, options, marks=2):
    """Question on its own slide; options only (no answer)."""
    slide = question_slide(f"Q{n}.  {topic}")
    qtx(slide, prompt, XL, Y0, CW, 1.35, 19)
    body(slide, f"{marks} marks", XL, Y0 + 1.40, CW, 0.30, 14)
    choices(slide, options, XL, Y0 + 1.85)
    add_notes(slide, f"Weekly Test 02 · Q{n} · {topic}")
    return slide


def paper_a(n, topic, steps, answer, trap=None, formula=None):
    """Easy solution on the next slide, class-method wording."""
    slide = answer_slide(f"Solution: Q{n}.  {topic}")
    y = Y0
    if formula:
        ncert(slide, formula, XL, y, CW, 0.78, 18)
        y += 0.95
    for step in steps:
        fml(slide, step, XL, y, CW, LH, 18)
        y += LH + NG
    result_card(slide, answer, XL, y + 0.08, CW, 0.50, 19)
    if trap:
        trap_card(slide, trap, XL, min(y + 0.70, 6.55), CW, 0.40)
    add_notes(slide, f"Solution Q{n}")
    return slide


# ── COVER ─────────────────────────────────────────────────────────────────────
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
tx(slide, "Paper Discussion", bx + bw / 2, by + 0.14, bw / 2, 0.50, 22, BLACK, True, PP_ALIGN.CENTER, spa=0, scaled=False)
tx(slide, "Basic Maths and Calculus", bx, by + 1.02, bw, 0.62, 30, WHITE, True, PP_ALIGN.CENTER, spa=0, scaled=False, font="Georgia")
tx(slide, "Weekly Test 02  ·  Lectures 01–10", bx, by + 1.62, bw, 0.40, 18, GOLD, True, PP_ALIGN.CENTER, spa=0, scaled=False)
tx(slide, "By - Excellence Programme", XL, 5.28, CW, 0.50, 25, WHITE, True, PP_ALIGN.CENTER, spa=0, scaled=False)
tx(slide, "20 September 2026", XL, 5.80, CW, 0.32, 18, MUTED, True, PP_ALIGN.CENTER, spa=0, scaled=False)
_footer(slide, 1)
add_notes(slide, "Paper discussion: each question, then its easy class-method solution.")
after_content_slide()

# Agenda
slide = prs.slides.add_slide(BLANK)
_bg(slide)
tx(slide, "How we will discuss the paper", XL, 0.35, CW, 0.64, 30, GOLD, True, spa=0, scaled=False)
_solid(slide, XL, 1.00, 4.20, 0.035, GOLD)
topics = [
    "30 MCQs from Weekly Test 02 (Lectures 01–10)",
    "Question on one slide  →  solution on the next",
    "Same methods we used in class (ASTC, √(a²+b²), Tn = Sn − Sₙ₋₁, …)",
    "Keep steps short and easy — board style",
    "Rough-work slides appear every few questions",
]
for i, item in enumerate(topics):
    yy = 1.25 + i * 0.85
    tx(slide, f"{i + 1:02d}", XL, yy, 0.70, 0.50, 23, GOLD, True, spa=0, scaled=False)
    tx(slide, item, XL + 0.90, yy, CW - 0.90, 0.70, 18, WHITE, False, spa=0, scaled=False)
_footer(slide, 2)
add_notes(slide, "Remind: methods match our lecture slides, not new tricks.")
after_content_slide()

slide = new_slide("Method reminder", "Lock lines before we start", TEAL)
ncert(slide, "ASTC: sign first  ·  a sinθ + b cosθ → √(a²+b²)  ·  arc = rθ (radian)", XL, Y0, CW, 0.90, 17)
body(slide, "Triangle ratios from the right triangle. Identities: (sec − tan)(sec + tan) = 1.", XL, Y0 + 1.15, CW, 0.55)
body(slide, "AP: Tn = Sn − Sₙ₋₁.  GP rebound: first fall + 2 × infinite series of ups.", XL, Y0 + 1.85, CW, 0.55)
body(slide, "Circle: centre (−g/2, −f/2).  Ellipse foci: 2ae.  Diff: product / quotient / chain.", XL, Y0 + 2.55, CW, 0.55)
trap_card(slide, "Do not invent a new method on the paper. Use the same board method from class.", XL, Y0 + 3.40, CW, 0.55)

# ── Q1–Q30 ────────────────────────────────────────────────────────────────────

paper_q(1, "L01 ratios",
        "In △ABC, ∠B = 90°. If tan A = 5/12, then sin A + cos A equals",
        ["17/13", "17/12", "7/13", "12/5"])
paper_a(1, "L01 ratios",
        ["tan A = opp/adj = 5/12  →  draw the right triangle",
         "hyp = √(5² + 12²) = √(25 + 144) = √169 = 13",
         "sin A = 5/13,   cos A = 12/13",
         "sin A + cos A = 5/13 + 12/13 = 17/13"],
        "Answer: A   17/13",
        trap="Do not forget to find the hypotenuse with Pythagoras first.")

paper_q(2, "L01 identities",
        "If sec θ − tan θ = 1/2, then sec θ + tan θ equals",
        ["1/2", "1", "2", "4"])
paper_a(2, "L01 identities",
        ["Use the identity (sec θ − tan θ)(sec θ + tan θ) = sec²θ − tan²θ = 1",
         "So (sec θ + tan θ) = 1 / (sec θ − tan θ)",
         "sec θ + tan θ = 1 / (1/2) = 2"],
        "Answer: C   2",
        formula="(sec θ − tan θ)(sec θ + tan θ) = 1")

paper_q(3, "L01 identities",
        "If sin θ + cos θ = √2, then tan θ + cot θ equals",
        ["1", "2", "√2", "0"])
paper_a(3, "L01 identities",
        ["Square both sides: (sin θ + cos θ)² = 2",
         "1 + 2 sin θ cos θ = 2  →  sin 2θ = 1  →  2θ = 90° (acute)  →  θ = 45° = π/4",
         "tan(π/4) = 1,   cot(π/4) = 1",
         "tan θ + cot θ = 1 + 1 = 2"],
        "Answer: B   2",
        trap="After squaring, take the acute angle that fits 0 < θ < π/2.")

paper_q(4, "L02 ASTC",
        "The value of sin 510° · cos 330° − cos 390° · sin 150° is",
        ["0", "1/2", "√3/2", "−1/2"])
paper_a(4, "L02 ASTC",
        ["Reduce each angle: sin 510° = sin(510 − 360) = sin 150° = 1/2",
         "cos 330° = + cos 30° = √3/2   (ASTC: 4th quadrant, cos +)",
         "cos 390° = cos 30° = √3/2",
         "Value = (1/2)(√3/2) − (√3/2)(1/2) = 0"],
        "Answer: A   0",
        trap="Sign from ASTC first, then the acute brother angle.")

paper_q(5, "L02 ASTC",
        "The value of cos(−1410°) is",
        ["1/2", "−1/2", "√3/2", "−√3/2"])
paper_a(5, "L02 ASTC",
        ["cos is even: cos(−1410°) = cos 1410°",
         "1410° = 3 × 360° + 330°  →  cos 1410° = cos 330°",
         "cos 330° = + cos 30° = √3/2   (ASTC 4th quadrant)"],
        "Answer: C   √3/2",
        formula="cos(−θ) = cos θ")

paper_q(6, "L02 radian",
        "An arc of a circle of radius 21 cm subtends 60° at the centre. The length of the arc is",
        ["7π cm", "11π cm", "14π cm", "22π cm"])
paper_a(6, "L02 radian",
        ["Convert to radian: 60° = 60 × π/180 = π/3",
         "Arc length ℓ = rθ = 21 × (π/3) = 7π cm"],
        "Answer: A   7π cm",
        formula="ℓ = rθ   (θ in radian)",
        trap="Never put degrees straight into ℓ = rθ.")

paper_q(7, "L02 height",
        "The angle of elevation of a tower from a point A is 30°. From a point B, 20 m closer to the tower on the same horizontal line, the elevation is 60°. The height of the tower is",
        ["10√3 m", "20√3 m", "10 m", "20 m"])
paper_a(7, "L02 height",
        ["Let height = h, and distance from A to base = x",
         "From A: tan 30° = h/x  →  1/√3 = h/x  →  x = h√3",
         "From B (x − 20): tan 60° = h/(x − 20)  →  √3 = h/(x − 20)",
         "h√3 − 20 = h/√3   →   multiply by √3:  3h − 20√3 = h  →  2h = 20√3  →  h = 10√3"],
        "Answer: A   10√3 m")

paper_q(8, "L03 compound",
        "If sin(A + B) = 1 and cos(A − B) = √3/2 with 0 < A, B < π/2, then A equals",
        ["π/3", "π/4", "π/6", "π/2"])
paper_a(8, "L03 compound",
        ["sin(A + B) = 1  →  A + B = π/2",
         "cos(A − B) = √3/2  →  A − B = π/6   (acute)",
         "Add the two equations: 2A = π/2 + π/6 = 2π/3",
         "A = π/3"],
        "Answer: A   π/3")

paper_q(9, "L03 double",
        "If tan θ = 1/2, then (2 tan θ)/(1 − tan² θ) equals",
        ["4/3", "3/4", "1", "4"])
paper_a(9, "L03 double",
        ["Recognise the double-angle formula: tan 2θ = (2 tan θ)/(1 − tan² θ)",
         "Put tan θ = 1/2:  numerator = 2 × (1/2) = 1",
         "denominator = 1 − (1/2)² = 1 − 1/4 = 3/4",
         "Value = 1 / (3/4) = 4/3"],
        "Answer: A   4/3",
        formula="tan 2θ = (2 tan θ)/(1 − tan² θ)")

paper_q(10, "L03 small angle",
        "A circular disc of radius 10 cm subtends 0.04 rad at the eye. Using the small-angle approximation, the diameter of the disc as seen is about",
        ["0.2 cm", "0.4 cm", "2 cm", "4 cm"])
paper_a(10, "L03 small angle",
        ["Small-angle: θ ≈ (linear size) / distance",
         "Here distance = radius to eye = 10 cm,  θ = 0.04 rad",
         "diameter ≈ θ × distance = 0.04 × 10 = 0.4 cm"],
        "Answer: B   0.4 cm",
        formula="θ (rad) ≈ size / distance",
        trap="Use diameter (full size), not radius of the disc, when the question asks for diameter.")

paper_q(11, "L04 max-min",
        "The maximum value of 5 sin θ + 12 cos θ − 7 is",
        ["6", "13", "20", "−7"])
paper_a(11, "L04 max-min",
        ["For a sin θ + b cos θ,  ymax = √(a² + b²)",
         "√(5² + 12²) = √(25 + 144) = √169 = 13",
         "So max of (5 sin θ + 12 cos θ − 7) = 13 − 7 = 6"],
        "Answer: A   6",
        formula="ymax of a sinθ + b cosθ = √(a² + b²)",
        trap="Do not write ymax = 5 + 12. Sin and cos cannot both be 1 at once.")

paper_q(12, "L04 max-min",
        "The range of f(θ) = 2 + 3 sin θ + 4 cos θ is",
        ["[−3, 7]", "[−5, 5]", "[2, 7]", "[−1, 5]"])
paper_a(12, "L04 max-min",
        ["3 sin θ + 4 cos θ lies between −√(9+16) and +√(9+16), i.e. [−5, 5]",
         "Add the constant 2: range of f = [2 − 5, 2 + 5] = [−3, 7]"],
        "Answer: A   [−3, 7]",
        formula="a sinθ + b cosθ ∈ [−√(a²+b²), +√(a²+b²)]")

paper_q(13, "L04 phase",
        "A particle has coordinates x = 3 cos(ωt) and y = 4 sin(ωt). Its maximum distance from the origin is",
        ["3", "4", "5", "7"])
paper_a(13, "L04 phase",
        ["Eliminate t: x²/9 + y²/16 = 1  (ellipse, axes-aligned)",
         "Distance from origin is largest along the major axis",
         "Semi-major axis = 4  (along y)",
         "Maximum distance = 4"],
        "Answer: B   4",
        trap="Do not write √(3²+4²)=5 — that corner is outside the ellipse.")

paper_q(14, "L05 phase",
        "y₁ = 4 sin(ωt) and y₂ = 3 sin(ωt + π/2). The amplitude of y₁ + y₂ is",
        ["5", "7", "1", "√7"])
paper_a(14, "L05 phase",
        ["Phase difference φ = π/2",
         "Resultant amplitude R = √(A₁² + A₂² + 2 A₁ A₂ cos φ)",
         "cos(π/2) = 0  →  R = √(4² + 3²) = √(16 + 9) = 5",
         "Same as perpendicular phasors (our board method)"],
        "Answer: A   5",
        formula="φ = 90°  →  R = √(A₁² + A₂²)",
        trap="Do not add 4 + 3 = 7 unless the waves are exactly in phase.")

paper_q(15, "L05 AP",
        "If the pth, qth and rth terms of an A.P. are a, b and c respectively, then a(q − r) + b(r − p) + c(p − q) equals",
        ["0", "1", "a + b + c", "pqr"])
paper_a(15, "L05 AP",
        ["Write a = A + (p − 1)D,   b = A + (q − 1)D,   c = A + (r − 1)D",
         "Put into a(q − r) + b(r − p) + c(p − q)",
         "The A terms cancel: A[(q−r)+(r−p)+(p−q)] = 0",
         "The D terms also cancel identically  →  whole expression = 0"],
        "Answer: A   0")

paper_q(16, "L06 AP",
        "The sum of n terms of an A.P. is 3n² + 5n. Its 10th term is",
        ["32", "59", "62", "65"])
paper_a(16, "L06 AP",
        ["Class formula: Tn = Sn − Sₙ₋₁",
         "S₁₀ = 3(100) + 5(10) = 300 + 50 = 350",
         "S₉ = 3(81) + 5(9) = 243 + 45 = 288",
         "T₁₀ = 350 − 288 = 62"],
        "Answer: C   62",
        formula="Tn = Sn − Sₙ₋₁")

paper_q(17, "L06 GP",
        "In a G.P., the 3rd term is 24 and the 6th term is 192. The first term is",
        ["3", "6", "8", "12"])
paper_a(17, "L06 GP",
        ["ar² = 24,   ar⁵ = 192",
         "Divide: (ar⁵)/(ar²) = r³ = 192/24 = 8  →  r = 2",
         "ar² = a · 4 = 24  →  a = 6"],
        "Answer: B   6")

paper_q(18, "L07 GP",
        "A ball rebounds to 4/5 of the height from which it falls. If it falls from 10 m, the total distance travelled before coming to rest is",
        ["50 m", "90 m", "100 m", "18 m"])
paper_a(18, "L07 GP",
        ["First fall = 10 m (only once)",
         "After that: up + down form an infinite GP with first up = 10 × (4/5) = 8",
         "Sum of all ups and downs after first fall = 2 × [8 / (1 − 4/5)] = 2 × 8 / (1/5) = 80",
         "Total distance = 10 + 80 = 90 m"],
        "Answer: B   90 m",
        formula="Total = h + 2 · (h r)/(1 − r)",
        trap="Do not forget the first fall is counted only once.")

paper_q(19, "L06 straight line",
        "The lines 2x − 3y + 6 = 0 and 4x − 6y − 5 = 0 are",
        ["coincident", "intersecting at one point", "parallel and distinct", "perpendicular"])
paper_a(19, "L06 straight line",
        ["Slope of first: 2x − 3y + 6 = 0  →  y = (2/3)x + 2  →  m₁ = 2/3",
         "Slope of second: 4x − 6y − 5 = 0  →  y = (2/3)x − 5/6  →  m₂ = 2/3",
         "Same slope, but second is not a scalar multiple of the first (constants 6 vs −5)",
         "Therefore parallel and distinct"],
        "Answer: C   parallel and distinct",
        trap="Same slope alone is not enough for coincident — check the constant term too.")

paper_q(20, "L07 straight line",
        "The equation of the line through (2, −3) and parallel to 3x + 4y − 7 = 0 is",
        ["3x + 4y + 6 = 0", "3x + 4y − 6 = 0", "4x − 3y − 17 = 0", "3x − 4y + 6 = 0"])
paper_a(20, "L07 straight line",
        ["Parallel line keeps the same x and y coefficients: 3x + 4y + k = 0",
         "It passes through (2, −3): 3(2) + 4(−3) + k = 0",
         "6 − 12 + k = 0  →  k = 6",
         "Equation: 3x + 4y + 6 = 0"],
        "Answer: A   3x + 4y + 6 = 0")

paper_q(21, "L08 graphs",
        "A particle starts from rest with constant acceleration and covers 8 m in the first 2 s. Distance covered in the next 2 s is",
        ["8 m", "16 m", "24 m", "32 m"])
paper_a(21, "L08 graphs",
        ["u = 0,   s = ut + ½ a t²  →  8 = ½ a (2)²  →  8 = 2a  →  a = 4 m/s²",
         "Distance in first 4 s: s₄ = ½(4)(4)² = 2 × 16 = 32 m",
         "Next 2 s = s₄ − s₂ = 32 − 8 = 24 m"],
        "Answer: C   24 m",
        formula="s = ut + ½ a t²")

paper_q(22, "L08 graphs",
        "On a straight-line v–t graph, velocity falls from 10 m/s to 0 in 5 s. Displacement in this interval is",
        ["10 m", "25 m", "50 m", "0 m"])
paper_a(22, "L08 graphs",
        ["On a v–t graph, displacement = area under the graph",
         "Here the graph is a straight triangle: base 5 s, height 10 m/s",
         "Area = ½ × 5 × 10 = 25 m"],
        "Answer: B   25 m",
        formula="displacement = area under v–t graph")

paper_q(23, "L08 graphs",
        "If a = −2 m/s² (constant) and initial velocity is +10 m/s, the particle comes to rest after",
        ["2 s", "5 s", "10 s", "20 s"])
paper_a(23, "L08 graphs",
        ["At rest: v = 0",
         "Use v = u + a t",
         "0 = 10 + (−2)t  →  2t = 10  →  t = 5 s"],
        "Answer: B   5 s",
        formula="v = u + a t")

paper_q(24, "L09 circle",
        "For the circle x² + y² − 4x + 6y − 3 = 0, the centre and radius are",
        ["(2, −3), 4", "(−2, 3), 4", "(2, −3), 2", "(2, 3), 4"])
paper_a(24, "L09 circle",
        ["General form: x² + y² + 2gx + 2fy + c = 0",
         "Here 2g = −4 → g = −2;   2f = 6 → f = 3;   c = −3",
         "Centre = (−g, −f) = (2, −3)",
         "Radius = √(g² + f² − c) = √(4 + 9 − (−3)) = √16 = 4"],
        "Answer: A   (2, −3), 4",
        formula="centre (−g, −f),   r = √(g² + f² − c)")

paper_q(25, "L09 ellipse",
        "For the ellipse x²/16 + y²/9 = 1, the distance between the foci is",
        ["√7", "2√7", "5", "√5"])
paper_a(25, "L09 ellipse",
        ["a² = 16 → a = 4;   b² = 9 → b = 3  (a > b, major along x)",
         "e = √(1 − b²/a²) = √(1 − 9/16) = √(7/16) = √7 / 4",
         "Distance between foci = 2ae = 2 · 4 · (√7/4) = 2√7"],
        "Answer: B   2√7",
        formula="foci distance = 2ae")

paper_q(26, "L09 exponential",
        "If y = e^{3x} + e^{−3x}, then d²y/dx² − 9y equals",
        ["0", "18e^{3x}", "−18e^{−3x}", "9"])
paper_a(26, "L09 exponential",
        ["dy/dx = 3 e^{3x} − 3 e^{−3x}",
         "d²y/dx² = 9 e^{3x} + 9 e^{−3x} = 9(e^{3x} + e^{−3x}) = 9y",
         "Therefore d²y/dx² − 9y = 0"],
        "Answer: A   0")

paper_q(27, "L09 differentiation",
        "If y = √(2x + 3), then dy/dx at x = 3 equals",
        ["1/√3", "1/√9", "1/3", "√3"])
paper_a(27, "L09 differentiation",
        ["Write y = (2x + 3)^{1/2}",
         "Chain rule: dy/dx = (1/2)(2x + 3)^{−1/2} · 2 = 1 / √(2x + 3)",
         "At x = 3:  1 / √(6 + 3) = 1 / √9 = 1/3"],
        "Answer: C   1/3",
        formula="d/dx [√u] = (1/(2√u)) · du/dx")

paper_q(28, "L10 rate of change",
        "The volume of a sphere increases at 4π cm³/s when the radius is 2 cm. The rate of increase of the radius is",
        ["1/4 cm/s", "1/2 cm/s", "1 cm/s", "2 cm/s"])
paper_a(28, "L10 rate of change",
        ["V = (4/3)π r³  →  dV/dt = 4π r² (dr/dt)",
         "Given dV/dt = 4π when r = 2",
         "4π = 4π (2)² (dr/dt) = 4π · 4 · (dr/dt)",
         "1 = 4 (dr/dt)  →  dr/dt = 1/4 cm/s"],
        "Answer: A   1/4 cm/s",
        formula="dV/dt = 4π r² dr/dt")

paper_q(29, "L10 differentiation",
        "If y = (x² + 1)/(x² − 1) with |x| ≠ 1, then dy/dx equals",
        ["−4x/(x² − 1)²", "4x/(x² − 1)²", "−2x/(x² − 1)²", "2x/(x² − 1)"])
paper_a(29, "L10 differentiation",
        ["Quotient rule: (u/v)' = (v u' − u v') / v²",
         "u = x² + 1, u' = 2x;   v = x² − 1, v' = 2x",
         "Numerator = (x² − 1)(2x) − (x² + 1)(2x) = 2x[(x² − 1) − (x² + 1)] = 2x(−2) = −4x",
         "dy/dx = −4x / (x² − 1)²"],
        "Answer: A   −4x/(x² − 1)²",
        formula="(u/v)' = (v u' − u v') / v²",
        trap="Keep the order v u' − u v' — swapping the sign is a common trap.")

paper_q(30, "L10 product / chain",
        "If y = sin(x²), then dy/dx equals",
        ["cos(x²)", "2x cos(x²)", "2x sin(x²)", "cos(2x)"])
paper_a(30, "L10 product / chain",
        ["Outer function sin(·), inner function x²",
         "Chain rule: dy/dx = cos(x²) · d(x²)/dx",
         "d(x²)/dx = 2x",
         "dy/dx = 2x cos(x²)"],
        "Answer: B   2x cos(x²)",
        formula="d/dx [sin u] = cos u · du/dx",
        trap="Do not forget to multiply by the derivative of the inside (2x).")

# Closing
slide = new_slide("Class summary", "Paper discussion lock lines", TEAL)
fml(slide, "Every question: write the class formula first, then substitute", XL, Y0, CW, LH, 17)
fml(slide, "ASTC / √(a²+b²) / Tn=Sn−Sₙ₋₁ / ℓ=rθ / quotient–chain — same as boards", XL, Y0 + LH + NG, CW, LH, 16)
fml(slide, "Read options after you finish; do not reverse-engineer from choices", XL, Y0 + 2 * LH + 2 * NG, CW, LH, 17)
body(slide, "Re-attempt any wrong question tomorrow without looking at the solution slide.", XL, Y0 + 3 * LH + 3 * NG, CW, 0.55)

slide = new_slide("After class", "Practice", AMBER)
body(slide, "1. Re-do Q1–Q10 from memory (ratios, ASTC, compound, small angle).", XL, Y0, CW, 0.52)
body(slide, "2. Re-do Q11–Q20 (max–min, phase, AP/GP, lines).", XL, Y0 + 0.72, CW, 0.52)
body(slide, "3. Re-do Q21–Q30 (graphs, circle, ellipse, differentiation).", XL, Y0 + 1.44, CW, 0.52)
body(slide, "4. Mark every trap in orange in your notebook.", XL, Y0 + 2.16, CW, 0.52)
trap_card(slide, "If a method feels new, go back to that lecture slide — we stay with class methods only.", XL, Y0 + 3.00, CW, 0.62)

OUTPUT.parent.mkdir(parents=True, exist_ok=True)
prs.save(OUTPUT)
print(f"Created {len(prs.slides)} slides")
print(OUTPUT)
