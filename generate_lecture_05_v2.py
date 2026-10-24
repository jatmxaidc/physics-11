"""
Lecture 05 v2 — Frame of Reference, Position, Distance, Displacement
Class 11 Physics  •  CBSE + NEET Foundation  •  2-hour session  •  18 slides
"""
from __future__ import annotations
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

BASE   = Path(__file__).resolve().parent
OUTPUT = BASE / "Lecture_05_v2_Frame_Reference_Distance_Displacement.pptx"

# ── Palette ───────────────────────────────────────────────────────────────────
NAVY       = RGBColor(8,   30,  63)
BLUE       = RGBColor(21,  95, 170)
BLUE_BG    = RGBColor(231, 243, 255)
TEAL       = RGBColor(0,  128, 138)
TEAL_BG    = RGBColor(220, 248, 252)
GREEN      = RGBColor(25, 108,  48)
GREEN_BG   = RGBColor(230, 248, 238)
AMBER      = RGBColor(160,  92,   0)
AMBER_BG   = RGBColor(255, 251, 222)
ORANGE     = RGBColor(200,  70,   0)
ORANGE_BG  = RGBColor(255, 242, 225)
RED        = RGBColor(172,  22,  22)
RED_BG     = RGBColor(255, 236, 236)
TEXT       = RGBColor(20,  26,  38)
MUTED      = RGBColor(90, 102, 120)
STROKE     = RGBColor(208, 218, 232)
WHITE      = RGBColor(255, 255, 255)
SOFT       = RGBColor(247, 250, 253)
DARK_PANEL = RGBColor(8,   30,  63)

# ── Canvas ────────────────────────────────────────────────────────────────────
prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

# Layout constants
XL = 0.65          # left margin
CW = 12.05         # content width
Y0 = 1.22          # content start
YN = 5.52          # notes zone starts
YF = 7.12          # footer y


# ═══════════════════════════════════════════════════════════════════════════════
# LOW-LEVEL PRIMITIVES
# ═══════════════════════════════════════════════════════════════════════════════

def _bg(slide, color=WHITE):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color


def _solid(slide, x, y, w, h, fill, no_line=True):
    s = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if no_line:
        s.line.fill.background()
    else:
        s.line.color.rgb = fill
    return s


def _card_bg(slide, x, y, w, h, fill, border, line_pt=1.0, rounded=True):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    s = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    s.line.color.rgb = border; s.line.width = Pt(line_pt)
    return s


def _accent(slide, x, y, h, color, w=0.065):
    s = _solid(slide, x, y, w, h, color)
    return s


def tx(slide, text, x, y, w, h,
       size=15, color=TEXT, bold=False,
       align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP,
       spa=1, lsp=1.1, font="Aptos"):
    box = slide.shapes.add_textbox(
        Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear(); tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(0)
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]
    p.alignment = align
    p.space_before = Pt(0); p.space_after = Pt(spa); p.line_spacing = lsp
    r = p.add_run()
    r.text = text
    r.font.name = font; r.font.size = Pt(size)
    r.font.bold = bold; r.font.color.rgb = color
    return box


def tx_two_runs(slide, text1, color1, bold1, text2, color2, bold2,
                x, y, w, h, size=15, align=PP_ALIGN.LEFT):
    """One paragraph with two differently-styled runs."""
    box = slide.shapes.add_textbox(
        Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear(); tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(0)
    p = tf.paragraphs[0]
    p.alignment = align
    p.space_before = Pt(0); p.space_after = Pt(0); p.line_spacing = 1.1
    for text, color, bold in ((text1, color1, bold1), (text2, color2, bold2)):
        r = p.add_run()
        r.text = text
        r.font.name = "Aptos"; r.font.size = Pt(size)
        r.font.bold = bold; r.font.color.rgb = color
    return box


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE LAYOUT COMPONENTS
# ═══════════════════════════════════════════════════════════════════════════════

def _bar(slide, label, color=BLUE):
    _solid(slide, 0, 0, 13.333, 0.29, color)
    tx(slide, label.upper(), XL, 0.05, 9, 0.20, 8, WHITE, True, spa=0)


def _title(slide, text):
    tx(slide, text, XL, 0.34, CW, 0.72, 27, NAVY, True, spa=0)


def _sep(slide, y=1.10):
    _solid(slide, XL, y, CW, 0.018, STROKE)


def _footer(slide, n):
    tx(slide, "LECTURE 05  •  CH. 2  MOTION IN A STRAIGHT LINE  •  CBSE + NEET  CLASS 11",
       XL, YF, 10, 0.25, 7.5, MUTED, spa=0)
    tx(slide, f"{n:02d}", 12.25, YF, 0.8, 0.25, 9, MUTED, True, PP_ALIGN.RIGHT, spa=0)


def _notes_zone(slide):
    _solid(slide, XL, YN, CW, 0.012, STROKE)
    tx(slide, "MY NOTES", XL, YN + 0.05, 1.6, 0.20, 7.5,
       RGBColor(205, 215, 228), True, spa=0)


def new_slide(section, title, bar_color=BLUE):
    slide = prs.slides.add_slide(BLANK)
    _bg(slide)
    _bar(slide, section, bar_color)
    _title(slide, title)
    _sep(slide)
    _footer(slide, len(prs.slides))
    _notes_zone(slide)
    return slide


def add_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


# ═══════════════════════════════════════════════════════════════════════════════
# CARD COMPONENTS
# ═══════════════════════════════════════════════════════════════════════════════

def concept_card(slide, label, headline, body, x, y, w, h,
                 fill=BLUE_BG, border=BLUE):
    _card_bg(slide, x, y, w, h, fill, border, 0.9)
    _accent(slide, x, y, h, border)
    cy = y + 0.10
    if label:
        tx(slide, label.upper(), x + 0.14, cy, w - 0.22, 0.20, 9, border, True, spa=1)
        cy += 0.21
    if headline:
        tx(slide, headline, x + 0.14, cy, w - 0.22, 0.30, 17, NAVY, True, spa=2)
        cy += 0.32
    if body:
        tx(slide, body, x + 0.14, cy, w - 0.22, h - (cy - y) - 0.06, 13, MUTED, spa=0)


def formula_card(slide, formula, note, x, y, w, h, fsize=18):
    _card_bg(slide, x, y, w, h, AMBER_BG, AMBER, 1.1)
    _accent(slide, x, y, h, AMBER)
    tx(slide, formula, x + 0.14, y + 0.10, w - 0.22, 0.36, fsize, NAVY, True, spa=2, lsp=1.05)
    if note:
        tx(slide, note, x + 0.14, y + 0.10 + 0.38, w - 0.22,
           h - 0.52, 12, MUTED, spa=0)


def trap_card(slide, text, x, y, w, h):
    _card_bg(slide, x, y, w, h, ORANGE_BG, ORANGE, 0.9)
    _accent(slide, x, y, h, ORANGE)
    tx(slide, "NEET TRAP", x + 0.14, y + 0.08, w - 0.22, 0.18, 8, ORANGE, True, spa=1)
    tx(slide, text, x + 0.14, y + 0.28, w - 0.22, h - 0.32, 13, TEXT, spa=0)


def result_card(slide, text, x, y, w, h, tsize=14):
    _card_bg(slide, x, y, w, h, GREEN_BG, GREEN, 0.9)
    _accent(slide, x, y, h, GREEN)
    tx(slide, text, x + 0.14, y + 0.08, w - 0.22, h - 0.14, tsize, GREEN, True, spa=0)


def warn_row(slide, wrong, correct, x, y, w):
    """Wrong / Correct pair — compact two-row block."""
    _card_bg(slide, x, y, w, 0.32, RED_BG, RED, 0.7)
    _accent(slide, x, y, 0.32, RED, 0.055)
    tx(slide, f"✗  {wrong}", x + 0.13, y + 0.07, w - 0.2, 0.22, 12, RED, True, spa=0)
    _card_bg(slide, x + 0.28, y + 0.35, w - 0.28, 0.32, GREEN_BG, GREEN, 0.7)
    _accent(slide, x + 0.28, y + 0.35, 0.32, GREEN, 0.055)
    tx(slide, f"✓  {correct}", x + 0.41, y + 0.42, w - 0.48, 0.22, 12, GREEN, True, spa=0)
    return y + 0.80


# ═══════════════════════════════════════════════════════════════════════════════
# DIAGRAM HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def number_line(slide, x, y, w, values, labels=None, colors=None):
    """
    Draw a number line.
    labels  : dict  value → label text (above the line)
    colors  : dict  value → RGBColor for that tick+number
    Returns bottom y of the element.
    """
    if labels is None: labels = {}
    if colors is None: colors = {}
    vmin, vmax = min(values), max(values)
    rng = max(1, vmax - vmin)

    def px(v): return x + (v - vmin) / rng * w

    # Main line
    _solid(slide, x, y, w, 0.030, NAVY)
    # Right arrowhead
    s = slide.shapes.add_shape(
        MSO_SHAPE.RIGHT_ARROW,
        Inches(x + w - 0.14), Inches(y - 0.038),
        Inches(0.18), Inches(0.106))
    s.fill.solid(); s.fill.fore_color.rgb = NAVY; s.line.fill.background()

    for v in values:
        cx = px(v)
        col = colors.get(v, NAVY)
        # Tick
        _solid(slide, cx - 0.010, y - 0.085, 0.020, 0.200, col)
        # Number below
        tx(slide, str(v), cx - 0.175, y + 0.12, 0.35, 0.22,
           11, col, v in colors, PP_ALIGN.CENTER, spa=0)
        # Label above
        if v in labels:
            tx(slide, labels[v], cx - 0.52, y - 0.52, 1.04, 0.28,
               11, col, True, PP_ALIGN.CENTER, spa=0)

    return y + 0.52   # bottom y (below numbers)


def path_arrow(slide, line_x, line_y, line_w, v_start, v_end,
               v_min, v_max, color=BLUE, label=None, above=True):
    """Draw a path arrow segment above or below the number line."""
    rng = max(1, v_max - v_min)
    x1 = line_x + (v_start - v_min) / rng * line_w
    x2 = line_x + (v_end   - v_min) / rng * line_w
    ay = (line_y - 0.68) if above else (line_y + 0.46)
    H  = 0.13

    if x2 >= x1:
        s = slide.shapes.add_shape(
            MSO_SHAPE.RIGHT_ARROW,
            Inches(x1), Inches(ay), Inches(max(0.15, x2 - x1)), Inches(H))
    else:
        s = slide.shapes.add_shape(
            MSO_SHAPE.LEFT_ARROW,
            Inches(x2), Inches(ay), Inches(max(0.15, x1 - x2)), Inches(H))
    s.fill.solid(); s.fill.fore_color.rgb = color
    s.line.fill.background()

    if label:
        mid = (x1 + x2) / 2
        ty_lbl = ay - 0.24 if above else ay + 0.17
        tx(slide, label, mid - 0.52, ty_lbl, 1.04, 0.22, 11, color, True, PP_ALIGN.CENTER, spa=0)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 01 — TITLE
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
_bg(slide)
_solid(slide, 0, 0, 4.35, 7.5, DARK_PANEL)

tx(slide, "LECTURE 05",         0.55, 0.70, 3.3, 0.26, 10, RGBColor(160, 195, 235), True, spa=0)
tx(slide, "Motion in a\nStraight Line", 0.55, 1.10, 3.3, 1.0,  26, WHITE, True, spa=4)
tx(slide, "Chapter 2  —  NCERT Physics Part I", 0.55, 5.75, 3.3, 0.30, 11, RGBColor(155, 185, 220), spa=0)
tx(slide, "Class 11  •  CBSE + NEET  •  2026–27", 0.55, 6.12, 3.3, 0.28, 11, WHITE, spa=0)

tx(slide, "Frame of Reference,\nPosition, Distance\nand Displacement",
   4.75, 1.20, 7.8, 2.0, 30, NAVY, True, spa=4)
tx(slide, "2-hour session", 4.75, 3.55, 3.5, 0.26, 13, MUTED, spa=0)

s = _card_bg(slide, 4.75, 4.05, 8.2, 0.38, BLUE_BG, BLUE, 1.0)
tx(slide, "Observe  →  Define  →  Represent  →  Apply",
   4.95, 4.17, 7.8, 0.22, 14, BLUE, True, PP_ALIGN.CENTER, spa=0)

_footer(slide, 1)
add_notes(slide, "0:00 — Welcome. Ask: 'If I say a ball is at position 5, is that complete?' (No — no unit, no origin, no direction.) Today we build the full vocabulary of kinematics.")


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 02 — SESSION OVERVIEW
# ═══════════════════════════════════════════════════════════════════════════════
slide = new_slide("Session overview", "What you will be able to do by the end", NAVY)
add_notes(slide, "0:02 — Read each outcome briefly. Explain structure: 45 min concepts → 5 min break → 45 min application → exit test.")

outcomes = [
    ("Decide",      "whether an object is at rest or in motion relative to a stated observer"),
    ("Describe",    "the location of any object using an origin, direction and position coordinate"),
    ("Calculate",   "distance and displacement for straight-line and reversal paths"),
    ("Distinguish", "scalar distance from vector displacement and state when they are equal"),
    ("Solve",       "NEET-pattern MCQs and CBSE 3-mark questions on this topic"),
]
for i, (key, rest) in enumerate(outcomes):
    y = Y0 + i * 0.70
    circ = slide.shapes.add_shape(MSO_SHAPE.OVAL,
        Inches(XL), Inches(y + 0.05), Inches(0.40), Inches(0.40))
    circ.fill.solid(); circ.fill.fore_color.rgb = BLUE; circ.line.fill.background()
    tx(slide, str(i + 1), XL + 0.10, y + 0.15, 0.20, 0.20, 10, WHITE, True, PP_ALIGN.CENTER, spa=0)
    tx_two_runs(slide, key + "  ", NAVY, True, rest, TEXT, False,
                XL + 0.56, y + 0.08, CW - 0.66, 0.30, 15)

result_card(slide, "Mastery check: explain why distance can be non-zero when displacement is zero.",
            XL, Y0 + 5 * 0.70 + 0.08, CW, 0.50)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 03 — HOOK
# ═══════════════════════════════════════════════════════════════════════════════
slide = new_slide("The central question", "Which instrument is telling the truth?", TEAL)
add_notes(slide, "0:06 — Show both cards. Ask students: 'GPS says 10 km, odometer reads 18 km — one is wrong. Which?' After class discussion reveal: NEITHER is wrong. GPS measures change in position (displacement). Odometer measures total path (distance). Both are valid measurements of different things. That is today's entire lesson in one example.")

tx(slide, "You drive from home to school by a winding road.",
   XL, Y0, CW, 0.28, 17, TEXT, True, spa=0)

for i, (col, bar_c, lbl, val, sub) in enumerate([
    (BLUE_BG,   BLUE,   "GPS / MAP",   "Reads: 10 km",  "Straight-line change in position from start to finish"),
    (AMBER_BG,  AMBER,  "ODOMETER",    "Reads: 18 km",  "Total length of road actually travelled"),
]):
    cx = XL + i * 6.2
    _card_bg(slide, cx, Y0 + 0.48, 5.85, 2.15, col, bar_c, 1.2)
    _accent(slide, cx, Y0 + 0.48, 2.15, bar_c)
    tx(slide, lbl, cx + 0.14, Y0 + 0.62, 5.5, 0.20, 9, bar_c, True, spa=1)
    tx(slide, val, cx + 0.14, Y0 + 0.88, 5.5, 0.36, 22, NAVY, True, spa=4)
    tx(slide, sub, cx + 0.14, Y0 + 1.32, 5.5, 0.30, 13, MUTED, spa=0)

tx(slide, "Neither is wrong. They measure different things — and naming those two things precisely is today's entire lecture.",
   XL, Y0 + 2.82, CW, 0.32, 15, TEAL, True, spa=0)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 04 — RECALL
# ═══════════════════════════════════════════════════════════════════════════════
slide = new_slide("Warm-up", "Recall — answer before we begin", RGBColor(55, 70, 90))
add_notes(slide, "0:12 — Give 25 seconds per question. Q1: value + unit. Q2: metre. Q3: derived (uses length ÷ time). Q4: opposite to the chosen positive direction. These build the vocabulary for the next 90 minutes.")

questions = [
    "What makes a physical quantity complete — not just a number?",
    "What is the SI base unit of length?",
    "Is speed a base quantity or a derived quantity? Justify in one sentence.",
    "In one-dimensional motion, what does a negative sign on a value indicate?",
]
for i, q in enumerate(questions):
    y = Y0 + i * 0.74
    circ = slide.shapes.add_shape(MSO_SHAPE.OVAL,
        Inches(XL), Inches(y + 0.04), Inches(0.40), Inches(0.40))
    circ.fill.solid(); circ.fill.fore_color.rgb = NAVY; circ.line.fill.background()
    tx(slide, str(i + 1), XL + 0.10, y + 0.14, 0.20, 0.22, 10, WHITE, True, PP_ALIGN.CENTER, spa=0)
    tx(slide, q, XL + 0.57, y + 0.07, CW - 0.67, 0.30, 16, TEXT, spa=0)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 05 — FRAME OF REFERENCE
# ═══════════════════════════════════════════════════════════════════════════════
slide = new_slide("Core concept", "Frame of reference", BLUE)
add_notes(slide, "0:18 — Draw the axis on the board as you speak. Point to each component. Key point: changing the frame changes the coordinates but NOT the actual separation between objects. Demo: use classroom door as origin, right wall as +x.")

concept_card(slide, "definition",
             "A frame of reference is the coordinate system relative to which position and motion are described.",
             "It sets the rules for measuring where and when things happen.",
             XL, Y0, CW, 0.95, BLUE_BG, BLUE)

elements = [
    (BLUE,  "ORIGIN  (O)",             "The reference point — assigned coordinate zero. Measurements start here."),
    (AMBER, "POSITIVE DIRECTION  (+x)", "One direction is chosen as positive. Opposite is negative."),
    (TEAL,  "CLOCK  (t)",              "Records when each observation is made. Position can change with time."),
]
for i, (col, head, body) in enumerate(elements):
    ex = XL + i * 4.05
    ew = 3.82
    _card_bg(slide, ex, Y0 + 1.12, ew, 1.62, SOFT, col, 1.0)
    _solid(slide, ex, Y0 + 1.12, ew, 0.06, col)
    tx(slide, head, ex + 0.14, Y0 + 1.25, ew - 0.22, 0.22, 10, col, True, spa=2)
    tx(slide, body, ex + 0.14, Y0 + 1.54, ew - 0.22, 0.68, 13, TEXT, spa=0)

tx(slide, "Key principle: shifting the origin changes every position coordinate but leaves the separation between any two points unchanged.",
   XL, Y0 + 2.88, CW, 0.30, 13, MUTED, False, spa=0)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 06 — POSITION
# ═══════════════════════════════════════════════════════════════════════════════
slide = new_slide("Definition", "Position", BLUE)
add_notes(slide, "0:26 — Ask: 'What is the position of A?' (−4 m). 'What is the position of B?' (+3 m). 'What does the negative sign tell us?' (It is left of the origin — nothing more about motion.)")

_accent(slide, XL, Y0, 0.56, BLUE, 0.065)
tx(slide, "Position specifies the location of an object relative to the chosen origin.",
   XL + 0.14, Y0 + 0.02, CW - 0.2, 0.30, 17, NAVY, True, spa=1)
tx(slide, "In one-dimensional motion, position is a single signed coordinate  x.",
   XL + 0.14, Y0 + 0.35, CW - 0.2, 0.24, 14, MUTED, spa=0)

NLx, NLy, NLw = XL + 0.2, Y0 + 1.32, CW - 0.4
number_line(slide, NLx, NLy, NLw,
            [-6, -4, -3, -2, -1, 0, 1, 2, 3, 4, 5, 6],
            labels={-4: "A", 3: "B", 0: "Origin O"},
            colors={-4: RED, 3: GREEN})

tx(slide, "A :  x = − 4 m", XL,          NLy + 0.60, 3.2, 0.28, 16, RED,   True, spa=0)
tx(slide, "B :  x = + 3 m", XL + 3.5,    NLy + 0.60, 3.2, 0.28, 16, GREEN, True, spa=0)
tx(slide, "SI unit: metre (m)   ·   Dimension: [L]", XL + 7.8, NLy + 0.60, 4.2, 0.28, 13, MUTED, True, spa=0)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 07 — DISTANCE
# ═══════════════════════════════════════════════════════════════════════════════
slide = new_slide("Definition", "Distance", AMBER)
add_notes(slide, "0:34 — Draw each path segment on the board as you speak. Emphasise: distance records every step you take regardless of direction — like the odometer. It can never decrease during a journey.")

concept_card(slide, "distance",
             "Total length of the actual path followed by an object.",
             "Scalar  ·  always ≥ 0  ·  increases with every step  ·  path-dependent  ·  SI unit: metre",
             XL, Y0, 9.5, 0.96, AMBER_BG, AMBER)

NLx2, NLy2, NLw2 = XL + 0.1, Y0 + 2.10, CW - 0.2
VL2, VH2 = -4, 9
number_line(slide, NLx2, NLy2, NLw2,
            list(range(VL2, VH2 + 1)),
            labels={2: "A", 8: "Turn", -3: "C"},
            colors={2: BLUE, 8: ORANGE, -3: RED})

path_arrow(slide, NLx2, NLy2, NLw2, 2, 8,   VL2, VH2, BLUE,   "6 m",  above=True)
path_arrow(slide, NLx2, NLy2, NLw2, 8, -3,  VL2, VH2, ORANGE, "11 m", above=False)

formula_card(slide, "Distance  =  6  +  11  =  17 m",
             "Add every segment length — direction does not matter",
             XL, NLy2 + 1.45, 7.5, 0.70)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 08 — DISPLACEMENT
# ═══════════════════════════════════════════════════════════════════════════════
slide = new_slide("Definition", "Displacement", BLUE)
add_notes(slide, "0:44 — Use the SAME path as the previous slide. Ask: 'Forget the journey — where did you start and where did you finish?' Draw one arrow from A to C. That arrow is displacement. Δx = xf − xi = (−3) − 2 = −5 m. The negative means the net change is in the negative direction.")

concept_card(slide, "displacement",
             "Change in position from the initial point to the final point.",
             "Vector  ·  can be positive, negative or zero  ·  path-independent  ·  SI unit: metre",
             XL, Y0, 9.5, 0.96, BLUE_BG, BLUE)

formula_card(slide, "Δx  =  x\u1da0inal  −  x\u1da2nitial",
             "Subtract initial position from final — the sign gives the direction of net change",
             XL, Y0 + 1.12, 8.5, 0.85)

NLx3, NLy3, NLw3 = XL + 0.1, Y0 + 2.52, CW - 0.2
VL3, VH3 = -4, 9
number_line(slide, NLx3, NLy3, NLw3,
            list(range(VL3, VH3 + 1)),
            labels={2: "xᵢ = 2", -3: "xf = −3"},
            colors={2: BLUE, -3: RED})

path_arrow(slide, NLx3, NLy3, NLw3, 2, -3, VL3, VH3, RED, "Δx = −5 m", above=False)

tx(slide, "Δx  =  (−3)  −  2  =  −5 m   →   5 m in the negative direction",
   XL, NLy3 + 1.05, CW, 0.28, 15, NAVY, True, spa=0)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 09 — DISTANCE vs DISPLACEMENT
# ═══════════════════════════════════════════════════════════════════════════════
slide = new_slide("Comparison", "Distance vs displacement — side by side", NAVY)
add_notes(slide, "0:53 — Go through each row. For NEET: most-tested rows are 'physical type' and 'possible values'. For CBSE: scalar/vector distinction and the example. Ask students to reproduce the table in their notes.")

headers   = ["Feature",         "Distance",                         "Displacement"]
rows_data = [
    ("Measures",      "Total path length",                "Change in position"),
    ("Type",          "Scalar",                           "Vector"),
    ("Path",          "Depends on path taken",            "Depends only on start and end"),
    ("Possible sign", "Zero or positive only",            "Positive, zero or negative"),
    ("SI unit",       "metre (m)",                        "metre (m)"),
]
col_x = [XL, XL + 3.2, XL + 7.6]
col_w = [3.0, 4.3, 4.6]

for j, (hx, hw, ht) in enumerate(zip(col_x, col_w, headers)):
    _solid(slide, hx, Y0, hw, 0.40, NAVY, False)
    tx(slide, ht, hx + 0.12, Y0 + 0.10, hw - 0.2, 0.22, 12, WHITE, True, PP_ALIGN.CENTER, spa=0)

for i, row in enumerate(rows_data):
    ry = Y0 + 0.40 + i * 0.48
    fill = WHITE if i % 2 == 0 else SOFT
    for j, (val, hx, hw) in enumerate(zip(row, col_x, col_w)):
        s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
            Inches(hx), Inches(ry), Inches(hw), Inches(0.48))
        s.fill.solid(); s.fill.fore_color.rgb = fill
        s.line.color.rgb = STROKE; s.line.width = Pt(0.5)
        bold = (j == 0)
        clr  = NAVY if j == 0 else TEXT
        tx(slide, val, hx + 0.12, ry + 0.12, hw - 0.2, 0.24,
           13, clr, bold, PP_ALIGN.CENTER, spa=0)

result_card(slide, "Distance ≥ |Displacement|    Equality holds only when motion is along a straight line without any reversal.",
            XL, Y0 + 0.40 + 5 * 0.48 + 0.08, CW, 0.52)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 10 — WORKED EXAMPLE A — STRAIGHT PATH
# ═══════════════════════════════════════════════════════════════════════════════
slide = new_slide("Worked example A", "Straight-line path — when distance equals |displacement|", GREEN)
add_notes(slide, "1:00 — Ask students to predict before you calculate. 'Are distance and displacement magnitude equal here? Why?' Because every step goes in the same direction — no path is wasted. Would they differ if the object went from 1 to 9 and back to 7? Yes — distance = 16, but Δx = 6.")

tx(slide, "An object moves from  x = 1 m  to  x = 9 m  in a straight line without reversing.",
   XL, Y0, CW, 0.28, 17, NAVY, True, spa=0)

NLxA, NLyA, NLwA = XL + 0.1, Y0 + 1.30, CW - 0.2
number_line(slide, NLxA, NLyA, NLwA,
            list(range(0, 11)),
            labels={1: "Start (xᵢ = 1)", 9: "Finish (xf = 9)"},
            colors={1: BLUE, 9: GREEN})
path_arrow(slide, NLxA, NLyA, NLwA, 1, 9, 0, 10, BLUE, "8 m", above=True)

ty = NLyA + 1.22
formula_card(slide, "Distance  =  8 m",
             "Straight path — one segment", XL, ty, 5.55, 0.82)
formula_card(slide, "Δx  =  9 − 1  =  + 8 m",
             "Direction: positive (+x)", XL + 6.2, ty, 5.55, 0.82)

result_card(slide, "Distance = |Displacement| = 8 m   ✓   This equality holds for any straight path with no reversal.",
            XL, ty + 1.00, CW, 0.52)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 11 — WORKED EXAMPLE B — REVERSAL
# ═══════════════════════════════════════════════════════════════════════════════
slide = new_slide("Worked example B", "Reversal of direction — when they differ", ORANGE)
add_notes(slide, "1:10 — Most commonly tested NEET scenario. Ask students to predict before calculating. A common wrong answer is distance = 3 m (using the signed sum). Stress: for distance you ADD all segment lengths regardless of direction. The signed approach gives displacement, not distance.")

tx(slide, "Path:  x = − 2 m  →  x = 6 m  →  x = 1 m",
   XL, Y0, CW, 0.28, 18, NAVY, True, spa=0)

NLxB, NLyB, NLwB = XL + 0.1, Y0 + 1.30, CW - 0.2
VLB, VHB = -3, 7
number_line(slide, NLxB, NLyB, NLwB,
            list(range(VLB, VHB + 1)),
            labels={-2: "Start", 6: "Turn", 1: "End"},
            colors={-2: BLUE, 6: ORANGE, 1: RED})
path_arrow(slide, NLxB, NLyB, NLwB, -2, 6,  VLB, VHB, BLUE,   "8 m",  above=True)
path_arrow(slide, NLxB, NLyB, NLwB,  6, 1,  VLB, VHB, ORANGE, "5 m",  above=False)

tyB = NLyB + 1.24
formula_card(slide, "Distance  =  8  +  5  =  13 m",
             "Sum both segment lengths — all positive", XL, tyB, 5.55, 0.82)
formula_card(slide, "Δx  =  1 − (−2)  =  + 3 m",
             "xf − xi only — direction is positive", XL + 6.2, tyB, 5.55, 0.82)

trap_card(slide,
          "Do NOT compute distance by adding signed displacements: (+8) + (−5) = 3 m is the displacement, not the distance. Distance = 13 m; |Δx| = 3 m.",
          XL, tyB + 1.00, CW, 0.70)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 12 — BREAK
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
_bg(slide, SOFT)
_solid(slide, 0, 0, 13.333, 0.06, BLUE)
tx(slide, "5-MINUTE BREAK", XL, 2.0, CW, 0.60, 36, NAVY, True, PP_ALIGN.CENTER, spa=4)
tx(slide, "Before you return: write one sentence distinguishing distance from displacement.",
   XL, 2.95, CW, 0.28, 14, MUTED, False, PP_ALIGN.CENTER, spa=0)
tx(slide, "Origin  ·  Position x  ·  Distance (path)  ·  Displacement Δx = xf − xi",
   XL, 3.50, CW, 0.28, 14, BLUE, True, PP_ALIGN.CENTER, spa=0)
_footer(slide, 12)
add_notes(slide, "Break — 5 minutes. After break review 2–3 student sentences from the prompt above.")


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 13 — SIGNS OF DISPLACEMENT
# ═══════════════════════════════════════════════════════════════════════════════
slide = new_slide("Direction", "What does the sign of displacement mean?", BLUE)
add_notes(slide, "1:18 — Stress: the sign depends entirely on which direction was chosen as positive. Δx = 0 does NOT mean the object did not move — it means it returned to its starting position. A large negative displacement may correspond to a very long path.")

cases = [
    (BLUE,   "Δx > 0", "Net movement in\nthe positive direction", "Finish is to the right of start"),
    (ORANGE, "Δx < 0", "Net movement in\nthe negative direction", "Finish is to the left of start"),
    (GREEN,  "Δx = 0", "Start = Finish\n(position unchanged)", "At rest OR returned to start"),
]
for i, (col, sym, meaning, note) in enumerate(cases):
    cx = XL + i * 4.05
    cw = 3.82
    _card_bg(slide, cx, Y0, cw, 3.0, WHITE, col, 1.5)
    _solid(slide, cx, Y0, cw, 0.52, col, False)
    tx(slide, sym, cx + 0.1, Y0 + 0.10, cw - 0.2, 0.30, 22, WHITE, True, PP_ALIGN.CENTER, spa=0)
    tx(slide, meaning, cx + 0.18, Y0 + 0.78, cw - 0.3, 0.60, 15, col, True, spa=3, lsp=1.15)
    tx(slide, note,    cx + 0.18, Y0 + 1.55, cw - 0.3, 0.58, 13, MUTED, spa=0)

result_card(slide, "Zero displacement does NOT mean zero distance. A complete revolution returns Δx = 0 but distance = full circumference.",
            XL, Y0 + 3.18, CW, 0.54)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 14 — NEET TRAPS
# ═══════════════════════════════════════════════════════════════════════════════
slide = new_slide("Exam readiness", "Three traps that cost marks — NEET and CBSE", ORANGE)
add_notes(slide, "1:28 — Ask students to read each wrong statement and correct it before you reveal. These appear almost every year as MCQ distractors.")

y_cur = Y0
y_cur = warn_row(slide,
    "Distance and displacement are always equal.",
    "Equal ONLY for a straight-line path with no reversal.",
    XL, y_cur, CW)
y_cur += 0.14
y_cur = warn_row(slide,
    "Negative displacement means the object moved in the wrong direction.",
    "Negative means net movement opposite to the chosen positive axis — it is not an error.",
    XL, y_cur, CW)
y_cur += 0.14
warn_row(slide,
    "After one complete revolution, both distance and displacement are zero.",
    "Displacement = 0 (returned to start), but distance = circumference ≠ 0.",
    XL, y_cur, CW)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 15 — PRACTICE
# ═══════════════════════════════════════════════════════════════════════════════
slide = new_slide("Student practice", "Solve — show full working", NAVY)
add_notes(slide, "1:36 — Students work individually for 10 minutes. Do not show answers yet. Walk the room and check: (a) path diagram drawn, (b) formula written, (c) units, (d) correct sign on displacement. After 10 min call on 3 students to explain working.")

tx(slide, "Show:  path diagram  →  formula  →  substitution  →  unit  →  direction",
   XL, Y0, CW, 0.25, 12, MUTED, False, spa=0)

problems = [
    "An object starts at x = −5 m, moves to x = +8 m, then returns to x = +2 m. Find distance and displacement.",
    "A ball rolls 12 m east then 5 m west. Find: (i) distance, (ii) displacement, (iii) is |displacement| < distance?",
    "A runner completes exactly one lap of a 400 m circular track. Find distance and displacement.",
]
for i, prob in enumerate(problems):
    y = Y0 + 0.38 + i * 0.78
    circ = slide.shapes.add_shape(MSO_SHAPE.OVAL,
        Inches(XL), Inches(y), Inches(0.40), Inches(0.40))
    circ.fill.solid(); circ.fill.fore_color.rgb = NAVY; circ.line.fill.background()
    tx(slide, str(i + 1), XL + 0.10, y + 0.12, 0.20, 0.20, 10, WHITE, True, PP_ALIGN.CENTER, spa=0)
    tx(slide, prob, XL + 0.57, y + 0.07, CW - 0.67, 0.30, 16, TEXT, spa=0)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 16 — NEET MCQs
# ═══════════════════════════════════════════════════════════════════════════════
slide = new_slide("NEET MCQ drill", "Choose first — justify second", ORANGE)
add_notes(slide, "1:50 — Answers: 1(B), 2(C), 3(B). Q1: equality only for straight path without reversal. Q2: complete revolution → xf = xi → Δx = 0. Q3: position coordinate is measured from origin, so it changes when origin shifts; distance and displacement between any two fixed events do not change.")

mcqs = [
    ("Distance equals the magnitude of displacement when an object",
     ["(A) moves in a curved path",
      "(B) moves in a straight line without reversing direction",
      "(C) completes a full circular revolution",
      "(D) moves in a straight line and reverses once"]),
    ("A particle completes one full revolution along a circular track. Its displacement is",
     ["(A) equal to the circumference",
      "(B) equal to the diameter",
      "(C) zero",
      "(D) equal to the radius"]),
    ("If only the origin of a reference frame is shifted, which quantity changes for a fixed object?",
     ["(A) Distance between two given points",
      "(B) Position coordinate of each point",
      "(C) Displacement between two given points",
      "(D) Total path length"]),
]
y_q = Y0
for i, (q, opts) in enumerate(mcqs):
    tx(slide, f"Q{i + 1}.  {q}", XL, y_q, CW, 0.25, 13, NAVY, True, spa=0)
    y_q += 0.26
    for opt in opts:
        tx(slide, opt, XL + 0.4, y_q, CW - 0.5, 0.20, 12, TEXT, spa=0)
        y_q += 0.21
    y_q += 0.08

_card_bg(slide, XL, y_q + 0.04, CW, 0.32, AMBER_BG, AMBER, 0.8)
tx(slide, "Choose your answer first — then justify in one sentence.",
   XL + 0.15, y_q + 0.12, CW - 0.2, 0.20, 12, AMBER, True, spa=0)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 17 — EXIT TEST
# ═══════════════════════════════════════════════════════════════════════════════
slide = new_slide("Exit test", "Answer individually — full working required", RED)
add_notes(slide, "1:58 — 7 minutes individual. Collect for topic-wise analysis. Model answer for Q3: distance = 8 + 5 = 13 m; displacement = 1 − (−2) = +3 m. For NEET: correct formula + correct sign = full marks. Q4 tests complete-revolution scenario.")

exit_qs = [
    "What minimum information is required before calling an object 'at rest'?",
    "Write the position of an object located 7 m in the negative direction from the origin.",
    "An object moves from x = −2 m to x = 6 m, then to x = 1 m. Calculate distance and displacement with full working.",
    "A car completes exactly half a circular lap of radius 7 m. Find (i) distance and (ii) magnitude of displacement.",
]
for i, q in enumerate(exit_qs):
    y = Y0 + i * 0.62
    circ = slide.shapes.add_shape(MSO_SHAPE.OVAL,
        Inches(XL), Inches(y), Inches(0.40), Inches(0.40))
    circ.fill.solid(); circ.fill.fore_color.rgb = RED; circ.line.fill.background()
    tx(slide, str(i + 1), XL + 0.10, y + 0.12, 0.20, 0.20, 10, WHITE, True, PP_ALIGN.CENTER, spa=0)
    tx(slide, q, XL + 0.57, y + 0.07, CW - 0.67, 0.30, 15, TEXT, spa=0)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 18 — SUMMARY
# ═══════════════════════════════════════════════════════════════════════════════
slide = new_slide("Closure", "Lecture summary and next session", NAVY)
add_notes(slide, "2:05 — One-minute close. Ask: 'Complete the sentence: Motion is relative because…' Read the four cards. Homework: NCERT §3.1–3.3; daily test (separate sheet); draw 3 original number-line journeys and calculate both quantities. Next: average speed, average velocity, x–t graphs.")

cards = [
    (BLUE,  "FRAME OF REFERENCE", "Origin + direction + clock — sets the coordinate system for observation"),
    (TEAL,  "POSITION  x",        "Location of an object relative to the origin  •  Unit: m  •  Dim: [L]"),
    (AMBER, "DISTANCE",           "Total path length  •  Scalar  •  ≥ 0  •  path-dependent"),
    (GREEN, "DISPLACEMENT  Δx",   "xf − xi  •  Vector  •  Path-independent  •  Can be +, − or 0"),
]
for i, (col, head, body) in enumerate(cards):
    r, c = divmod(i, 2)
    sx, sy = XL + c * 6.1, Y0 + r * 1.55
    sw = 5.82
    _card_bg(slide, sx, sy, sw, 1.30, WHITE, col, 1.3)
    _solid(slide, sx, sy, 0.07, 1.30, col)
    tx(slide, head, sx + 0.17, sy + 0.14, sw - 0.28, 0.22, 10, col, True, spa=3)
    tx(slide, body, sx + 0.17, sy + 0.45, sw - 0.28, 0.62, 13, TEXT, spa=0)

result_card(slide, "Next lecture:  Average speed, average velocity and the position–time graph  (x–t graph)",
            XL, Y0 + 3.22, CW, 0.52)


prs.save(OUTPUT)
print(f"Created {len(prs.slides)} slides")
print(OUTPUT)
