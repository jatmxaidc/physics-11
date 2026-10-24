"""
Programme homepage — Class 11 Physics lecture plan
CBSE 2026–27 + NEET Foundation
"""
from __future__ import annotations
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

BASE   = Path(__file__).resolve().parent
OUTPUT = BASE / "Lecture_Plan_Homepage.pptx"

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
RED        = RGBColor(172,  22,  22)
TEXT       = RGBColor(20,  26,  38)
MUTED      = RGBColor(90, 102, 120)
STROKE     = RGBColor(208, 218, 232)
WHITE      = RGBColor(255, 255, 255)
SOFT       = RGBColor(247, 250, 253)
GOLD       = RGBColor(212, 175,  55)

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]


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


def _card(slide, x, y, w, h, fill, border, line_pt=1.0):
    s = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    s.line.color.rgb = border
    s.line.width = Pt(line_pt)
    return s


def tx(slide, text, x, y, w, h,
       size=15, color=TEXT, bold=False,
       align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP,
       spa=0, lsp=1.08, font="Aptos"):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
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


def add_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def _footer(slide, n, total=2):
    _solid(slide, 0, 7.28, 13.333, 0.22, NAVY)
    tx(slide, "CLASS 11 PHYSICS  •  CBSE 2026–27 + NEET FOUNDATION  •  1 SEP 2026 – 31 DEC 2026",
       0.45, 7.32, 11.2, 0.16, 8, RGBColor(190, 210, 230), spa=0)
    tx(slide, f"{n:02d} / {total:02d}", 11.9, 7.32, 1.0, 0.16, 8, WHITE, True,
       PP_ALIGN.RIGHT, spa=0)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 01 — PROGRAMME HOMEPAGE
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
_bg(slide)
_solid(slide, 0, 0, 13.333, 7.5, NAVY)
_solid(slide, 0, 0, 0.12, 7.5, GOLD)

tx(slide, "CLASS 11 PHYSICS EXCELLENCE PROGRAMME",
   0.70, 0.42, 12.0, 0.28, 11, GOLD, True, spa=0)

tx(slide, "Lecture Plan  2026–27",
   0.70, 0.82, 12.0, 0.48, 28, WHITE, True, spa=0)

tx(slide, "CBSE annual examination  +  NEET foundation   •   one student, one clear path",
   0.70, 1.34, 12.0, 0.28, 14, RGBColor(170, 200, 230), spa=0)

# Four facts
facts = [
    ("1 SEP 2026",   "We start today"),
    ("133 TOPICS",   "Every PW Class 11 topic kept"),
    ("99 DAYS",      "Packed only where time allows"),
    ("31 DEC 2026",  "Class 11 complete"),
]
for i, (big, small) in enumerate(facts):
    x = 0.70 + i * 3.10
    _card(slide, x, 1.82, 2.92, 1.05, RGBColor(14, 42, 82), RGBColor(40, 80, 130), 1.0)
    tx(slide, big,   x + 0.16, 1.94, 2.60, 0.42, 18, WHITE, True, spa=0)
    tx(slide, small, x + 0.16, 2.40, 2.60, 0.28, 12, RGBColor(160, 190, 220), spa=0)

# How we work
tx(slide, "HOW EACH WEEK WORKS", 0.70, 3.10, 6.0, 0.22, 11, GOLD, True, spa=0)

rows = [
    ("Week 1",      "Starts Tuesday 1 Sep  •  then Monday–Saturday, 60 min"),
    ("After class", "Daily test of 5–10 questions  •  must be done before next lecture"),
    ("Next day",    "Teacher solves the weak questions first  •  then new topic"),
    ("Sunday",      "Formula sheet  •  error-log update  •  one timed mixed set"),
]
for i, (k, v) in enumerate(rows):
    y = 3.40 + i * 0.48
    _card(slide, 0.70, y, 7.55, 0.42, RGBColor(14, 42, 82), RGBColor(40, 80, 130), 0.75)
    tx(slide, k, 0.86, y + 0.10, 1.55, 0.24, 13, GOLD, True, spa=0)
    tx(slide, v, 2.50, y + 0.10, 5.55, 0.24, 13, WHITE, spa=0)

# Right column — where we begin
_card(slide, 8.50, 3.10, 4.35, 2.22, RGBColor(14, 42, 82), GOLD, 1.25)
tx(slide, "WE START HERE", 8.70, 3.24, 4.0, 0.22, 11, GOLD, True, spa=0)
tx(slide, "Lecture 01  •  Tue 1 Sep", 8.70, 3.50, 4.0, 0.28, 14, WHITE, True, spa=0)
tx(slide, "Trigonometry", 8.70, 3.82, 4.0, 0.32, 18, WHITE, True, spa=0)
tx(slide, "Ratios, identities, right-triangle\nproblems  •  PW lecture 1",
   8.70, 4.22, 3.95, 0.55, 13, RGBColor(180, 205, 230), spa=0)
tx(slide, "Then: Algebra → Vectors → Units",
   8.70, 4.88, 3.95, 0.28, 12, GOLD, spa=0)

# Promise strip
_card(slide, 0.70, 5.50, 12.15, 0.72, RGBColor(14, 42, 82), RGBColor(40, 80, 130), 0.75)
tx(slide, "Promise of this plan", 0.90, 5.58, 4.0, 0.22, 11, GOLD, True, spa=0)
tx(slide, "Every PW topic is taught. Hard ideas get a full hour. Related PW titles share a day only when 60 minutes is enough. Class 11 done by 31 December.",
   0.90, 5.84, 11.75, 0.28, 13, WHITE, spa=0)

tx(slide, "Handwritten notes", 0.70, 6.40, 3.0, 0.18, 9, RGBColor(110, 140, 175), spa=0)
_solid(slide, 0.70, 6.62, 12.15, 0.01, RGBColor(40, 80, 130))

_footer(slide, 1)
add_notes(slide,
    "Open with this homepage. Say: this is our year, not a copied coaching batch. "
    "Point to the four dates. Point to daily test. Point to Lecture 01 Mathematical Tools. "
    "Do not spend more than 3 minutes here.")


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 02 — YEAR MAP
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
_bg(slide)
_solid(slide, 0, 0, 13.333, 0.38, NAVY)
tx(slide, "YEAR MAP", 0.55, 0.10, 3.0, 0.20, 11, GOLD, True, spa=0)
tx(slide, "Class 11 Physics  •  133 PW topics  •  1 Sep → 31 Dec 2026",
   3.40, 0.10, 9.4, 0.20, 13, WHITE, spa=0)

tx(slide, "Full hour for hard topics. Shared hour only when the second title fits.",
   0.55, 0.52, 12.2, 0.28, 16, NAVY, True, spa=0)

phases = [
    (BLUE,  "PHASE 1  •  1 SEP – 12 OCT",
     "Tools, vectors, units, kinematics",
     [
         "Maths     11 topics · 9 days",
         "Vectors    7 topics · 5 days",
         "Units     10 topics · 8 days",
         "1D motion 14 topics · 9 days",
         "Plane      7 topics · 5 days",
         "",
         "L001  1 Sep   Trigonometry (full hour)",
         "L036  12 Oct  Plane practice",
     ]),
    (TEAL,  "PHASE 2  •  13 OCT – 24 NOV",
     "Laws to solids",
     [
         "Laws      16 topics · 11 days",
         "WEP        8 topics · 5 days",
         "CoM        7 topics · 5 days",
         "Rotation  12 topics · 9 days",
         "Gravitation 6 topics · 5 days",
         "Solids     3 topics · 2 days",
         "",
         "L073  24 Nov  Solids practice",
     ]),
    (AMBER, "PHASE 3  •  25 NOV – 31 DEC",
     "Fluids to waves + lock",
     [
         "Fluids     8 topics · 5 days",
         "Thermal    6 topics · 5 days",
         "KTG+thermo 6 topics · 6 days",
         "Oscillations 6 topics · 5 days",
         "Waves      6 topics · 5 days",
         "Lock tests 6 days (25–31 Dec)",
         "",
         "L105  31 Dec  Class 11 complete",
     ]),
]

for i, (col, head, sub, items) in enumerate(phases):
    x = 0.45 + i * 4.28
    _card(slide, x, 0.92, 4.10, 4.72, SOFT, col, 1.15)
    _solid(slide, x, 0.92, 4.10, 0.72, col)
    tx(slide, head, x + 0.16, 0.98, 3.80, 0.28, 12, WHITE, True, spa=0)
    tx(slide, sub,  x + 0.16, 1.28, 3.80, 0.26, 12, WHITE, spa=0)
    for j, item in enumerate(items):
        ty = 1.78 + j * 0.42
        tx(slide, item, x + 0.22, ty, 3.70, 0.32, 14, TEXT, spa=0)

tx(slide, "Daily test after every lecture   •   70% chapter gate before moving on   •   weak topics reteach within 7 days",
   0.55, 5.80, 12.2, 0.26, 13, MUTED, spa=0)

tx(slide, "Handwritten notes", 0.55, 6.18, 3.0, 0.18, 9, RGBColor(160, 170, 185), spa=0)
_solid(slide, 0.55, 6.40, 12.25, 0.01, STROKE)

_footer(slide, 2)
add_notes(slide,
    "Say: every PW topic is in the plan. Banking, FBD, rolling, 2D collision each get a full hour. "
    "Shared days are only where PW split one idea. Theory 24 Dec. Done 31 Dec.")

prs.save(OUTPUT)
print(f"Created {len(prs.slides)} slides")
print(OUTPUT)
