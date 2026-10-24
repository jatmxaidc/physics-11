from pathlib import Path
import math

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


BASE = Path(__file__).resolve().parent
OUTPUT = BASE / "Lecture_05_Frame_of_Reference_Position_Distance_Displacement.pptx"

NAVY = RGBColor(24, 54, 93)
BLUE = RGBColor(31, 98, 166)
LIGHT_BLUE = RGBColor(231, 241, 252)
TEAL = RGBColor(0, 125, 138)
GREEN = RGBColor(38, 128, 84)
LIGHT_GREEN = RGBColor(232, 247, 239)
RED = RGBColor(180, 45, 45)
LIGHT_RED = RGBColor(253, 237, 237)
GOLD = RGBColor(190, 126, 0)
LIGHT_GOLD = RGBColor(255, 247, 224)
TEXT = RGBColor(34, 40, 49)
MUTED = RGBColor(95, 105, 118)
GRID = RGBColor(215, 221, 229)
WHITE = RGBColor(255, 255, 255)
SOFT = RGBColor(247, 249, 252)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank = prs.slide_layouts[6]


def set_background(slide, color=WHITE):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_text(
    slide,
    text,
    x,
    y,
    w,
    h,
    size=24,
    color=TEXT,
    bold=False,
    align=PP_ALIGN.LEFT,
    font="Aptos",
    margin=0.03,
    valign=MSO_ANCHOR.TOP,
    preserve_size=False,
):
    if not preserve_size:
        size = min(size, 18)
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = box.text_frame
    frame.clear()
    frame.margin_left = Inches(margin)
    frame.margin_right = Inches(margin)
    frame.margin_top = Inches(margin)
    frame.margin_bottom = Inches(margin)
    frame.vertical_anchor = valign
    paragraph = frame.paragraphs[0]
    paragraph.alignment = align
    run = paragraph.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return box


def add_bullets(slide, items, x, y, w, h, size=23, color=TEXT, spacing=8):
    size = min(size, 18)
    spacing = min(spacing, 4)
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.margin_left = Inches(0.12)
    frame.margin_right = Inches(0.08)
    for index, item in enumerate(items):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.text = item
        paragraph.level = 0
        paragraph.font.name = "Aptos"
        paragraph.font.size = Pt(size)
        paragraph.font.color.rgb = color
        paragraph.space_after = Pt(spacing)
        paragraph.text = "• " + paragraph.text
    return box


def add_rect(
    slide,
    x,
    y,
    w,
    h,
    fill=SOFT,
    line=GRID,
    radius=True,
):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(
        shape_type, Inches(x), Inches(y), Inches(w), Inches(h)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = line
    shape.line.width = Pt(1.2)
    return shape


def add_header(slide, title, section=None):
    if section:
        add_text(slide, section.upper(), 0.65, 0.28, 4.2, 0.35, 11, BLUE, True)
    add_text(
        slide,
        title,
        0.65,
        0.65,
        12.0,
        0.72,
        28,
        NAVY,
        True,
        preserve_size=True,
    )
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.65), Inches(1.38), Inches(12.0), Inches(0.02)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = GRID
    line.line.fill.background()


def add_footer(slide, number):
    add_text(
        slide,
        "LECTURE 05  •  MOTION IN A STRAIGHT LINE",
        0.65,
        7.05,
        7.0,
        0.22,
        8,
        MUTED,
    )
    add_text(slide, f"{number:02d}", 11.9, 7.02, 0.75, 0.25, 9, MUTED, True, PP_ALIGN.RIGHT)


def add_notes(slide, notes):
    slide.notes_slide.notes_text_frame.text = notes


def new_slide(title, section, notes, background=WHITE):
    slide = prs.slides.add_slide(blank)
    set_background(slide, background)
    add_header(slide, title, section)
    add_footer(slide, len(prs.slides))
    add_notes(slide, notes)
    return slide


def formula_box(slide, formula, caption, x, y, w, color=LIGHT_BLUE):
    add_rect(slide, x, y, w, 1.15, color, BLUE)
    add_text(
        slide,
        formula,
        x + 0.15,
        y + 0.12,
        w - 0.3,
        0.48,
        20,
        NAVY,
        True,
        PP_ALIGN.CENTER,
        preserve_size=True,
    )
    add_text(slide, caption, x + 0.15, y + 0.67, w - 0.3, 0.28, 12, MUTED, False, PP_ALIGN.CENTER)


def number_line(slide, x, y, w, values, labels=None, highlight=None):
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(0.035)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = NAVY
    line.line.fill.background()
    minimum, maximum = min(values), max(values)
    if minimum == maximum:
        maximum += 1
    labels = labels or {}
    for value in values:
        px = x + (value - minimum) / (maximum - minimum) * w
        tick = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(px - 0.012),
            Inches(y - 0.12),
            Inches(0.024),
            Inches(0.27),
        )
        tick.fill.solid()
        tick.fill.fore_color.rgb = NAVY
        tick.line.fill.background()
        color = RED if value == highlight else TEXT
        add_text(slide, str(value), px - 0.28, y + 0.16, 0.56, 0.3, 12, color, value == highlight, PP_ALIGN.CENTER)
        if value in labels:
            add_text(slide, labels[value], px - 0.6, y - 0.62, 1.2, 0.35, 13, color, True, PP_ALIGN.CENTER)


def apply_compact_card_style():
    """Apply the compact left-aligned card treatment across the deck."""
    for slide_number, slide in enumerate(prs.slides, start=1):
        original_shapes = list(slide.shapes)
        cards = [
            shape
            for shape in original_shapes
            if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
            and shape.auto_shape_type == MSO_SHAPE.ROUNDED_RECTANGLE
            and shape.width >= Inches(2.0)
            and shape.height >= Inches(0.5)
            and shape.height <= Inches(3.0)
        ]

        for card in cards:
            contained_text = []
            for shape in original_shapes:
                if shape is card or not getattr(shape, "has_text_frame", False):
                    continue
                if not shape.text.strip():
                    continue
                center_x = shape.left + shape.width // 2
                center_y = shape.top + shape.height // 2
                if (
                    card.left <= center_x <= card.left + card.width
                    and card.top <= center_y <= card.top + card.height
                ):
                    contained_text.append(shape)

            if not contained_text:
                continue

            for text_shape in contained_text:
                for paragraph in text_shape.text_frame.paragraphs:
                    paragraph.alignment = PP_ALIGN.LEFT
                    paragraph.space_before = Pt(0)
                    paragraph.space_after = Pt(0)
                    paragraph.line_spacing = 1.0
                text_shape.text_frame.margin_left = Inches(0.02)
                text_shape.text_frame.margin_right = Inches(0.02)
                text_shape.text_frame.margin_top = Inches(0.01)
                text_shape.text_frame.margin_bottom = Inches(0.01)

                point_sizes = [
                    run.font.size.pt
                    for paragraph in text_shape.text_frame.paragraphs
                    for run in paragraph.runs
                    if run.font.size is not None
                ]
                point_size = max(point_sizes, default=15)
                width_inches = max(0.5, text_shape.width / Inches(1))
                characters_per_line = max(
                    12, int(width_inches * 72 / (point_size * 0.52))
                )
                estimated_lines = 0
                for paragraph in text_shape.text_frame.paragraphs:
                    explicit_lines = paragraph.text.splitlines() or [""]
                    estimated_lines += sum(
                        max(1, math.ceil(len(line) / characters_per_line))
                        for line in explicit_lines
                    )
                fitted_text_height = max(
                    0.22, estimated_lines * point_size * 1.15 / 72 + 0.04
                )
                text_shape.height = Inches(fitted_text_height)

            # Compact vertically while preserving text arranged on the same row.
            if slide_number != 4:
                rows = []
                for text_shape in sorted(contained_text, key=lambda item: item.top):
                    if rows and abs(text_shape.top - rows[-1][0].top) <= Inches(0.14):
                        rows[-1].append(text_shape)
                    else:
                        rows.append([text_shape])

                current_top = card.top + Inches(0.13)
                for row in rows:
                    row_height = max(item.height for item in row)
                    for text_shape in row:
                        text_shape.top = current_top
                    current_top += row_height + Inches(0.04)

                fitted_height = current_top - card.top + Inches(0.06)
                card.height = max(Inches(0.52), fitted_height)

            card.line.width = Pt(0.9)

            # Slide 4 already has handcrafted accent edges.
            if slide_number == 4:
                continue
            try:
                accent_color = card.line.color.rgb or BLUE
            except (AttributeError, TypeError):
                accent_color = BLUE
            if accent_color in (GRID, WHITE):
                accent_color = BLUE

            accent = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                card.left,
                card.top,
                Inches(0.065),
                card.height,
            )
            accent.fill.solid()
            accent.fill.fore_color.rgb = accent_color
            accent.line.fill.background()


# 1 — Title
slide = prs.slides.add_slide(blank)
set_background(slide, WHITE)
add_rect(slide, 0, 0, 4.35, 7.5, NAVY, NAVY, False)
add_text(slide, "LECTURE 05", 0.65, 0.8, 3.1, 0.4, 14, RGBColor(185, 212, 242), True)
add_text(
    slide,
    "Motion begins\nwith a reference",
    0.65,
    1.5,
    3.15,
    1.8,
    25,
    WHITE,
    True,
    preserve_size=True,
)
add_text(slide, "Chapter 2 • Motion in a Straight Line", 0.65, 5.85, 3.1, 0.55, 13, WHITE)
add_text(
    slide,
    "Frame of Reference,\nPosition, Distance\nand Displacement",
    5.05,
    1.45,
    7.3,
    2.3,
    25,
    NAVY,
    True,
    preserve_size=True,
)
add_text(slide, "Class 11 Physics • CBSE + NEET Foundation • 2 hours", 5.05, 4.2, 6.8, 0.45, 15, MUTED)
add_rect(slide, 5.05, 5.0, 6.9, 1.0, LIGHT_BLUE, BLUE)
add_text(slide, "Observe → Define → Represent → Compare → Apply", 5.3, 5.29, 6.4, 0.4, 18, BLUE, True, PP_ALIGN.CENTER)
add_footer(slide, 1)
add_notes(slide, """0:00–0:03 — Welcome students. Say: “Today we begin kinematics, the study of motion without discussing its cause.” Ask: “Is a passenger sitting in a moving train at rest or in motion?” Accept both answers, then say the answer depends on the observer. State that this single idea will lead to frame of reference, position, distance and displacement. Do not define everything immediately.""")

# 2 — Session map
slide = new_slide(
    "What we will accomplish",
    "Session roadmap",
    """0:03–0:06 — Read the outcomes briefly. Explain the two-hour structure: concept blocks, quick checks, a short break, worked examples, NEET MCQs and an exit test. Tell students they should answer aloud and show working, not merely copy.""",
)
add_bullets(
    slide,
    [
        "Decide whether an object is at rest or in motion relative to an observer",
        "Choose an origin, direction and coordinate to describe position",
        "Calculate distance travelled and displacement",
        "Distinguish scalar distance from vector displacement",
        "Solve CBSE numericals and NEET-style conceptual questions",
    ],
    0.9,
    1.75,
    11.5,
    3.35,
    23,
)
add_rect(slide, 0.9, 5.35, 11.5, 0.55, LIGHT_GOLD, GOLD)
add_text(slide, "Success check: explain why distance can be non-zero while displacement is zero.", 1.1, 5.48, 11.0, 0.25, 14, GOLD, True)

# 3 — Recall
slide = new_slide(
    "Recall before we start",
    "Warm-up",
    """0:06–0:10 — Give 20 seconds per prompt. Expected: physical quantity = numerical value × unit; length is a base quantity; metre is SI unit; speed has m/s. Connect units to motion: position, distance and displacement all have the dimension of length and SI unit metre.""",
)
questions = [
    ("1", "What makes a physical quantity complete?"),
    ("2", "What is the SI base unit of length?"),
    ("3", "Is speed a base or derived quantity?"),
    ("4", "What does a negative sign usually indicate in one-dimensional motion?"),
]
for i, (n, q) in enumerate(questions):
    x, y = 1.0, 1.72 + i * 0.68
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(0.42), Inches(0.42))
    circle.fill.solid(); circle.fill.fore_color.rgb = BLUE; circle.line.fill.background()
    add_text(slide, n, x + 0.03, y + 0.1, 0.36, 0.2, 11, WHITE, True, PP_ALIGN.CENTER)
    add_text(slide, q, x + 0.62, y + 0.04, 10.5, 0.3, 16, TEXT, True)

# 4 — Hook
slide = new_slide(
    "At rest or in motion?",
    "The central question",
    """0:10–0:15 — Ask three students to answer independently. A passenger is at rest relative to the train but in motion relative to a person on the platform or Earth. Emphasise: rest and motion are relative concepts. Avoid saying “everything is always moving”; the scientific answer must name the reference frame.""",
)
add_rect(slide, 1.0, 1.75, 5.35, 1.6, RGBColor(246, 250, 255), RGBColor(184, 207, 232))
add_rect(slide, 1.0, 1.75, 0.09, 1.6, BLUE, BLUE, False)
add_text(slide, "TRAIN FRAME", 1.35, 1.98, 2.0, 0.22, 10, BLUE, True)
add_text(slide, "Passenger sitting inside a moving train", 1.35, 2.31, 4.65, 0.32, 17, NAVY, True)
add_text(slide, "Result: at rest relative to the train", 1.35, 2.82, 4.65, 0.28, 15, GREEN, True)
add_rect(slide, 6.98, 1.75, 5.35, 1.6, RGBColor(255, 251, 241), RGBColor(229, 204, 143))
add_rect(slide, 6.98, 1.75, 0.09, 1.6, GOLD, GOLD, False)
add_text(slide, "GROUND FRAME", 7.33, 1.98, 2.2, 0.22, 10, GOLD, True)
add_text(slide, "Observer standing on the platform", 7.33, 2.31, 4.65, 0.32, 17, NAVY, True)
add_text(slide, "Result: passenger is moving", 7.33, 2.82, 4.65, 0.28, 15, RED, True)
add_rect(slide, 1.0, 3.78, 11.33, 0.62, RGBColor(242, 250, 250), RGBColor(171, 213, 217))
add_text(slide, "KEY IDEA", 1.3, 3.98, 1.0, 0.22, 10, TEAL, True)
add_text(slide, "Same passenger • different observers • different descriptions", 2.28, 3.92, 9.7, 0.28, 16, TEAL, True)

# 5 — Rest and motion
slide = new_slide(
    "Rest and motion are relative",
    "Core concept",
    """0:15–0:20 — Define rest and motion using change of position with time relative to a chosen observer. Ask: “Is this classroom at rest?” Expected: relative to Earth/classroom, yes; relative to the Sun, it moves. Stress that the frame must be stated whenever ambiguity exists.""",
)
add_rect(slide, 1.15, 1.75, 5.0, 2.25, LIGHT_GREEN, GREEN)
add_text(slide, "REST", 1.4, 2.02, 4.5, 0.35, 18, GREEN, True, PP_ALIGN.CENTER)
add_text(slide, "Position does not change with time\nrelative to the chosen reference frame.", 1.4, 2.7, 4.5, 0.72, 17, TEXT, False, PP_ALIGN.CENTER)
add_rect(slide, 7.18, 1.75, 5.0, 2.25, LIGHT_RED, RED)
add_text(slide, "MOTION", 7.43, 2.02, 4.5, 0.35, 18, RED, True, PP_ALIGN.CENTER)
add_text(slide, "Position changes with time\nrelative to the chosen reference frame.", 7.43, 2.7, 4.5, 0.72, 17, TEXT, False, PP_ALIGN.CENTER)

# 6 — Frame definition
slide = new_slide(
    "Frame of reference",
    "Definition",
    """0:20–0:26 — Say: “A frame of reference is the coordinate system, together with a clock, relative to which observations are made.” For this introductory lecture, focus on origin, positive direction and time. Demonstrate with the classroom door as origin and right as positive.""",
)
formula_box(slide, "Reference frame = origin + axes + clock", "A system used to specify position and motion", 1.8, 1.75, 9.75)
elements = [
    ("Origin O", "Where position is measured from"),
    ("Direction / axes", "Which way is taken as positive"),
    ("Clock", "When each observation is made"),
]
for i, (head, desc) in enumerate(elements):
    x = 0.9 + i * 4.15
    add_rect(slide, x, 3.25, 3.7, 1.35, SOFT, GRID)
    add_text(slide, head, x + 0.2, 3.46, 3.3, 0.3, 17, BLUE, True, PP_ALIGN.CENTER)
    add_text(slide, desc, x + 0.25, 3.94, 3.2, 0.38, 14, TEXT, False, PP_ALIGN.CENTER)

# 7 — Choose a frame
slide = new_slide(
    "Build a one-dimensional reference frame",
    "Representation",
    """0:26–0:31 — Draw a long horizontal line on the board. Mark origin O = 0, right as positive, left as negative. Explain that changing the origin changes coordinate values but not the actual separation or physical path. Ask where an object 3 m left of O is located: x = −3 m.""",
)
number_line(slide, 1.1, 3.0, 11.1, [-6, -4, -2, 0, 2, 4, 6], {0: "Origin O"}, 0)
add_text(slide, "Negative direction  ←", 1.1, 1.95, 4.2, 0.4, 20, RED, True, PP_ALIGN.CENTER)
add_text(slide, "→  Positive direction", 8.0, 1.95, 4.2, 0.4, 20, GREEN, True, PP_ALIGN.CENTER)
add_rect(slide, 2.1, 4.65, 9.1, 0.75, LIGHT_GOLD, GOLD)
add_text(slide, "Coordinate tells location relative to the selected origin—not absolute location.", 2.35, 4.86, 8.6, 0.3, 16, GOLD, True, PP_ALIGN.CENTER)

# 8 — Same event, different origins
slide = new_slide(
    "Position depends on the chosen origin",
    "Frame of reference",
    """0:31–0:36 — Suppose a tree is physically fixed. Observer A selects a point 2 m left of the tree as origin, so x = +2 m. Observer B selects a point 5 m right, so x = −5 m. Both are correct. Position coordinate depends on frame; the object itself has not moved.""",
)
add_rect(slide, 1.15, 1.95, 5.0, 2.25, LIGHT_BLUE, BLUE)
add_text(slide, "Observer A", 1.4, 2.16, 4.5, 0.32, 18, BLUE, True, PP_ALIGN.CENTER)
add_text(slide, "Origin is 2 m left of the tree", 1.4, 2.72, 4.5, 0.32, 16, TEXT, False, PP_ALIGN.CENTER)
add_text(slide, "Position of tree:  x = +2 m", 1.4, 3.38, 4.5, 0.35, 18, NAVY, True, PP_ALIGN.CENTER)
add_rect(slide, 7.18, 1.95, 5.0, 2.25, LIGHT_GOLD, GOLD)
add_text(slide, "Observer B", 7.43, 2.16, 4.5, 0.32, 18, GOLD, True, PP_ALIGN.CENTER)
add_text(slide, "Origin is 5 m right of the tree", 7.43, 2.72, 4.5, 0.32, 16, TEXT, False, PP_ALIGN.CENTER)
add_text(slide, "Position of tree:  x = −5 m", 7.43, 3.38, 4.5, 0.35, 18, NAVY, True, PP_ALIGN.CENTER)
add_text(slide, "Different coordinates can describe the same physical location.", 1.4, 4.8, 10.5, 0.38, 18, TEAL, True, PP_ALIGN.CENTER)

# 9 — Position
slide = new_slide(
    "Position",
    "Definition",
    """0:36–0:42 — Define position as the location of an object relative to the chosen origin at a specified time. In one dimension, use the coordinate x. Point to A and B on the number line and ask students to state their positions. Emphasise that the sign identifies the side of the origin. Position has SI unit metre and dimension [L].""",
)
add_rect(slide, 0.95, 1.72, 0.07, 0.72, BLUE, BLUE, False)
add_text(slide, "Position specifies where an object is located relative to the chosen origin.", 1.25, 1.76, 10.7, 0.34, 18, TEXT, True)
add_text(slide, "In one-dimensional motion, position is represented by the coordinate x.", 1.25, 2.16, 10.7, 0.28, 15, MUTED)
number_line(
    slide,
    1.15,
    3.35,
    10.9,
    [-6, -4, -2, 0, 2, 3, 4, 6],
    {-4: "A", 0: "Origin O", 3: "B"},
)
add_text(slide, "A:  x = −4 m", 1.25, 4.42, 2.6, 0.3, 15, BLUE, True)
add_text(slide, "B:  x = +3 m", 4.1, 4.42, 2.6, 0.3, 15, GREEN, True)
add_text(slide, "SI unit: metre (m)    •    Dimension: [L]", 7.0, 4.42, 4.8, 0.3, 14, MUTED, True)

# 10 — Position examples
slide = new_slide(
    "Reading position on a number line",
    "Guided examples",
    """0:42–0:48 — Ask students to answer before revealing verbally. A at −4 m, B at +2 m, C at +6 m. Ask which is farthest from origin: C, because |6| is greatest. Ask distance between A and B: 6 m. Keep separation distinct from displacement until later.""",
)
number_line(slide, 1.1, 2.8, 11.1, [-6, -4, -2, 0, 2, 4, 6], {-4: "A", 2: "B", 6: "C", 0: "O"})
add_bullets(
    slide,
    [
        "State the position of A, B and C.",
        "Which object is farthest from the origin?",
        "What is the separation between A and B?",
    ],
    1.2,
    4.0,
    10.8,
    1.3,
    19,
)

# 11 — Change in position
slide = new_slide(
    "Initial and final position",
    "Preparing for displacement",
    """0:48–0:53 — Introduce xi and xf. Use x-axis: initial x_i = −2 m and final x_f = +5 m. Ask what changed. Do not call it distance yet. Write final minus initial and evaluate +7 m. Explain the positive sign means the net change is toward the chosen positive direction.""",
)
add_rect(slide, 1.0, 1.8, 5.2, 1.35, LIGHT_BLUE, BLUE)
add_text(slide, "Initial position", 1.25, 2.02, 4.7, 0.3, 17, BLUE, True, PP_ALIGN.CENTER)
add_text(slide, "xᵢ = −2 m", 1.25, 2.52, 4.7, 0.38, 18, NAVY, True, PP_ALIGN.CENTER)
add_rect(slide, 7.1, 1.8, 5.2, 1.35, LIGHT_GREEN, GREEN)
add_text(slide, "Final position", 7.35, 2.02, 4.7, 0.3, 17, GREEN, True, PP_ALIGN.CENTER)
add_text(slide, "xᶠ = +5 m", 7.35, 2.52, 4.7, 0.38, 18, NAVY, True, PP_ALIGN.CENTER)
formula_box(slide, "Change in position = xᶠ − xᵢ", "Always subtract the initial position from the final position", 2.15, 3.75, 9.0)

# 12 — Checkpoint
slide = new_slide(
    "Checkpoint 1",
    "Five-minute retrieval",
    """0:53–0:58 — Students answer individually, then compare with a partner. Answers: (1) relative to ground: motion; relative to bus: rest. (2) x = −3 m. (3) No; position can be negative while velocity is positive. Use responses to decide whether to repeat reference-frame explanation.""",
)
add_bullets(
    slide,
    [
        "A student sits inside a moving bus. State one frame in which the student is at rest and one in which the student is moving.",
        "An object is 3 m left of the origin. Write its position.",
        "Does negative position necessarily mean the object is moving in the negative direction? Explain.",
    ],
    1.0,
    1.75,
    11.2,
    3.35,
    21,
    TEXT,
    16,
)
add_rect(slide, 2.0, 5.2, 9.3, 0.65, LIGHT_GOLD, GOLD)
add_text(slide, "PAIR • EXPLAIN • CORRECT", 2.3, 5.38, 8.7, 0.3, 16, GOLD, True, PP_ALIGN.CENTER)

# 13 — Break
slide = new_slide(
    "Short break",
    "Reset",
    """0:58–1:03 — Five-minute break. Keep this slide visible. After the break, recap in one sentence: position tells where; now we study how much path is covered and how position changes.""",
    SOFT,
)
add_text(
    slide,
    "5 minutes",
    2.5,
    2.0,
    8.3,
    1.0,
    25,
    NAVY,
    True,
    PP_ALIGN.CENTER,
    preserve_size=True,
)
add_text(slide, "Return ready to compare distance and displacement.", 2.0, 3.45, 9.3, 0.55, 23, TEAL, True, PP_ALIGN.CENTER)
add_text(slide, "Recall during the break: origin • direction • position", 2.3, 4.55, 8.7, 0.45, 17, MUTED, False, PP_ALIGN.CENTER)

# 14 — Distance
slide = new_slide(
    "Distance travelled",
    "Definition",
    """1:03–1:10 — Define distance as total length of actual path travelled. It is scalar, non-negative and path-dependent. SI unit metre. Demonstrate by walking from one side of room to another and partially returning; every segment contributes positively to distance.""",
)
add_text(slide, "Total length of the actual path followed by an object.", 1.0, 1.8, 11.3, 0.65, 25, TEXT, True, PP_ALIGN.CENTER)
formula_box(slide, "Distance = sum of lengths of all path segments", "Scalar • path-dependent • never negative • SI unit: m", 1.8, 2.7, 9.75, LIGHT_GOLD)
add_rect(slide, 1.35, 4.2, 10.65, 0.9, LIGHT_RED, RED)
add_text(slide, "Common error: distance is not always |final position − initial position|.", 1.65, 4.45, 10.05, 0.35, 17, RED, True, PP_ALIGN.CENTER)

# 15 — Displacement
slide = new_slide(
    "Displacement",
    "Definition",
    """1:10–1:17 — Define displacement as change in position. In 1D, Δx = xf − xi. It is a vector; sign gives direction. It depends only on initial and final positions, not on path. Magnitude is |Δx|. Make students pronounce delta x as “change in x”.""",
)
add_text(slide, "Change in position from the initial point to the final point.", 1.0, 1.8, 11.3, 0.65, 25, TEXT, True, PP_ALIGN.CENTER)
formula_box(slide, "Δx = xᶠ − xᵢ", "Vector • path-independent • can be positive, negative or zero", 2.05, 2.7, 9.25, LIGHT_BLUE)
add_rect(slide, 1.35, 4.2, 10.65, 0.9, LIGHT_GREEN, GREEN)
add_text(slide, "Magnitude of displacement = |Δx|  and is always non-negative.", 1.65, 4.45, 10.05, 0.35, 17, GREEN, True, PP_ALIGN.CENTER)

# 16 — Comparison
slide = new_slide(
    "Distance versus displacement",
    "Compare",
    """1:17–1:24 — Work through each row. Distance: actual path, scalar, non-negative, path dependent. Displacement: shortest directed change, vector, signed in 1D, path independent. Both have SI unit metre and dimension [L]. Ask students to name the one similarity and three differences.""",
)
headers = ["Feature", "Distance", "Displacement"]
rows = [
    ("Meaning", "Actual path length", "Change in position"),
    ("Nature", "Scalar", "Vector"),
    ("Path", "Depends on path", "Only endpoints matter"),
    ("Possible value", "Zero or positive", "Positive, negative or zero"),
    ("SI unit", "metre (m)", "metre (m)"),
]
x_positions = [0.8, 3.5, 7.8]
widths = [2.7, 4.3, 4.7]
for j, head in enumerate(headers):
    add_rect(slide, x_positions[j], 1.75, widths[j], 0.65, NAVY, NAVY, False)
    add_text(slide, head, x_positions[j] + 0.1, 1.92, widths[j] - 0.2, 0.27, 15, WHITE, True, PP_ALIGN.CENTER)
for i, row in enumerate(rows):
    y = 2.4 + i * 0.62
    fill = WHITE if i % 2 == 0 else SOFT
    for j, value in enumerate(row):
        add_rect(slide, x_positions[j], y, widths[j], 0.62, fill, GRID, False)
        add_text(slide, value, x_positions[j] + 0.12, y + 0.13, widths[j] - 0.24, 0.3, 13, TEXT, j == 0, PP_ALIGN.CENTER)

# 17 — Example 1
slide = new_slide(
    "Worked example 1: straight path",
    "Distance and displacement",
    """1:24–1:30 — Object moves from x = 2 m to x = 9 m without reversing. Distance = 7 m. Displacement = 9 − 2 = +7 m. Point out equality occurs because motion is along a straight line without change of direction. Ask what the plus sign means.""",
)
number_line(slide, 1.1, 2.55, 11.1, [0, 2, 4, 6, 8, 9, 10], {2: "Start", 9: "Finish"})
add_text(slide, "Motion:  x = 2 m  →  x = 9 m", 2.2, 1.65, 8.9, 0.45, 24, NAVY, True, PP_ALIGN.CENTER)
add_rect(slide, 1.5, 4.0, 4.8, 0.95, LIGHT_GOLD, GOLD)
add_text(slide, "Distance = 7 m", 1.75, 4.27, 4.3, 0.35, 18, GOLD, True, PP_ALIGN.CENTER)
add_rect(slide, 7.0, 4.0, 4.8, 0.95, LIGHT_BLUE, BLUE)
add_text(slide, "Δx = 9 − 2 = +7 m", 7.25, 4.27, 4.3, 0.35, 18, BLUE, True, PP_ALIGN.CENTER)

# 18 — Example 2
slide = new_slide(
    "Worked example 2: reversal of direction",
    "Distance and displacement",
    """1:30–1:38 — Start at x = 2 m, move to x = 8 m, then to x = −3 m. First path = 6 m; second = 11 m; distance = 17 m. Displacement = −3 − 2 = −5 m. Explain why adding signed displacement segments also gives +6 + (−11) = −5 m.""",
)
add_text(slide, "Path:  2 m  →  8 m  →  −3 m", 2.0, 1.65, 9.3, 0.48, 26, NAVY, True, PP_ALIGN.CENTER)
number_line(slide, 1.1, 2.8, 11.1, [-4, -3, 0, 2, 4, 6, 8], {-3: "Finish", 2: "Start", 8: "Turn"})
add_rect(slide, 1.15, 4.4, 5.2, 0.95, LIGHT_GOLD, GOLD)
add_text(slide, "Distance = 6 + 11 = 17 m", 1.35, 4.67, 4.8, 0.35, 18, GOLD, True, PP_ALIGN.CENTER)
add_rect(slide, 6.95, 4.4, 5.2, 0.95, LIGHT_BLUE, BLUE)
add_text(slide, "Δx = −3 − 2 = −5 m", 7.15, 4.67, 4.8, 0.35, 18, BLUE, True, PP_ALIGN.CENTER)

# 19 — Round trip
slide = new_slide(
    "Worked example 3: round trip",
    "Special case",
    """1:38–1:44 — A runner completes one 400 m lap and returns to start. Distance = 400 m; displacement = 0. Ask: “Did the runner move?” Yes—distance proves path was covered. Zero displacement does not imply zero distance or no motion.""",
)
add_rect(slide, 1.0, 1.75, 11.3, 1.45, SOFT, GRID)
add_text(slide, "Runner completes one 400 m lap and returns to the starting point.", 1.4, 2.0, 10.5, 0.4, 18, NAVY, True, PP_ALIGN.CENTER)
add_text(slide, "Initial position = Final position", 2.8, 2.58, 7.7, 0.32, 17, TEAL, True, PP_ALIGN.CENTER)
add_rect(slide, 1.35, 3.8, 4.9, 0.95, LIGHT_GOLD, GOLD)
add_text(slide, "Distance = 400 m", 1.6, 4.07, 4.4, 0.35, 18, GOLD, True, PP_ALIGN.CENTER)
add_rect(slide, 7.05, 3.8, 4.9, 0.95, LIGHT_BLUE, BLUE)
add_text(slide, "Displacement = 0", 7.3, 4.07, 4.4, 0.35, 18, BLUE, True, PP_ALIGN.CENTER)

# 20 — Relation
slide = new_slide(
    "The fundamental relation",
    "Distance and displacement",
    """1:44–1:49 — State that distance is always greater than or equal to magnitude of displacement. Equality when path is straight with no reversal. Strict inequality when extra path or reversal occurs. Never compare signed displacement directly with distance; compare its magnitude.""",
)
formula_box(slide, "Distance ≥ |Displacement|", "For every journey", 2.2, 1.75, 8.9, LIGHT_GREEN)
cases = [
    ("Equal", "Straight-line motion\nwithout reversing direction", GREEN, LIGHT_GREEN),
    ("Distance is greater", "Curved path, detour\nor reversal of direction", RED, LIGHT_RED),
]
for i, (head, desc, color, fill) in enumerate(cases):
    x = 1.1 + i * 6.05
    add_rect(slide, x, 3.35, 5.5, 1.55, fill, color)
    add_text(slide, head, x + 0.2, 3.58, 5.1, 0.32, 18, color, True, PP_ALIGN.CENTER)
    add_text(slide, desc, x + 0.3, 4.08, 4.9, 0.5, 16, TEXT, False, PP_ALIGN.CENTER)

# 21 — Signs
slide = new_slide(
    "What does the sign of displacement mean?",
    "Direction",
    """1:49–1:54 — Positive displacement: final position lies in positive direction relative to initial. Negative: final lies in negative direction. Zero: same initial and final position. The sign depends on the selected positive axis, not on “good/bad” or distance travelled.""",
)
items = [
    ("Δx > 0", "Net change in positive direction", GREEN, LIGHT_GREEN),
    ("Δx < 0", "Net change in negative direction", RED, LIGHT_RED),
    ("Δx = 0", "Initial and final positions coincide", BLUE, LIGHT_BLUE),
]
for i, (formula, meaning, color, fill) in enumerate(items):
    x = 0.8 + i * 4.25
    add_rect(slide, x, 2.0, 3.8, 2.1, fill, color)
    add_text(slide, formula, x + 0.25, 2.3, 3.3, 0.4, 18, color, True, PP_ALIGN.CENTER)
    add_text(slide, meaning, x + 0.35, 3.0, 3.1, 0.55, 16, TEXT, False, PP_ALIGN.CENTER)

# 22 — Real life
slide = new_slide(
    "Real-life instruments tell different stories",
    "Application",
    """1:54–1:58 — Explain odometer measures accumulated path length approximately—distance. A navigation app can show straight-line change from start to destination—magnitude/direction of displacement—while also suggesting road distance. This anchors abstract definitions.""",
)
add_rect(slide, 1.15, 1.8, 5.0, 2.35, LIGHT_GOLD, GOLD)
add_text(slide, "ODOMETER", 1.4, 2.07, 4.5, 0.35, 18, GOLD, True, PP_ALIGN.CENTER)
add_text(slide, "Records total path length\ntravelled by the vehicle", 1.4, 2.68, 4.5, 0.65, 17, TEXT, False, PP_ALIGN.CENTER)
add_text(slide, "DISTANCE", 1.4, 3.55, 4.5, 0.3, 17, GOLD, True, PP_ALIGN.CENTER)
add_rect(slide, 7.18, 1.8, 5.0, 2.35, LIGHT_BLUE, BLUE)
add_text(slide, "MAP / GPS", 7.43, 2.07, 4.5, 0.35, 18, BLUE, True, PP_ALIGN.CENTER)
add_text(slide, "Can compare start and finish\nwith direction", 7.43, 2.68, 4.5, 0.65, 17, TEXT, False, PP_ALIGN.CENTER)
add_text(slide, "DISPLACEMENT", 7.43, 3.55, 4.5, 0.3, 17, BLUE, True, PP_ALIGN.CENTER)

# 23 — Guided practice
slide = new_slide(
    "Guided practice",
    "Solve on the board",
    """1:58–2:06 — If strictly limited to 120 minutes, start this at 1:55 and compress application slide. Q1: distance 13 m, displacement +3 m. Q2: distance 16 m, displacement 0. Q3: distance 20 m, displacement −20 m. Ask for formula and signs, not only answers.""",
)
add_bullets(
    slide,
    [
        "A moves from x = −5 m to x = +4 m, then returns to x = −2 m. Find distance and displacement.",
        "A child walks 8 m east and then 8 m west. Find distance and displacement.",
        "An object moves directly from x = +12 m to x = −8 m. Find distance and displacement.",
    ],
    0.95,
    1.7,
    11.35,
    3.45,
    21,
    TEXT,
    16,
)
add_rect(slide, 2.3, 5.25, 8.7, 0.55, LIGHT_BLUE, BLUE)
add_text(slide, "Show: path segments • formula • substitution • unit • direction", 2.55, 5.4, 8.2, 0.25, 14, BLUE, True, PP_ALIGN.CENTER)

# 24 — NEET MCQs
slide = new_slide(
    "NEET concept check",
    "Multiple choice",
    """2:06–2:12 — Answers: 1(B), 2(C), 3(B). Q1 equality only for straight motion without reversal. Q2 one full circle returns to start. Q3 coordinate depends on origin. Require students to justify each choice; guessing is not mastery.""",
)
add_text(slide, "1. Distance equals magnitude of displacement when motion is:", 0.9, 1.65, 11.6, 0.4, 18, NAVY, True)
add_text(slide, "A. circular   B. straight without reversal   C. random   D. always", 1.25, 2.15, 11.0, 0.38, 16, TEXT)
add_text(slide, "2. After one complete revolution on a circular track, displacement is:", 0.9, 2.95, 11.6, 0.4, 18, NAVY, True)
add_text(slide, "A. circumference   B. radius   C. zero   D. diameter", 1.25, 3.45, 11.0, 0.38, 16, TEXT)
add_text(slide, "3. Which quantity changes if only the origin is shifted?", 0.9, 4.25, 11.6, 0.4, 18, NAVY, True)
add_text(slide, "A. distance travelled   B. position coordinate   C. path length   D. elapsed time", 1.25, 4.75, 11.0, 0.38, 16, TEXT)
add_rect(slide, 3.0, 5.15, 7.3, 0.65, LIGHT_GOLD, GOLD)
add_text(slide, "Choose first. Justify second.", 3.25, 5.34, 6.8, 0.3, 16, GOLD, True, PP_ALIGN.CENTER)

# 25 — CBSE answer
slide = new_slide(
    "CBSE written-answer practice",
    "Construct a complete response",
    """2:12–2:16 — Students verbally frame a 3-mark answer: definitions, two differences, one example. Model structure: distance is actual path length and scalar; displacement is directed change and vector; 400 m lap gives distance 400 m and displacement zero. Correct incomplete phrases.""",
)
add_rect(slide, 0.9, 1.7, 11.5, 1.15, LIGHT_BLUE, BLUE)
add_text(slide, "Question: Distinguish between distance and displacement with one suitable example.", 1.2, 2.02, 10.9, 0.48, 20, NAVY, True, PP_ALIGN.CENTER)
steps = [
    ("1 mark", "Correct definitions"),
    ("1 mark", "Scalar/vector and path dependence"),
    ("1 mark", "Example with values and units"),
]
for i, (mark, point) in enumerate(steps):
    x = 1.0 + i * 4.1
    add_rect(slide, x, 3.35, 3.65, 1.35, SOFT, GRID)
    add_text(slide, mark, x + 0.2, 3.57, 3.25, 0.3, 17, BLUE, True, PP_ALIGN.CENTER)
    add_text(slide, point, x + 0.3, 4.08, 3.05, 0.35, 15, TEXT, False, PP_ALIGN.CENTER)

# 26 — Error check
slide = new_slide(
    "Common errors to eliminate",
    "Exam readiness",
    """2:16–2:20 — Use this as rapid true/false if time remains. Correct statements: position can be negative; distance cannot; displacement can be zero after motion; distance is at least magnitude of displacement. Stress units and direction in final answers.""",
)
errors = [
    "Treating distance and displacement as identical",
    "Forgetting: displacement = final position − initial position",
    "Using a negative sign for distance",
    "Assuming zero displacement means no motion occurred",
    "Writing a numerical answer without unit or direction",
]
add_bullets(slide, errors, 1.0, 1.75, 11.2, 3.65, 18, RED, 10)

# 27 — Exit test
slide = new_slide(
    "Exit test",
    "Individual response",
    """Final 5–7 minutes or assign as daily test if the session has reached two hours. Answers: 1 reference frame; 2 −6 m; 3 distance 14 m, displacement +2 m; 4 zero; 5 straight line with no reversal. Collect answers for topic-wise analysis.""",
)
add_bullets(
    slide,
    [
        "What must be specified before calling an object at rest or in motion?",
        "Write the position of an object 6 m left of the origin.",
        "A moves 8 m east and 6 m west. Find distance and displacement.",
        "What is displacement after returning to the starting point?",
        "When does distance equal magnitude of displacement?",
    ],
    0.95,
    1.6,
    11.4,
    3.75,
    20,
    TEXT,
    12,
)

# 28 — Summary/homework
slide = new_slide(
    "Takeaway and daily practice",
    "Closure",
    """Close in one minute. Ask students to complete the sentence: “Motion is relative because…” Then recap the four boxed statements. Homework: NCERT reading for Motion in a Straight Line introduction; 5–10 question daily test; draw three number-line journeys and calculate distance/displacement. Next lecture: average speed, average velocity and graphs.""",
)
summary = [
    ("Frame", "origin + direction + clock"),
    ("Position", "location relative to origin"),
    ("Distance", "total actual path length"),
    ("Displacement", "Δx = xᶠ − xᵢ"),
]
for i, (head, detail) in enumerate(summary):
    row, col = divmod(i, 2)
    x, y = 0.9 + col * 6.05, 1.75 + row * 1.55
    add_rect(slide, x, y, 5.55, 1.15, SOFT, GRID)
    add_text(slide, head, x + 0.2, y + 0.18, 1.85, 0.4, 17, BLUE, True)
    add_text(slide, detail, x + 2.05, y + 0.18, 3.2, 0.55, 15, TEXT)
add_rect(slide, 1.15, 4.85, 11.0, 0.85, LIGHT_GOLD, GOLD)
add_text(slide, "Next: average speed, average velocity and motion graphs", 1.45, 5.08, 10.4, 0.32, 18, GOLD, True, PP_ALIGN.CENTER)


apply_compact_card_style()
prs.save(OUTPUT)
print(f"Created {len(prs.slides)} slides")
print(OUTPUT)
