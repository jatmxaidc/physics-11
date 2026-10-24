from pathlib import Path

from reportlab.lib.colors import HexColor, white
from reportlab.pdfgen import canvas

from generate_lecture_pdf import (
    BLUE,
    BLUE_SOFT,
    GOLD,
    GOLD_SOFT,
    GREEN,
    GREEN_SOFT,
    GRID,
    H,
    INK,
    MUTED,
    NAVY,
    PAGE,
    RED,
    RED_SOFT,
    W,
    base,
    blank_lines,
    kicker,
    rounded_rect,
    tag,
    title,
    wrap,
)


OUT = Path(__file__).resolve().parent / "lecture_materials" / "Physical_Quantities_Whiteboard_Handwritten_Template.pdf"


def writing_area(c, x=70, y=85, width=820, height=260, label="WRITE / DRAW HERE"):
    c.setFillColor(HexColor("#FFFFFF"))
    c.setStrokeColor(HexColor("#C8CED5"))
    c.setLineWidth(1.1)
    c.setDash(6, 5)
    c.roundRect(x, y, width, height, 14, fill=1, stroke=1)
    c.setDash()
    c.setFillColor(HexColor("#A1A9B2"))
    c.setFont("Helvetica-Bold", 9)
    c.drawString(x + 18, y + height - 24, label)


def page_1(c):
    base(c, 1, "WHITEBOARD TEMPLATE")
    kicker(c, "Class 11 Physics / Demo lecture")
    wrap(c, "Physical Quantities,<br/>SI Units & Unit Conversion", 70, 425, 610, size=48, leading=52, bold=True)
    tag(c, "Handwritten lecture", 72, 205, BLUE_SOFT, BLUE)
    tag(c, "CBSE + NEET", 205, 205, GREEN_SOFT, GREEN)
    c.setStrokeColor(BLUE)
    c.setLineWidth(3)
    c.circle(790, 280, 105, fill=0, stroke=1)
    c.setFillColor(BLUE)
    c.setFont("Helvetica-Bold", 48)
    c.drawCentredString(790, 262, "SI")
    wrap(c, "Teacher: ______________________________", 70, 105, 500, size=13, color=MUTED)
    wrap(c, "Date: _________________________________", 520, 105, 350, size=13, color=MUTED)


def page_2(c):
    base(c, 2, "OPENING HOOK")
    kicker(c, "Opening story")
    title(c, "Why do units matter?", 432, 46)
    rounded_rect(c, 70, 330, 820, 67, fill=RED_SOFT, stroke=HexColor("#F5BFC5"))
    c.setFillColor(RED)
    c.setFont("Helvetica-Bold", 24)
    c.drawString(95, 354, "A $125 million unit mistake")
    writing_area(c, y=80, height=225, label="TELL THE STORY / DRAW THE SPACECRAFT / COLLECT STUDENT ANSWERS")


def page_3(c):
    base(c, 3, "PHYSICAL QUANTITIES")
    kicker(c, "Topic 1")
    title(c, "Physical Quantities", 432, 46)
    wrap(c, "Definition:", 72, 352, 180, size=16, bold=True)
    blank_lines(c, 175, 344, 690, 2, 35)
    wrap(c, "Examples:", 72, 268, 180, size=16, bold=True)
    blank_lines(c, 175, 260, 690, 2, 35)
    wrap(c, "Not physical quantities:", 72, 182, 220, size=16, bold=True)
    blank_lines(c, 285, 174, 580, 2, 35)
    rounded_rect(c, 70, 66, 820, 52, fill=BLUE_SOFT, stroke=HexColor("#B8D1F5"))
    c.setFillColor(BLUE)
    c.setFont("Helvetica-Bold", 19)
    c.drawCentredString(480, 84, "Physical quantity = __________________ x __________________")


def page_4(c):
    base(c, 4, "BASE + DERIVED")
    kicker(c, "Topic 2")
    title(c, "Fundamental & Derived Quantities", 432, 43)
    rounded_rect(c, 70, 105, 385, 260, fill=GREEN_SOFT, stroke=HexColor("#B8E5D0"))
    tag(c, "Fundamental", 92, 326, GREEN_SOFT, GREEN)
    wrap(c, "Meaning:", 92, 294, 120, size=14, bold=True)
    blank_lines(c, 92, 266, 320, 2, 34)
    wrap(c, "Examples:", 92, 188, 120, size=14, bold=True)
    blank_lines(c, 92, 160, 320, 2, 34)

    rounded_rect(c, 505, 105, 385, 260, fill=BLUE_SOFT, stroke=HexColor("#B8D1F5"))
    tag(c, "Derived", 527, 326, BLUE_SOFT, BLUE)
    wrap(c, "Meaning:", 527, 294, 120, size=14, bold=True)
    blank_lines(c, 527, 266, 320, 2, 34)
    wrap(c, "Examples:", 527, 188, 120, size=14, bold=True)
    blank_lines(c, 527, 160, 320, 2, 34)


def page_5(c):
    base(c, 5, "SI BASE UNITS")
    kicker(c, "Topic 3")
    title(c, "Seven SI Base Quantities", 432, 45)
    x, y, width, height = 100, 82, 760, 300
    c.setFillColor(white)
    c.setStrokeColor(HexColor("#BAC1C9"))
    c.setLineWidth(1)
    c.roundRect(x, y, width, height, 10, fill=1, stroke=1)
    row_h = height / 8
    col_widths = [310, 280, 170]
    headers = ["Physical quantity", "SI unit", "Symbol"]
    current_x = x
    c.setFillColor(BLUE_SOFT)
    c.roundRect(x, y + height - row_h, width, row_h, 10, fill=1, stroke=0)
    c.rect(x, y + height - row_h, width, row_h / 2, fill=1, stroke=0)
    for index, col_width in enumerate(col_widths):
        c.setFillColor(BLUE)
        c.setFont("Helvetica-Bold", 12)
        c.drawString(current_x + 14, y + height - row_h + 13, headers[index])
        current_x += col_width
        if index < 2:
            c.setStrokeColor(HexColor("#BAC1C9"))
            c.line(current_x, y, current_x, y + height)
    for row in range(1, 8):
        c.setStrokeColor(HexColor("#D4D8DD"))
        c.line(x, y + row_h * row, x + width, y + row_h * row)
        c.setFillColor(HexColor("#A2AAB3"))
        c.setFont("Helvetica", 9)
        c.drawString(x + 12, y + row_h * row + 13, str(8 - row))


def page_6(c):
    base(c, 6, "DERIVED QUANTITIES")
    kicker(c, "Topic 4")
    title(c, "Derived Quantities & Units", 432, 45)
    prompts = [
        "Speed =",
        "Acceleration =",
        "Force =",
        "Work =",
    ]
    y = 345
    for prompt in prompts:
        c.setFillColor(INK)
        c.setFont("Helvetica-Bold", 17)
        c.drawString(78, y, prompt)
        blank_lines(c, 220, y - 3, 645, 1)
        y -= 65
    writing_area(c, 70, 62, 820, 80, "DERIVE ONE SI UNIT USING BASE UNITS")


def page_7(c):
    base(c, 7, "UNIT CONVERSION")
    kicker(c, "Topic 5")
    title(c, "Unit Conversion Method", 432, 46)
    rounded_rect(c, 70, 322, 820, 63, fill=GOLD_SOFT, stroke=HexColor("#EECF83"))
    c.setFillColor(GOLD)
    c.setFont("Helvetica-Bold", 21)
    c.drawCentredString(480, 345, "72 km h^-1 = __________________ m s^-1")
    writing_area(c, y=83, height=210, label="WRITE THE COMPLETE CONVERSION AND CANCEL UNITS")
    wrap(c, "Shortcut discovered by students:  km/h -> m/s = x __________", 75, 67, 760, size=13, color=MUTED)


def page_8(c):
    base(c, 8, "NEET CHECK")
    kicker(c, "45-second student challenge")
    title(c, "90 km h^-1 = ? m s^-1", 432, 48)
    options = [("A", "15"), ("B", "20"), ("C", "25"), ("D", "32.4")]
    positions = [(90, 240), (500, 240), (90, 145), (500, 145)]
    for (letter, value), (x, y) in zip(options, positions):
        rounded_rect(c, x, y, 365, 68, fill=white, stroke=HexColor("#C6CCD3"))
        c.setFillColor(BLUE_SOFT)
        c.circle(x + 38, y + 34, 19, fill=1, stroke=0)
        c.setFillColor(BLUE)
        c.setFont("Helvetica-Bold", 14)
        c.drawCentredString(x + 38, y + 29, letter)
        c.setFillColor(INK)
        c.setFont("Helvetica-Bold", 19)
        c.drawString(x + 75, y + 26, f"{value} m s^-1")
    wrap(c, "Student method:", 90, 110, 180, size=13, bold=True)
    blank_lines(c, 210, 102, 645, 1)


def page_9(c):
    base(c, 9, "EXAM TRAPS")
    kicker(c, "CBSE presentation + NEET traps")
    title(c, "Common Mistakes", 432, 46)
    prompts = [
        "1. Number without a unit",
        "2. Incorrect unit symbol",
        "3. Wrong conversion direction",
        "4. Unreasonable final answer",
    ]
    y = 345
    for prompt in prompts:
        rounded_rect(c, 72, y - 25, 815, 53, fill=white, stroke=GRID)
        c.setFillColor(RED)
        c.setFont("Helvetica-Bold", 15)
        c.drawString(93, y - 4, prompt)
        c.setFillColor(GREEN)
        c.drawRightString(855, y - 4, "Correction: __________________________")
        y -= 67


def page_10(c):
    base(c, 10, "PRACTICE BOARD")
    kicker(c, "Write / discuss / correct")
    title(c, "Practice Board", 432, 46)
    questions = [
        "1. Convert 36 km h^-1 to m s^-1.",
        "2. Write the SI unit of force using base units.",
        "3. A measurement is written as 12.5. What is missing?",
    ]
    y = 355
    for question in questions:
        c.setFillColor(INK)
        c.setFont("Helvetica-Bold", 17)
        c.drawString(76, y, question)
        blank_lines(c, 76, y - 30, 805, 2, 30)
        y -= 112


def page_11(c):
    base(c, 11, "EXIT CHECK")
    kicker(c, "One-minute understanding check")
    title(c, "Complete Before You Leave", 432, 44)
    prompts = [
        "A physical quantity must contain",
        "The SI base unit of mass is",
        "To convert km/h to m/s, I",
    ]
    y = 345
    for index, prompt in enumerate(prompts, 1):
        c.setFillColor(BLUE_SOFT)
        c.circle(92, y + 2, 18, fill=1, stroke=0)
        c.setFillColor(BLUE)
        c.setFont("Helvetica-Bold", 13)
        c.drawCentredString(92, y - 3, str(index))
        c.setFillColor(INK)
        c.setFont("Helvetica-Bold", 17)
        c.drawString(130, y - 4, prompt)
        blank_lines(c, 130, y - 37, 730, 1)
        y -= 98


def page_12(c):
    base(c, 12, "DAILY LEARNING LOOP")
    kicker(c, "After-class system")
    title(c, "The Lecture Ends. Learning Continues.", 432, 43)
    steps = [
        ("1", "Learn"),
        ("2", "Attempt"),
        ("3", "Analyse"),
        ("4", "Correct"),
        ("5", "Continue"),
    ]
    x = 45
    for index, (number, name) in enumerate(steps):
        rounded_rect(c, x, 218, 160, 100, fill=white, stroke=GRID)
        c.setFillColor(BLUE_SOFT)
        c.circle(x + 80, 286, 17, fill=1, stroke=0)
        c.setFillColor(BLUE)
        c.setFont("Helvetica-Bold", 12)
        c.drawCentredString(x + 80, 282, number)
        c.setFillColor(INK)
        c.setFont("Helvetica-Bold", 17)
        c.drawCentredString(x + 80, 247, name)
        if index < 4:
            c.setFillColor(BLUE)
            c.setFont("Helvetica-Bold", 20)
            c.drawString(x + 166, 260, "->")
        x += 185
    writing_area(c, 90, 85, 780, 95, "EXPLAIN THE DAILY TEST + NEXT-DAY REVIEW CYCLE")


def generate():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(OUT), pagesize=PAGE)
    c.setTitle("Physical Quantities - Whiteboard Handwritten Lecture Template")
    c.setAuthor("Orbit Physics")
    c.setSubject("Sparse Microsoft Whiteboard lecture template for Class 11 Physics")
    for page in [page_1, page_2, page_3, page_4, page_5, page_6, page_7, page_8, page_9, page_10, page_11, page_12]:
        page(c)
        c.showPage()
    c.save()
    print(OUT)


if __name__ == "__main__":
    generate()
