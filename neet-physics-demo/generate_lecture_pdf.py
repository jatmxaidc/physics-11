from pathlib import Path

from reportlab.lib.colors import Color, HexColor, white
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph


OUT = Path(__file__).resolve().parent / "lecture_materials" / "Physical_Quantities_SI_Units_Demo_Lecture.pdf"
PAGE = landscape((7.5 * inch, 13.333333 * inch))
W, H = PAGE

INK = HexColor("#111923")
MUTED = HexColor("#5E6975")
PAPER = HexColor("#FFFDF7")
GRID = HexColor("#E8E3D7")
BLUE = HexColor("#1769E0")
BLUE_SOFT = HexColor("#E9F1FF")
GREEN = HexColor("#138A59")
GREEN_SOFT = HexColor("#E7F7EF")
RED = HexColor("#E64444")
RED_SOFT = HexColor("#FFEDED")
GOLD = HexColor("#E4A11B")
GOLD_SOFT = HexColor("#FFF5D8")
NAVY = HexColor("#07111F")


def rounded_rect(c, x, y, width, height, radius=12, fill=PAPER, stroke=GRID, line=1):
    c.setFillColor(fill)
    c.setStrokeColor(stroke)
    c.setLineWidth(line)
    c.roundRect(x, y, width, height, radius, fill=1, stroke=1)


def wrap(c, text, x, y, width, font="Helvetica", size=16, color=INK, leading=None, bold=False):
    style = ParagraphStyle(
        "slide",
        fontName="Helvetica-Bold" if bold else font,
        fontSize=size,
        leading=leading or size * 1.3,
        textColor=color,
        spaceAfter=0,
    )
    p = Paragraph(text, style)
    _, height = p.wrap(width, H)
    p.drawOn(c, x, y - height)
    return height


def title(c, text, y=430, size=36, max_width=820):
    return wrap(c, text, 70, y, max_width, size=size, color=INK, leading=size * 1.08, bold=True)


def kicker(c, text):
    c.setFillColor(BLUE)
    c.roundRect(70, 486, 24, 4, 2, fill=1, stroke=0)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(104, 484, text.upper())


def tag(c, text, x, y, fill=BLUE_SOFT, color=BLUE):
    pad = 9
    width = stringWidth(text.upper(), "Helvetica-Bold", 8) + pad * 2
    c.setFillColor(fill)
    c.setStrokeColor(fill)
    c.roundRect(x, y, width, 20, 10, fill=1, stroke=0)
    c.setFillColor(color)
    c.setFont("Helvetica-Bold", 8)
    c.drawCentredString(x + width / 2, y + 6.5, text.upper())
    return width


def grid(c):
    c.setFillColor(PAPER)
    c.rect(0, 0, W, H, fill=1, stroke=0)
    c.setStrokeColor(GRID)
    c.setLineWidth(0.25)
    for x in range(0, int(W) + 1, 24):
        c.line(x, 0, x, H)
    for y in range(0, int(H) + 1, 24):
        c.line(0, y, W, y)


def footer(c, page, label="PHYSICAL QUANTITIES + SI UNITS"):
    c.setStrokeColor(GRID)
    c.line(70, 40, W - 70, 40)
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 8)
    c.drawRightString(W - 70, 24, f"{page:02d}")


def base(c, page, label="PHYSICAL QUANTITIES + SI UNITS"):
    grid(c)
    footer(c, page, label)


def blank_lines(c, x, y, width, count=3, gap=34):
    c.setStrokeColor(HexColor("#B9C0C8"))
    c.setLineWidth(0.8)
    for i in range(count):
        c.line(x, y - i * gap, x + width, y - i * gap)


def slide_1(c):
    base(c, 1, "DEMO LECTURE")
    kicker(c, "Class 11 Physics / Demo lecture")
    wrap(c, "Every measurement<br/>tells a story.", 70, 425, 560, size=54, leading=55, bold=True)
    wrap(c, "Physical quantities, SI units and the art of converting between them.", 72, 285, 470, size=18, color=MUTED, leading=27)
    tag(c, "CBSE foundation", 72, 190, GREEN_SOFT, GREEN)
    tag(c, "NEET ready", 202, 190, RED_SOFT, RED)
    tag(c, "Interactive demo", 302, 190, BLUE_SOFT, BLUE)

    c.setStrokeColor(BLUE)
    c.setLineWidth(3)
    c.circle(770, 265, 120, fill=0, stroke=1)
    c.ellipse(610, 218, 930, 312, fill=0, stroke=1)
    c.ellipse(650, 145, 890, 385, fill=0, stroke=1)
    c.setFillColor(BLUE)
    c.setFont("Helvetica-Bold", 58)
    c.drawCentredString(770, 245, "SI")
    wrap(c, "Teacher: ____________________   Date: ____________________", 70, 92, 650, size=11, color=MUTED)


def slide_2(c):
    base(c, 2, "WHY UNITS MATTER")
    kicker(c, "Opening story / Real consequence")
    title(c, "A $125 million unit mistake", 432, 44)
    wrap(c, "The Mars Climate Orbiter was lost because two engineering teams used different unit systems.", 72, 352, 520, size=18, color=MUTED, leading=27)

    rounded_rect(c, 70, 128, 400, 150, fill=RED_SOFT, stroke=HexColor("#F6BBBB"))
    c.setFillColor(RED)
    c.setFont("Helvetica-Bold", 28)
    c.drawString(95, 236, "Team A")
    c.setFillColor(INK)
    c.setFont("Helvetica-Bold", 18)
    c.drawString(95, 190, "pound-force seconds")
    wrap(c, "Imperial measurement output", 95, 165, 320, size=12, color=MUTED)

    rounded_rect(c, 490, 128, 400, 150, fill=BLUE_SOFT, stroke=HexColor("#B8D1F5"))
    c.setFillColor(BLUE)
    c.setFont("Helvetica-Bold", 28)
    c.drawString(515, 236, "Team B")
    c.setFillColor(INK)
    c.setFont("Helvetica-Bold", 18)
    c.drawString(515, 190, "newton seconds")
    wrap(c, "SI measurement expected", 515, 165, 320, size=12, color=MUTED)

    c.setFillColor(RED)
    c.setFont("Helvetica-Bold", 26)
    c.drawCentredString(480, 190, "!=")
    wrap(c, "<b>Discuss:</b> Where could a unit mistake be dangerous in everyday life?", 70, 90, 800, size=15, color=INK)


def slide_3(c):
    base(c, 3, "PHYSICAL QUANTITY")
    kicker(c, "Quick class challenge")
    title(c, "Can it be measured objectively?", 432, 42)
    wrap(c, "Circle the physical quantities. Cross out the others.", 72, 372, 700, size=17, color=MUTED)

    items = [("Length", True), ("Happiness", False), ("Temperature", True), ("Beauty", False), ("Mass", True), ("Time", True)]
    x_positions = [72, 345, 618]
    y_positions = [240, 125]
    for index, (name, _) in enumerate(items):
        x = x_positions[index % 3]
        y = y_positions[index // 3]
        rounded_rect(c, x, y, 240, 86, fill=white, stroke=GRID)
        c.setFillColor(INK)
        c.setFont("Helvetica-Bold", 20)
        c.drawString(x + 22, y + 48, name)
        c.setStrokeColor(HexColor("#B9C0C8"))
        c.circle(x + 205, y + 44, 17, fill=0, stroke=1)

    rounded_rect(c, 72, 63, 786, 42, fill=BLUE_SOFT, stroke=HexColor("#B8D1F5"))
    c.setFillColor(BLUE)
    c.setFont("Helvetica-Bold", 18)
    c.drawCentredString(465, 77, "Physical quantity = numerical value x unit")


def slide_4(c):
    base(c, 4, "BASE + DERIVED")
    kicker(c, "The quantity family")
    title(c, "Base quantities build the language. Derived quantities build the physics.", 432, 37)

    rounded_rect(c, 70, 145, 385, 205, fill=GREEN_SOFT, stroke=HexColor("#B8E5D0"))
    tag(c, "Independent", 92, 303, GREEN_SOFT, GREEN)
    wrap(c, "Base quantities", 92, 286, 320, size=26, bold=True)
    wrap(c, "Length<br/>Mass<br/>Time<br/>Electric current<br/>Temperature", 92, 235, 290, size=15, color=MUTED, leading=24)

    rounded_rect(c, 505, 145, 385, 205, fill=BLUE_SOFT, stroke=HexColor("#B8D1F5"))
    tag(c, "Constructed", 527, 303, BLUE_SOFT, BLUE)
    wrap(c, "Derived quantities", 527, 286, 320, size=26, bold=True)
    wrap(c, "Area = length x length<br/>Speed = distance / time<br/>Force = mass x acceleration", 527, 235, 310, size=15, color=MUTED, leading=28)

    wrap(c, "<b>Your turn:</b> Is velocity a base quantity or a derived quantity? Explain.", 70, 105, 810, size=15)
    blank_lines(c, 70, 76, 820, 1)


def slide_5(c):
    base(c, 5, "SEVEN SI BASE UNITS")
    kicker(c, "International System of Units")
    title(c, "Seven agreed foundations. One scientific language.", 432, 42)

    units = [
        ("m", "Length", "metre"),
        ("kg", "Mass", "kilogram"),
        ("s", "Time", "second"),
        ("A", "Current", "ampere"),
        ("K", "Temperature", "kelvin"),
        ("mol", "Amount", "mole"),
        ("cd", "Intensity", "candela"),
    ]
    start_x = 70
    card_w = 196
    card_h = 112
    gap = 18
    for i, (symbol, quantity, unit) in enumerate(units):
        row = 0 if i < 4 else 1
        col = i if i < 4 else i - 4
        x = start_x + col * (card_w + gap)
        if row == 1:
            x += 107
        y = 242 if row == 0 else 105
        rounded_rect(c, x, y, card_w, card_h, fill=white, stroke=GRID)
        c.setFillColor(BLUE_SOFT)
        c.roundRect(x + 16, y + 52, 45, 45, 10, fill=1, stroke=0)
        c.setFillColor(BLUE)
        c.setFont("Helvetica-Bold", 16 if len(symbol) < 3 else 12)
        c.drawCentredString(x + 38.5, y + 67, symbol)
        c.setFillColor(INK)
        c.setFont("Helvetica-Bold", 14)
        c.drawString(x + 74, y + 75, quantity)
        c.setFillColor(MUTED)
        c.setFont("Helvetica", 11)
        c.drawString(x + 74, y + 56, unit)
        c.setFont("Helvetica", 9)
        c.drawString(x + 16, y + 20, "Example: __________________")

    wrap(c, "Board note: Kelvin uses K, not degree K. Unit symbols are never plural.", 70, 82, 820, size=12, color=MUTED)


def slide_6(c):
    base(c, 6, "DERIVED UNITS")
    kicker(c, "Build, do not memorise")
    title(c, "A derived unit comes directly from its equation.", 432, 42)

    rounded_rect(c, 145, 230, 670, 82, fill=BLUE_SOFT, stroke=BLUE, line=1.5)
    c.setFillColor(BLUE)
    c.setFont("Helvetica-Bold", 30)
    c.drawCentredString(480, 258, "Force = mass x acceleration")

    steps = [("mass", "kg"), ("acceleration", "m s^-2"), ("force", "kg m s^-2"), ("name", "newton (N)")]
    x = 75
    for index, (label, value) in enumerate(steps):
        rounded_rect(c, x, 125, 190, 83, fill=white, stroke=GRID)
        c.setFillColor(MUTED)
        c.setFont("Helvetica-Bold", 9)
        c.drawString(x + 16, 184, label.upper())
        c.setFillColor(INK)
        c.setFont("Helvetica-Bold", 17)
        c.drawString(x + 16, 151, value)
        if index < len(steps) - 1:
            c.setFillColor(BLUE)
            c.setFont("Helvetica-Bold", 20)
            c.drawString(x + 202, 157, "->")
        x += 220

    wrap(c, "<b>Board task:</b> Derive the SI unit of speed and acceleration.", 70, 95, 800, size=15)
    blank_lines(c, 70, 65, 820, 1)


def slide_7(c):
    base(c, 7, "UNIT CONVERSION")
    kicker(c, "Conversion without guesswork")
    title(c, "Multiply by conversion factors equal to one.", 432, 43)
    wrap(c, "Write the units on every line. Cancel them like algebraic factors.", 72, 325, 730, size=17, color=MUTED)

    c.setFillColor(INK)
    c.setFont("Helvetica-Bold", 30)
    c.drawString(74, 260, "72 km h^-1")
    c.setFillColor(BLUE)
    c.setFont("Helvetica-Bold", 24)
    c.drawString(272, 263, "x")

    factors = [
        (310, "1000 m", "1 km"),
        (520, "1 h", "3600 s"),
    ]
    for x, top, bottom in factors:
        c.setFillColor(INK)
        c.setFont("Helvetica-Bold", 19)
        c.drawCentredString(x + 70, 280, top)
        c.setStrokeColor(INK)
        c.line(x, 265, x + 140, 265)
        c.drawCentredString(x + 70, 240, bottom)

    c.setFillColor(BLUE)
    c.setFont("Helvetica-Bold", 24)
    c.drawString(480, 263, "x")
    c.drawString(683, 263, "=")
    c.setFillColor(GREEN)
    c.setFont("Helvetica-Bold", 31)
    c.drawString(725, 259, "20 m s^-1")

    rounded_rect(c, 70, 130, 820, 83, fill=GOLD_SOFT, stroke=HexColor("#EECF83"))
    c.setFillColor(GOLD)
    c.setFont("Helvetica-Bold", 17)
    c.drawString(95, 180, "SPEED SHORTCUT")
    c.setFillColor(INK)
    c.setFont("Helvetica-Bold", 17)
    c.drawString(95, 150, "km/h -> m/s: multiply by 5/18     |     m/s -> km/h: multiply by 18/5")
    wrap(c, "Use the whiteboard pen to cancel km and h above.", 70, 103, 800, size=12, color=MUTED)


def slide_8(c):
    base(c, 8, "NEET SPEED CHECK")
    kicker(c, "Student challenge / 45 seconds")
    title(c, "A car moves at 90 km h^-1. What is its speed in m s^-1?", 432, 40)

    options = [("A", "15 m s^-1"), ("B", "20 m s^-1"), ("C", "25 m s^-1"), ("D", "32.4 m s^-1")]
    positions = [(90, 225), (500, 225), (90, 120), (500, 120)]
    for (letter, value), (x, y) in zip(options, positions):
        rounded_rect(c, x, y, 365, 77, fill=white, stroke=GRID, line=1.2)
        c.setFillColor(BLUE_SOFT)
        c.circle(x + 39, y + 38, 21, fill=1, stroke=0)
        c.setFillColor(BLUE)
        c.setFont("Helvetica-Bold", 15)
        c.drawCentredString(x + 39, y + 33, letter)
        c.setFillColor(INK)
        c.setFont("Helvetica-Bold", 19)
        c.drawString(x + 78, y + 30, value)

    wrap(c, "<b>Show your work:</b>", 90, 92, 300, size=13)
    blank_lines(c, 230, 79, 625, 1)


def slide_9(c):
    base(c, 9, "CBSE + NEET TRAPS")
    kicker(c, "Four mistakes to eliminate")
    title(c, "Correct method earns marks and prevents traps.", 432, 41)

    mistakes = [
        ("1", "Number without a unit", "Write 20 m s^-1"),
        ("2", "Writing kgs or secs", "Use kg and s"),
        ("3", "Wrong conversion direction", "km/h to m/s: x 5/18"),
        ("4", "Too many calculator digits", "Respect given precision"),
    ]
    y = 305
    for number, wrong, fix in mistakes:
        rounded_rect(c, 72, y, 816, 62, fill=white, stroke=GRID)
        c.setFillColor(RED_SOFT)
        c.roundRect(90, y + 13, 36, 36, 9, fill=1, stroke=0)
        c.setFillColor(RED)
        c.setFont("Helvetica-Bold", 14)
        c.drawCentredString(108, y + 24, number)
        c.setFillColor(INK)
        c.setFont("Helvetica-Bold", 15)
        c.drawString(150, y + 24, wrong)
        c.setFillColor(GREEN)
        c.setFont("Helvetica-Bold", 13)
        c.drawRightString(860, y + 24, fix)
        y -= 75

    rounded_rect(c, 72, 69, 816, 40, fill=BLUE_SOFT, stroke=HexColor("#B8D1F5"))
    c.setFillColor(BLUE)
    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString(480, 83, "Estimate first: the converted answer should make physical sense.")


def slide_10(c):
    base(c, 10, "WHITEBOARD PRACTICE")
    kicker(c, "Practice board / Write directly here")
    title(c, "Three questions. Show every unit.", 432, 42)

    questions = [
        "1. Convert 36 km h^-1 to m s^-1.",
        "2. Write the SI unit of force using base units.",
        "3. A measurement is written as 12.5. What essential information is missing?",
    ]
    y = 350
    for question in questions:
        c.setFillColor(INK)
        c.setFont("Helvetica-Bold", 16)
        c.drawString(76, y, question)
        blank_lines(c, 76, y - 28, 805, 2, 29)
        y -= 110
    tag(c, "Write / discuss / correct", 76, 60, GREEN_SOFT, GREEN)


def slide_11(c):
    base(c, 11, "EXIT CHECK")
    kicker(c, "One-minute understanding check")
    title(c, "Finish these three sentences.", 432, 42)

    prompts = [
        "A physical quantity must contain ...",
        "The SI base unit of mass is ...",
        "To convert km/h to m/s, I ...",
    ]
    y = 340
    for index, prompt in enumerate(prompts, 1):
        c.setFillColor(BLUE_SOFT)
        c.circle(93, y + 3, 19, fill=1, stroke=0)
        c.setFillColor(BLUE)
        c.setFont("Helvetica-Bold", 13)
        c.drawCentredString(93, y - 2, str(index))
        c.setFillColor(INK)
        c.setFont("Helvetica-Bold", 18)
        c.drawString(130, y - 3, prompt)
        blank_lines(c, 130, y - 37, 720, 1)
        y -= 100

    rounded_rect(c, 70, 69, 820, 48, fill=GOLD_SOFT, stroke=HexColor("#EECF83"))
    c.setFillColor(GOLD)
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(480, 86, "Do not reveal the answers until students commit to a response.")


def slide_12(c):
    base(c, 12, "DAILY LEARNING LOOP")
    kicker(c, "Orbit learning system")
    title(c, "The lecture ends. The learning loop does not.", 432, 43)

    steps = [
        ("1", "Learn", "60-minute lesson"),
        ("2", "Attempt", "5-10 questions"),
        ("3", "Analyse", "Topic-wise data"),
        ("4", "Correct", "Next-day review"),
        ("5", "Continue", "Confident progress"),
    ]
    x = 45
    for index, (number, name, detail) in enumerate(steps):
        rounded_rect(c, x, 195, 160, 120, fill=white, stroke=GRID)
        c.setFillColor(BLUE_SOFT)
        c.circle(x + 80, 285, 17, fill=1, stroke=0)
        c.setFillColor(BLUE)
        c.setFont("Helvetica-Bold", 12)
        c.drawCentredString(x + 80, 281, number)
        c.setFillColor(INK)
        c.setFont("Helvetica-Bold", 18)
        c.drawCentredString(x + 80, 243, name)
        c.setFillColor(MUTED)
        c.setFont("Helvetica", 10)
        c.drawCentredString(x + 80, 220, detail)
        if index < len(steps) - 1:
            c.setFillColor(BLUE)
            c.setFont("Helvetica-Bold", 20)
            c.drawString(x + 166, 250, "->")
        x += 185

    rounded_rect(c, 155, 94, 650, 62, fill=NAVY, stroke=NAVY)
    c.setFillColor(white)
    c.setFont("Helvetica-Bold", 19)
    c.drawCentredString(480, 119, "Next: complete the daily test in the Orbit Physics app")
    wrap(c, "Tomorrow's lesson begins with evidence, not assumptions.", 70, 75, 820, size=13, color=MUTED, bold=True)


def generate():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(OUT), pagesize=PAGE)
    c.setTitle("Physical Quantities, SI Units and Unit Conversion - Demo Lecture")
    c.setAuthor("Orbit Physics")
    c.setSubject("Class 11 CBSE and NEET Physics demo lecture")
    slides = [
        slide_1,
        slide_2,
        slide_3,
        slide_4,
        slide_5,
        slide_6,
        slide_7,
        slide_8,
        slide_9,
        slide_10,
        slide_11,
        slide_12,
    ]
    for slide in slides:
        slide(c)
        c.showPage()
    c.save()
    print(OUT)


if __name__ == "__main__":
    generate()
