from pathlib import Path

from reportlab.lib.colors import black, HexColor, white
from reportlab.lib.pagesizes import landscape
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas


OUT = Path(__file__).resolve().parent / "lecture_materials" / "Simple_Physical_Quantities_Lecture_Slides.pdf"
PAGE = landscape((7.5 * inch, 13.333333 * inch))
W, H = PAGE
GREY = HexColor("#666666")
LIGHT = HexColor("#D8D8D8")


def page(c, number, heading, subheading=None):
    c.setFillColor(white)
    c.rect(0, 0, W, H, fill=1, stroke=0)
    c.setFillColor(black)
    c.setFont("Helvetica-Bold", 32)
    c.drawString(55, H - 65, heading)
    if subheading:
        c.setFillColor(GREY)
        c.setFont("Helvetica", 15)
        c.drawString(57, H - 94, subheading)
    c.setStrokeColor(LIGHT)
    c.setLineWidth(1)
    c.line(55, 36, W - 55, 36)
    c.setFillColor(GREY)
    c.setFont("Helvetica", 9)
    c.drawRightString(W - 55, 20, str(number))


def line(c, x, y, width, count=1, gap=32):
    c.setStrokeColor(LIGHT)
    c.setLineWidth(1)
    for index in range(count):
        c.line(x, y - index * gap, x + width, y - index * gap)


def label(c, text, x, y, size=18):
    c.setFillColor(black)
    c.setFont("Helvetica-Bold", size)
    c.drawString(x, y, text)


def bullet(c, text, x, y, size=17):
    c.setFillColor(black)
    c.setFont("Helvetica", size)
    c.drawString(x, y, f"- {text}")


def generate():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(OUT), pagesize=PAGE)
    c.setTitle("Physical Quantities, SI Units and Unit Conversion")

    page(c, 1, "Physical Quantities, SI Units and Unit Conversion", "Class 11 Physics")
    label(c, "Today's topics", 60, 370, 22)
    bullet(c, "Physical quantities", 75, 325)
    bullet(c, "Fundamental and derived quantities", 75, 285)
    bullet(c, "SI system of units", 75, 245)
    bullet(c, "Unit conversion", 75, 205)
    label(c, "Teacher:", 60, 105, 14)
    line(c, 130, 102, 320)

    c.showPage()
    page(c, 2, "Why do we need measurement?")
    label(c, "Question:", 60, 390)
    bullet(c, "Can we describe a distance by saying only '5'?", 75, 350)
    bullet(c, "What information is missing?", 75, 310)
    label(c, "Discussion / examples", 60, 250, 16)
    line(c, 60, 220, 840, 5, 36)

    c.showPage()
    page(c, 3, "Physical Quantity")
    label(c, "Definition", 60, 390)
    line(c, 60, 355, 840, 2, 38)
    label(c, "Examples", 60, 270)
    line(c, 60, 235, 840, 2, 38)
    label(c, "Physical quantity =", 60, 135)
    line(c, 250, 132, 650)

    c.showPage()
    page(c, 4, "Fundamental and Derived Quantities")
    c.setStrokeColor(black)
    c.line(W / 2, 90, W / 2, 400)
    label(c, "Fundamental quantities", 75, 380, 20)
    label(c, "Meaning", 75, 330, 15)
    line(c, 75, 300, 360, 2, 38)
    label(c, "Examples", 75, 205, 15)
    line(c, 75, 175, 360, 2, 38)
    label(c, "Derived quantities", W / 2 + 35, 380, 20)
    label(c, "Meaning", W / 2 + 35, 330, 15)
    line(c, W / 2 + 35, 300, 360, 2, 38)
    label(c, "Examples", W / 2 + 35, 205, 15)
    line(c, W / 2 + 35, 175, 360, 2, 38)

    c.showPage()
    page(c, 5, "SI Base Quantities and Units")
    x, y, width, height = 95, 80, 770, 330
    row_h = height / 8
    cols = [310, 290, 170]
    c.setStrokeColor(black)
    c.rect(x, y, width, height, fill=0, stroke=1)
    c.setFont("Helvetica-Bold", 13)
    c.drawString(x + 12, y + height - 28, "Physical quantity")
    c.drawString(x + cols[0] + 12, y + height - 28, "SI unit")
    c.drawString(x + cols[0] + cols[1] + 12, y + height - 28, "Symbol")
    c.line(x + cols[0], y, x + cols[0], y + height)
    c.line(x + cols[0] + cols[1], y, x + cols[0] + cols[1], y + height)
    for index in range(1, 8):
        c.line(x, y + row_h * index, x + width, y + row_h * index)

    c.showPage()
    page(c, 6, "Rules for Writing Units")
    bullet(c, "Write the numerical value followed by the unit.", 70, 385)
    bullet(c, "Use the correct unit symbol.", 70, 335)
    bullet(c, "Unit symbols are not written in plural.", 70, 285)
    bullet(c, "Leave a space between the value and the unit.", 70, 235)
    bullet(c, "Always include the unit in the final answer.", 70, 185)
    label(c, "Examples / corrections", 60, 125, 16)
    line(c, 240, 122, 660)

    c.showPage()
    page(c, 7, "Derived Quantities and Units")
    label(c, "Speed =", 65, 390)
    line(c, 170, 387, 720)
    label(c, "Acceleration =", 65, 315)
    line(c, 215, 312, 675)
    label(c, "Force =", 65, 240)
    line(c, 160, 237, 730)
    label(c, "Work =", 65, 165)
    line(c, 150, 162, 740)
    label(c, "Derive one SI unit here", 65, 95, 15)
    line(c, 255, 92, 635)

    c.showPage()
    page(c, 8, "Unit Conversion")
    label(c, "Convert:", 60, 390, 20)
    c.setFont("Helvetica-Bold", 28)
    c.drawString(60, 335, "72 km h^-1 = ____________ m s^-1")
    label(c, "Method / calculation", 60, 270, 16)
    line(c, 60, 235, 840, 4, 40)
    label(c, "Shortcut:", 60, 70, 15)
    line(c, 145, 67, 755)

    c.showPage()
    page(c, 9, "NEET Practice Question")
    c.setFont("Helvetica-Bold", 25)
    c.drawString(60, 390, "A car moves at 90 km h^-1. Its speed in m s^-1 is:")
    bullet(c, "A. 15 m s^-1", 90, 320)
    bullet(c, "B. 20 m s^-1", 490, 320)
    bullet(c, "C. 25 m s^-1", 90, 260)
    bullet(c, "D. 32.4 m s^-1", 490, 260)
    label(c, "Working", 60, 190, 16)
    line(c, 60, 155, 840, 3, 38)

    c.showPage()
    page(c, 10, "Practice")
    bullet(c, "1. Convert 36 km h^-1 to m s^-1.", 60, 390)
    line(c, 60, 350, 840, 2, 36)
    bullet(c, "2. Write the SI unit of force using base units.", 60, 265)
    line(c, 60, 225, 840, 2, 36)
    bullet(c, "3. A measurement is written as 12.5. What is missing?", 60, 140)
    line(c, 60, 100, 840, 1)

    c.showPage()
    page(c, 11, "Summary")
    label(c, "Physical quantity =", 60, 390, 18)
    line(c, 250, 387, 640)
    label(c, "Fundamental quantity =", 60, 315, 18)
    line(c, 285, 312, 605)
    label(c, "Derived quantity =", 60, 240, 18)
    line(c, 245, 237, 645)
    label(c, "km/h to m/s =", 60, 165, 18)
    line(c, 220, 162, 670)
    label(c, "One question I still have:", 60, 90, 15)
    line(c, 260, 87, 630)

    c.showPage()
    page(c, 12, "Daily Test After Class", "5-10 questions / 10-15 minutes")
    bullet(c, "Complete the test before the next lecture.", 70, 365)
    bullet(c, "Incorrect questions are tracked topic-wise.", 70, 315)
    bullet(c, "Important mistakes are discussed before the next topic.", 70, 265)
    bullet(c, "Retry incorrect questions within 48 hours.", 70, 215)
    label(c, "Learn  ->  Attempt  ->  Analyse  ->  Correct  ->  Continue", 100, 125, 23)

    c.showPage()
    c.save()
    print(OUT)


if __name__ == "__main__":
    generate()
