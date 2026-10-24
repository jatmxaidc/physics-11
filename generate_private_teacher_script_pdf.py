from html import escape
from pathlib import Path
import re

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
)


BASE = Path(__file__).resolve().parent
SOURCE = BASE / "SI_Units_Demo_Lecture_Private_Teacher_Script.md"
OUTPUT = BASE / "SI_Units_Demo_Lecture_Private_Teacher_Script.pdf"

pdfmetrics.registerFont(TTFont("Arial", r"C:\Windows\Fonts\arial.ttf"))
pdfmetrics.registerFont(TTFont("Arial-Bold", r"C:\Windows\Fonts\arialbd.ttf"))


def inline_markup(text):
    text = escape(text)
    text = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", text)
    text = re.sub(
        r"`(.+?)`",
        r'<font name="Courier" backColor="#EEEEEE">\1</font>',
        text,
    )
    return text


def add_page_number(canvas, document):
    canvas.saveState()
    canvas.setFont("Arial", 8)
    canvas.setFillColor(colors.HexColor("#666666"))
    canvas.drawString(18 * mm, 10 * mm, "PRIVATE TEACHER COPY")
    canvas.drawRightString(192 * mm, 10 * mm, str(document.page))
    canvas.restoreState()


def build_pdf():
    styles = getSampleStyleSheet()
    body = ParagraphStyle(
        "Body",
        parent=styles["BodyText"],
        fontName="Arial",
        fontSize=10.5,
        leading=15,
        spaceAfter=5,
    )
    title = ParagraphStyle(
        "Title",
        parent=body,
        fontName="Arial-Bold",
        fontSize=20,
        leading=24,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#17365D"),
        spaceAfter=12,
    )
    heading = ParagraphStyle(
        "Heading",
        parent=body,
        fontName="Arial-Bold",
        fontSize=15,
        leading=19,
        textColor=colors.HexColor("#17365D"),
        spaceBefore=10,
        spaceAfter=7,
    )
    subheading = ParagraphStyle(
        "Subheading",
        parent=body,
        fontName="Arial-Bold",
        fontSize=11,
        textColor=colors.HexColor("#8B1E1E"),
        spaceBefore=6,
        spaceAfter=3,
    )
    bullet = ParagraphStyle(
        "Bullet",
        parent=body,
        leftIndent=12,
        firstLineIndent=-8,
        bulletIndent=3,
    )

    story = []
    first_slide_heading = True
    for raw_line in SOURCE.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line == "---":
            story.append(Spacer(1, 3 * mm))
        elif line.startswith("# "):
            story.append(Paragraph(inline_markup(line[2:]), title))
        elif line.startswith("## "):
            if line.startswith("## Slide ") and not first_slide_heading:
                story.append(PageBreak())
            if line.startswith("## Slide "):
                first_slide_heading = False
            story.append(Paragraph(inline_markup(line[3:]), heading))
        elif line.startswith("### "):
            story.append(Paragraph(inline_markup(line[4:]), subheading))
        elif line.startswith("- "):
            story.append(
                Paragraph("• " + inline_markup(line[2:]), bullet)
            )
        else:
            story.append(Paragraph(inline_markup(line), body))

    document = SimpleDocTemplate(
        str(OUTPUT),
        pagesize=A4,
        rightMargin=18 * mm,
        leftMargin=18 * mm,
        topMargin=16 * mm,
        bottomMargin=17 * mm,
        title="Private Teacher Script — SI Units and Unit Conversion",
        author="Class 11 Physics",
    )
    document.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
    print(OUTPUT)


if __name__ == "__main__":
    build_pdf()
