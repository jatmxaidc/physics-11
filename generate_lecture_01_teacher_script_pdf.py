"""Build the private Lecture 01 teacher reading-script PDF from markdown."""
from html import escape
from pathlib import Path
import re

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer

BASE = Path(__file__).resolve().parent
SOURCE = BASE / "Lecture_01_Trigonometry_Teacher_Reading_Script.md"
OUTPUT = BASE / "Lecture_01_Trigonometry_Teacher_Reading_Script.pdf"

pdfmetrics.registerFont(TTFont("Calibri", r"C:\Windows\Fonts\calibri.ttf"))
pdfmetrics.registerFont(TTFont("Calibri-Bold", r"C:\Windows\Fonts\calibrib.ttf"))


def inline_markup(text):
    text = escape(text)
    text = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", text)
    text = re.sub(
        r"`(.+?)`",
        r'<font name="Courier" backColor="#F3F0E4">\1</font>',
        text,
    )
    return text


def add_page_number(canvas, document):
    canvas.saveState()
    canvas.setFillColor(colors.HexColor("#8B1E1E"))
    canvas.setFont("Calibri-Bold", 8)
    canvas.drawString(18 * mm, 11 * mm, "PRIVATE TEACHER COPY  —  do not project this PDF")
    canvas.setFillColor(colors.HexColor("#555555"))
    canvas.setFont("Calibri", 8)
    canvas.drawRightString(192 * mm, 11 * mm, str(document.page))
    canvas.restoreState()


def build_pdf():
    styles = getSampleStyleSheet()
    body = ParagraphStyle(
        "Body",
        parent=styles["BodyText"],
        fontName="Calibri",
        fontSize=10.5,
        leading=14.5,
        alignment=TA_JUSTIFY,
        spaceAfter=4,
    )
    title = ParagraphStyle(
        "Title",
        parent=body,
        fontName="Calibri-Bold",
        fontSize=18,
        leading=22,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#17365D"),
        spaceAfter=8,
    )
    heading = ParagraphStyle(
        "Heading",
        parent=body,
        fontName="Calibri-Bold",
        fontSize=14,
        leading=18,
        textColor=colors.HexColor("#17365D"),
        spaceBefore=2,
        spaceAfter=6,
        alignment=TA_CENTER,
        backColor=colors.HexColor("#E8EEF6"),
        borderPadding=6,
    )
    subheading = ParagraphStyle(
        "Subheading",
        parent=body,
        fontName="Calibri-Bold",
        fontSize=11,
        textColor=colors.HexColor("#8B1E1E"),
        spaceBefore=7,
        spaceAfter=2,
        alignment=TA_JUSTIFY,
    )
    bullet = ParagraphStyle(
        "Bullet",
        parent=body,
        leftIndent=14,
        firstLineIndent=-10,
        bulletIndent=4,
        alignment=TA_JUSTIFY,
    )

    story = []
    for raw_line in SOURCE.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line:
            story.append(Spacer(1, 1.5 * mm))
        elif line == "---":
            story.append(Spacer(1, 3 * mm))
        elif line.startswith("# "):
            story.append(Paragraph(inline_markup(line[2:]), title))
        elif line.startswith("## "):
            if line.startswith("## Slide "):
                story.append(PageBreak())
            story.append(Paragraph(inline_markup(line[3:]), heading))
        elif line.startswith("### "):
            story.append(Paragraph(inline_markup(line[4:]), subheading))
        elif line.startswith("- "):
            story.append(Paragraph("• " + inline_markup(line[2:]), bullet))
        else:
            story.append(Paragraph(inline_markup(line), body))

    document = SimpleDocTemplate(
        str(OUTPUT),
        pagesize=A4,
        rightMargin=16 * mm,
        leftMargin=16 * mm,
        topMargin=14 * mm,
        bottomMargin=16 * mm,
        title="Lecture 01 — Teacher Reading Script (Trigonometry)",
        author="Class 11 Physics Excellence Programme",
    )
    document.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
    print(OUTPUT)


if __name__ == "__main__":
    build_pdf()
