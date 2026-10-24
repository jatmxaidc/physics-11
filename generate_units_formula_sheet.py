# -*- coding: utf-8 -*-
"""Units & Dimensions Formula Notes — full chapter (Excellence Units 01–06).

Same warm notebook style as Basic Maths and Vectors formula sheets.
"""
from __future__ import annotations

import re
from pathlib import Path

import matplotlib
from reportlab.lib.colors import HexColor, white
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    KeepTogether,
    PageBreak,
    HRFlowable,
)

BASE = Path(__file__).resolve().parent
OUT = BASE / "class notes" / "Units_and_Dimensions_Formula_Sheet.pdf"
OUT2 = BASE / "lecture slides" / "Units_and_Dimensions_Formula_Sheet.pdf"

INK = HexColor("#1A2332")
MUTED = HexColor("#5C6B7A")
NAVY = HexColor("#0E3A5F")
GOLD = HexColor("#B8860B")
TEAL = HexColor("#0D7377")
CREAM = HexColor("#FFFBF5")
CARD = HexColor("#F0F7F7")
SOFT = HexColor("#F7F9FC")
LINE = HexColor("#D0D8E0")
ORANGE = HexColor("#B45309")
TIP_BG = HexColor("#FFF7ED")
EX_BG = HexColor("#ECFDF5")
EX_LINE = HexColor("#059669")

PAGE_W, PAGE_H = A4
MARGIN = 15 * mm


def _register_fonts():
    folder = Path(matplotlib.__file__).resolve().parent / "mpl-data" / "fonts" / "ttf"
    mapping = {
        "BodyFont": "DejaVuSans.ttf",
        "BodyBold": "DejaVuSans-Bold.ttf",
        "BodyItalic": "DejaVuSans-Oblique.ttf",
    }
    for face, name in mapping.items():
        pdfmetrics.registerFont(TTFont(face, str(folder / name)))


_register_fonts()
F, FB, FI = "BodyFont", "BodyBold", "BodyItalic"


def _powers(text: str) -> str:
    text = re.sub(r"\^\(([^)]+)\)", r"<super>\1</super>", text)
    text = re.sub(r"\^(-?[0-9]+)", r"<super>\1</super>", text)
    text = re.sub(r"\^([A-Za-z])", r"<super>\1</super>", text)
    return text


def m(text: str) -> str:
    parts = re.split(r"(</?(?:b|i|br|super|sub|font|/font|u)[^>]*>)", text)
    out = []
    for part in parts:
        if part.startswith("<"):
            out.append(part)
        else:
            esc = part.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            out.append(_powers(esc))
    return "".join(out)


def styles():
    s = getSampleStyleSheet()
    s.add(ParagraphStyle(
        "CoverBig", fontName=FB, fontSize=22, textColor=NAVY,
        alignment=TA_CENTER, leading=28, spaceAfter=4,
    ))
    s.add(ParagraphStyle(
        "CoverMid", fontName=F, fontSize=12, textColor=MUTED,
        alignment=TA_CENTER, leading=17, spaceAfter=3,
    ))
    s.add(ParagraphStyle(
        "Sec", fontName=FB, fontSize=13, textColor=NAVY,
        spaceBefore=10, spaceAfter=4, leading=17,
    ))
    s.add(ParagraphStyle(
        "Sub", fontName=FB, fontSize=10.5, textColor=TEAL,
        spaceBefore=7, spaceAfter=3, leading=14,
    ))
    s.add(ParagraphStyle(
        "Talk", fontName=F, fontSize=10, textColor=INK,
        leading=14.5, spaceAfter=4, alignment=TA_JUSTIFY,
    ))
    s.add(ParagraphStyle(
        "Formula", fontName=FB, fontSize=11, textColor=NAVY,
        leading=18, alignment=TA_CENTER, spaceBefore=1, spaceAfter=1,
    ))
    s.add(ParagraphStyle(
        "CellL", fontName=F, fontSize=9, textColor=INK,
        leading=12.5, alignment=TA_LEFT,
    ))
    s.add(ParagraphStyle(
        "Tip", fontName=FI, fontSize=9, textColor=ORANGE,
        leading=12.5, spaceBefore=2, spaceAfter=4,
    ))
    s.add(ParagraphStyle(
        "Ex", fontName=F, fontSize=9.5, textColor=INK,
        leading=13.5, spaceAfter=2,
    ))
    s.add(ParagraphStyle(
        "TOC", fontName=F, fontSize=10.5, textColor=INK,
        leading=16, leftIndent=2,
    ))
    return s


S = styles()


def header_footer(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(NAVY)
    canvas.roundRect(0, PAGE_H - 11 * mm, PAGE_W, 11 * mm, 0, fill=1, stroke=0)
    canvas.setFillColor(GOLD)
    canvas.rect(0, PAGE_H - 11.8 * mm, PAGE_W, 1.2, fill=1, stroke=0)
    canvas.setFillColor(white)
    canvas.setFont(FB, 8.5)
    canvas.drawString(MARGIN, PAGE_H - 7 * mm, "Excellence  ·  Class 11 Physics")
    canvas.setFont(F, 8)
    canvas.drawRightString(PAGE_W - MARGIN, PAGE_H - 7 * mm, "Units & Dimensions — formula notes")
    canvas.setStrokeColor(LINE)
    canvas.setLineWidth(0.5)
    canvas.line(MARGIN, 12 * mm, PAGE_W - MARGIN, 12 * mm)
    canvas.setFillColor(MUTED)
    canvas.setFont(F, 8)
    canvas.drawCentredString(
        PAGE_W / 2, 7 * mm,
        f"Page {doc.page}   ·   Full chapter (Units 01–06)   ·   Write these in your own words too",
    )
    canvas.restoreState()


def para(text, style="Talk"):
    return Paragraph(m(text), S[style])


def formula_card(text, blank_after=3):
    """One formula per line. Pass a list to stack separate formulas (never join with ·)."""
    if isinstance(text, (list, tuple)):
        text = "<br/>".join(text)
    data = [[para(text, "Formula")]]
    t = Table(data, colWidths=[PAGE_W - 2 * MARGIN])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), CARD),
        ("BOX", (0, 0), (-1, -1), 1.2, TEAL),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return KeepTogether([t, Spacer(1, blank_after * mm)])


def tip_card(text):
    data = [[para(f"<b>Remember:</b> {text}", "Tip")]]
    t = Table(data, colWidths=[PAGE_W - 2 * MARGIN])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), TIP_BG),
        ("BOX", (0, 0), (-1, -1), 0.9, ORANGE),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    return KeepTogether([t, Spacer(1, 3 * mm)])


def meaning_card(lines):
    """Explain symbols / words before the formula box."""
    if isinstance(lines, str):
        lines = [lines]
    body = "<br/>".join([f"<b>What the words mean</b>"] + [f"• {ln}" for ln in lines])
    data = [[para(body, "Ex")]]
    t = Table(data, colWidths=[PAGE_W - 2 * MARGIN])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), SOFT),
        ("BOX", (0, 0), (-1, -1), 0.9, TEAL),
        ("LEFTPADDING", (0, 0), (-1, -1), 9),
        ("RIGHTPADDING", (0, 0), (-1, -1), 9),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))
    return KeepTogether([t, Spacer(1, 3 * mm)])


def example(title, lines, answer):
    body = "<br/>".join(
        [f"<b>Worked example — {title}</b>"]
        + [f"• {ln}" for ln in lines]
        + [f"<b>Final answer:</b> {answer}"]
    )
    data = [[para(body, "Ex")]]
    t = Table(data, colWidths=[PAGE_W - 2 * MARGIN])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), EX_BG),
        ("BOX", (0, 0), (-1, -1), 1.0, EX_LINE),
        ("LEFTPADDING", (0, 0), (-1, -1), 9),
        ("RIGHTPADDING", (0, 0), (-1, -1), 9),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))
    return KeepTogether([t, Spacer(1, 4 * mm)])


def section(num_title):
    return KeepTogether([
        para(num_title, "Sec"),
        HRFlowable(width="100%", thickness=1.4, color=GOLD, spaceAfter=5, spaceBefore=0),
    ])


def two_cards(left_title, left_lines, right_title, right_lines):
    gap = 4 * mm
    w = (PAGE_W - 2 * MARGIN - gap) / 2

    def side(title, lines):
        bits = [para(f"<b>{title}</b>", "Sub")] + [para(x, "CellL") for x in lines]
        rows = [[b] for b in bits]
        tb = Table(rows, colWidths=[w - 4])
        wrap = Table([[tb]], colWidths=[w])
        wrap.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), SOFT),
            ("BOX", (0, 0), (-1, -1), 0.8, LINE),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        return wrap

    outer = Table(
        [[side(left_title, left_lines), side(right_title, right_lines)]],
        colWidths=[w, w],
    )
    outer.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (0, 0), gap / 2),
        ("LEFTPADDING", (1, 0), (1, 0), gap / 2),
    ]))
    return KeepTogether([outer, Spacer(1, 3 * mm)])


def nice_table(headers, rows, size=8.5):
    n = len(headers)
    col_w = (PAGE_W - 2 * MARGIN) / n
    hs = ParagraphStyle(
        "HCell", fontName=FB, fontSize=size, textColor=white,
        leading=size + 3, alignment=TA_CENTER,
    )
    cs = ParagraphStyle(
        "CCell", fontName=F, fontSize=size, textColor=INK,
        leading=size + 3, alignment=TA_CENTER,
    )
    data = [[Paragraph(m(h), hs) for h in headers]]
    for row in rows:
        data.append([Paragraph(m(c), cs) for c in row])
    t = Table(data, colWidths=[col_w] * n, repeatRows=1)
    cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("BACKGROUND", (0, 1), (-1, -1), CREAM),
        ("BOX", (0, 0), (-1, -1), 0.9, NAVY),
        ("INNERGRID", (0, 0), (-1, -1), 0.4, LINE),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]
    for i in range(2, len(data), 2):
        cmds.append(("BACKGROUND", (0, i), (-1, i), white))
    cmds.append(("BACKGROUND", (0, 1), (0, -1), HexColor("#E8F1F8")))
    t.setStyle(TableStyle(cmds))
    return KeepTogether([t, Spacer(1, 3 * mm)])


def build():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT2.parent.mkdir(parents=True, exist_ok=True)

    doc = BaseDocTemplate(
        str(OUT),
        pagesize=A4,
        leftMargin=MARGIN,
        rightMargin=MARGIN,
        topMargin=14 * mm,
        bottomMargin=15 * mm,
        title="Units and Dimensions Formula Notes — Excellence Class 11",
        author="Excellence Programme",
    )
    frame = Frame(MARGIN, 15 * mm, PAGE_W - 2 * MARGIN, PAGE_H - 29 * mm, id="main")
    doc.addPageTemplates([PageTemplate(id="main", frames=frame, onPage=header_footer)])

    story = []

    # ===== COVER =====
    story.append(Spacer(1, 8 * mm))
    story.append(para("EXCELLENCE", "CoverBig"))
    story.append(para("Class 11 Physics  ·  Session 2026–27", "CoverMid"))
    story.append(Spacer(1, 2 * mm))
    story.append(HRFlowable(width="50%", thickness=2, color=GOLD, spaceBefore=2, spaceAfter=6))
    story.append(para("Units and Measurement", "CoverBig"))
    story.append(para("Clear formula notes — every word explained", "CoverMid"))
    story.append(Spacer(1, 3 * mm))
    story.append(para(
        "This notebook is written for a student who wants every idea clear. "
        "First we explain in plain English. Then we explain every symbol. "
        "Then we show the formula on its own line. Then a small example.",
        "CoverMid",
    ))
    story.append(Spacer(1, 3 * mm))
    story.append(formula_card([
        "How to use each page",
        "1. Read “In plain English”",
        "2. Read “What the words mean”",
        "3. Copy the formula box",
        "4. Try the green example",
    ]))
    story.append(para("<b>Chapter map</b>", "Sub"))
    for line in [
        "0.  Symbol dictionary (read this first)",
        "1.  Number and unit",
        "2.  Dimensions (M, L, T)",
        "3.  Changing one unit into another",
        "4.  Checking if a formula is possible",
        "5.  Finding a formula using dimensions",
        "6.  Significant figures (which digits to keep)",
        "7.  Accuracy, precision and kinds of error",
        "8.  Absolute, relative and percentage error",
        "9.  How error grows when we use a formula",
        "10. Metre scale, vernier, screw gauge",
        "11. Revision lines to say aloud",
    ]:
        story.append(para(line, "TOC"))
    story.append(PageBreak())

    # ===== 0 SYMBOLS =====
    story.append(section("0. Symbol dictionary — read this first"))
    story.append(para(
        "Many students get stuck only because a short letter is not explained. "
        "Keep this page open while you study the rest.",
        "Talk",
    ))
    story.append(nice_table(
        ["Symbol / short name", "Full meaning in school language"],
        [
            ["Q", "any physical quantity (mass, length, force…)"],
            ["n", "the number part (like 5 in 5 kg)"],
            ["u", "the unit part (like kg in 5 kg)"],
            ["M", "mass dimension (from kilogram)"],
            ["L", "length dimension (from metre)"],
            ["T", "time dimension (from second)"],
            ["[Q]", "dimensional formula of Q (powers of M, L, T)"],
            ["Δx  (delta x)", "absolute error in x — how many units wrong"],
            ["x̄  (x bar)", "mean = average of many readings"],
            ["L.C.", "least count = smallest step the tool can show"],
            ["M.S.D.", "main scale division = gap between two main marks"],
            ["V.S.D.", "vernier scale division = gap between two vernier marks"],
            ["M.S.R.", "main scale reading = last main mark before vernier zero"],
            ["V.S.R.", "vernier scale reading = vernier mark that matches a main mark"],
            ["Pitch", "how far a screw moves in one full turn"],
            ["N", "number of equal marks on the circular (round) scale"],
        ],
        size=8.5,
    ))
    story.append(tip_card(
        "If a formula uses a letter you do not know, come back to this table first. "
        "Do not memorise a formula with unknown letters."
    ))

    # ===== 1 =====
    story.append(section("1. Number and unit"))
    story.append(para(
        "In plain English: whenever we measure something, we write two parts — "
        "how many, and of what. “5” alone is not enough. “5 kg” is complete.",
        "Talk",
    ))
    story.append(meaning_card([
        "Physical quantity = the thing we measure (mass, length, time, force…)",
        "Number (n) = the digit part",
        "Unit (u) = the agreed size we compare with (kg, m, s…)",
        "SI = the world standard unit system used in school and exams",
    ]))
    story.append(formula_card([
        "Physical quantity = number × unit",
        "Q = n × u",
    ]))
    story.append(para(
        "SI starts with 7 base quantities. Other units (newton, joule, watt) are built from these.",
        "Talk",
    ))
    story.append(nice_table(
        ["What we measure", "SI unit name", "We write"],
        [
            ["Length", "metre", "m"],
            ["Mass", "kilogram", "kg"],
            ["Time", "second", "s"],
            ["Electric current", "ampere", "A"],
            ["Temperature", "kelvin", "K"],
            ["Amount of substance", "mole", "mol"],
            ["Light intensity", "candela", "cd"],
        ],
        size=9,
    ))
    story.append(example(
        "writing a quantity",
        [
            "A bag has mass five kilograms.",
            "Number n = 5. Unit u = kg.",
            "So Q = 5 × kg.",
        ],
        "5 kg",
    ))

    story.append(PageBreak())

    # ===== 2 =====
    story.append(section("2. Dimensions — what sits inside a quantity"))
    story.append(para(
        "In plain English: dimension tells us which base things hide inside a quantity. "
        "Force is made from mass, length and time. We write that as powers of M, L and T.",
        "Talk",
    ))
    story.append(meaning_card([
        "M = mass, L = length, T = time",
        "[force] means “dimensional formula of force”",
        "L T^(-1) means length to power 1, time to power −1 (like speed = distance/time)",
        "A negative power means “divided by” that quantity",
    ]))
    story.append(formula_card([
        "Dimension = powers of M, L, T inside a quantity",
        "We write it inside square brackets: [force], [speed]",
    ]))
    story.append(para("<b>Table to learn (most asked)</b>", "Sub"))
    story.append(para(
        "Read the middle column first. It tells you how to remember the left column.",
        "Talk",
    ))
    story.append(nice_table(
        ["Quantity", "Think like this", "Dimensional formula"],
        [
            ["Speed / velocity", "distance ÷ time", "[v] = L T^(-1)"],
            ["Acceleration", "speed ÷ time", "[a] = L T^(-2)"],
            ["Force", "mass × acceleration", "[F] = M L T^(-2)"],
            ["Momentum", "mass × speed", "[p] = M L T^(-1)"],
            ["Work / energy", "force × distance", "[W] = M L^2 T^(-2)"],
            ["Power", "work ÷ time", "[P] = M L^2 T^(-3)"],
            ["Pressure", "force ÷ area", "[P] = M L^(-1) T^(-2)"],
            ["Density", "mass ÷ volume", "[ρ] = M L^(-3)"],
            ["Frequency", "1 ÷ time", "[f] = T^(-1)"],
            ["Angular speed ω", "angle ÷ time (angle has no dimension)", "[ω] = T^(-1)"],
            ["Torque", "force × distance (turning)", "[τ] = M L^2 T^(-2)"],
            ["G", "from F = G m₁ m₂ / r²", "[G] = M^(-1) L^3 T^(-2)"],
        ],
        size=8,
    ))
    story.append(two_cards(
        "Same formula — different names",
        [
            "Work, energy and torque",
            "all have M L^2 T^(-2)",
            "Same dimensions does NOT mean",
            "they are the same quantity",
        ],
        "No dimension?",
        [
            "Angle in radian: has unit, no M/L/T",
            "Strain = change in length ÷ length",
            "Inside sin, cos, e^x, log: treat as 1",
            "“No dimension” ≠ always “no unit”",
        ],
    ))
    story.append(example(
        "find [force]",
        [
            "Force = mass × acceleration.",
            "Mass → M. Acceleration → L T^(-2).",
            "So [F] = M × L T^(-2) = M L T^(-2).",
        ],
        "[F] = M L T^(-2)",
    ))

    # ===== 3 =====
    story.append(section("3. Changing one unit into another"))
    story.append(para(
        "In plain English: the real amount of stuff does not change when you change the unit label. "
        "If the unit becomes smaller, the number must become larger.",
        "Talk",
    ))
    story.append(meaning_card([
        "n₁ = old number, u₁ = old unit",
        "n₂ = new number, u₂ = new unit",
        "n₁ u₁ = n₂ u₂ means “same physical amount”",
    ]))
    story.append(formula_card([
        "n₁ u₁ = n₂ u₂",
        "n₂ = n₁ × (u₁ / u₂)",
    ]))
    story.append(tip_card(
        "For area, square the length factor. For volume, cube it. "
        "1 m = 100 cm, so 1 m² = 100² cm² = 10,000 cm²."
    ))
    story.append(example(
        "5 m² into cm²",
        [
            "1 m = 100 cm, so 1 m² = (100)² cm² = 10^4 cm².",
            "5 m² = 5 × 10^4 cm².",
        ],
        "5 × 10^4 cm²",
    ))

    story.append(PageBreak())

    # ===== 4 =====
    story.append(section("4. Checking if a formula is possible"))
    story.append(para(
        "In plain English: you can add only like with like. "
        "You cannot add metres to kilograms. "
        "So in a correct physics equation, every term that is added or subtracted "
        "must have the same dimensions. This rule is called homogeneity.",
        "Talk",
    ))
    story.append(meaning_card([
        "Homogeneity = “same dimensions on both sides / every added term”",
        "A, B, C here are any physical terms in an equation",
        "[A] means dimensional formula of A",
    ]))
    story.append(formula_card([
        "If A + B = C",
        "then [A] = [B] = [C]",
    ]))
    story.append(two_cards(
        "You can use this to",
        [
            "Check if a given formula can be correct",
            "Find the dimension of an unknown letter",
            "Find a missing power",
            "For checking: treat +, − like =",
        ],
        "Dimensions cannot",
        [
            "Find a pure number (2π, ½, 6π)",
            "Tell work apart from torque",
            "(they share the same dimensions)",
            "Build a formula from angle alone",
        ],
    ))
    story.append(tip_card(
        "Inside sin( ), cos( ), exp( ), log( ), the thing inside must have no dimension. "
        "If the inside has length or time sitting alone, the formula is wrong."
    ))

    # ===== 5 =====
    story.append(section("5. Finding a formula using dimensions"))
    story.append(para(
        "In plain English: suppose time period T depends on length l and gravity g. "
        "We guess T = k × l^a × g^b. Then we match M, L, T powers on both sides "
        "and find a and b. The constant k (like 2π) is found from experiment, not from dimensions.",
        "Talk",
    ))
    story.append(meaning_card([
        "∝ means “is proportional to” (same as × a constant)",
        "a, b = unknown powers we must find",
        "k = a pure number (dimensions cannot find it)",
        "Equate = match powers of M, then L, then T",
    ]))
    story.append(formula_card([
        "Write: Q = k A^a B^b C^c",
        "Match powers of M on both sides",
        "Match powers of L on both sides",
        "Match powers of T on both sides",
        "Solve for a, b, c",
    ]))
    story.append(para("<b>Pendulum example (step by step)</b>", "Sub"))
    story.append(formula_card([
        "Assume: T = k l^a g^b",
        "Result of matching: a = 1/2 , b = −1/2",
        "So T = k √(l/g)",
        "Experiment / theory gives k = 2π",
    ]))
    story.append(example(
        "pendulum powers",
        [
            "Left: [T] = T^1. Right: l has L, g has L T^(-2).",
            "T^1 = L^a × (L T^(-2))^b = L^(a+b) T^(-2b).",
            "Time: −2b = 1 → b = −1/2. Length: a + b = 0 → a = 1/2.",
        ],
        "T ∝ √(l/g)",
    ))

    # ===== 6 =====
    story.append(section("6. Significant figures — which digits to keep"))
    story.append(para(
        "In plain English: significant figures are the digits we trust in a measurement. "
        "We keep all sure digits, and we keep one last digit that may be a little unsure. "
        "More significant digits means a finer (more detailed) measurement.",
        "Talk",
    ))
    story.append(meaning_card([
        "Significant digit = a digit that counts in the measurement",
        "Leading zeros (0.0025) only show place value — they do not count",
        "Trailing zeros after a decimal (2.500) do count — they were measured",
        "10^3 in 2.5 × 10^3 is only a place label — not an extra digit",
    ]))
    story.append(nice_table(
        ["Rule in simple words", "Example", "How many digits?"],
        [
            ["Digits 1–9 always count", "123", "3"],
            ["Zero between digits counts", "1002", "4"],
            ["Zeros only for place do not count", "0.0025", "2"],
            ["Zeros after decimal at end count", "2.500", "4"],
            ["Power of ten is not a digit", "2.5 × 10^3", "2"],
        ],
        size=8.5,
    ))
    story.append(two_cards(
        "After you calculate",
        [
            "Add or subtract:",
            "keep the least number of decimal places",
            "Multiply or divide:",
            "keep the least number of significant figures",
        ],
        "Rounding",
        [
            "Look at the next digit",
            "If it is 0–4 → leave the keep-digit",
            "If it is 5–9 → raise the keep-digit by 1",
            "(school rule for exactly 5: raise by 1)",
        ],
    ))

    story.append(PageBreak())

    # ===== 7 =====
    story.append(section("7. Accuracy, precision and kinds of error"))
    story.append(para(
        "In plain English: accuracy means your answer is close to the true value. "
        "Precision means your repeated readings are close to each other. "
        "You can hit the same wrong spot every time — that is precise, but not accurate.",
        "Talk",
    ))
    story.append(meaning_card([
        "True value = the correct value (or the best accepted value)",
        "Error = difference between your reading and the true value",
        "Systematic = always biased the same way (wrong zero, wrong scale)",
        "Random = sometimes high, sometimes low (eye, shake)",
    ]))
    story.append(nice_table(
        ["Kind of error", "What it looks like in the lab"],
        [
            ["Systematic error", "Every reading is shifted the same way"],
            ["Random error", "Readings scatter up and down around the mean"],
            ["Least count error", "Tool cannot show smaller than its smallest mark"],
            ["Gross error", "A human mistake (wrong writing, wrong scale used)"],
        ],
        size=9,
    ))

    # ===== 8 =====
    story.append(section("8. Absolute, relative and percentage error"))
    story.append(para(
        "In plain English: absolute error says “how many units wrong”. "
        "Relative error says “what fraction is wrong”. "
        "Percentage error says “how many parts in 100 are wrong”.",
        "Talk",
    ))
    story.append(meaning_card([
        "Δx (say “delta x”) = absolute error in x",
        "measured = the value you got",
        "true = the correct / accepted value",
        "x in Δx/x is usually the true value (or the measured value if true is not known)",
        "Absolute error has the same unit as x. Relative error has no unit.",
    ]))
    story.append(formula_card([
        "Absolute error:   Δx = |measured − true|",
        "Relative error:   Δx / x",
        "Percentage error: (Δx / x) × 100",
    ]))
    story.append(formula_card([
        "How to write the final answer",
        "(value ± absolute error) unit",
        "Example: (5.0 ± 0.1) cm",
    ]))
    story.append(two_cards(
        "Many readings of one quantity",
        [
            "Step 1: find the mean (average)",
            "Step 2: each error = |reading − mean|",
            "Step 3: mean absolute error =",
            "average of those errors",
            "Write (mean ± mean absolute error)",
        ],
        "Only the tool is known",
        [
            "If the question gives no better info,",
            "take absolute error ≈ least count",
            "Metre scale → about ± 1 mm",
            "Vernier → ± its least count",
        ],
    ))
    story.append(example(
        "percentage error",
        [
            "True length = 50.0 cm. Measured = 49.0 cm.",
            "Absolute error Δx = |49.0 − 50.0| = 1.0 cm.",
            "Percentage error = (1.0 / 50.0) × 100 = 2%.",
        ],
        "2%",
    ))

    # ===== 9 =====
    story.append(section("9. How error grows when we use a formula"))
    story.append(para(
        "In plain English: if A and B both have errors, and you make Y from them, "
        "Y also has error. The rule depends on what you do with A and B.",
        "Talk",
    ))
    story.append(meaning_card([
        "Y = the final quantity you calculate",
        "ΔY = absolute error in Y",
        "ΔY/Y = relative error in Y",
        "|m| means take the power as a positive number (for maximum error)",
        "Exact numbers like ½, 2, π add no error of their own",
    ]))
    story.append(para("<b>Four separate rules (learn one at a time)</b>", "Sub"))
    story.append(formula_card([
        "If Y = A + B   or   Y = A − B",
        "then ΔY = ΔA + ΔB",
        "(add the absolute errors)",
    ]))
    story.append(formula_card([
        "If Y = A × B   or   Y = A ÷ B",
        "then ΔY/Y = ΔA/A + ΔB/B",
        "(add the relative errors)",
    ]))
    story.append(formula_card([
        "If Y = A^m × B^n",
        "then ΔY/Y = |m|(ΔA/A) + |n|(ΔB/B)",
        "(power multiplies that relative error)",
    ]))
    story.append(formula_card([
        "If Y = (exact number) × A",
        "then ΔY/Y = ΔA/A",
        "(the exact number adds no error)",
    ]))
    story.append(tip_card(
        "Memory trick: + or − → start with absolute error. "
        "× , ÷ or power → start with relative error. "
        "Then multiply relative error by Y if the question asks for absolute error."
    ))
    story.append(example(
        "error in kinetic energy",
        [
            "m = (5.0 ± 0.5) kg, v = (20.0 ± 0.4) m/s. Formula K = (1/2) m v².",
            "K = 1000 J. Relative error = 0.5/5 + 2×(0.4/20) = 0.10 + 0.04 = 0.14.",
            "Absolute error ΔK = 0.14 × 1000 = 140 J.",
        ],
        "(1000 ± 140) J",
    ))

    story.append(PageBreak())

    # ===== 10 =====
    story.append(section("10. Metre scale, vernier and screw gauge"))
    story.append(para(
        "In plain English: least count is the smallest length the tool can show. "
        "Zero error is the false reading when nothing is being measured. "
        "We always remove that false reading from the reading we took.",
        "Talk",
    ))
    story.append(meaning_card([
        "Least count (L.C.) = smallest step the instrument can show",
        "Main scale = the fixed long scale",
        "Vernier scale = the short sliding scale",
        "M.S.R. = main scale reading (mark just before vernier zero)",
        "V.S.R. = vernier scale reading (which vernier mark lines up)",
        "Pitch = distance the screw moves in one full turn",
        "Zero error = reading when jaws are closed with nothing inside",
        "Zero correction = opposite of zero error",
    ]))
    story.append(nice_table(
        ["Tool", "How to find least count", "How to take reading"],
        [
            ["Metre scale", "usually 1 mm", "reading ± 1 mm"],
            ["Vernier", "L.C. = 1 M.S.D. − 1 V.S.D.", "M.S.R. + (V.S.R. × L.C.)"],
            ["School vernier (10 marks)", "often L.C. = 0.1 mm", "same reading rule"],
            ["Screw gauge", "L.C. = Pitch / N", "linear reading + (round mark × L.C.)"],
        ],
        size=8,
    ))
    story.append(formula_card([
        "Reading taken = M.S.R. + (V.S.R. × L.C.)",
        "True reading = reading taken − zero error",
        "Zero correction = − (zero error)",
    ]))
    story.append(two_cards(
        "Positive zero error",
        [
            "Jaws closed",
            "Vernier zero is to the RIGHT of main zero",
            "Zero error is positive (+)",
            "Size = (matching vernier mark) × L.C.",
        ],
        "Negative zero error",
        [
            "Jaws closed",
            "Vernier zero is to the LEFT of main zero",
            "Zero error is negative (−)",
            "Size = −(N − matching mark) × L.C.",
            "N = total vernier marks (often 10)",
        ],
    ))
    story.append(example(
        "vernier least count",
        [
            "25 vernier marks match 24 main marks. One main gap = 1 mm.",
            "1 V.S.D. = 24/25 mm = 0.96 mm.",
            "L.C. = 1 − 0.96 = 0.04 mm.",
        ],
        "0.04 mm",
    ))
    story.append(example(
        "true reading with zero error",
        [
            "Reading taken = 8.56 cm. Zero error = +0.02 cm.",
            "True reading = 8.56 − 0.02 = 8.54 cm.",
        ],
        "8.54 cm",
    ))

    # ===== 11 =====
    story.append(section("11. Revision lines — say these aloud"))
    story.append(para(
        "Close the book. If you can say these in your own words, you understand the chapter.",
        "Talk",
    ))
    story.append(nice_table(
        ["Topic", "Say this in one breath"],
        [
            ["Quantity", "Always write number × unit."],
            ["Dimension", "Force is M L T^(-2). Energy is M L^2 T^(-2)."],
            ["Convert", "n₁ u₁ = n₂ u₂. Same amount, new number."],
            ["Check formula", "Every added term must have the same dimensions."],
            ["Derive", "Assume powers, match M, L, T, solve."],
            ["Limit", "Dimensions never give a pure number like 2π."],
            ["Sig figs", "+/− → decimal places. ×/÷ → significant figures."],
            ["Accuracy", "Close to true. Precision = readings close together."],
            ["Absolute error", "Δx = |measured − true| (same unit as x)."],
            ["Relative error", "Δx / x (no unit). ×100 gives percentage."],
            ["Error grow", "+/− add Δ. ×/÷/power add relative errors."],
            ["Vernier L.C.", "1 M.S.D. − 1 V.S.D."],
            ["Vernier reading", "M.S.R. + V.S.R. × L.C."],
            ["True reading", "reading taken − zero error."],
            ["Screw gauge", "L.C. = Pitch / N."],
        ],
        size=8.5,
    ))
    story.append(Spacer(1, 3 * mm))
    story.append(formula_card([
        "Last tip",
        "If a sentence uses a short name you forgot, open page 0 (symbol dictionary).",
        "If you cannot say a formula in one simple sentence, that page needs one more read.",
    ]))
    story.append(Spacer(1, 2 * mm))
    story.append(para(
        "From Excellence Units lectures 01–06. For revision and class notes — not a question paper.",
        "CoverMid",
    ))


    doc.build(story)
    OUT2.write_bytes(OUT.read_bytes())
    print(f"Created {OUT}")
    print(f"Copied  {OUT2}")
    print(f"Size {OUT.stat().st_size} bytes")


if __name__ == "__main__":
    build()
