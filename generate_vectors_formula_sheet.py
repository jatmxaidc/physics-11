# -*- coding: utf-8 -*-
"""Vectors Formula Notes — student-friendly printable PDF (Vector 01–05).

Same clarity + tight layout as Basic_Maths_Formula_Sheet_L01_L18.
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
OUT = BASE / "class notes" / "Vectors_Formula_Sheet_V01_V05.pdf"
OUT2 = BASE / "lecture slides" / "Vectors_Formula_Sheet_V01_V05.pdf"

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
    return mapping


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
        "Sec", fontName=FB, fontSize=12.5, textColor=NAVY,
        spaceBefore=5, spaceAfter=2, leading=16,
    ))
    s.add(ParagraphStyle(
        "Sub", fontName=FB, fontSize=10.5, textColor=TEAL,
        spaceBefore=4, spaceAfter=2, leading=13,
    ))
    s.add(ParagraphStyle(
        "Talk", fontName=F, fontSize=10, textColor=INK,
        leading=13.5, spaceAfter=2, alignment=TA_JUSTIFY,
    ))
    s.add(ParagraphStyle(
        "Formula", fontName=FB, fontSize=10.5, textColor=NAVY,
        leading=15, alignment=TA_CENTER, spaceBefore=0, spaceAfter=0,
    ))
    s.add(ParagraphStyle(
        "Cell", fontName=F, fontSize=9, textColor=INK,
        leading=12, alignment=TA_CENTER,
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
        leading=15, leftIndent=2,
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
    canvas.drawRightString(PAGE_W - MARGIN, PAGE_H - 7 * mm, "Vectors — formula notes")
    canvas.setStrokeColor(LINE)
    canvas.setLineWidth(0.5)
    canvas.line(MARGIN, 12 * mm, PAGE_W - MARGIN, 12 * mm)
    canvas.setFillColor(MUTED)
    canvas.setFont(F, 8)
    canvas.drawCentredString(
        PAGE_W / 2, 7 * mm,
        f"Page {doc.page}   ·   Vector lectures 01–05   ·   Write these in your own words too",
    )
    canvas.restoreState()


def para(text, style="Talk"):
    return Paragraph(m(text), S[style])


def _card_table(body_para, bg, border, pad=6, blank_after=1.5):
    """Tight card. No KeepTogether — large KeepTogether blocks leave empty page gaps."""
    t = Table([[body_para]], colWidths=[PAGE_W - 2 * MARGIN])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), bg),
        ("BOX", (0, 0), (-1, -1), 1.0, border),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), pad),
        ("BOTTOMPADDING", (0, 0), (-1, -1), pad),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return [t, Spacer(1, blank_after * mm)]


def formula_card(text, blank_after=1.5):
    if isinstance(text, (list, tuple)):
        text = "<br/>".join(text)
    return _card_table(para(text, "Formula"), CARD, TEAL, pad=5, blank_after=blank_after)


def tip_card(text):
    return _card_table(para(f"<b>Remember:</b> {text}", "Tip"), TIP_BG, ORANGE, pad=4, blank_after=1.5)


def meaning_card(lines):
    if isinstance(lines, str):
        lines = [lines]
    body = "<br/>".join([f"<b>What the words mean</b>"] + [f"• {ln}" for ln in lines])
    return _card_table(para(body, "Ex"), SOFT, TEAL, pad=5, blank_after=1.5)


def example(title, lines, answer):
    body = "<br/>".join(
        [f"<b>Worked example — {title}</b>"]
        + [f"• {ln}" for ln in lines]
        + [f"<b>Final answer:</b> {answer}"]
    )
    return _card_table(para(body, "Ex"), EX_BG, EX_LINE, pad=5, blank_after=2)


def section(num_title):
    return KeepTogether([
        para(num_title, "Sec"),
        HRFlowable(width="100%", thickness=1.2, color=GOLD, spaceAfter=3, spaceBefore=0),
    ])


def two_cards(left_title, left_lines, right_title, right_lines):
    gap = 3 * mm
    w = (PAGE_W - 2 * MARGIN - gap) / 2

    def side(title, lines):
        bits = [para(f"<b>{title}</b>", "Sub")] + [para(x, "CellL") for x in lines]
        rows = [[b] for b in bits]
        tb = Table(rows, colWidths=[w - 4])
        wrap = Table([[tb]], colWidths=[w])
        wrap.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), SOFT),
            ("BOX", (0, 0), (-1, -1), 0.8, LINE),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
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
    return [outer, Spacer(1, 1.5 * mm)]


def nice_table(headers, rows, size=8.5):
    """Return [table, spacer]. Use story.extend(...). Never KeepTogether wide tables."""
    n = len(headers)
    col_w = (PAGE_W - 2 * MARGIN) / n
    pad = 2 if n >= 8 else 3
    hs = ParagraphStyle(
        "HCell", fontName=FB, fontSize=size, textColor=white,
        leading=size + 2, alignment=TA_CENTER,
    )
    cs = ParagraphStyle(
        "CCell", fontName=F, fontSize=size, textColor=INK,
        leading=size + 2, alignment=TA_CENTER,
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
        ("LEFTPADDING", (0, 0), (-1, -1), pad),
        ("RIGHTPADDING", (0, 0), (-1, -1), pad),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]
    for i in range(2, len(data), 2):
        cmds.append(("BACKGROUND", (0, i), (-1, i), white))
    cmds.append(("BACKGROUND", (0, 1), (0, -1), HexColor("#E8F1F8")))
    t.setStyle(TableStyle(cmds))
    return [t, Spacer(1, 1.5 * mm)]


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
        title="Vectors Formula Notes — Excellence Class 11",
        author="Excellence Programme",
    )
    frame = Frame(MARGIN, 15 * mm, PAGE_W - 2 * MARGIN, PAGE_H - 29 * mm, id="main")
    doc.addPageTemplates([PageTemplate(id="main", frames=frame, onPage=header_footer)])

    story = []

    # ===== COVER =====
    story.append(Spacer(1, 3 * mm))
    story.append(para("EXCELLENCE", "CoverBig"))
    story.append(para("Class 11 Physics  ·  Session 2026–27", "CoverMid"))
    story.append(HRFlowable(width="50%", thickness=2, color=GOLD, spaceBefore=2, spaceAfter=4))
    story.append(para("Vectors", "CoverBig"))
    story.append(para("Complete formula notes — Lectures 01 to 05", "CoverMid"))
    story.append(para(
        "Full coverage of the Vectors block: types, components, resultant, "
        "dot product, cross product, work and torque. "
        "Every formula is on its own line. Symbols are explained before use.",
        "CoverMid",
    ))
    story.append(Spacer(1, 1 * mm))
    story.extend(formula_card([
        "How to use each page",
        "1. Read “In plain English”",
        "2. Read “What the words mean”",
        "3. Copy each formula line separately",
        "4. Try the green example",
    ]))
    story.append(para("<b>Full chapter map</b>", "Sub"))
    for line in [
        "0.  Symbol dictionary",
        "1.  Scalar vs vector and notation",
        "2.  Types of vectors",
        "3.  Magnitude, unit vector, scalar multiply",
        "4.  Rectangular components and compass angles",
        "5.  Triangle law and resultant",
        "6.  Subtraction, component add, Δv",
        "7.  Dot product, projection, angle tricks",
        "8.  Cross product and î ĵ k̂",
        "9.  Physics uses — work, torque, more",
        "10. Revision lock lines",
    ]:
        story.append(para(line, "TOC"))
    story.append(PageBreak())

    # ===== 0 =====
    story.append(section("0. Symbol dictionary — read this first"))
    story.extend(nice_table(
        ["Symbol / short name", "Full meaning in school language"],
        [
            ["A →  or bold A", "a vector (size + direction)"],
            ["|A| or A", "magnitude = length of the arrow (a number ≥ 0)"],
            ["Â", "unit vector of A (size 1, same direction)"],
            ["î , ĵ , k̂", "unit vectors along +x, +y, +z"],
            ["A<sub>x</sub>, A<sub>y</sub>, A<sub>z</sub>", "components of A along the axes"],
            ["θ", "angle between two vectors (usually 0° to 180°)"],
            ["R = A + B", "resultant = vector sum"],
            ["α", "angle that R makes with A (in triangle law)"],
            ["A · B", "dot (scalar) product — answer is a number"],
            ["A × B", "cross (vector) product — answer is a vector"],
            ["kA", "scalar multiply — scale size; reverse if k < 0"],
            ["Δv", "change in velocity = v<sub>f</sub> − v<sub>i</sub> (a vector)"],
            ["W , τ", "work (scalar), torque (vector)"],
        ],
        size=8.5,
    ))
    story.extend(tip_card(
        "If a formula uses a letter you forgot, open this page first."
    ))

    # ===== 1 =====
    story.append(section("1. Scalar vs vector and notation"))
    story.append(para(
        "In plain English: a scalar has only size. A vector has size and direction, "
        "and it must obey the triangle law when you add it. Having a “direction word” "
        "is not enough — electric current has a sense of flow but is not a vector.",
        "Talk",
    ))
    story.extend(meaning_card([
        "Magnitude = how long the arrow is",
        "Direction = which way the arrow points",
        "Triangle law = head-to-tail joining when you add vectors",
        "Negative of A (−A) = same size, opposite direction",
    ]))
    story.extend(formula_card([
        "Vector = magnitude × unit vector in its direction",
        "A → = |A| Â",
    ]))
    story.extend(two_cards(
        "Scalars (size only)",
        [
            "Mass, time, temperature, energy",
            "Speed (how fast, no direction)",
            "Work and power are scalars",
            "Current is not a vector",
        ],
        "Vectors (size + direction)",
        [
            "Displacement, velocity, acceleration",
            "Force, momentum, torque",
            "Electric field, magnetic field",
            "Reverse direction → reverse the arrow",
        ],
    ))
    story.extend(tip_card(
        "î, ĵ, k̂ have size 1. They only carry direction along the axes."
    ))

    # ===== 2 =====
    story.append(section("2. Types of vectors — say these in one line"))
    story.extend(nice_table(
        ["Name", "Meaning in plain words"],
        [
            ["Equal vectors", "Same kind, same size, same direction"],
            ["Parallel", "Same direction (sizes may differ)"],
            ["Anti-parallel", "Opposite directions"],
            ["Negative of A", "Same size, opposite direction: −A"],
            ["Unit vector", "Size = 1, keeps the direction"],
            ["Null (zero)", "Size = 0; no useful direction"],
            ["Collinear", "On the same line (or parallel lines)"],
            ["Coplanar", "Lie in the same plane"],
            ["Orthogonal", "Perpendicular (angle = 90°)"],
            ["Concurrent", "Lines of action meet at one point"],
            ["Axial", "Along an axis (τ, ω, L) — sense by RHR"],
        ],
        size=8.5,
    ))
    story.extend(formula_card([
        "Angle between two vectors = smaller angle between directions (0° to 180°)",
        "Put both arrows tail-to-tail (or tip-to-tip) before you measure θ",
    ]))

    # ===== 3 =====
    story.append(section("3. Magnitude, unit vector, scalar multiply"))
    story.append(para(
        "In plain English: magnitude is the length. Unit vector keeps only direction. "
        "Multiplying by a number k stretches or shrinks; negative k flips the arrow.",
        "Talk",
    ))
    story.extend(meaning_card([
        "|A| = length of A (never negative)",
        "Â has no physical unit — only direction",
        "k > 0 → same direction; k < 0 → opposite direction",
    ]))
    story.extend(formula_card([
        "|A| = √(A<sub>x</sub>^2 + A<sub>y</sub>^2 + A<sub>z</sub>^2)",
        "Â = A / |A|",
        "Â = (A<sub>x</sub>/|A|) î + (A<sub>y</sub>/|A|) ĵ + (A<sub>z</sub>/|A|) k̂",
        "kA has magnitude |k| |A|",
    ]))
    story.extend(example(
        "unit vector",
        [
            "A = 3 î + 4 ĵ. Magnitude |A| = √(9 + 16) = 5.",
            "Unit vector Â = (3/5) î + (4/5) ĵ.",
        ],
        "Â = (3/5) î + (4/5) ĵ",
    ))

    # ===== 4 =====
    story.append(section("4. Rectangular components and compass angles"))
    story.append(para(
        "In plain English: resolve means break a vector into axis pieces. "
        "θ is usually measured from +x. Watch the quadrant for signs. "
        "Compass wording: the last word is the reference line you start from.",
        "Talk",
    ))
    story.extend(meaning_card([
        "A<sub>x</sub> = piece along x",
        "A<sub>y</sub> = piece along y",
        "“30° east of north” → start from north, turn 30° toward east",
        "If θ is measured from +y, cos goes with y and sin with x",
    ]))
    story.extend(formula_card([
        "A<sub>x</sub> = A cos θ",
        "A<sub>y</sub> = A sin θ",
        "A = A<sub>x</sub> î + A<sub>y</sub> ĵ",
        "A = √(A<sub>x</sub>^2 + A<sub>y</sub>^2)",
        "tan θ = A<sub>y</sub> / A<sub>x</sub>   (choose the correct quadrant)",
    ]))
    story.extend(two_cards(
        "Handy physics angles",
        [
            "sin 37° ≈ 3/5 ,  cos 37° ≈ 4/5",
            "sin 53° ≈ 4/5 ,  cos 53° ≈ 3/5",
            "sin 30° = 1/2 ,  cos 30° = √3/2",
            "sin 45° = cos 45° = 1/√2",
        ],
        "Sign by quadrant",
        [
            "1st: A<sub>x</sub> > 0 , A<sub>y</sub> > 0",
            "2nd: A<sub>x</sub> < 0 , A<sub>y</sub> > 0",
            "3rd: A<sub>x</sub> < 0 , A<sub>y</sub> < 0",
            "4th: A<sub>x</sub> > 0 , A<sub>y</sub> < 0",
        ],
    ))
    story.extend(example(
        "resolve a force",
        [
            "F = 50 N at 37° above +x.",
            "F<sub>x</sub> = 50 × (4/5) = 40 N ,  F<sub>y</sub> = 50 × (3/5) = 30 N.",
        ],
        "F = 40 î + 30 ĵ  (newton)",
    ))

    # ===== 5 =====
    story.append(section("5. Triangle law and the resultant"))
    story.append(para(
        "In plain English: put the tail of B at the tip of A. "
        "The arrow from the free start of A to the free end of B is R = A + B. "
        "Drawing order does not change R. A + B = B + A.",
        "Talk",
    ))
    story.extend(meaning_card([
        "θ = angle between A and B",
        "α = angle that R makes with A",
        "Maximum R when same direction; minimum when opposite",
    ]))
    story.extend(formula_card([
        "R = √(A^2 + B^2 + 2 A B cos θ)",
        "tan α = (B sin θ) / (A + B cos θ)",
        "|A − B| ≤ R ≤ A + B",
        "|A + B| ≤ |A| + |B|   (never bigger than the sum of sizes)",
    ]))
    story.extend(two_cards(
        "Special cases",
        [
            "θ = 0°: R = A + B (maximum)",
            "θ = 180°: R = |A − B| (minimum)",
            "θ = 90°: R = √(A^2 + B^2)",
            "If |A| = |B| = A: R = 2 A cos(θ/2)",
            "If |A| = |B| = A: |A − B| = 2 A sin(θ/2)",
        ],
        "Equal vectors tip",
        [
            "When |A| = |B|, R bisects the angle",
            "Three forces close a triangle if each",
            "size lies between |diff| and sum of others",
            "Polygon law: many vectors head-to-tail;",
            "resultant joins first start to last tip",
        ],
    ))
    story.extend(tip_card(
        "Cos θ is negative when θ is between 90° and 180°. That makes R smaller — which matches the picture."
    ))

    # ===== 6 =====
    story.append(section("6. Subtraction, component add, Δv"))
    story.append(para(
        "In plain English: subtraction is A + (−B). Flip B, then add. "
        "The safest everyday method: add x with x, y with y, z with z.",
        "Talk",
    ))
    story.extend(formula_card([
        "A − B = A + (−B)",
        "|A − B| = √(A^2 + B^2 − 2 A B cos θ)",
        "R<sub>x</sub> = A<sub>x</sub> + B<sub>x</sub>",
        "R<sub>y</sub> = A<sub>y</sub> + B<sub>y</sub>",
        "R<sub>z</sub> = A<sub>z</sub> + B<sub>z</sub>",
        "R = R<sub>x</sub> î + R<sub>y</sub> ĵ + R<sub>z</sub> k̂",
        "Δv = v<sub>f</sub> − v<sub>i</sub>",
    ]))
    story.extend(tip_card(
        "Rotation example: 90° anti-clockwise in xy-plane sends î → ĵ and ĵ → −î."
    ))
    story.extend(example(
        "add by columns",
        [
            "A = 2 î + 3 ĵ ,  B = − î + 4 ĵ.",
            "R<sub>x</sub> = 2 − 1 = 1 ,  R<sub>y</sub> = 3 + 4 = 7.",
            "R = î + 7 ĵ ,  |R| = √(1 + 49) = √50 = 5√2.",
        ],
        "R = î + 7 ĵ",
    ))

    # ===== 7 =====
    story.append(section("7. Dot product, projection, angle tricks"))
    story.append(para(
        "In plain English: the dot product mixes two vectors and gives a number. "
        "It is large when they point the same way, zero when perpendicular, "
        "and negative when the angle is obtuse.",
        "Talk",
    ))
    story.extend(meaning_card([
        "A · B is a scalar (a number, with units of product of the two)",
        "Projection of B along A = (A · B) / |A|",
        "A · B = B · A (order does not matter)",
    ]))
    story.extend(formula_card([
        "A · B = |A| |B| cos θ",
        "A · B = A<sub>x</sub>B<sub>x</sub> + A<sub>y</sub>B<sub>y</sub> + A<sub>z</sub>B<sub>z</sub>",
        "cos θ = (A · B) / (|A| |B|)",
        "A · A = |A|^2",
        "Projection of B on A = (A · B) / |A|",
    ]))
    story.extend(nice_table(
        ["Situation", "What you get"],
        [
            ["θ = 0° (same way)", "A · B = |A| |B|  (maximum)"],
            ["θ acute", "A · B > 0"],
            ["θ = 90°", "A · B = 0"],
            ["θ obtuse", "A · B < 0"],
            ["θ = 180°", "A · B = −|A| |B|"],
            ["Unit axes", "î · î = 1 ;  î · ĵ = 0"],
        ],
        size=8.5,
    ))
    story.append(para("<b>Angle-condition shortcuts (equal sizes often)</b>", "Sub"))
    story.extend(nice_table(
        ["Given", "Angle between A and B"],
        [
            ["|A| = |B| = |A + B|", "120°"],
            ["|A| = |B| = |A − B|", "60°"],
            ["|A + B| = |A − B|", "90°"],
            ["|A + B| = |A| + |B|", "0°"],
            ["|A − B| = |A| + |B|", "180°"],
        ],
        size=8.5,
    ))
    story.extend(example(
        "angle between two vectors",
        [
            "A = î + ĵ ,  B = î − ĵ.",
            "A · B = 1 − 1 = 0 , so θ = 90°.",
        ],
        "θ = 90°",
    ))

    # ===== 8 =====
    story.append(section("8. Cross product and î ĵ k̂"))
    story.append(para(
        "In plain English: the cross product gives another vector. "
        "Size uses sin θ. Direction is perpendicular to both (right-hand rule). "
        "Flip the order and the arrow flips.",
        "Talk",
    ))
    story.extend(meaning_card([
        "|A × B| = area of the parallelogram made by A and B",
        "Maximum size when θ = 90°: |A × B| = |A| |B|",
        "Parallel or anti-parallel → cross product is zero",
    ]))
    story.extend(formula_card([
        "|A × B| = |A| |B| sin θ",
        "A × B = − (B × A)",
        "Unit vector along A × B = (A × B) / |A × B|",
        "In xy-plane: A × B = (A<sub>x</sub>B<sub>y</sub> − A<sub>y</sub>B<sub>x</sub>) k̂",
    ]))
    story.extend(formula_card([
        "Full 3D form",
        "A × B = (A<sub>y</sub>B<sub>z</sub> − A<sub>z</sub>B<sub>y</sub>) î",
        "+ (A<sub>z</sub>B<sub>x</sub> − A<sub>x</sub>B<sub>z</sub>) ĵ",
        "+ (A<sub>x</sub>B<sub>y</sub> − A<sub>y</sub>B<sub>x</sub>) k̂",
    ]))
    story.extend(two_cards(
        "Unit-vector cycle (must know)",
        [
            "î × ĵ = k̂",
            "ĵ × k̂ = î",
            "k̂ × î = ĵ",
            "Flip order → flip sign: ĵ × î = −k̂",
            "Same vector: î × î = 0",
        ],
        "Dot vs cross — one glance",
        [
            "Dot → a number (scalar)",
            "Cross → a vector",
            "Dot uses cos θ",
            "Cross uses sin θ",
            "Parallel → cross is zero",
            "Perpendicular → dot is zero",
        ],
    ))
    story.extend(tip_card(
        "Right-hand rule: fingers from A toward B (smaller angle). Thumb shows A × B."
    ))
    story.extend(example(
        "simple cross product",
        [
            "A = 2 î ,  B = 3 ĵ.",
            "A × B = 2 î × 3 ĵ = 6 (î × ĵ) = 6 k̂.",
        ],
        "6 k̂",
    ))

    # ===== 9 =====
    story.append(section("9. Physics uses — work, torque, more"))
    story.append(para(
        "In plain English: work uses the dot product. Torque uses the cross product. "
        "Same idea appears in magnetic force and angular momentum.",
        "Talk",
    ))
    story.extend(two_cards(
        "Work (dot product)",
        [
            "W = F · s = F s cos θ",
            "θ = angle between F and displacement",
            "If F ⟂ s , work = 0",
            "Unit: joule (J)",
        ],
        "Torque (cross product)",
        [
            "τ = r × F",
            "|τ| = r F sin θ",
            "θ = angle between r and F",
            "Direction: perpendicular to the plane",
        ],
    ))
    story.extend(formula_card([
        "Magnetic force: F = q (v × B)",
        "Angular momentum: L = r × p",
    ]))

    # ===== 10 =====
    story.append(section("10. Revision lock lines — say these aloud"))
    story.extend(nice_table(
        ["Topic", "Say this in one breath"],
        [
            ["Write a vector", "A = |A| Â ; also A<sub>x</sub> î + A<sub>y</sub> ĵ + A<sub>z</sub> k̂"],
            ["Magnitude", "|A| = √(A<sub>x</sub>^2 + A<sub>y</sub>^2 + A<sub>z</sub>^2)"],
            ["Unit / scale", "Â = A/|A| ; kA flips if k is negative"],
            ["Resolve", "A<sub>x</sub> = A cos θ , A<sub>y</sub> = A sin θ"],
            ["Resultant", "R = √(A^2 + B^2 + 2AB cos θ) ; |A−B| ≤ R ≤ A+B"],
            ["Equal sizes", "R = 2A cos(θ/2) ; |A−B| = 2A sin(θ/2)"],
            ["Add safely", "Add x with x, y with y, z with z"],
            ["Dot", "A·B = AB cos θ = sum of component products"],
            ["Projection", "(A·B)/|A| is the piece of B along A"],
            ["Perpendicular", "A·B = 0 (neither zero)"],
            ["Cross", "|A×B| = AB sin θ ; A×B = −B×A ; RHR"],
            ["î ĵ k̂", "î×ĵ=k̂ , ĵ×k̂=î , k̂×î=ĵ"],
            ["Work / torque", "W = F·s ; τ = r×F"],
        ],
        size=8.5,
    ))
    story.append(Spacer(1, 1 * mm))
    story.extend(formula_card([
        "Last tip",
        "Draw a small arrow sketch before you calculate.",
        "Ask: do I need cos (dot / components) or sin (cross / perpendicular piece)?",
        "If you cannot say a formula in one spoken sentence, that page needs one more read.",
    ]))
    story.append(para(
        "Source: Excellence Vector lectures 01–05. "
        "For revision and class notes — not a question paper.",
        "CoverMid",
    ))

    doc.build(story)
    OUT2.write_bytes(OUT.read_bytes())
    print(f"Created {OUT}")
    print(f"Copied  {OUT2}")
    print(f"Size {OUT.stat().st_size} bytes")


if __name__ == "__main__":
    build()
