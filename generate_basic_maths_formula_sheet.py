# -*- coding: utf-8 -*-
"""Basic Maths Formula Notes — student-friendly printable PDF (L01–L18).

Plain language, clear layout, proper superscripts. Not a machine dump.
"""
from __future__ import annotations

import re
from pathlib import Path

import matplotlib
from reportlab.lib.colors import HexColor, white, black, Color
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
    ListFlowable,
    ListItem,
)

BASE = Path(__file__).resolve().parent
OUT = BASE / "class notes" / "Basic_Maths_Formula_Sheet_L01_L18.pdf"
OUT2 = BASE / "lecture slides" / "Basic_Maths_Formula_Sheet_L01_L18.pdf"

# Warm, readable palette (not dark-mode machine UI)
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
GREEN = HexColor("#166534")
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
        "HeadCell", fontName=FB, fontSize=9, textColor=white,
        leading=12, alignment=TA_CENTER,
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
    s.add(ParagraphStyle(
        "PlainBullet", fontName=F, fontSize=10, textColor=INK,
        leading=14, leftIndent=4, spaceAfter=2,
    ))
    return s


S = styles()


def header_footer(canvas, doc):
    canvas.saveState()
    # soft top bar
    canvas.setFillColor(NAVY)
    canvas.roundRect(0, PAGE_H - 11 * mm, PAGE_W, 11 * mm, 0, fill=1, stroke=0)
    canvas.setFillColor(GOLD)
    canvas.rect(0, PAGE_H - 11.8 * mm, PAGE_W, 1.2, fill=1, stroke=0)
    canvas.setFillColor(white)
    canvas.setFont(FB, 8.5)
    canvas.drawString(MARGIN, PAGE_H - 7 * mm, "Excellence  ·  Class 11 Physics")
    canvas.setFont(F, 8)
    canvas.drawRightString(PAGE_W - MARGIN, PAGE_H - 7 * mm, "Basic Maths — formula notes")
    # footer
    canvas.setStrokeColor(LINE)
    canvas.setLineWidth(0.5)
    canvas.line(MARGIN, 12 * mm, PAGE_W - MARGIN, 12 * mm)
    canvas.setFillColor(MUTED)
    canvas.setFont(F, 8)
    canvas.drawCentredString(
        PAGE_W / 2, 7 * mm,
        f"Page {doc.page}   ·   Lectures 01–18   ·   Write these in your own words too",
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
    """One formula per line. Pass a list for separate formulas (never join with ·)."""
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
    """Return [table, spacer]. Use story.extend(...). Do not KeepTogether wide tables
    (they can exceed one frame and vanish from the PDF)."""
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
        title="Basic Maths Formula Notes — Excellence Class 11",
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
    story.append(para("Basic Maths & Calculus", "CoverBig"))
    story.append(para("Complete formula notes — Lectures 01 to 18", "CoverMid"))
    story.append(para(
        "Full coverage of the Basic Maths block: trigonometry tables, identities, ASTC, "
        "compound angles, AP–GP, graphs, differentiation and integration. "
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
        "1.  Trig ratios and height–distance",
        "2.  Full trig tables (0°–90° and 90°–180°)",
        "3.  Identities, range, brother / allied angles",
        "4.  ASTC, radian, arc length",
        "5.  Compound, double, half and small angles",
        "6.  Max–min of a sin + b cos (phasor)",
        "7.  AP and GP (with physics uses)",
        "8.  Coordinates, line, graph shapes, circle, ellipse",
        "9.  Differentiation — full rule set",
        "10. Applications: motion, related rates, max–min",
        "11. Integration — indefinite, definite, area, average",
        "12. Revision lock lines",
    ]:
        story.append(para(line, "TOC"))
    story.append(PageBreak())

    # ===== 0 =====
    story.append(section("0. Symbol dictionary — read this first"))
    story.extend(nice_table(
        ["Symbol / short name", "Full meaning in school language"],
        [
            ["θ (theta)", "an angle (degree or radian — check the question)"],
            ["P, B, H", "opposite side, adjacent side, hypotenuse"],
            ["sin, cos, tan", "sine, cosine, tangent"],
            ["cosec, sec, cot", "1/sin, 1/cos, 1/tan"],
            ["ND", "not defined (division by zero)"],
            ["rad", "radian — angle unit used in calculus and arc = Rθ"],
            ["π (pi)", "≈ 3.14 ; also π radian = 180°"],
            ["s, R", "arc length, radius"],
            ["a, d, n (AP)", "first term, common difference, number of terms"],
            ["a, r, n (GP)", "first term, common ratio, number of terms"],
            ["m, c", "slope of a line, y-intercept"],
            ["dy/dx , y' , f'(x)", "derivative = rate of change / tangent slope"],
            ["d²y/dx²", "second derivative (used for max / min test)"],
            ["∫ f(x) dx", "indefinite integral (anti-derivative) + C"],
            ["∫_a^b f(x) dx", "definite integral = F(b) − F(a) (no +C)"],
            ["ω, φ", "angular frequency, phase angle"],
        ],
        size=8.5,
    ))
    story.extend(tip_card(
        "If a formula uses a letter you forgot, open this page first."
    ))

    # ===== 1 =====
    story.append(section("1. Trig ratios and height–distance"))
    story.append(para(
        "In plain English: pick an acute angle θ in a right triangle. "
        "Name opposite (P), adjacent (B) and hypotenuse (H). Six ratios are fractions from these.",
        "Talk",
    ))
    story.extend(meaning_card([
        "Opposite (P) = side facing θ",
        "Adjacent (B) = side next to θ (not hypotenuse)",
        "Hypotenuse (H) = longest side",
        "Angle of elevation = looking up from horizontal",
        "Angle of depression = looking down from horizontal",
    ]))
    story.extend(nice_table(
        ["Name", "Write it as", "Also equal to"],
        [
            ["sine (sin θ)", "P / H", "1 / cosec θ"],
            ["cosine (cos θ)", "B / H", "1 / sec θ"],
            ["tangent (tan θ)", "P / B", "sin θ / cos θ"],
            ["cosecant (cosec θ)", "H / P", "1 / sin θ"],
            ["secant (sec θ)", "H / B", "1 / cos θ"],
            ["cotangent (cot θ)", "B / P", "1 / tan θ"],
        ],
        size=9,
    ))
    story.extend(formula_card([
        "sin θ × cosec θ = 1",
        "cos θ × sec θ = 1",
        "tan θ × cot θ = 1",
    ]))
    story.extend(formula_card([
        "Height and distance (right triangle)",
        "height = (distance along ground) × tan(elevation)",
        "If elevation from A is α and from B is β (same tower),",
        "use two tan equations and solve for height",
    ]))
    story.extend(tip_card(
        "Tan is never P/H. Tan is always P/B. Elevation and depression are measured from the horizontal."
    ))
    story.extend(example(
        "find other ratios",
        [
            "cos θ = 8/17 (acute). Adjacent = 8, hypotenuse = 17.",
            "Opposite = √(17² − 8²) = 15.",
            "sin θ = 15/17 , tan θ = 15/8.",
        ],
        "sin θ = 15/17 , tan θ = 15/8",
    ))

    # ===== 2 TABLES =====
    story.append(section("2. Full trig tables — do not skip this page"))
    story.append(para(
        "In plain English: learn the pattern. ND means not defined. "
        "The table is split into two parts so every cell prints clearly on the page.",
        "Talk",
    ))
    story.append(para("<b>Table A — 0° to 90°</b>", "Sub"))
    story.extend(nice_table(
        ["θ", "0°", "30°", "45°", "60°", "90°"],
        [
            ["sin", "0", "1/2", "1/√2", "√3/2", "1"],
            ["cos", "1", "√3/2", "1/√2", "1/2", "0"],
            ["tan", "0", "1/√3", "1", "√3", "ND"],
            ["cosec", "ND", "2", "√2", "2/√3", "1"],
            ["sec", "1", "2/√3", "√2", "2", "ND"],
            ["cot", "ND", "√3", "1", "1/√3", "0"],
        ],
        size=9,
    ))
    story.append(para("<b>Table B — 90° to 180°</b>", "Sub"))
    story.extend(nice_table(
        ["θ", "90°", "120°", "135°", "150°", "180°"],
        [
            ["sin", "1", "√3/2", "1/√2", "1/2", "0"],
            ["cos", "0", "−1/2", "−1/√2", "−√3/2", "−1"],
            ["tan", "ND", "−√3", "−1", "−1/√3", "0"],
            ["cosec", "1", "2/√3", "√2", "2", "ND"],
            ["sec", "ND", "−2", "−√2", "−2/√3", "−1"],
            ["cot", "0", "−1/√3", "−1", "−√3", "ND"],
        ],
        size=9,
    ))
    story.append(para(
        "Brother angles: 120° ↔ 60°, 135° ↔ 45°, 150° ↔ 30° (from 180°). "
        "Physics favourites: sin 37° ≈ 3/5, cos 37° ≈ 4/5; sin 53° ≈ 4/5, cos 53° ≈ 3/5.",
        "Talk",
    ))
    story.extend(formula_card([
        "Beyond 180°: use ASTC for the sign, then reduce to an acute angle",
        "Example idea: sin(210°) = sin(180°+30°) = −sin 30° = −1/2",
    ]))
    story.extend(example(
        "read the table",
        ["sin 150° = 1/2 , sec 120° = −2."],
        "sin 150° = 1/2 , sec 120° = −2",
    ))


    # ===== 3 =====
    story.append(section("3. Identities, range, brother / allied angles"))
    story.append(para(
        "In plain English: identities are always true (where defined). "
        "Start from sin²θ + cos²θ = 1. The other two are the same line divided by cos²θ or sin²θ.",
        "Talk",
    ))
    story.extend(meaning_card([
        "sin²θ means (sin θ)×(sin θ), not sin(sin θ)",
        "Brother / allied angle = rewrite using 90°±θ, 180°±θ, 270°±θ, 360°±θ",
        "At 90° and 270°, sin ↔ cos (function may change). At 180° and 360°, function stays, sign may change",
    ]))
    story.extend(formula_card([
        "sin^2θ + cos^2θ = 1",
        "1 + tan^2θ = sec^2θ",
        "1 + cot^2θ = cosec^2θ",
    ]))
    story.extend(formula_card([
        "Range: −1 ≤ sin θ ≤ 1",
        "Range: −1 ≤ cos θ ≤ 1",
        "tan θ can be any real number (where defined)",
    ]))
    story.extend(two_cards(
        "Useful brother lines",
        [
            "sin(90° − θ) = cos θ",
            "cos(90° − θ) = sin θ",
            "sin(180° − θ) = sin θ",
            "cos(180° − θ) = −cos θ",
            "sin(180° + θ) = −sin θ",
            "cos(180° + θ) = −cos θ",
        ],
        "More brother lines",
        [
            "sin(360° − θ) = −sin θ",
            "cos(360° − θ) = cos θ",
            "sin(−θ) = −sin θ  (odd)",
            "cos(−θ) = cos θ  (even)",
            "tan(−θ) = −tan θ  (odd)",
            "tan(180° + θ) = tan θ",
        ],
    ))

    # ===== 4 =====
    story.append(section("4. ASTC, radian and arc length"))
    story.append(para(
        "In plain English: ASTC tells which ratios stay positive in each quadrant "
        "(anti-clockwise from the first). Radian is another angle unit; physics needs it for arc and calculus.",
        "Talk",
    ))
    story.extend(meaning_card([
        "1st quadrant (0°–90°): All positive (A)",
        "2nd (90°–180°): only Sin / cosec positive (S)",
        "3rd (180°–270°): only Tan / cot positive (T)",
        "4th (270°–360°): only Cos / sec positive (C)",
        "s = arc length, R = radius, θ in radian for s = Rθ",
    ]))
    story.extend(formula_card([
        "ASTC (anti-clockwise from 1st quadrant)",
        "A = All positive",
        "S = only Sin positive",
        "T = only Tan positive",
        "C = only Cos positive",
    ]))
    story.extend(formula_card([
        "π radian = 180°",
        "degree → radian:  multiply by π/180",
        "radian → degree:  multiply by 180/π",
        "Arc length: s = R × θ   (θ in radian)",
    ]))
    story.extend(tip_card(
        "For differentiation / integration of sin, cos, tan — angle must be in radian. "
        "Example: π/3 rad = 60°."
    ))
    story.extend(example(
        "radian to degree",
        ["(5π/6) × (180°/π) = 150°."],
        "150°",
    ))

    # ===== 5 =====
    story.append(section("5. Compound, double, half and small angles"))
    story.append(para(
        "In plain English: expand sin/cos/tan of a sum or difference. "
        "Double angle is the same with both angles equal. Half angle and small-angle shortcuts are below.",
        "Talk",
    ))
    story.extend(two_cards(
        "Sine and cosine of A ± B",
        [
            "sin(A+B) = sin A cos B + cos A sin B",
            "sin(A−B) = sin A cos B − cos A sin B",
            "cos(A+B) = cos A cos B − sin A sin B",
            "cos(A−B) = cos A cos B + sin A sin B",
        ],
        "Tangent of A ± B",
        [
            "tan(A+B) = (tan A + tan B)/(1 − tan A tan B)",
            "tan(A−B) = (tan A − tan B)/(1 + tan A tan B)",
            "If tan A × tan B = 1, tan(A+B) is ND",
        ],
    ))
    story.extend(formula_card([
        "sin 2θ = 2 sin θ cos θ",
        "cos 2θ = cos^2θ − sin^2θ",
        "cos 2θ = 2 cos^2θ − 1",
        "cos 2θ = 1 − 2 sin^2θ",
        "tan 2θ = 2 tan θ / (1 − tan^2θ)",
    ]))
    story.extend(formula_card([
        "sin^2θ = (1 − cos 2θ)/2",
        "cos^2θ = (1 + cos 2θ)/2",
        "sin(θ/2) and cos(θ/2) use half-angle forms from these",
    ]))
    story.extend(formula_card([
        "Small angle (θ in radian, about 5° or less)",
        "sin θ ≈ θ",
        "tan θ ≈ θ",
        "cos θ ≈ 1",
        "Never put degrees inside this shortcut",
    ]))
    story.extend(example(
        "compound angle",
        [
            "sin 75° = sin(45°+30°) = sin45 cos30 + cos45 sin30",
            "= (√6 + √2)/4.",
        ],
        "(√6 + √2)/4",
    ))


    # ===== 6 =====
    story.append(section("6. Max–min of a sin θ + b cos θ (phasor)"))
    story.append(para(
        "In plain English: pack a sine and a cosine into one sine with amplitude R. "
        "Used in SHM, waves and AC.",
        "Talk",
    ))
    story.extend(meaning_card([
        "a, b = fixed numbers",
        "R = √(a² + b²) = biggest size of the sum",
        "φ = phase; tan φ = b/a (choose correct quadrant)",
    ]))
    story.extend(formula_card([
        "y = a sin θ + b cos θ",
        "Maximum of y = √(a^2 + b^2)",
        "Minimum of y = −√(a^2 + b^2)",
    ]))
    story.extend(formula_card([
        "a sin ωt + b cos ωt = R sin(ωt + φ)",
        "R = √(a^2 + b^2)",
        "tan φ = b / a",
    ]))

    # ===== 7 =====
    story.append(section("7. AP and GP"))
    story.append(para(
        "In plain English: AP = equal adding. GP = equal multiplying.",
        "Talk",
    ))
    story.extend(meaning_card([
        "AP: a = first term, d = common difference, n = number of terms",
        "GP: a = first term, r = common ratio",
        "Tₙ = n-th term, Sₙ = sum of n terms, S∞ = endless sum (|r|<1 only)",
    ]))
    story.extend(two_cards(
        "Arithmetic progression (AP)",
        [
            "Tₙ = a + (n−1)d",
            "Sₙ = n/2 [2a + (n−1)d]",
            "Sₙ = n/2 (first + last)",
            "1+2+…+n = n(n+1)/2",
        ],
        "Geometric progression (GP)",
        [
            "Tₙ = a r^(n−1)",
            "Sₙ = a(r^n − 1)/(r − 1)  (r≠1)",
            "S∞ = a/(1−r)  only if |r|<1",
            "Physics: repeating resistors, decaying series",
        ],
    ))
    story.extend(tip_card(
        "Do not use S∞ when |r| ≥ 1. The endless sum does not settle."
    ))

    # ===== 8 =====
    story.append(section("8. Coordinates, line, graph shapes, circle, ellipse"))
    story.append(para(
        "In plain English: graphs are stories on paper. Always label quantity and unit on each axis.",
        "Talk",
    ))
    story.extend(formula_card([
        "Distance between (x₁,y₁) and (x₂,y₂)",
        "√[(x₂ − x₁)^2 + (y₂ − y₁)^2]",
    ]))
    story.extend(formula_card([
        "Slope m = (y₂ − y₁)/(x₂ − x₁) = Δy/Δx = tan θ",
        "Straight line: y = mx + c",
        "Parallel lines: same m",
        "Perpendicular lines: m₁ m₂ = −1",
    ]))
    story.extend(two_cards(
        "Shapes to recognise",
        [
            "Modulus: y = |x|  (V-shape)",
            "Hyperbola: y = k/x  or xy = k",
            "Parabola: y = ax^2 + bx + c",
            "Square root: y = √x",
            "Exponential: y = e^x or e^(−x)",
        ],
        "Physics graph links",
        [
            "F = −kx  → straight line through origin (3rd/1st)",
            "K ∝ v^2  → parabola in v",
            "T ∝ √L  → square-root shape",
            "Ideal gas: pV = constant → hyperbola",
            "v–t area = displacement",
        ],
    ))
    story.extend(two_cards(
        "Circle",
        [
            "x^2 + y^2 = R^2  (centre origin)",
            "(x−h)^2 + (y−k)^2 = R^2",
        ],
        "Ellipse / SHM",
        [
            "x^2/a^2 + y^2/b^2 = 1",
            "SHM: |v| = ω √(A^2 − x^2)",
        ],
    ))
    story.extend(formula_card([
        "Quadratic: a x^2 + b x + c = 0",
        "x = (−b ± √(b^2 − 4ac)) / (2a)",
        "Discriminant D = b^2 − 4ac",
        "D>0 two real roots; D=0 one; D<0 no real root",
    ]))


    # ===== 9 =====
    story.append(section("9. Differentiation — full rule set"))
    story.append(para(
        "In plain English: differentiation finds how fast y changes when x changes a tiny bit. "
        "On a graph it is the tangent slope.",
        "Talk",
    ))
    story.extend(meaning_card([
        "dy/dx = derivative of y with respect to x",
        "Also written y' or f'(x)",
        "Trig rules need the angle in radian",
    ]))
    story.extend(formula_card([
        "dy/dx = limit of (Δy/Δx) as Δx → 0",
    ]))
    story.extend(nice_table(
        ["If y is…", "Then dy/dx is…", "In plain words"],
        [
            ["constant C", "0", "Flat — no change"],
            ["x^n", "n x^(n−1)", "Power rule"],
            ["k × f(x)", "k × f'(x)", "Constant walks out"],
            ["A + B", "A' + B'", "Term by term"],
            ["A − B", "A' − B'", "Term by term"],
            ["e^x", "e^x", "Stays the same"],
            ["ln|x|", "1/x", "Natural log"],
            ["sin x", "cos x", "x in radian"],
            ["cos x", "−sin x", "x in radian"],
            ["tan x", "sec^2 x", "Where defined"],
            ["sec x", "sec x tan x", ""],
            ["cosec x", "−cosec x cot x", ""],
            ["cot x", "−cosec^2 x", ""],
        ],
        size=8.5,
    ))
    story.extend(formula_card([
        "Product rule: (AB)' = A B' + B A'",
        "Quotient rule: (A/B)' = (B A' − A B') / B^2",
        "Chain rule: dy/dx = (dy/du)×(du/dx)",
        "If y = e^u then dy/dx = e^u × (du/dx)",
        "If y = sin u then dy/dx = cos u × (du/dx)",
    ]))
    story.extend(formula_card([
        "Parametric: if x=x(t), y=y(t)",
        "dy/dx = (dy/dt) / (dx/dt)",
        "Partial derivative: vary one variable, hold others fixed",
        "Example: ∂V/∂x means differentiate V w.r.t. x only",
    ]))

    # ===== 10 =====
    story.append(section("10. Applications — motion, related rates, max–min"))
    story.append(para(
        "In plain English: the same derivative idea appears in motion and in “how fast is this changing”.",
        "Talk",
    ))
    story.extend(formula_card([
        "v = dx/dt",
        "a = dv/dt = d^2x/dt^2",
        "Related rates: differentiate both sides with respect to time",
    ]))
    story.extend(formula_card([
        "Stationary point: dy/dx = 0",
        "If d^2y/dx^2 > 0 → local minimum (valley)",
        "If d^2y/dx^2 < 0 → local maximum (hill)",
        "If d^2y/dx^2 = 0 → test further (may be inflection)",
    ]))
    story.extend(example(
        "related rates idea",
        [
            "Circle area A = π r^2. Differentiate w.r.t. time t:",
            "dA/dt = 2 π r (dr/dt).",
        ],
        "dA/dt = 2πr (dr/dt)",
    ))


    # ===== 11 =====
    story.append(section("11. Integration — indefinite, definite, area, average"))
    story.append(para(
        "In plain English: integration undoes differentiation. "
        "Indefinite integral has +C. Definite integral has limits and gives a number (no +C).",
        "Talk",
    ))
    story.extend(meaning_card([
        "F(x) = anti-derivative of f (so F' = f)",
        "C = constant of integration",
        "a = lower limit, b = upper limit",
    ]))
    story.extend(formula_card([
        "∫ f(x) dx = F(x) + C",
        "∫ x^n dx = x^(n+1)/(n+1) + C   (n ≠ −1)",
        "∫_a^b f(x) dx = F(b) − F(a)",
        "(no +C in a definite integral)",
    ]))
    story.extend(nice_table(
        ["Function f(x)", "∫ f(x) dx"],
        [
            ["k (constant)", "kx + C"],
            ["x^n  (n ≠ −1)", "x^(n+1)/(n+1) + C"],
            ["1/x", "ln|x| + C"],
            ["e^x", "e^x + C"],
            ["sin x", "−cos x + C"],
            ["cos x", "sin x + C"],
            ["sec^2 x", "tan x + C"],
            ["sec x tan x", "sec x + C"],
            ["cosec^2 x", "−cot x + C"],
        ],
        size=9,
    ))
    story.extend(formula_card([
        "Area under curve (above axis) = ∫_a^b f(x) dx",
        "Area between curves = ∫ (upper − lower) dx",
        "Average value = (1/(b−a)) × ∫_a^b f(x) dx",
        "Displacement = area under v–t graph",
        "v = ∫ a dt + C₁ ,   x = ∫ v dt + C₂",
    ]))
    story.extend(formula_card([
        "Substitution (chain rule backwards)",
        "1. Choose u = inside piece",
        "2. Write du = (…) dx",
        "3. Rewrite the whole integral in u",
        "4. For definite integrals, change both limits to u values",
    ]))
    story.extend(example(
        "power rule",
        [
            "∫ 5x^2 dx = 5 × (x^3/3) + C = (5/3)x^3 + C.",
            "Check: differentiate to get 5x^2 back.",
        ],
        "(5/3) x^3 + C",
    ))

    # ===== 12 =====
    story.append(section("12. Revision lock lines — say these aloud"))
    story.extend(nice_table(
        ["Topic", "Say this in one breath"],
        [
            ["Ratios", "sin=P/H, cos=B/H, tan=P/B. Tan is never P/H."],
            ["Table", "Know 0°–180° for all six ratios. ND = not defined."],
            ["Identity", "sin^2θ + cos^2θ = 1; then 1+tan^2θ=sec^2θ."],
            ["ASTC", "All, Sin, Tan, Cos — positive in that order around the circle."],
            ["Radian", "s = Rθ with θ in radian; π rad = 180°."],
            ["Double", "sin2θ=2sinθcosθ; store three forms of cos2θ."],
            ["Small angle", "sinθ≈θ, tanθ≈θ, cosθ≈1 (θ in radian)."],
            ["a sin+b cos", "Biggest value √(a^2+b^2)."],
            ["AP / GP", "Equal add vs equal multiply; S∞ needs |r|<1."],
            ["Line", "y=mx+c; perpendicular slopes multiply to −1."],
            ["Differentiate", "Power, product, quotient, chain."],
            ["Max / min", "y'=0; y''>0 valley; y''<0 hill."],
            ["Integrate", "Undo derivative; definite = F(b)−F(a)."],
            ["Area / average", "Area = definite integral; average = that ÷ (b−a)."],
        ],
        size=8.5,
    ))
    story.append(Spacer(1, 2 * mm))
    story.extend(formula_card([
        "Last tip",
        "Keep the printed trig tables on your desk while you revise.",
        "Use radian for calculus trig. Label graph axes with units.",
        "If a formula is not on this sheet but was in class, add it by hand on the blank margin.",
    ]))
    story.append(para(
        "Source: Excellence Basic Maths & Calculus lectures 01–18. For revision — not a question paper.",
        "CoverMid",
    ))

    doc.build(story)
    OUT2.write_bytes(OUT.read_bytes())
    print(f"Created {OUT}")
    print(f"Copied  {OUT2}")
    print(f"Size {OUT.stat().st_size} bytes")


if __name__ == "__main__":
    build()
