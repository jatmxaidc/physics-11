"""Student-facing Class 11 Physics lecture planner PDF (Dec or Feb track)."""
from __future__ import annotations

from collections import OrderedDict
from dataclasses import dataclass
from datetime import date
from html import escape
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

BASE = Path(__file__).resolve().parent

pdfmetrics.registerFont(TTFont("Arial", r"C:\Windows\Fonts\arial.ttf"))
pdfmetrics.registerFont(TTFont("Arial-Bold", r"C:\Windows\Fonts\arialbd.ttf"))

NAVY = colors.HexColor("#081E3F")
GOLD = colors.HexColor("#B8860B")
BLUE = colors.HexColor("#155FAA")
SOFT = colors.HexColor("#F4F7FB")
LINE = colors.HexColor("#D0DAE8")
MUTED = colors.HexColor("#5A6678")
WHITE = colors.white
TEXT = colors.HexColor("#141A26")

TYPE_LABEL = {
    "Concept": "Concept",
    "Practice": "Practice / test",
    "Mock": "Mock test",
    "Revision": "Review",
}


@dataclass
class StudentPlan:
    lec: list
    days: list
    holidays: dict
    out: Path
    footer_range: str
    cover_kicker: str
    cover_title: str
    facts: list
    shared_note: str
    duties_html: str
    chapter_note: str
    after_title: str
    after_body: str
    pdf_title: str


def dec_plan() -> StudentPlan:
    from generate_detailed_lecture_plan import LEC, days, md_day  # noqa: F401 — md_day used by build
    return StudentPlan(
        lec=LEC,
        days=days,
        holidays={
            date(2026, 10, 2): "Possible holiday — if there is no class, these topics move to the next class.",
            date(2026, 11, 9): "Diwali week — attend if class is held; otherwise complete the same topics from the recording / notes.",
            date(2026, 11, 10): "Diwali week — attend if class is held; otherwise complete the same topics from the recording / notes.",
            date(2026, 12, 25): "Christmas — if there is no class, sit the lock paper at home and bring it on 26 Dec.",
            date(2026, 12, 31): "Class 11 Physics complete.",
        },
        out=BASE / "Class_11_Physics_Student_Lecture_Planner_2026-27.pdf",
        footer_range="1 Sep 2026 – 31 Dec 2026",
        cover_kicker="CBSE 2026–27  +  NEET Foundation",
        cover_title="Student Lecture Planner",
        facts=[
            ["Starts", "Tuesday 1 September 2026"],
            ["Classes", "Monday–Saturday  •  60 minutes"],
            ["Week 1", "Tuesday 1 September to Saturday 5 September"],
            ["New topics end", "Thursday 24 December 2026"],
            ["Class 11 complete", "Thursday 31 December 2026"],
            ["Total classes", "105  (99 teaching + 6 lock / revision)"],
            ["After every class", "Daily test of 5–10 questions, before the next lecture"],
            ["Chapter complete", "70% on the chapter practice test"],
        ],
        duties_html=(
            "1. Complete the <b>daily test (DT)</b> before the next lecture. It takes 10–15 minutes.<br/>"
            "2. On a day with two topic lines, the test covers <b>both</b> topics.<br/>"
            "3. Read the matching NCERT section before the next class.<br/>"
            "4. Reattempt every wrong numerical within <b>48 hours</b>.<br/>"
            "5. A chapter is complete only after NCERT examples, exercises, formula sheet, practice class, and <b>70%</b> on the chapter test.<br/>"
            "6. Below 70% means extra practice that week before we add a new chapter load.<br/>"
            "7. One missed daily test may be recovered on Sunday."
        ),
        shared_note="Some days cover two related titles in one hour. Both titles are listed. Come prepared for both.",
        chapter_note="New theory ends on 24 December. 25–31 December is lock and revision — no new chapter.",
        after_title="After 31 December",
        after_body=(
            "Class 11 Physics teaching for this programme is complete. January is for school examinations, "
            "error-log repair and formula revision. Keep this planner for revision: every topic you were taught is listed by date."
        ),
        pdf_title="Class 11 Physics Student Lecture Planner 2026–27",
    )


def feb_plan() -> StudentPlan:
    from generate_detailed_lecture_plan_feb import LEC, days, n_teach

    theory = days[n_teach - 1]
    lock_start = days[n_teach]
    done = days[-1]
    return StudentPlan(
        lec=LEC,
        days=days,
        holidays={
            date(2026, 10, 2): "Possible holiday — if there is no class, this topic moves to the next class.",
            date(2026, 11, 9): "Diwali week — attend if class is held; otherwise complete the same topic from the recording / notes.",
            date(2026, 11, 10): "Diwali week — attend if class is held; otherwise complete the same topic from the recording / notes.",
            date(2026, 12, 25): "Christmas — if there is no class, this topic moves to Monday 28 December.",
            date(2027, 1, 1): "New Year — if there is no class, this topic moves to Monday 4 January.",
            date(2027, 1, 26): "Republic Day — if there is no class, this topic moves to 27 January.",
            date(2027, 2, 10): "Class 11 Physics complete.",
        },
        out=BASE / "Class_11_Physics_Student_Lecture_Planner_to_10_Feb_2027.pdf",
        footer_range="1 Sep 2026 – 10 Feb 2027  •  Mon–Fri",
        cover_kicker="Monday–Friday  •  finishes 10 February",
        cover_title="Student Lecture Planner",
        facts=[
            ["Starts", "Tuesday 1 September 2026"],
            ["Classes", "Monday–Friday  •  60 minutes"],
            ["Week 1", "Tuesday 1 September to Friday 4 September"],
            ["Saturday", "No lecture — daily-test catch-up, NCERT and error log"],
            ["New topics end", theory.strftime("%A %d %B %Y")],
            ["Class 11 complete", done.strftime("%A %d %B %Y")],
            ["Total classes", f"{len(LEC)}  ({n_teach} teaching + {len(LEC) - n_teach} lock / revision)"],
            ["After every class", "Daily test of 5–10 questions, before the next lecture"],
            ["Chapter complete", "70% on the chapter practice test"],
        ],
        duties_html=(
            "1. Complete the <b>daily test (DT)</b> before the next lecture. It takes 10–15 minutes.<br/>"
            "2. On a day with two topic lines, the test covers <b>both</b> topics.<br/>"
            "3. Read the matching NCERT section before the next class.<br/>"
            "4. Reattempt every wrong numerical within <b>48 hours</b>.<br/>"
            "5. A chapter is complete only after NCERT examples, exercises, formula sheet, practice class, and <b>70%</b> on the chapter test.<br/>"
            "6. Below 70% means extra practice that week before we add a new chapter load.<br/>"
            "7. One missed daily test may be recovered on <b>Saturday or Sunday</b>."
        ),
        shared_note="Some days cover two related titles in one hour. Both titles are listed. Come prepared for both. Saturday is not a class day.",
        chapter_note=(
            f"New theory ends on {theory.strftime('%d %B')}. "
            f"{lock_start.strftime('%d %B')}–{done.strftime('%d %B')} is lock and revision — no new chapter. "
            "Classes run Monday to Friday only."
        ),
        after_title="After 10 February",
        after_body=(
            "Class 11 Physics teaching for this programme is complete. Mid-February onward is for school examinations, "
            "error-log repair and formula revision. Keep this planner for revision: every topic you were taught is listed by date."
        ),
        pdf_title="Class 11 Physics Student Lecture Planner to 10 Feb 2027",
    )


def footer(canvas, doc, plan: StudentPlan):
    canvas.saveState()
    canvas.setFillColor(NAVY)
    canvas.rect(0, 0, A4[0], 12 * mm, fill=1, stroke=0)
    canvas.setFillColor(WHITE)
    canvas.setFont("Arial", 8)
    canvas.drawString(
        16 * mm, 4.5 * mm,
        f"Class 11 Physics  •  CBSE 2026–27 + NEET Foundation  •  {plan.footer_range}",
    )
    canvas.drawRightString(A4[0] - 16 * mm, 4.5 * mm, str(doc.page))
    canvas.restoreState()


def cover_footer(canvas, doc, plan: StudentPlan):
    footer(canvas, doc, plan)
    canvas.saveState()
    canvas.setFillColor(NAVY)
    canvas.rect(0, A4[1] - 48 * mm, A4[0], 48 * mm, fill=1, stroke=0)
    canvas.setFillColor(GOLD)
    canvas.rect(0, A4[1] - 48 * mm, 3.2 * mm, 48 * mm, fill=1, stroke=0)
    canvas.setFillColor(GOLD)
    canvas.setFont("Arial-Bold", 10)
    canvas.drawString(18 * mm, A4[1] - 18 * mm, "CLASS 11 PHYSICS EXCELLENCE PROGRAMME")
    canvas.setFillColor(WHITE)
    canvas.setFont("Arial-Bold", 22)
    canvas.drawString(18 * mm, A4[1] - 32 * mm, plan.cover_title)
    canvas.setFont("Arial", 11)
    canvas.drawString(18 * mm, A4[1] - 42 * mm, plan.cover_kicker)
    canvas.restoreState()


def styles():
    title = ParagraphStyle(
        "t", fontName="Arial-Bold", fontSize=16, leading=20,
        textColor=NAVY, spaceAfter=8,
    )
    h = ParagraphStyle(
        "h", fontName="Arial-Bold", fontSize=13, leading=17,
        textColor=NAVY, spaceBefore=8, spaceAfter=6,
    )
    body = ParagraphStyle(
        "b", fontName="Arial", fontSize=10, leading=14,
        textColor=TEXT, spaceAfter=4,
    )
    small = ParagraphStyle(
        "s", fontName="Arial", fontSize=8, leading=11, textColor=TEXT,
    )
    small_b = ParagraphStyle(
        "sb", fontName="Arial-Bold", fontSize=8, leading=11, textColor=NAVY,
    )
    muted = ParagraphStyle(
        "m", fontName="Arial", fontSize=9, leading=12, textColor=MUTED, spaceAfter=4,
    )
    cell = ParagraphStyle(
        "c", fontName="Arial", fontSize=7.5, leading=10, textColor=TEXT,
    )
    cell_b = ParagraphStyle(
        "cb", fontName="Arial-Bold", fontSize=7.5, leading=10, textColor=NAVY,
    )
    week = ParagraphStyle(
        "w", fontName="Arial-Bold", fontSize=11, leading=14,
        textColor=BLUE, spaceBefore=8, spaceAfter=4,
    )
    center = ParagraphStyle(
        "ctr", fontName="Arial", fontSize=10, leading=14,
        alignment=TA_CENTER, textColor=TEXT,
    )
    return title, h, body, small, small_b, muted, cell, cell_b, week, center


def fact_table(plan: StudentPlan):
    t = Table(plan.facts, colWidths=[42 * mm, 130 * mm])
    t.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (0, -1), "Arial-Bold"),
        ("FONTNAME", (1, 0), (1, -1), "Arial"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("TEXTCOLOR", (0, 0), (0, -1), NAVY),
        ("TEXTCOLOR", (1, 0), (1, -1), TEXT),
        ("BACKGROUND", (0, 0), (-1, -1), SOFT),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("GRID", (0, 0), (-1, -1), 0.4, LINE),
    ]))
    return t


def chapter_map_white_header(plan: StudentPlan):
    _, _, _, _, _, _, cell, cell_b, _, _ = styles()
    head = ParagraphStyle(
        "ch", fontName="Arial-Bold", fontSize=8, leading=11, textColor=WHITE,
    )
    ranges = OrderedDict()
    for i, row in enumerate(plan.lec):
        ch = row[1]
        if ch not in ranges:
            ranges[ch] = [i, i]
        else:
            ranges[ch][1] = i
    data = [[
        Paragraph("Chapter", head),
        Paragraph("Lectures", head),
        Paragraph("Dates", head),
        Paragraph("Days", head),
    ]]
    for ch, (a, b) in ranges.items():
        da, db = plan.days[a], plan.days[b]
        if da.month == db.month and da.year == db.year:
            span = f"{da.strftime('%d')}–{db.strftime('%d %b')}"
        else:
            span = f"{da.strftime('%d %b')} – {db.strftime('%d %b')}"
        data.append([
            Paragraph(escape(ch), cell),
            Paragraph(f"L{a+1:03d} – L{b+1:03d}", cell_b),
            Paragraph(span, cell),
            Paragraph(str(b - a + 1), cell),
        ])
    t = Table(data, colWidths=[78 * mm, 32 * mm, 42 * mm, 20 * mm], repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, SOFT]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.3, LINE),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("ALIGN", (3, 1), (3, -1), "CENTER"),
    ]))
    return t


def week_table(plan: StudentPlan, idxs: list[int]):
    _, _, _, _, _, _, cell, cell_b, _, _ = styles()
    head = ParagraphStyle(
        "wh", fontName="Arial-Bold", fontSize=7.5, leading=10, textColor=WHITE,
    )
    data = [[
        Paragraph("L", head),
        Paragraph("Date", head),
        Paragraph("Chapter", head),
        Paragraph("Topics for the class", head),
        Paragraph("Type", head),
        Paragraph("After class", head),
    ]]
    for i in idxs:
        _pw, ch, topics, typ, dt, _pws, _why, _pc = plan.lec[i]
        dte = plan.days[i]
        date_s = dte.strftime("%a %d %b")
        if dte.year != 2026:
            date_s = dte.strftime("%a %d %b %Y")
        if dte in plan.holidays:
            date_s += " *"
        topic_html = (
            escape(topics)
            .replace("  ·  ", "<br/>")
            .replace("(1) ", "• ")
            .replace("(2) ", "• ")
            .replace("(3) ", "• ")
        )
        dt_short = f"DT-{i+1:03d}"
        if "No extra DT" in dt or "No DT" in dt:
            dt_short = "No daily test"
        data.append([
            Paragraph(f"{i+1:03d}", cell_b),
            Paragraph(escape(date_s), cell),
            Paragraph(escape(ch), cell),
            Paragraph(topic_html, cell),
            Paragraph(escape(TYPE_LABEL.get(typ, typ)), cell),
            Paragraph(escape(dt_short), cell),
        ])
    t = Table(
        data,
        colWidths=[12 * mm, 28 * mm, 36 * mm, 66 * mm, 22 * mm, 22 * mm],
        repeatRows=1,
    )
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, SOFT]),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.3, LINE),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    return t


def build(plan: StudentPlan):
    from generate_detailed_lecture_plan import md_day, week_key

    title, h, body, small, small_b, muted, cell, cell_b, week, center = styles()
    story = []
    story.append(Spacer(1, 42 * mm))
    story.append(Paragraph(
        "This booklet is your class timetable for Class 11 Physics. "
        "Bring it to every lecture. Tick each class after you finish the daily test.",
        body,
    ))
    story.append(Spacer(1, 4 * mm))
    story.append(fact_table(plan))
    story.append(Spacer(1, 6 * mm))
    story.append(Paragraph("How every class works", h))
    story.append(Paragraph(
        "<b>0–10 min</b> — We correct yesterday’s weak daily-test questions.<br/>"
        "<b>10–40 min</b> — New concept, diagram and derivation.<br/>"
        "<b>40–55 min</b> — Guided numerical and NEET-style items.<br/>"
        "<b>55–60 min</b> — Formula recap and today’s daily-test instruction.",
        body,
    ))
    story.append(Paragraph("What you must do", h))
    story.append(Paragraph(plan.duties_html, body))
    story.append(Paragraph(escape(plan.shared_note), muted))

    story.append(PageBreak())
    story.append(Paragraph("Chapter map", title))
    story.append(Paragraph(plan.chapter_note, muted))
    story.append(chapter_map_white_header(plan))
    story.append(Spacer(1, 5 * mm))
    story.append(Paragraph("If a class is cancelled", h))
    for hd, note in plan.holidays.items():
        story.append(Paragraph(f"<b>{hd.strftime('%a %d %b %Y')}</b> — {escape(note)}", small))
    story.append(Spacer(1, 3 * mm))
    story.append(Paragraph(
        "If there is no class on a listed date, the same lecture number and topics move to the next working day. Topics are not dropped.",
        body,
    ))

    story.append(PageBreak())
    story.append(Paragraph("Lecture calendar", title))
    story.append(Paragraph(
        "L = lecture number. DT = daily test to be finished before the next class. "
        "A star (*) means a possible holiday — see page 2.",
        muted,
    ))

    weeks = OrderedDict()
    for i, dte in enumerate(plan.days):
        weeks.setdefault(week_key(dte), []).append(i)

    for w_i, (_monday, idxs) in enumerate(weeks.items(), start=1):
        a, b = plan.days[idxs[0]], plan.days[idxs[-1]]
        if a.month == b.month and a.year == b.year:
            wr = f"{a.strftime('%d')}–{b.strftime('%d %b %Y')}"
        else:
            wr = f"{a.strftime('%d %b')} – {b.strftime('%d %b %Y')}"
        ch0, ch1 = plan.lec[idxs[0]][1], plan.lec[idxs[-1]][1]
        subtitle = ch0 if ch0 == ch1 else f"{ch0}  →  {ch1}"
        block = [
            Paragraph(f"Week {w_i}: {wr}  —  {escape(subtitle)}", week),
            week_table(plan, idxs),
        ]
        notes = [i for i in idxs if plan.days[i] in plan.holidays]
        for i in notes:
            block.append(Paragraph(
                f"* {md_day(plan.days[i])}: {escape(plan.holidays[plan.days[i]])}",
                muted,
            ))
        story.append(KeepTogether(block))

    story.append(Spacer(1, 6 * mm))
    story.append(Paragraph(plan.after_title, h))
    story.append(Paragraph(plan.after_body, body))

    def on_first(c, d):
        cover_footer(c, d, plan)

    def on_later(c, d):
        footer(c, d, plan)

    doc = SimpleDocTemplate(
        str(plan.out),
        pagesize=A4,
        leftMargin=16 * mm,
        rightMargin=16 * mm,
        topMargin=16 * mm,
        bottomMargin=18 * mm,
        title=plan.pdf_title,
        author="Class 11 Physics Excellence Programme",
    )
    doc.build(story, onFirstPage=on_first, onLaterPages=on_later)
    print(f"Wrote {plan.out}")
    print(f"Lectures {len(plan.lec)}")


if __name__ == "__main__":
    import sys
    track = sys.argv[1] if len(sys.argv) > 1 else "dec"
    if track in ("feb", "february", "10feb"):
        build(feb_plan())
    else:
        build(dec_plan())
