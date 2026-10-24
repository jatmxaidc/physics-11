"""Rebuild lecture_planner seed from the Feb 2027 student planner."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from generate_detailed_lecture_plan_feb import HOLIDAYS, LEC, days  # noqa: E402

DEST = Path(__file__).resolve().parents[1] / "supabase" / "seed_lecture_planner.sql"


def sql_text(value: str | None) -> str:
    if value is None or str(value).strip() == "":
        return "null"
    return "'" + str(value).replace("'", "''") + "'"


def phase_for(index: int, session_type: str) -> str:
    if index >= 111:
        return "Lock / revision"
    if "Practice" in session_type or "Mock" in session_type:
        return "Practice"
    return "Teaching"


rows: list[str] = []
for index, (pw, chapter, topics, session_type, daily_test, *_rest) in enumerate(LEC):
    planned = days[index]
    holiday = HOLIDAYS.get(planned)
    homework = f"DT-{index + 1:03d} · {daily_test}"
    if holiday:
        homework = f"{homework} · Note: {holiday}"
    delivery = session_type if session_type else "Concept"
    if pw and pw != "—":
        delivery = f"{delivery} · PW {pw}"

    rows.append(
        "("
        + ", ".join(
            [
                str(index + 1),
                f"'{planned.isoformat()}'::date",
                sql_text(planned.strftime("%A")),
                sql_text(phase_for(index, session_type)),
                sql_text(chapter),
                sql_text(topics),
                sql_text(delivery),
                sql_text(homework),
            ]
        )
        + ")"
    )

sql = f"""-- Student Lecture Planner to 10 Feb 2027
-- Source PDF: Class_11_Physics_Student_Lecture_Planner_to_10_Feb_2027.pdf
-- Generated from generate_detailed_lecture_plan_feb.py (117 Mon–Fri lectures)

delete from public.lecture_planner where lecture_number > {len(LEC)};

insert into public.lecture_planner (
  lecture_number, planned_date, day_name, phase, chapter_title, topic, delivery_focus, assessment_or_homework
) values
{',\n'.join(rows)}
on conflict (lecture_number) do update set
  planned_date = excluded.planned_date,
  day_name = excluded.day_name,
  phase = excluded.phase,
  chapter_title = excluded.chapter_title,
  topic = excluded.topic,
  delivery_focus = excluded.delivery_focus,
  assessment_or_homework = excluded.assessment_or_homework;
"""

DEST.write_text(sql, encoding="utf-8")
print(f"Wrote {DEST}")
print(f"Lectures: {len(LEC)}  start={days[0]}  end={days[-1]}")
