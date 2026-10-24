import {
  CalendarDays,
  CheckCircle2,
  CircleDashed,
} from "lucide-react";
import {
  markPlannerCompletedThrough,
  markPlannerLectureComplete,
  markPlannerLectureIncomplete,
} from "@/app/actions/teaching";
import { Notice, PageHeading } from "@/components/production-ui";
import type { PlannerBoardRow } from "@/lib/queries";

function formatPlanDate(value: string | null) {
  if (!value) return "Date TBA";
  return new Intl.DateTimeFormat("en-IN", {
    weekday: "short",
    day: "numeric",
    month: "short",
    year: "numeric",
  }).format(new Date(`${value}T00:00:00`));
}

function formatPlanDateShort(value: string | null) {
  if (!value) return "Date TBA";
  const date = new Date(`${value}T00:00:00`);
  const day = date.getDate();
  const month = new Intl.DateTimeFormat("en-GB", { month: "short" }).format(date);
  const year = date.getFullYear();
  return `${day} ${month}, ${year}`;
}

export default function LecturePlannerBoard({
  role,
  rows,
  notice,
}: {
  role: "student" | "teacher";
  rows: PlannerBoardRow[];
  notice?: { error?: string; success?: string };
}) {
  const total = rows.length;
  const completed = rows.filter((row) => row.completion === "completed").length;
  const upcoming = total - completed;
  const percent = total ? Math.round((completed / total) * 100) : 0;
  const lastCompleted = [...rows].reverse().find((row) => row.completion === "completed");
  const nextUp = rows.find((row) => row.completion === "upcoming");

  return (
    <>
      <PageHeading
        eyebrow="Programme map"
        title="Lecture planner"
        description={
          role === "student"
            ? "Official timetable to 10 Feb 2027 (Mon–Fri). This planner is separate from My classes — your teacher marks lectures complete as the programme advances; class PDFs stay under My classes."
            : "Official timetable only — not linked to live class uploads. Mark lectures complete as you finish them. Class notes stay under Classes & materials."
        }
      />
      {notice?.error && <Notice kind="error">{notice.error}</Notice>}
      {notice?.success && <Notice kind="success">{notice.success}</Notice>}

      <section className="planner-hero">
        <div>
          <span className="small-label">CLASS 11 PHYSICS · 1 SEP 2026 – 10 FEB 2027</span>
          <h2>{completed} of {total} lectures completed</h2>
          <p>
            {lastCompleted
              ? `Latest completed: Lecture ${String(lastCompleted.lecture_number).padStart(2, "0")}`
              : "No lecture marked complete yet"}
            {nextUp
              ? ` · Next expected: ${formatPlanDate(nextUp.planned_date)}`
              : ""}
          </p>
          <div className="planner-progress-track" aria-hidden>
            <span style={{ width: `${percent}%` }} />
          </div>
        </div>
        <div className="planner-hero-stats">
          <article>
            <strong>{completed}</strong>
            <span>Completed</span>
          </article>
          <article>
            <strong>{upcoming}</strong>
            <span>Upcoming</span>
          </article>
          <article>
            <strong>{total}</strong>
            <span>Planned</span>
          </article>
        </div>
      </section>

      {role === "teacher" && (
        <details className="production-create-panel" open>
          <summary>
            <span><CheckCircle2 size={17} /> Mark completed through…</span>
            <small>Sets Lectures 01 to N as complete and clears anything after N</small>
          </summary>
          <form action={markPlannerCompletedThrough} className="form-grid">
            <label>
              Completed through lecture
              <select name="lectureNumber" defaultValue={lastCompleted?.lecture_number ?? 1}>
                {rows.map((row) => (
                  <option value={row.lecture_number} key={row.lecture_number}>
                    L{String(row.lecture_number).padStart(3, "0")} · {formatPlanDate(row.planned_date)} · {row.topic.slice(0, 48)}
                  </option>
                ))}
              </select>
            </label>
            <button className="button primary" type="submit">Update progress</button>
          </form>
        </details>
      )}

      <section className="planner-board">
        <div className="panel-head">
          <div>
            <span className="small-label">{total} LECTURES</span>
            <h3>Official plan · expected class dates</h3>
          </div>
          <CalendarDays size={18} />
        </div>

        <div className="planner-list">
          {rows.map((row) => (
            <div key={row.lecture_number} className={`planner-row ${row.completion}`}>
              <span className={`planner-number ${row.completion}`}>
                {String(row.lecture_number).padStart(2, "0")}
              </span>
              <div className="planner-copy">
                <span className="planner-meta">
                  <span className="planner-meta-chapter">{row.chapter_title}</span>
                  <span className="planner-meta-sep">·</span>
                  <span className="planner-meta-when">{row.day_name ?? "Day TBA"}</span>
                  <span className="planner-meta-sep">·</span>
                  <span className="planner-meta-when">{formatPlanDateShort(row.planned_date)}</span>
                </span>
                <strong>{row.topic}</strong>
              </div>
              <span className={`status-chip ${row.completion === "completed" ? "released" : "draft"}`}>
                {row.completion === "completed" ? "Completed" : "Upcoming"}
              </span>
              <div className="planner-action">
                {row.completion === "completed" ? <CheckCircle2 size={18} /> : <CircleDashed size={18} />}
                {role === "teacher" && (
                  row.completion === "completed" ? (
                    <form action={markPlannerLectureIncomplete}>
                      <input type="hidden" name="lectureNumber" value={row.lecture_number} />
                      <button className="button secondary" type="submit">Undo</button>
                    </form>
                  ) : (
                    <form action={markPlannerLectureComplete}>
                      <input type="hidden" name="lectureNumber" value={row.lecture_number} />
                      <button className="button primary" type="submit">Mark done</button>
                    </form>
                  )
                )}
              </div>
            </div>
          ))}
        </div>
      </section>
    </>
  );
}
