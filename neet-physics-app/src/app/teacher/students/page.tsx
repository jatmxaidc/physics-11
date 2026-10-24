import Link from "next/link";
import { ArrowRight, CheckCircle2, Target, Users } from "lucide-react";
import { markAttendance } from "@/app/actions/teaching";
import { EmptyState, Notice, PageHeading, formatDate } from "@/components/production-ui";
import { requireProfile } from "@/lib/auth";
import {
  getLinkedStudents,
  getStudentAttempts,
  getStudentSummaryData,
  getTeacherClasses,
  getTeacherLinkedAttempts,
} from "@/lib/queries";

function profileFromLink(link: Awaited<ReturnType<typeof getLinkedStudents>>[number]) {
  const relation = link.profiles as unknown;
  return (Array.isArray(relation) ? relation[0] : relation) as {
    id: string;
    full_name: string;
  } | null;
}

function scorePercent(totalScore: number | null, totalMarks: number | null | undefined) {
  if (totalScore == null || !totalMarks) return null;
  return Math.round((Number(totalScore) / Number(totalMarks)) * 100);
}

export default async function TeacherStudentsPage({
  searchParams,
}: {
  searchParams: Promise<{ error?: string; success?: string; student?: string }>;
}) {
  const teacher = await requireProfile(["teacher"]);
  const [links, classes, allAttempts, params] = await Promise.all([
    getLinkedStudents(teacher.id),
    getTeacherClasses(),
    getTeacherLinkedAttempts(),
    searchParams,
  ]);

  const students = links
    .map(profileFromLink)
    .filter((student): student is { id: string; full_name: string } => Boolean(student?.id));

  const selectedId = params.student && students.some((s) => s.id === params.student)
    ? params.student
    : students[0]?.id ?? null;
  const selected = students.find((s) => s.id === selectedId) ?? null;

  const [selectedAttempts, summary] = selected
    ? await Promise.all([getStudentAttempts(selected.id), getStudentSummaryData(selected.id)])
    : [[], null];

  const finished = selectedAttempts.filter((a) => a.status !== "in_progress");
  const withScore = finished.filter((a) => a.total_score != null || a.objective_score != null);
  const avg =
    withScore.length === 0
      ? null
      : Math.round(
          withScore.reduce((sum, attempt) => {
            const assessment = Array.isArray(attempt.assessments)
              ? attempt.assessments[0]
              : attempt.assessments;
            const score = Number(attempt.total_score ?? attempt.objective_score ?? 0);
            const total = Number(assessment?.total_marks ?? 0);
            return sum + (total ? (score / total) * 100 : 0);
          }, 0) / withScore.length,
        );

  return (
    <>
      <PageHeading
        eyebrow="Students & results"
        title="Student responses and performance"
        description="Open any attempt to see every answer, marks awarded, and correct options. Scores are visible to you even before student release."
      />
      {params.error && <Notice kind="error">{params.error}</Notice>}
      {params.success && <Notice kind="success">{params.success}</Notice>}

      {students.length === 0 ? (
        <EmptyState
          title="No linked students"
          description="Ask the project administrator to add rows in student_links so your teacher account can see this class."
        />
      ) : (
        <>
          <section className="metric-grid">
            <article className="metric-card">
              <span className="metric-icon blue">
                <Users size={20} />
              </span>
              <div>
                <span>Linked students</span>
                <strong>{students.length}</strong>
                <p>Visible under your account</p>
              </div>
            </article>
            <article className="metric-card">
              <span className="metric-icon violet">
                <Target size={20} />
              </span>
              <div>
                <span>Submitted attempts</span>
                <strong>{allAttempts.length}</strong>
                <p>Across all linked students</p>
              </div>
            </article>
            <article className="metric-card">
              <span className="metric-icon green">
                <CheckCircle2 size={20} />
              </span>
              <div>
                <span>Selected average</span>
                <strong>{avg == null ? "—" : `${avg}%`}</strong>
                <p>{selected?.full_name ?? "Choose a student"}</p>
              </div>
            </article>
          </section>

          {students.length > 1 && (
            <div className="production-panel" style={{ marginBottom: "1rem" }}>
              <div className="panel-head">
                <div>
                  <span className="small-label">ROSTER</span>
                  <h3>Choose a student</h3>
                </div>
              </div>
              <div className="production-class-rows">
                {students.map((student) => (
                  <Link
                    href={`/teacher/students?student=${student.id}`}
                    key={student.id}
                    className={student.id === selectedId ? "is-active" : undefined}
                  >
                    <span className="class-dot">
                      {student.full_name
                        .split(" ")
                        .map((part) => part[0])
                        .slice(0, 2)
                        .join("")}
                    </span>
                    <div>
                      <strong>{student.full_name}</strong>
                      <small>
                        {allAttempts.filter((a) => a.student_id === student.id).length} submissions
                      </small>
                    </div>
                    <ArrowRight size={15} />
                  </Link>
                ))}
              </div>
            </div>
          )}

          {selected && (
            <>
              <div className="student-profile-banner">
                <span className="avatar">
                  {selected.full_name
                    .split(" ")
                    .map((part) => part[0])
                    .slice(0, 2)
                    .join("")}
                </span>
                <div>
                  <h2>{selected.full_name}</h2>
                  <p>
                    {finished.length} submitted · {summary?.attempt_released ?? 0} released to student
                    {summary ? ` · attendance ${summary.attendance_percent}%` : ""}
                    {summary ? ` · homework ${summary.homework_percent}%` : ""}
                  </p>
                </div>
              </div>

              <form action={markAttendance} className="attendance-form">
                <input type="hidden" name="studentId" value={selected.id} />
                <label>
                  Class
                  <select name="classId" required>
                    {classes
                      .filter((item) => item.status === "published")
                      .map((item) => (
                        <option value={item.id} key={item.id}>
                          Lecture {item.lecture_number} · {item.title}
                        </option>
                      ))}
                  </select>
                </label>
                <label>
                  Status
                  <select name="status">
                    <option value="present">Present</option>
                    <option value="absent">Absent</option>
                    <option value="excused">Excused</option>
                  </select>
                </label>
                <button className="button secondary">Save attendance</button>
              </form>

              <div className="production-panel">
                <div className="panel-head">
                  <div>
                    <span className="small-label">ATTEMPTS</span>
                    <h3>Open any paper to see answers and marks</h3>
                  </div>
                </div>
                {finished.length === 0 ? (
                  <p className="inline-empty">This student has not submitted an assessment yet.</p>
                ) : (
                  <div className="production-result-rows">
                    {finished.map((attempt) => {
                      const assessment = Array.isArray(attempt.assessments)
                        ? attempt.assessments[0]
                        : attempt.assessments;
                      const displayScore = attempt.total_score ?? attempt.objective_score;
                      const pct = scorePercent(displayScore, assessment?.total_marks);
                      return (
                        <Link href={`/teacher/students/${attempt.id}`} key={attempt.id}>
                          <article>
                            <span
                              className={`status-chip ${
                                attempt.status === "released" ? "released" : "draft"
                              }`}
                            >
                              {attempt.status.replace(/_/g, " ")}
                            </span>
                            <div>
                              <strong>{assessment?.title ?? "Assessment"}</strong>
                              <small>
                                {formatDate(attempt.submitted_at)}
                                {pct != null ? ` · ${pct}%` : ""}
                                {Number((attempt as { focus_leave_count?: number }).focus_leave_count ?? 0) > 0
                                  ? ` · ${Number((attempt as { focus_leave_count?: number }).focus_leave_count)} leaves`
                                  : ""}
                              </small>
                            </div>
                            <b>
                              {displayScore == null
                                ? "View answers"
                                : `${displayScore} / ${assessment?.total_marks}`}
                            </b>
                          </article>
                        </Link>
                      );
                    })}
                  </div>
                )}
              </div>
            </>
          )}

          <div className="production-panel" style={{ marginTop: "1rem" }}>
            <div className="panel-head">
              <div>
                <span className="small-label">ALL RECENT SUBMISSIONS</span>
                <h3>Across the class</h3>
              </div>
            </div>
            {allAttempts.length === 0 ? (
              <p className="inline-empty">No submitted assessments yet.</p>
            ) : (
              <div className="production-result-rows">
                {allAttempts.slice(0, 20).map((attempt) => {
                  const assessment = Array.isArray(attempt.assessments)
                    ? attempt.assessments[0]
                    : attempt.assessments;
                  const studentRel = attempt.profiles as unknown;
                  const student = (Array.isArray(studentRel) ? studentRel[0] : studentRel) as
                    | { full_name?: string }
                    | null;
                  const displayScore = attempt.total_score ?? attempt.objective_score;
                  return (
                    <Link href={`/teacher/students/${attempt.id}`} key={attempt.id}>
                      <article>
                        <span
                          className={`status-chip ${
                            attempt.status === "released" ? "released" : "draft"
                          }`}
                        >
                          {attempt.status.replace(/_/g, " ")}
                        </span>
                        <div>
                          <strong>{assessment?.title ?? "Assessment"}</strong>
                          <small>
                            {student?.full_name ?? "Student"} · {formatDate(attempt.submitted_at)}
                          </small>
                        </div>
                        <b>
                          {displayScore == null
                            ? "Open"
                            : `${displayScore} / ${assessment?.total_marks}`}
                        </b>
                      </article>
                    </Link>
                  );
                })}
              </div>
            )}
          </div>
        </>
      )}
    </>
  );
}
