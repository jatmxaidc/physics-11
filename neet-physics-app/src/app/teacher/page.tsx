import Link from "next/link";
import { ArrowRight, BookOpen, FileText, MessageSquareText, PencilLine, Target } from "lucide-react";
import { PageHeading, formatDate } from "@/components/production-ui";
import { requireProfile } from "@/lib/auth";
import {
  getPendingAssessmentAttendance,
  getTeacherAssessments,
  getTeacherClasses,
  getTeacherLinkedAttempts,
  getTeacherReviewAttempts,
  splitTeacherReviewQueue,
} from "@/lib/queries";

function greetingForIst(now = new Date()) {
  const hour = Number(
    new Intl.DateTimeFormat("en-GB", {
      timeZone: "Asia/Kolkata",
      hour: "numeric",
      hour12: false,
    }).format(now),
  );
  if (hour < 12) return "Good morning";
  if (hour < 17) return "Good afternoon";
  return "Good evening";
}

export default async function TeacherDashboard() {
  const profile = await requireProfile(["teacher"]);
  const [classes, assessments, reviews, recentAttempts, attendanceRequests] = await Promise.all([
    getTeacherClasses(),
    getTeacherAssessments(),
    getTeacherReviewAttempts(),
    getTeacherLinkedAttempts(),
    getPendingAssessmentAttendance(),
  ]);
  const { shortAnswerQueue, readyToRelease } = splitTeacherReviewQueue(reviews);
  const actionable =
    shortAnswerQueue.length + readyToRelease.length + attendanceRequests.length;

  return (
    <>
      <PageHeading
        eyebrow="Teacher workspace"
        title={`${greetingForIst()}, ${profile.full_name}`}
        description="Publish materials, mark work, release results, and open any student attempt to see answers and scores."
        actions={
          <Link className="button primary" href="/teacher/students">
            Students & results <ArrowRight size={16} />
          </Link>
        }
      />
      <section className="metric-grid">
        <article className="metric-card">
          <span className="metric-icon blue">
            <BookOpen size={20} />
          </span>
          <div>
            <span>Classes</span>
            <strong>{classes.length}</strong>
            <p>{classes.filter((item) => item.status === "published").length} published</p>
          </div>
        </article>
        <article className="metric-card">
          <span className="metric-icon green">
            <FileText size={20} />
          </span>
          <div>
            <span>Materials</span>
            <strong>{classes.reduce((sum, item) => sum + item.materials.length, 0)}</strong>
            <p>Private secure storage</p>
          </div>
        </article>
        <article className="metric-card">
          <span className="metric-icon violet">
            <PencilLine size={20} />
          </span>
          <div>
            <span>Assessments</span>
            <strong>{assessments.length}</strong>
            <p>{assessments.filter((item) => item.status === "published").length} open</p>
          </div>
        </article>
        <article className="metric-card">
          <span className="metric-icon amber">
            <MessageSquareText size={20} />
          </span>
          <div>
            <span>Needs action</span>
            <strong>{actionable}</strong>
            <p>
              {attendanceRequests.length} attendance · {shortAnswerQueue.length} to mark ·{" "}
              {readyToRelease.length} to release
            </p>
          </div>
        </article>
      </section>
      <section className="two-column">
        <article className="production-panel">
          <div className="panel-head">
            <div>
              <span className="small-label">RECENT STUDENT RESULTS</span>
              <h3>Open an attempt to see every answer</h3>
            </div>
            <Link className="text-button" href="/teacher/students">
              All results <ArrowRight size={15} />
            </Link>
          </div>
          {recentAttempts.length === 0 ? (
            <p className="inline-empty">No submitted assessments yet from linked students.</p>
          ) : (
            <div className="production-result-rows">
              {recentAttempts.slice(0, 6).map((attempt) => {
                const assessment = Array.isArray(attempt.assessments)
                  ? attempt.assessments[0]
                  : attempt.assessments;
                const studentRel = attempt.profiles as unknown;
                const student = (Array.isArray(studentRel) ? studentRel[0] : studentRel) as
                  | { full_name?: string }
                  | null;
                const score = attempt.total_score ?? attempt.objective_score;
                return (
                  <Link href={`/teacher/students/${attempt.id}`} key={attempt.id}>
                    <article>
                      <span className={`status-chip ${attempt.status === "released" ? "released" : "draft"}`}>
                        {attempt.status.replace(/_/g, " ")}
                      </span>
                      <div>
                        <strong>{assessment?.title ?? "Assessment"}</strong>
                        <small>
                          {student?.full_name ?? "Student"} · {formatDate(attempt.submitted_at)}
                        </small>
                      </div>
                      <b>{score == null ? "Open" : `${score} / ${assessment?.total_marks}`}</b>
                    </article>
                  </Link>
                );
              })}
            </div>
          )}
        </article>
        <article className="production-panel">
          <div className="panel-head">
            <div>
              <span className="small-label">MARKING QUEUE</span>
              <h3>Review and release</h3>
            </div>
            <Link className="text-button" href="/teacher/review">
              Open queue <ArrowRight size={15} />
            </Link>
          </div>
          {actionable === 0 ? (
            <p className="inline-empty">No submitted work needs marking or release.</p>
          ) : (
            <div className="production-review-summary">
              {attendanceRequests.slice(0, 2).map((row) => {
                const assessment = Array.isArray(row.assessments) ? row.assessments[0] : row.assessments;
                const student = Array.isArray(row.profiles) ? row.profiles[0] : row.profiles;
                return (
                  <div key={`${row.assessment_id}-${row.student_id}`}>
                    <span className="status-chip draft">Attendance</span>
                    <div>
                      <strong>{(assessment as { title?: string } | null)?.title ?? "Assessment"}</strong>
                      <small>
                        {(student as { full_name?: string } | null)?.full_name ?? "Student"} waiting
                        to start
                      </small>
                    </div>
                  </div>
                );
              })}
              {readyToRelease.slice(0, 2).map((review) => {
                const relation = review.assessments as unknown;
                const assessment = (Array.isArray(relation) ? relation[0] : relation) as {
                  title: string;
                } | null;
                return (
                  <div key={review.id}>
                    <span className="status-chip draft">Release</span>
                    <div>
                      <strong>{assessment?.title ?? "Assessment"}</strong>
                      <small>Ready to release to student</small>
                    </div>
                  </div>
                );
              })}
              {shortAnswerQueue
                .slice(
                  0,
                  Math.max(
                    0,
                    4 - Math.min(2, attendanceRequests.length) - Math.min(2, readyToRelease.length),
                  ),
                )
                .map((item) => (
                  <div key={item.response.id}>
                    <span className="status-chip draft">Mark</span>
                    <div>
                      <strong>{item.assessment?.title ?? "Assessment"}</strong>
                      <small>Short answer waiting</small>
                    </div>
                  </div>
                ))}
            </div>
          )}
          <div style={{ padding: "0 18px 18px" }}>
            <Link className="button secondary" href="/teacher/students">
              <Target size={16} /> View performance
            </Link>
          </div>
        </article>
      </section>
    </>
  );
}
