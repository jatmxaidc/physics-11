import Link from "next/link";
import { CheckCircle2, RotateCcw, Target } from "lucide-react";
import { EmptyState, Notice, PageHeading, formatDate } from "@/components/production-ui";
import { requireProfile } from "@/lib/auth";
import { getStudentAttempts, getStudentOpenErrors } from "@/lib/queries";

export default async function StudentProgressPage() {
  const profile = await requireProfile(["student"]);
  const attempts = await getStudentAttempts(profile.id);
  let errors: Awaited<ReturnType<typeof getStudentOpenErrors>> = [];
  let errorsLoadFailed = false;
  try {
    errors = await getStudentOpenErrors(profile.id);
  } catch {
    errorsLoadFailed = true;
  }
  const finished = attempts.filter((attempt) =>
    ["released", "submitted", "teacher_review"].includes(attempt.status),
  );
  const released = finished.filter((attempt) => attempt.status === "released");

  return (
    <>
      <PageHeading
        eyebrow="Learning evidence"
        title="My progress"
        description="Open any finished attempt for per-question answers, explanations (when released), and teacher feedback."
        actions={
          <Link className="button secondary" href="/student/retry">
            <RotateCcw size={16} /> 48-hour retry ({errors.length})
          </Link>
        }
      />
      {errorsLoadFailed && (
        <Notice kind="error">Could not load open retries. Open the 48-hour retry page to try again.</Notice>
      )}
      {finished.length === 0 ? (
        <EmptyState
          title="No assessment results yet"
          description="Complete your first published quiz to begin building a progress record."
        />
      ) : (
        <>
          <section className="metric-grid">
            <article className="metric-card">
              <span className="metric-icon blue">
                <Target size={20} />
              </span>
              <div>
                <span>Finished attempts</span>
                <strong>{finished.length}</strong>
                <p>Submitted or released</p>
              </div>
            </article>
            <article className="metric-card">
              <span className="metric-icon green">
                <CheckCircle2 size={20} />
              </span>
              <div>
                <span>Released</span>
                <strong>{released.length}</strong>
                <p>Ready to review</p>
              </div>
            </article>
            <article className="metric-card">
              <span className="metric-icon amber">
                <RotateCcw size={20} />
              </span>
              <div>
                <span>Open retries</span>
                <strong>{errorsLoadFailed ? "—" : errors.length}</strong>
                <p>
                  <Link href="/student/retry">Review correction list</Link>
                </p>
              </div>
            </article>
          </section>
          <div className="production-panel">
            <div className="panel-head">
              <div>
                <span className="small-label">RESULT HISTORY</span>
                <h3>Assessment attempts</h3>
              </div>
            </div>
            <div className="production-result-rows">
              {finished.map((attempt) => {
                const assessment = Array.isArray(attempt.assessments)
                  ? attempt.assessments[0]
                  : attempt.assessments;
                return (
                  <Link href={`/student/progress/${attempt.id}`} key={attempt.id}>
                    <article>
                      <span
                        className={`status-chip ${
                          attempt.status === "released" ? "released" : "draft"
                        }`}
                      >
                        {attempt.status.replace("_", " ")}
                      </span>
                      <div>
                        <strong>{assessment?.title ?? "Assessment"}</strong>
                        <small>Submitted {formatDate(attempt.submitted_at)} · Open review</small>
                      </div>
                      <b>
                        {attempt.total_score == null
                          ? "Pending"
                          : `${attempt.total_score} / ${assessment?.total_marks}`}
                      </b>
                    </article>
                  </Link>
                );
              })}
            </div>
          </div>
        </>
      )}
    </>
  );
}
