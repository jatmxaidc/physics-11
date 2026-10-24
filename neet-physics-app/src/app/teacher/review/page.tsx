import {
  releaseAttemptResults,
  reviewAssessmentAttendance,
  reviewResponse,
} from "@/app/actions/teaching";
import { MathContent } from "@/components/math-content";
import { EmptyState, Notice, PageHeading, formatDate } from "@/components/production-ui";
import {
  getPendingAssessmentAttendance,
  getTeacherReviewAttempts,
  splitTeacherReviewQueue,
} from "@/lib/queries";

export default async function TeacherReviewPage({
  searchParams,
}: {
  searchParams: Promise<{ error?: string; success?: string }>;
}) {
  const [attempts, attendanceRequests, params] = await Promise.all([
    getTeacherReviewAttempts(),
    getPendingAssessmentAttendance(),
    searchParams,
  ]);
  const { shortAnswerQueue, readyToRelease } = splitTeacherReviewQueue(attempts);
  const queueEmpty =
    shortAnswerQueue.length === 0 && readyToRelease.length === 0 && attendanceRequests.length === 0;

  return (
    <>
      <PageHeading
        eyebrow="Teacher marking"
        title="Attendance, review and release"
        description="Verify test attendance before students can start weekly/chapter papers. Then mark short answers and release results."
      />
      {params.error && <Notice kind="error">{params.error}</Notice>}
      {params.success && <Notice kind="success">{params.success}</Notice>}

      <section className="production-panel" style={{ marginBottom: "1.25rem" }}>
        <div className="panel-head">
          <div>
            <span className="small-label">TEST ATTENDANCE</span>
            <h3>Verify student is present before unlocking the paper</h3>
          </div>
        </div>
        {attendanceRequests.length === 0 ? (
          <p className="inline-empty">No attendance requests waiting.</p>
        ) : (
          <div className="manual-review-list">
            {attendanceRequests.map((row) => {
              const student = Array.isArray(row.profiles) ? row.profiles[0] : row.profiles;
              const assessment = Array.isArray(row.assessments) ? row.assessments[0] : row.assessments;
              return (
                <article
                  className="manual-review-card"
                  key={`${row.assessment_id}-${row.student_id}`}
                >
                  <header>
                    <div>
                      <span className="small-label">
                        {(assessment as { assessment_type?: string } | null)?.assessment_type?.replace(
                          /_/g,
                          " ",
                        ) ?? "test"}{" "}
                        · requested {formatDate(row.requested_at)}
                      </span>
                      <h2>{(assessment as { title?: string } | null)?.title ?? "Assessment"}</h2>
                      <p>
                        {(student as { full_name?: string } | null)?.full_name ?? "Student"} wants to
                        start this test
                      </p>
                    </div>
                    <span className="status-chip draft">Verify attendance</span>
                  </header>
                  <div className="publish-class-row" style={{ gap: "10px", flexWrap: "wrap" }}>
                    <form action={reviewAssessmentAttendance}>
                      <input type="hidden" name="assessmentId" value={row.assessment_id} />
                      <input type="hidden" name="studentId" value={row.student_id} />
                      <input type="hidden" name="decision" value="verified" />
                      <button className="button primary" type="submit">
                        Verify & unlock test
                      </button>
                    </form>
                    <form action={reviewAssessmentAttendance}>
                      <input type="hidden" name="assessmentId" value={row.assessment_id} />
                      <input type="hidden" name="studentId" value={row.student_id} />
                      <input type="hidden" name="decision" value="rejected" />
                      <button className="button secondary" type="submit">
                        Reject
                      </button>
                    </form>
                  </div>
                </article>
              );
            })}
          </div>
        )}
      </section>

      <section className="production-panel" style={{ marginBottom: "1.25rem" }}>
        <div className="panel-head">
          <div>
            <span className="small-label">READY TO RELEASE</span>
            <h3>Objective papers waiting for teacher release</h3>
          </div>
        </div>
        {readyToRelease.length === 0 ? (
          <p className="inline-empty">No papers are ready to release.</p>
        ) : (
          <div className="manual-review-list">
            {readyToRelease.map((attempt) => {
              const student = Array.isArray(attempt.profiles) ? attempt.profiles[0] : attempt.profiles;
              const assessment = Array.isArray(attempt.assessments)
                ? attempt.assessments[0]
                : attempt.assessments;
              return (
                <article className="manual-review-card" key={attempt.id}>
                  <header>
                    <div>
                      <span className="small-label">
                        {(assessment as { assessment_type?: string } | null)?.assessment_type?.replace(
                          /_/g,
                          " ",
                        ) ?? "assessment"}{" "}
                        · {formatDate(attempt.submitted_at)}
                      </span>
                      <h2>{(assessment as { title?: string } | null)?.title ?? "Assessment"}</h2>
                      <p>
                        {student?.full_name ?? "Student"} · objective score{" "}
                        {attempt.objective_score ?? "—"} /{" "}
                        {(assessment as { total_marks?: number } | null)?.total_marks ?? "—"}
                      </p>
                    </div>
                    <span className="status-chip draft">Awaiting release</span>
                  </header>
                  <form action={releaseAttemptResults} className="publish-class-row">
                    <input type="hidden" name="attemptId" value={attempt.id} />
                    <div>
                      <p>
                        <strong>Release results to the student</strong>
                        <span>
                          Daily quizzes still wait for their unlock time if it has not passed yet.
                        </span>
                      </p>
                    </div>
                    <button className="button primary">Release results</button>
                  </form>
                </article>
              );
            })}
          </div>
        )}
      </section>

      {shortAnswerQueue.length === 0 ? (
        queueEmpty ? (
          <EmptyState
            title="Nothing needs action"
            description="Attendance requests, short answers and papers waiting for release will appear here."
          />
        ) : (
          <p className="inline-empty">No short answers need marking right now.</p>
        )
      ) : (
        <div className="manual-review-list">
          {shortAnswerQueue.map(({ attempt, response, question, student, assessment }) => (
            <article className="manual-review-card" key={response.id}>
              <header>
                <div>
                  <span className="small-label">
                    {question.topic} · {assessment?.title}
                  </span>
                  <MathContent as="h2" content={question.prompt} />
                </div>
                <span className="status-chip draft">Teacher review</span>
              </header>
              <div className="student-response">
                <span>
                  STUDENT RESPONSE · {student?.full_name ?? "Student"} ·{" "}
                  {formatDate(attempt.submitted_at)}
                </span>
                <p>
                  {typeof response.response === "string"
                    ? response.response
                    : JSON.stringify(response.response)}
                </p>
              </div>
              <form action={reviewResponse} className="form-grid">
                <input type="hidden" name="responseId" value={response.id} />
                <label>
                  Marks awarded
                  <input
                    name="marksAwarded"
                    type="number"
                    min="0"
                    max={question.marks}
                    step="0.5"
                    required
                  />
                </label>
                <label className="full-field">
                  Feedback
                  <textarea name="feedback" placeholder="Explain what was correct and what to improve." />
                </label>
                <button className="button primary">Save review</button>
              </form>
            </article>
          ))}
        </div>
      )}
    </>
  );
}
