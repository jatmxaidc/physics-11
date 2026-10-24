import Link from "next/link";
import { ArrowLeft, CheckCircle2, CircleAlert, XCircle } from "lucide-react";
import { MathContent } from "@/components/math-content";
import { EmptyState, Notice, PageHeading, formatDate } from "@/components/production-ui";
import { requireProfile } from "@/lib/auth";
import { createClient } from "@/lib/supabase/server";
import { getAttemptReview } from "@/lib/queries";

function optionList(options: unknown): string[] {
  if (!Array.isArray(options)) return [];
  return options.map(String);
}

function optionIndex(value: unknown): number | null {
  if (typeof value === "number" && Number.isInteger(value)) return value;
  if (typeof value === "string" && /^\d+$/.test(value)) return Number(value);
  return null;
}

export default async function TeacherAttemptReviewPage({
  params,
}: {
  params: Promise<{ attemptId: string }>;
}) {
  await requireProfile(["teacher"]);
  const { attemptId } = await params;

  let review;
  try {
    review = await getAttemptReview(attemptId);
  } catch (error) {
    const message = error instanceof Error ? error.message : "Unable to load this attempt.";
    return (
      <>
        <PageHeading eyebrow="Student response" title="Unable to open attempt" description={message} />
        <Notice kind="error">{message}</Notice>
        <Link className="button secondary" href="/teacher/students">
          Back to students
        </Link>
      </>
    );
  }

  const supabase = await createClient();
  const { data: attemptMeta } = await supabase
    .from("attempts")
    .select("student_id, focus_leave_count, last_focus_left_at, profiles!attempts_student_id_fkey(full_name)")
    .eq("id", attemptId)
    .single();
  const studentRelation = attemptMeta?.profiles as unknown;
  const student = (Array.isArray(studentRelation) ? studentRelation[0] : studentRelation) as
    | { full_name?: string }
    | null;
  const studentName = student?.full_name ?? "Student";
  const focusLeaves = Number(attemptMeta?.focus_leave_count ?? 0);

  const { attempt, assessment, items, results_visible: resultsVisible = true } = review;
  const scoreLabel =
    attempt.total_score == null && attempt.objective_score == null
      ? "Pending"
      : `${attempt.total_score ?? attempt.objective_score ?? "—"} / ${assessment.total_marks}`;
  const correctCount = items.filter((item) => item.is_correct === true).length;
  const incorrectCount = items.filter((item) => item.is_correct === false).length;
  const pendingCount = items.filter(
    (item) => item.question_type === "short_answer" && item.marks_awarded == null,
  ).length;

  return (
    <>
      <PageHeading
        eyebrow="Student response"
        title={assessment.title}
        description={`${studentName} · submitted ${formatDate(attempt.submitted_at)} · ${attempt.status.replace(/_/g, " ")}`}
        actions={
          <Link className="button secondary" href="/teacher/students">
            <ArrowLeft size={16} /> All students
          </Link>
        }
      />

      <section className="metric-grid">
        <article className="metric-card">
          <span className="metric-icon green">
            <CheckCircle2 size={20} />
          </span>
          <div>
            <span>Score</span>
            <strong>{scoreLabel}</strong>
            <p>{assessment.assessment_type.replace(/_/g, " ")}</p>
          </div>
        </article>
        <article className="metric-card">
          <span className="metric-icon blue">
            <CircleAlert size={20} />
          </span>
          <div>
            <span>Objective</span>
            <strong>
              {correctCount}/{correctCount + incorrectCount > 0 ? correctCount + incorrectCount : items.length}
            </strong>
            <p>
              {correctCount} correct · {incorrectCount} incorrect
              {pendingCount ? ` · ${pendingCount} to mark` : ""}
            </p>
          </div>
        </article>
        <article className="metric-card">
          <span className="metric-icon amber">
            <CircleAlert size={20} />
          </span>
          <div>
            <span>Status</span>
            <strong>{attempt.status.replace(/_/g, " ")}</strong>
            <p>{resultsVisible ? "Full answer key visible to you" : "Waiting on release rules"}</p>
          </div>
        </article>
        <article className="metric-card">
          <span className="metric-icon violet">
            <CircleAlert size={20} />
          </span>
          <div>
            <span>Tab / app leaves</span>
            <strong>{focusLeaves}</strong>
            <p>
              {attemptMeta?.last_focus_left_at
                ? `Last leave ${formatDate(attemptMeta.last_focus_left_at as string)}`
                : "Recorded during the attempt"}
            </p>
          </div>
        </article>
      </section>

      {pendingCount > 0 && (
        <Notice kind="info">
          Short answers still need marks. Open the{" "}
          <Link href="/teacher/review">review queue</Link> to award marks and release.
        </Notice>
      )}

      {attempt.status === "teacher_review" && pendingCount === 0 && (
        <Notice kind="info">
          This paper is ready to release. Open the{" "}
          <Link href="/teacher/review">review queue</Link> and click Release results.
        </Notice>
      )}

      {items.length === 0 ? (
        <EmptyState title="No responses saved" description="This attempt has no question responses." />
      ) : (
        <div className="production-quiz-layout review-exam-layout">
          <section>
            {items.map((item, index) => {
              const options = optionList(item.options);
              const selected = optionIndex(item.response);
              const correct = optionIndex(item.correct_answer);
              const verdict =
                item.is_correct === true
                  ? "correct"
                  : item.is_correct === false
                    ? "incorrect"
                    : "pending";

              return (
                <article
                  className={`question-card production-question review-exam-card ${verdict}`}
                  key={item.question_id}
                >
                  <div className="question-meta">
                    <span>
                      QUESTION {index + 1} OF {items.length}
                    </span>
                    <b>
                      {item.topic} · {item.marks} marks
                      {item.marks_awarded != null ? ` · awarded ${item.marks_awarded}` : ""}
                    </b>
                  </div>
                  <MathContent as="h3" content={item.prompt} />

                  {item.question_type === "short_answer" ? (
                    <div className="review-written">
                      <p>
                        <strong>Student answer</strong>
                      </p>
                      <p>{item.response == null ? "No answer" : String(item.response)}</p>
                      {item.teacher_feedback && (
                        <blockquote>
                          <strong>Your feedback</strong>
                          <br />
                          {item.teacher_feedback}
                        </blockquote>
                      )}
                    </div>
                  ) : (
                    <div className="option-list review-options">
                      {options.map((option, optIndex) => {
                        const isSelected = selected === optIndex;
                        const isCorrect = correct === optIndex;
                        const className = [
                          isSelected ? "selected" : "",
                          isCorrect ? "is-correct" : "",
                          isSelected && !isCorrect ? "is-wrong" : "",
                        ]
                          .filter(Boolean)
                          .join(" ");
                        return (
                          <div className={`review-option ${className}`} key={optIndex}>
                            <span className="option-letter">{String.fromCharCode(65 + optIndex)}</span>
                            <MathContent as="p" content={option} />
                            {isCorrect && <small>Correct</small>}
                            {isSelected && !isCorrect && <small>Student choice</small>}
                            {isSelected && isCorrect && <small>Student choice</small>}
                          </div>
                        );
                      })}
                    </div>
                  )}

                  <div className="review-verdict">
                    {verdict === "correct" && (
                      <p className="ok">
                        <CheckCircle2 size={15} /> Correct
                      </p>
                    )}
                    {verdict === "incorrect" && (
                      <p className="bad">
                        <XCircle size={15} /> Incorrect
                      </p>
                    )}
                    {verdict === "pending" && <p>Awaiting marks</p>}
                    {item.explanation && (
                      <blockquote>
                        <strong>Explanation</strong>
                        <br />
                        <MathContent as="span" content={item.explanation} />
                      </blockquote>
                    )}
                  </div>
                </article>
              );
            })}
          </section>
        </div>
      )}
    </>
  );
}
