import Link from "next/link";
import { ArrowLeft, CheckCircle2, CircleAlert, XCircle } from "lucide-react";
import { MathContent } from "@/components/math-content";
import { EmptyState, Notice, PageHeading, formatDate } from "@/components/production-ui";
import { requireProfile } from "@/lib/auth";
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

export default async function AttemptReviewPage({
  params,
}: {
  params: Promise<{ attemptId: string }>;
}) {
  await requireProfile(["student"]);
  const { attemptId } = await params;

  let review;
  try {
    review = await getAttemptReview(attemptId);
  } catch (error) {
    const message = error instanceof Error ? error.message : "Unable to load this result.";
    return (
      <>
        <PageHeading eyebrow="Result review" title="Unable to open result" description={message} />
        <Notice kind="error">{message}</Notice>
        <Link className="button secondary" href="/student/progress">
          Back to progress
        </Link>
      </>
    );
  }

  const {
    attempt,
    assessment,
    items,
    show_explanations: showExplanations,
    results_visible: resultsVisible = showExplanations,
    unlock_at: unlockAt,
  } = review;
  const scoreLabel =
    !resultsVisible || attempt.total_score == null
      ? "Pending"
      : `${attempt.total_score} / ${assessment.total_marks}`;
  const unlockLabel = unlockAt
    ? new Date(unlockAt).toLocaleString("en-IN", {
        day: "numeric",
        month: "short",
        year: "numeric",
        hour: "numeric",
        minute: "2-digit",
      })
    : null;
  const isDaily = assessment.assessment_type === "daily_quiz";

  return (
    <>
      <PageHeading
        eyebrow={resultsVisible ? "Answer review" : "Submission"}
        title={assessment.title}
        description={`Submitted ${formatDate(attempt.submitted_at)} · ${attempt.status.replace("_", " ")}`}
        actions={
          <Link className="button secondary" href="/student/progress">
            <ArrowLeft size={16} /> All results
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
            <p>{assessment.assessment_type.replace("_", " ")}</p>
          </div>
        </article>
        <article className="metric-card">
          <span className="metric-icon blue">
            <CircleAlert size={20} />
          </span>
          <div>
            <span>Answer key</span>
            <strong>{resultsVisible ? "Visible" : "Locked"}</strong>
            <p>
              {resultsVisible
                ? "Correct options and explanations are shown on each question"
                : isDaily
                  ? unlockLabel
                    ? `Unlocks after ${unlockLabel}`
                    : "Unlocks after the quiz deadline"
                  : "Your teacher will release this result later"}
            </p>
          </div>
        </article>
      </section>

      {!resultsVisible && (
        <Notice kind="info">
          {isDaily
            ? `Answers and explanations stay hidden until the deadline ends${unlockLabel ? ` (${unlockLabel})` : ""}. Your selected answers are saved.`
            : "This weekly/chapter test is one-time. Results appear here after your teacher releases them."}
        </Notice>
      )}

      {attempt.teacher_feedback && resultsVisible && (
        <Notice kind="info">
          <strong>Teacher feedback:</strong> {attempt.teacher_feedback}
        </Notice>
      )}

      {items.length === 0 ? (
        <EmptyState title="No question detail" description="Responses were not saved for this attempt." />
      ) : (
        <div className="production-quiz-layout review-exam-layout">
          <section>
            {items.map((item, index) => {
              const options = optionList(item.options);
              const selected = optionIndex(item.response);
              const correct = optionIndex(item.correct_answer);
              const verdict =
                !resultsVisible
                  ? null
                  : item.is_correct === true
                    ? "correct"
                    : item.is_correct === false
                      ? "incorrect"
                      : "pending";

              return (
                <article className={`question-card production-question review-exam-card ${verdict ?? ""}`} key={item.question_id}>
                  <div className="question-meta">
                    <span>QUESTION {index + 1} OF {items.length}</span>
                    <b>
                      {item.topic} · {item.marks} marks
                      {resultsVisible && item.marks_awarded != null ? ` · awarded ${item.marks_awarded}` : ""}
                    </b>
                  </div>
                  <MathContent as="h3" content={item.prompt} />

                  {item.question_type === "short_answer" ? (
                    <div className="review-written">
                      <p><strong>Your answer</strong></p>
                      <p>{item.response == null ? "No answer" : String(item.response)}</p>
                      {resultsVisible && item.teacher_feedback && (
                        <blockquote>
                          <strong>Teacher note</strong>
                          <br />
                          {item.teacher_feedback}
                        </blockquote>
                      )}
                    </div>
                  ) : (
                    <div className="option-list review-options">
                      {options.map((option, optionIndex) => {
                        const isSelected = selected === optionIndex;
                        const isCorrect = resultsVisible && correct === optionIndex;
                        const className = [
                          isSelected ? "selected" : "",
                          isCorrect ? "is-correct" : "",
                          resultsVisible && isSelected && !isCorrect ? "is-wrong" : "",
                        ]
                          .filter(Boolean)
                          .join(" ");
                        return (
                          <div className={`review-option ${className}`} key={optionIndex}>
                            <span className="option-letter">{String.fromCharCode(65 + optionIndex)}</span>
                            <MathContent as="p" content={option} />
                            {resultsVisible && isCorrect && <small>Correct</small>}
                            {resultsVisible && isSelected && !isCorrect && <small>Your choice</small>}
                            {!resultsVisible && isSelected && <small>Your choice</small>}
                          </div>
                        );
                      })}
                    </div>
                  )}

                  {resultsVisible && (
                    <div className="review-verdict">
                      {verdict === "correct" && (
                        <p className="ok"><CheckCircle2 size={15} /> Correct</p>
                      )}
                      {verdict === "incorrect" && (
                        <p className="bad"><XCircle size={15} /> Incorrect</p>
                      )}
                      {verdict === "pending" && <p>Awaiting teacher marks</p>}
                      {item.explanation && (
                        <blockquote>
                          <strong>Why this is correct</strong>
                          <br />
                          <MathContent as="span" content={item.explanation} />
                        </blockquote>
                      )}
                      {item.teacher_feedback && item.question_type !== "short_answer" && (
                        <blockquote>
                          <strong>Teacher note</strong>
                          <br />
                          {item.teacher_feedback}
                        </blockquote>
                      )}
                    </div>
                  )}
                </article>
              );
            })}
          </section>
        </div>
      )}
    </>
  );
}
