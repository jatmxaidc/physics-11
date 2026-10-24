import Link from "next/link";
import { CheckCircle2, RotateCcw } from "lucide-react";
import { resolveErrorItem } from "@/app/actions/student";
import { MathContent } from "@/components/math-content";
import { EmptyState, Notice, PageHeading, formatDate } from "@/components/production-ui";
import { requireProfile } from "@/lib/auth";
import { getStudentOpenErrors } from "@/lib/queries";

export default async function StudentRetryPage({
  searchParams,
}: {
  searchParams: Promise<{ error?: string; success?: string }>;
}) {
  const profile = await requireProfile(["student"]);
  const [errors, params] = await Promise.all([
    getStudentOpenErrors(profile.id),
    searchParams,
  ]);
  const dueSoon = errors.filter((item) => new Date(item.retry_due_at).getTime() <= Date.now());

  return (
    <>
      <PageHeading
        eyebrow="Correction loop"
        title="48-hour retry"
        description="Wrong MCQ answers from daily quizzes stay here until you review them. Aim to clear each item within 48 hours."
      />
      {params.error && <Notice kind="error">{params.error}</Notice>}
      {params.success && <Notice kind="success">{params.success}</Notice>}
      {errors.length === 0 ? (
        <EmptyState
          title="Nothing to retry"
          description="When you miss a quiz question, it appears here with the topic and any teacher explanation."
        />
      ) : (
        <>
          <section className="metric-grid">
            <article className="metric-card">
              <span className="metric-icon amber"><RotateCcw size={20} /></span>
              <div>
                <span>Open errors</span>
                <strong>{errors.length}</strong>
                <p>{dueSoon.length} due now or overdue</p>
              </div>
            </article>
            <article className="metric-card">
              <span className="metric-icon green"><CheckCircle2 size={20} /></span>
              <div>
                <span>How it works</span>
                <strong>Review</strong>
                <p>
                  Revisit the concept, then mark the item reviewed.{" "}
                  <Link href="/student/flashcards">Flashcards</Link> help for notes you saved.
                </p>
              </div>
            </article>
          </section>
          <div className="teacher-class-stack">
            {errors.map((item) => {
              const overdue = new Date(item.retry_due_at).getTime() <= Date.now();
              return (
                <article className="production-panel" key={item.id}>
                  <div className="panel-head">
                    <div>
                      <span className={`status-chip ${overdue ? "current" : "scheduled"}`}>
                        {overdue ? "Due now" : "Retry window"}
                      </span>
                      <h3>{item.topic}</h3>
                      <small>Retry by {formatDate(item.retry_due_at)}</small>
                    </div>
                    <form action={resolveErrorItem}>
                      <input type="hidden" name="errorId" value={item.id} />
                      <button className="button secondary" type="submit">
                        Mark reviewed
                      </button>
                    </form>
                  </div>
                  <MathContent as="p" content={item.prompt} />
                  {item.explanation && (
                    <blockquote>
                      <strong>Explanation</strong>
                      <br />
                      <MathContent as="span" content={item.explanation} />
                    </blockquote>
                  )}
                </article>
              );
            })}
          </div>
        </>
      )}
    </>
  );
}
