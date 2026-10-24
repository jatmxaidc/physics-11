import Link from "next/link";
import { ArrowRight, Clock3, FileText } from "lucide-react";
import { EmptyState, PageHeading, formatDate } from "@/components/production-ui";
import { requireProfile } from "@/lib/auth";
import { getStudentAssessmentCards } from "@/lib/queries";

export default async function StudentAssessmentsPage() {
  const profile = await requireProfile(["student"]);
  const assessments = (await getStudentAssessmentCards(profile.id)).filter(
    (item) => item.assessment_type === "daily_quiz",
  );
  return (
    <>
      <PageHeading
        eyebrow="Daily learning loop"
        title="Daily quizzes"
        description="Complete each class quiz before the next lesson. Scores and explanations unlock after the quiz deadline."
      />
      {assessments.length === 0 ? (
        <EmptyState title="No daily quiz is open" description="Your teacher’s next published quiz will appear here." />
      ) : (
        <div className="production-assessment-grid">
          {assessments.map((assessment) => {
            const chip =
              assessment.action === "review"
                ? assessment.attempt?.status === "released"
                  ? "released"
                  : "draft"
                : assessment.action === "retry"
                  ? "current"
                : assessment.action === "resume"
                  ? "current"
                  : "scheduled";
            const label =
              assessment.action === "review"
                ? assessment.attempt?.status === "released"
                  ? "Result ready"
                  : "Submitted"
                : assessment.action === "retry"
                  ? "Retry open"
                : assessment.action === "resume"
                  ? "In progress"
                  : "Open";
            const href =
              assessment.action === "review" && assessment.attempt?.id
                ? `/student/progress/${assessment.attempt.id}`
                : assessment.action === "review"
                  ? "/student/progress"
                  : `/student/assessments/${assessment.id}`;
            const cta =
              assessment.action === "review"
                ? assessment.attempt?.status === "released"
                  ? "Review answers"
                  : "View status"
                : assessment.action === "retry"
                  ? "Try again before deadline"
                : assessment.action === "resume"
                  ? "Resume"
                  : "Start securely";
            return (
              <article className="test-card" key={assessment.id}>
                <div className="test-card-top">
                  <span className="metric-icon blue"><FileText size={20} /></span>
                  <span className={`status-chip ${chip}`}>{label}</span>
                </div>
                <span className="small-label">DAILY QUIZ</span>
                <h3>{assessment.title}</h3>
                <div className="test-facts">
                  <span><Clock3 size={15} />{assessment.duration_minutes} minutes</span>
                  <span>{assessment.total_marks} marks</span>
                  {assessment.opens_at && <span>Opens {formatDate(assessment.opens_at)}</span>}
                  <span>Due {formatDate(assessment.due_at)}</span>
                </div>
                <Link className="button primary full" href={href}>
                  {cta} <ArrowRight size={16} />
                </Link>
              </article>
            );
          })}
        </div>
      )}
    </>
  );
}
