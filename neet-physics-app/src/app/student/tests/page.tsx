import Link from "next/link";
import { ArrowRight, CalendarDays, Clock3 } from "lucide-react";
import { EmptyState, PageHeading, formatDate } from "@/components/production-ui";
import { requireProfile } from "@/lib/auth";
import { getStudentAssessmentCards } from "@/lib/queries";

export default async function StudentTestsPage() {
  const profile = await requireProfile(["student"]);
  const assessments = (await getStudentAssessmentCards(profile.id)).filter(
    (item) => item.assessment_type !== "daily_quiz",
  );
  return (
    <>
      <PageHeading
        eyebrow="Cumulative assessment"
        title="Weekly and chapter tests"
        description="Mark attendance first. After your teacher verifies it, you get one attempt. Results are released later by your teacher."
      />
      {assessments.length === 0 ? (
        <EmptyState
          title="No weekly test is open"
          description="Scheduled tests will appear here when your teacher publishes them."
        />
      ) : (
        <div className="production-assessment-grid">
          {assessments.map((item) => {
            const chip =
              item.action === "review"
                ? item.attempt?.status === "released"
                  ? "released"
                  : "draft"
                : item.action === "resume"
                  ? "current"
                  : item.action === "awaiting_attendance"
                    ? "draft"
                    : item.action === "mark_attendance"
                      ? "scheduled"
                      : "scheduled";
            const label =
              item.action === "review"
                ? item.attempt?.status === "released"
                  ? "Result ready"
                  : "Awaiting teacher"
                : item.action === "resume"
                  ? "In progress"
                  : item.action === "awaiting_attendance"
                    ? "Awaiting verify"
                    : item.action === "mark_attendance"
                      ? "Mark attendance"
                      : "Ready to start";
            const href =
              item.action === "review" && item.attempt?.id
                ? `/student/progress/${item.attempt.id}`
                : item.action === "review"
                  ? "/student/progress"
                  : `/student/assessments/${item.id}`;
            const cta =
              item.action === "review"
                ? item.attempt?.status === "released"
                  ? "Review answers"
                  : "View status"
                : item.action === "resume"
                  ? "Resume"
                  : item.action === "awaiting_attendance"
                    ? "Check status"
                    : item.action === "mark_attendance"
                      ? "Mark attendance"
                      : "Open test";
            return (
              <article className="test-card" key={item.id}>
                <div className="test-card-top">
                  <span className="metric-icon violet">
                    <CalendarDays size={20} />
                  </span>
                  <span className={`status-chip ${chip}`}>{label}</span>
                </div>
                <span className="small-label">{item.assessment_type.replace("_", " ")}</span>
                <h3>{item.title}</h3>
                <div className="test-facts">
                  <span>
                    <Clock3 size={15} />
                    {item.duration_minutes} minutes
                  </span>
                  <span>{item.total_marks} marks</span>
                  {item.opens_at && <span>Opens {formatDate(item.opens_at)}</span>}
                  <span>Due {formatDate(item.due_at)}</span>
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
