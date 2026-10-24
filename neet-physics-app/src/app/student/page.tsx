import Link from "next/link";
import {
  ArrowRight,
  BookOpen,
  CalendarDays,
  CheckCircle2,
  Clock3,
  FileText,
  Target,
} from "lucide-react";
import { requireProfile } from "@/lib/auth";
import {
  getPublishedClasses,
  getStudentAssessmentCards,
  getStudentAttempts,
  getStudentOpenErrors,
} from "@/lib/queries";
import { EmptyState, Notice, PageHeading, formatDate } from "@/components/production-ui";

export default async function StudentDashboard() {
  const profile = await requireProfile(["student"]);
  let errors: Awaited<ReturnType<typeof getStudentOpenErrors>> = [];
  let errorsLoadFailed = false;
  try {
    errors = await getStudentOpenErrors(profile.id);
  } catch {
    errorsLoadFailed = true;
  }
  const [classes, assessments, attempts] = await Promise.all([
    getPublishedClasses(),
    getStudentAssessmentCards(profile.id),
    getStudentAttempts(profile.id),
  ]);
  const released = attempts.filter((attempt) => attempt.status === "released");
  const finished = attempts.filter((attempt) =>
    ["released", "submitted", "teacher_review"].includes(attempt.status),
  );
  const average = released.length
    ? Math.round(
        released.reduce((sum, attempt) => {
          const assessment = Array.isArray(attempt.assessments)
            ? attempt.assessments[0]
            : attempt.assessments;
          const total = Number(assessment?.total_marks ?? 0);
          return sum + (total ? (Number(attempt.total_score ?? 0) / total) * 100 : 0);
        }, 0) / released.length,
      )
    : 0;
  const currentClass = classes.at(-1);
  const nextAssessment = assessments.find((assessment) =>
    ["start", "resume", "retry"].includes(assessment.action),
  );
  const openCount = assessments.filter((item) =>
    ["start", "resume", "retry"].includes(item.action),
  ).length;

  return (
    <>
      <PageHeading
        eyebrow="Student workspace"
        title={`Welcome back, ${profile.full_name.split(" ")[0]}`}
        description="Continue from your latest published class and complete due assessments."
      />
      {errorsLoadFailed && (
        <Notice kind="error">Could not load the 48-hour retry list. Open Retry later or refresh.</Notice>
      )}
      {classes.length === 0 ? (
        <EmptyState
          title="Your teacher has not published a class yet"
          description="Published slides, class notes, short notes and quizzes will appear here automatically."
        />
      ) : (
        <>
          <section className="production-hero-grid">
            <article className="production-current-class">
              <div className="class-hero-meta">
                <span className="status-chip current"><span /> Latest class</span>
                <span>{formatDate(currentClass?.class_date ?? null)}</span>
              </div>
              <span className="content-label">{currentClass?.chapters?.title ?? "Physics"}</span>
              <h2>Lecture {String(currentClass?.lecture_number).padStart(2, "0")} · {currentClass?.title}</h2>
              <p>{currentClass?.description}</p>
              <div className="production-resource-count">
                <FileText size={15} /> {currentClass?.materials.length ?? 0} published resources
              </div>
              <Link
                className="button primary"
                href={`/student/classes?lecture=${currentClass?.lecture_number}`}
              >
                Open class materials <ArrowRight size={16} />
              </Link>
            </article>
            <article className="production-next-action">
              <span className="metric-icon violet"><Target size={21} /></span>
              <span className="small-label">NEXT ASSESSMENT</span>
              {nextAssessment ? (
                <>
                  <h3>{nextAssessment.title}</h3>
                  <p>{nextAssessment.duration_minutes} minutes · {nextAssessment.total_marks} marks</p>
                  <small><Clock3 size={14} /> Due {formatDate(nextAssessment.due_at)}</small>
                  <Link className="button primary full" href={`/student/assessments/${nextAssessment.id}`}>
                    {nextAssessment.action === "resume"
                      ? "Resume assessment"
                      : nextAssessment.action === "retry"
                        ? "Try quiz again"
                        : "Start assessment"}{" "}
                    <ArrowRight size={16} />
                  </Link>
                </>
              ) : (
                <><h3>You are up to date</h3><p>No open assessment is waiting.</p></>
              )}
            </article>
          </section>
          <section className="metric-grid">
            <article className="metric-card">
              <span className="metric-icon blue"><BookOpen size={20} /></span>
              <div>
                <span>Published classes</span>
                <strong>{classes.length}</strong>
                <p>{classes.reduce((sum, item) => sum + item.materials.length, 0)} learning resources</p>
              </div>
            </article>
            <article className="metric-card">
              <span className="metric-icon green"><CheckCircle2 size={20} /></span>
              <div>
                <span>Assessments completed</span>
                <strong>{finished.length}</strong>
                <p>{released.length} results released</p>
              </div>
            </article>
            <article className="metric-card">
              <span className="metric-icon amber"><Target size={20} /></span>
              <div>
                <span>Average accuracy</span>
                <strong>{average}%</strong>
                <p>Across released results</p>
              </div>
            </article>
            <article className="metric-card">
              <span className="metric-icon violet"><CalendarDays size={20} /></span>
              <div>
                <span>Open assessments</span>
                <strong>{openCount}</strong>
                <p>
                  <Link href="/student/retry">
                    {errorsLoadFailed ? "Retry list unavailable" : `${errors.length} items in 48-hour retry`}
                  </Link>
                </p>
              </div>
            </article>
          </section>
          <section className="production-panel">
            <div className="panel-head">
              <div>
                <span className="small-label">RECENT CLASSES</span>
                <h3>Continue learning</h3>
              </div>
              <Link className="text-button" href="/student/classes">View all <ArrowRight size={15} /></Link>
            </div>
            <div className="production-class-rows">
              {classes.slice(-4).reverse().map((item) => (
                <Link href={`/student/classes?lecture=${item.lecture_number}`} key={item.id}>
                  <span className="class-dot">{item.lecture_number}</span>
                  <div>
                    <strong>{item.title}</strong>
                    <small>{item.chapters?.title} · {item.materials.length} resources</small>
                  </div>
                  <span>{formatDate(item.class_date)}</span>
                  <ArrowRight size={16} />
                </Link>
              ))}
            </div>
          </section>
        </>
      )}
    </>
  );
}
