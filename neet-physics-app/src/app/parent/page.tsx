import { BookOpen, CheckCircle2, CircleAlert, Target } from "lucide-react";
import { EmptyState, PageHeading, formatDate } from "@/components/production-ui";
import { requireProfile } from "@/lib/auth";
import { getLinkedStudent, getStudentAttempts, getStudentSummaryData } from "@/lib/queries";

export default async function ParentDashboard() {
  const parent = await requireProfile(["parent"]);
  const link = await getLinkedStudent(parent.id);
  const relation = link?.profiles as unknown;
  const student = (Array.isArray(relation) ? relation[0] : relation) as { id: string; full_name: string } | null;
  if (!student) {
    return <><PageHeading eyebrow="Parent workspace" title="Weekly summary" description="Progress summaries for your linked student." /><EmptyState title="No linked student" description="Ask the programme administrator to connect your parent account." /></>;
  }
  const [summary, attempts] = await Promise.all([
    getStudentSummaryData(student.id),
    getStudentAttempts(student.id),
  ]);
  const latestFeedback = attempts.find((attempt) => attempt.teacher_feedback);
  return (
    <>
      <PageHeading eyebrow="Parent workspace" title={`${student.full_name}’s progress summary`} description="A privacy-safe overview of attendance, consistency and released assessment results." />
      <section className="parent-production-hero">
        <div><span className="status-chip current"><span /> {summary.average_accuracy >= 70 ? "On track" : "Support recommended"}</span><h2>{summary.average_accuracy >= 70 ? "Consistent progress across the programme" : "One focused review will help restore progress"}</h2><p>This summary is generated from verified class attendance, homework records and released assessments.</p></div>
        <div className="parent-score"><strong>{summary.average_accuracy}%</strong><span>average accuracy</span></div>
      </section>
      <section className="metric-grid">
        <article className="metric-card"><span className="metric-icon green"><CheckCircle2 size={20} /></span><div><span>Attendance</span><strong>{summary.attendance_percent}%</strong><p>{summary.attendance_present} of {summary.attendance_total} classes</p></div></article>
        <article className="metric-card"><span className="metric-icon blue"><Target size={20} /></span><div><span>Assessments</span><strong>{summary.attempt_released}</strong><p>{summary.attempt_total} attempts recorded</p></div></article>
        <article className="metric-card"><span className="metric-icon amber"><BookOpen size={20} /></span><div><span>Homework</span><strong>{summary.homework_percent}%</strong><p>{summary.homework_complete} of {summary.homework_total} complete</p></div></article>
        <article className="metric-card"><span className="metric-icon violet"><CircleAlert size={20} /></span><div><span>Unresolved errors</span><strong>{summary.open_errors}</strong><p>Correction work pending</p></div></article>
      </section>
      <section className="two-column">
        <article className="production-panel"><div className="panel-head"><div><span className="small-label">RELEASED RESULTS</span><h3>Recent assessments</h3></div></div><div className="production-result-rows">{attempts.filter((attempt) => attempt.status === "released").slice(0, 5).map((attempt) => { const assessment = Array.isArray(attempt.assessments) ? attempt.assessments[0] : attempt.assessments; return <article key={attempt.id}><span className="status-chip released">Released</span><div><strong>{assessment?.title}</strong><small>{formatDate(attempt.released_at)}</small></div><b>{attempt.total_score} / {assessment?.total_marks}</b></article>; })}</div></article>
        <article className="production-panel"><div className="panel-head"><div><span className="small-label">TEACHER FEEDBACK</span><h3>Latest observation</h3></div></div>{latestFeedback ? <blockquote>{latestFeedback.teacher_feedback}</blockquote> : <p className="inline-empty">No teacher feedback has been released yet.</p>}<p className="parent-privacy-note">Private student notes and question-level responses are never shown in the parent workspace.</p></article>
      </section>
    </>
  );
}
