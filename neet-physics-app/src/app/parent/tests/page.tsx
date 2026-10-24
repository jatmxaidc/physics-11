import { EmptyState, PageHeading, formatDate } from "@/components/production-ui";
import { requireProfile } from "@/lib/auth";
import { getLinkedStudent, getStudentAttempts } from "@/lib/queries";

export default async function ParentTestsPage() {
  const parent = await requireProfile(["parent"]);
  const link = await getLinkedStudent(parent.id);
  const relation = link?.profiles as unknown;
  const student = (Array.isArray(relation) ? relation[0] : relation) as { id: string; full_name: string } | null;
  const attempts = student ? await getStudentAttempts(student.id) : [];
  return (
    <>
      <PageHeading eyebrow="Assessment record" title="Tests and results" description="Only released totals and teacher feedback are available to parent accounts." />
      {!student || attempts.length === 0 ? <EmptyState title="No test record available" description="Released assessment results will appear here." /> : <div className="production-panel"><div className="production-result-rows">{attempts.map((attempt) => { const assessment = Array.isArray(attempt.assessments) ? attempt.assessments[0] : attempt.assessments; return <article key={attempt.id}><span className={`status-chip ${attempt.status === "released" ? "released" : "draft"}`}>{attempt.status.replace("_", " ")}</span><div><strong>{assessment?.title}</strong><small>{formatDate(attempt.submitted_at)}</small></div><b>{attempt.status === "released" ? `${attempt.total_score} / ${assessment?.total_marks}` : "Not released"}</b></article>; })}</div></div>}
    </>
  );
}
