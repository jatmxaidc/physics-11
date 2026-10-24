import Link from "next/link";
import { ShieldCheck } from "lucide-react";
import AssessmentRunner from "@/components/assessment-runner";
import { Notice, PageHeading } from "@/components/production-ui";
import { getAssessmentForStudent } from "@/lib/queries";

export default async function AssessmentPage({
  params,
  searchParams,
}: {
  params: Promise<{ id: string }>;
  searchParams: Promise<{ error?: string; success?: string }>;
}) {
  const { id } = await params;
  const notices = await searchParams;
  let payload;
  try {
    payload = await getAssessmentForStudent(id);
  } catch (error) {
    const message = error instanceof Error ? error.message : "This assessment is unavailable.";
    return (
      <>
        <PageHeading
          eyebrow="Assessment"
          title="Unable to open assessment"
          description="The paper may be closed, not yet open, or past its due time."
        />
        <Notice kind="error">{message}</Notice>
        <Link className="button secondary" href="/student/assessments">
          Back to quizzes
        </Link>
      </>
    );
  }

  const isWeeklyStyle = payload.assessment.assessment_type !== "daily_quiz";

  const questions = payload.attendance.canStart
    ? payload.questions.flatMap((mapping) => {
        const relation = (mapping as { questions?: unknown }).questions as unknown;
        const question = Array.isArray(relation) ? relation[0] : relation;
        if (!question || typeof question !== "object") return [];
        const row = question as {
          id: string;
          prompt: string;
          question_type: string;
          options: unknown;
          topic: string;
        };
        return [
          {
            id: row.id,
            prompt: row.prompt,
            question_type: row.question_type,
            options: Array.isArray(row.options) ? row.options.map(String) : null,
            topic: row.topic,
            marks: Number((mapping as { marks: number }).marks),
          },
        ];
      })
    : [];

  const initialAttempt = payload.activeAttempt
    ? {
        attempt_id: payload.activeAttempt.id,
        started_at: payload.activeAttempt.started_at,
        duration_minutes: payload.assessment.duration_minutes,
        focus_leave_count: payload.activeAttempt.focus_leave_count ?? 0,
      }
    : null;

  return (
    <>
      {notices.error && <Notice kind="error">{notices.error}</Notice>}
      {notices.success && <Notice kind="success">{notices.success}</Notice>}
      {isWeeklyStyle && payload.attendance.status === "verified" && !initialAttempt && (
        <Notice kind="success">
          <ShieldCheck size={16} /> Attendance verified. You can begin the test.
        </Notice>
      )}
      <AssessmentRunner
        assessment={{
          id: payload.assessment.id,
          title: payload.assessment.title,
          duration_minutes: payload.assessment.duration_minutes,
          total_marks: Number(payload.assessment.total_marks),
          instructions: payload.assessment.instructions,
          opens_at: payload.assessment.opens_at,
          due_at: payload.assessment.due_at,
          assessment_type: payload.assessment.assessment_type,
        }}
        questions={questions}
        initialAttempt={initialAttempt}
        attendance={payload.attendance}
      />
    </>
  );
}
