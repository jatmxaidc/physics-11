import LecturePlannerBoard from "@/components/lecture-planner-board";
import { requireProfile } from "@/lib/auth";
import { getPlannerBoard } from "@/lib/queries";

export default async function TeacherPlannerPage({
  searchParams,
}: {
  searchParams: Promise<{ error?: string; success?: string }>;
}) {
  await requireProfile(["teacher"]);
  const [rows, params] = await Promise.all([getPlannerBoard(), searchParams]);
  return (
    <LecturePlannerBoard
      role="teacher"
      rows={rows}
      notice={{ error: params.error, success: params.success }}
    />
  );
}
