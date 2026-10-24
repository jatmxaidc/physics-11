import LecturePlannerBoard from "@/components/lecture-planner-board";
import { requireProfile } from "@/lib/auth";
import { getPlannerBoard } from "@/lib/queries";

export default async function StudentPlannerPage() {
  await requireProfile(["student"]);
  const rows = await getPlannerBoard();
  return <LecturePlannerBoard role="student" rows={rows} />;
}
