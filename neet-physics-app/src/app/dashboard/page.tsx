import { redirect } from "next/navigation";
import { requireProfile } from "@/lib/auth";

export const dynamic = "force-dynamic";

export default async function DashboardPage() {
  const profile = await requireProfile();
  redirect(`/${profile.role}`);
}
