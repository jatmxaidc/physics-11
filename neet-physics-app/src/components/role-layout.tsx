import ProductionShell from "@/components/production-shell";
import { requireProfile } from "@/lib/auth";
import type { Role } from "@/lib/data";

export default async function RoleLayout({
  role,
  children,
}: {
  role: Role;
  children: React.ReactNode;
}) {
  const profile = await requireProfile([role]);
  return <ProductionShell profile={profile}>{children}</ProductionShell>;
}
