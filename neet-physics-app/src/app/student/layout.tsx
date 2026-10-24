import RoleLayout from "@/components/role-layout";

export const dynamic = "force-dynamic";

export default function StudentLayout({ children }: { children: React.ReactNode }) {
  return <RoleLayout role="student">{children}</RoleLayout>;
}
