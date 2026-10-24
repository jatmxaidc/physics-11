import RoleLayout from "@/components/role-layout";

export const dynamic = "force-dynamic";

export default function TeacherLayout({ children }: { children: React.ReactNode }) {
  return <RoleLayout role="teacher">{children}</RoleLayout>;
}
