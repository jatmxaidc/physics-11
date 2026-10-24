import RoleLayout from "@/components/role-layout";

export const dynamic = "force-dynamic";

export default function ParentLayout({ children }: { children: React.ReactNode }) {
  return <RoleLayout role="parent">{children}</RoleLayout>;
}
