import { redirect } from "next/navigation";
import PhysicsPlatform from "@/components/physics-platform";

export const metadata = {
  title: "Interface Preview — SMXPhysics",
};

export default function PreviewPage() {
  const configured = Boolean(
    process.env.NEXT_PUBLIC_SUPABASE_URL && process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY,
  );
  if (configured) redirect("/login");
  return (
    <>
      <div className="preview-banner">
        Static UI preview only — metrics and buttons are fake until Supabase is configured.
      </div>
      <PhysicsPlatform />
    </>
  );
}
