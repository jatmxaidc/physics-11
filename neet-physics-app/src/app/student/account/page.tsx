import { KeyRound, ShieldCheck } from "lucide-react";
import { changeAccountPassword } from "@/app/actions/auth";
import PasswordInput from "@/components/password-input";
import { Notice, PageHeading } from "@/components/production-ui";
import { requireProfile } from "@/lib/auth";

export default async function StudentAccountPage({
  searchParams,
}: {
  searchParams: Promise<{ error?: string; success?: string }>;
}) {
  const profile = await requireProfile(["student"]);
  const params = await searchParams;

  return (
    <>
      <PageHeading
        eyebrow="Account settings"
        title="Security"
        description={`Signed in as ${profile.full_name}. Change your password anytime — this is separate from the first-login reset.`}
      />
      {params.error && <Notice kind="error">{params.error}</Notice>}
      {params.success && <Notice kind="success">{params.success}</Notice>}
      <article className="production-panel account-panel">
        <div className="panel-head">
          <div>
            <span className="small-label">PASSWORD</span>
            <h3>Update password</h3>
          </div>
          <span className="metric-icon blue">
            <KeyRound size={20} />
          </span>
        </div>
        <div className="setup-warning">
          <ShieldCheck size={16} />
          <p>Use at least 6 characters. Do not reuse a temporary password shared by your teacher.</p>
        </div>
        <form action={changeAccountPassword} className="form-grid">
          <label>
            New password
            <PasswordInput name="password" autoComplete="new-password" minLength={6} />
          </label>
          <label>
            Confirm new password
            <PasswordInput name="confirmPassword" autoComplete="new-password" minLength={6} />
          </label>
          <button className="button primary" type="submit">
            Save new password
          </button>
        </form>
      </article>
    </>
  );
}
