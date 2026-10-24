import { KeyRound, ShieldCheck } from "lucide-react";
import { changeInitialPassword } from "@/app/actions/auth";
import PasswordInput from "@/components/password-input";

export const dynamic = "force-dynamic";

export default async function InitialPasswordPage({
  searchParams,
}: {
  searchParams: Promise<{ error?: string }>;
}) {
  const params = await searchParams;
  return (
    <main className="password-page">
      <form action={changeInitialPassword} className="login-form">
        <span className="login-lock"><KeyRound size={21} /></span>
        <div><span className="eyebrow">FIRST LOGIN SECURITY</span><h2>Create your private password</h2><p>The temporary password cannot be used again after this step.</p></div>
        <div className="setup-warning"><ShieldCheck size={16} /><p>Use at least 6 characters. Choose a private password that you do not share with another user.</p></div>
        {params.error && <div className="login-error">{params.error}</div>}
        <label>New password<PasswordInput name="password" autoComplete="new-password" minLength={6} /></label>
        <label>Confirm new password<PasswordInput name="confirmPassword" autoComplete="new-password" minLength={6} /></label>
        <button className="button primary full">Save password and continue</button>
      </form>
    </main>
  );
}
