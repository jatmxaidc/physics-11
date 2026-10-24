import { ArrowRight, CheckCircle2, LockKeyhole } from "lucide-react";
import Image from "next/image";
import { login } from "@/app/actions/auth";
import PasswordInput from "@/components/password-input";

export default async function LoginPage({
  searchParams,
}: {
  searchParams: Promise<{ error?: string; next?: string }>;
}) {
  const params = await searchParams;
  const configured = Boolean(
    process.env.NEXT_PUBLIC_SUPABASE_URL &&
      process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY,
  );

  return (
    <main className="login-page">
      <section className="login-story">
        <div className="login-brand">
          <span><Image src="/smxphysics-logo.png" alt="" width={39} height={39} priority /></span>
          <div><strong>SMXPhysics</strong><small>Excellence Programme</small></div>
        </div>
        <div className="login-copy">
          <span>CLASS 11 · CBSE + NEET</span>
          <h1>Learn. Attempt.<br />Analyse. Improve.</h1>
          <p>A structured Physics programme connecting every class with notes, assessments and meaningful progress.</p>
          <div className="login-benefits">
            <p><CheckCircle2 size={17} /> Class-wise slides and short notes</p>
            <p><CheckCircle2 size={17} /> Daily quizzes and 48-hour correction</p>
            <p><CheckCircle2 size={17} /> Teacher analytics and parent summaries</p>
          </div>
        </div>
        <small className="login-footer">Class 11 Physics Excellence Programme · 2026–27</small>
      </section>
      <section className="login-panel">
        <form action={login} className="login-form">
          <span className="login-lock"><LockKeyhole size={20} /></span>
          <div><span className="eyebrow">SECURE SIGN IN</span><h2>Welcome back</h2><p>Use the account provided by your teacher.</p></div>
          {!configured && (
            <div className="setup-warning">
              <strong>Backend setup required</strong>
              <p>Add the Supabase values from <code>.env.example</code>. Until then, use the <a href="/preview">interface preview</a>.</p>
            </div>
          )}
          {params.error && <div className="login-error">{params.error}</div>}
          <input type="hidden" name="next" value={params.next ?? "/dashboard"} />
          <label>Email address<input name="email" type="email" autoComplete="email" placeholder="you@example.com" required /></label>
          <label>Password<PasswordInput name="password" autoComplete="current-password" placeholder="At least 6 characters" minLength={6} /></label>
          <div className="login-options"><span>Contact teacher for access</span></div>
          <button className="button primary full" disabled={!configured}>Sign in securely <ArrowRight size={17} /></button>
          <p className="privacy-note">Your learning records stay private to your linked teacher. Parent summaries appear only when a parent account is linked.</p>
        </form>
      </section>
    </main>
  );
}
