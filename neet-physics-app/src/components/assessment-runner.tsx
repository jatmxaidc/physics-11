"use client";

import { ArrowRight, CheckCircle2, Clock3, EyeOff, ShieldCheck } from "lucide-react";
import { useEffect, useMemo, useRef, useState } from "react";
import { requestAssessmentAttendance } from "@/app/actions/student";
import { MathContent } from "@/components/math-content";
import { createClient } from "@/lib/supabase/client";

type RunnerQuestion = {
  id: string;
  prompt: string;
  question_type: string;
  options: string[] | null;
  topic: string;
  marks: number;
};

type StartResult = {
  attempt_id: string;
  started_at: string;
  duration_minutes: number;
  focus_leave_count?: number;
};

type SubmitResult = {
  attempt_id: string;
  status: "submitted" | "teacher_review" | "released";
  results_visible?: boolean;
  unlock_at?: string | null;
  assessment_type?: string;
  objective_score: number | null;
  correct_count: number | null;
  objective_count: number;
};

const GRACE_SECONDS = 120;

function storageKey(attemptId: string) {
  return `orbit-assessment-answers:${attemptId}`;
}

function formatUnlock(value: string | null | undefined) {
  if (!value) return null;
  return new Date(value).toLocaleString("en-IN", {
    day: "numeric",
    month: "short",
    year: "numeric",
    hour: "numeric",
    minute: "2-digit",
  });
}

export default function AssessmentRunner({
  assessment,
  questions,
  initialAttempt = null,
  attendance = null,
}: {
  assessment: {
    id: string;
    title: string;
    duration_minutes: number;
    total_marks: number;
    instructions: string | null;
    opens_at?: string | null;
    due_at?: string | null;
    assessment_type?: string;
  };
  questions: RunnerQuestion[];
  initialAttempt?: StartResult | null;
  attendance?: {
    required: boolean;
    status: "none" | "requested" | "verified" | "rejected";
    canStart: boolean;
  } | null;
}) {
  const [attempt, setAttempt] = useState<StartResult | null>(initialAttempt);
  const [answers, setAnswers] = useState<Record<string, unknown>>({});
  const [draftsReady, setDraftsReady] = useState(false);
  const [result, setResult] = useState<SubmitResult | null>(null);
  const [remaining, setRemaining] = useState(assessment.duration_minutes * 60);
  const [graceLeft, setGraceLeft] = useState(GRACE_SECONDS);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [focusLeaves, setFocusLeaves] = useState(initialAttempt?.focus_leave_count ?? 0);
  const [isAway, setIsAway] = useState(false);
  const autoSubmitted = useRef(false);
  const answersRef = useRef(answers);
  const awayRef = useRef(false);
  const attemptIdRef = useRef<string | null>(initialAttempt?.attempt_id ?? null);
  const recordingRef = useRef(false);
  answersRef.current = answers;
  attemptIdRef.current = attempt?.attempt_id ?? null;

  useEffect(() => {
    if (!attempt || result) return;
    const officialEnd = new Date(attempt.started_at).getTime() + attempt.duration_minutes * 60_000;
    const hardEnd = officialEnd + GRACE_SECONDS * 1000;
    const tick = () => {
      const now = Date.now();
      setRemaining(Math.max(0, Math.ceil((officialEnd - now) / 1000)));
      setGraceLeft(Math.max(0, Math.ceil((hardEnd - now) / 1000)));
    };
    tick();
    const timer = window.setInterval(tick, 1000);
    return () => window.clearInterval(timer);
  }, [attempt, result]);

  useEffect(() => {
    if (!attempt) {
      setDraftsReady(false);
      return;
    }
    let cancelled = false;
    try {
      const raw = window.localStorage.getItem(storageKey(attempt.attempt_id));
      if (raw) {
        const parsed = JSON.parse(raw) as Record<string, unknown>;
        if (!cancelled) setAnswers(parsed);
      }
    } catch {
      // ignore corrupt drafts
    } finally {
      if (!cancelled) setDraftsReady(true);
    }
    return () => {
      cancelled = true;
    };
  }, [attempt]);

  useEffect(() => {
    if (!attempt || result || !draftsReady) return;
    window.localStorage.setItem(storageKey(attempt.attempt_id), JSON.stringify(answers));
  }, [answers, attempt, result, draftsReady]);

  useEffect(() => {
    if (!attempt || result || !draftsReady || autoSubmitted.current) return;
    if (graceLeft > 0) return;
    autoSubmitted.current = true;
    void submit(true, answersRef.current);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [graceLeft, attempt, result, draftsReady]);

  useEffect(() => {
    if (!attempt || result) return;

    const armedAt = Date.now() + 900;

    async function recordLeave() {
      const attemptId = attemptIdRef.current;
      if (!attemptId || recordingRef.current) return;
      recordingRef.current = true;
      setFocusLeaves((count) => count + 1);
      try {
        const supabase = createClient();
        const { data, error: leaveError } = await supabase.rpc("record_attempt_focus_leave", {
          p_attempt_id: attemptId,
        });
        if (!leaveError && typeof data === "number") {
          setFocusLeaves(data);
        }
      } catch {
        // Keep local count if the network hiccups; next leave will sync again.
      } finally {
        recordingRef.current = false;
      }
    }

    function markLeft() {
      if (Date.now() < armedAt) return;
      if (awayRef.current || !attemptIdRef.current) return;
      awayRef.current = true;
      setIsAway(true);
      void recordLeave();
    }

    function markBack() {
      if (!awayRef.current) return;
      awayRef.current = false;
      setIsAway(false);
    }

    function onVisibility() {
      if (document.hidden) markLeft();
      else markBack();
    }

    function onBlur() {
      markLeft();
    }

    function onFocus() {
      if (!document.hidden) markBack();
    }

    document.addEventListener("visibilitychange", onVisibility);
    window.addEventListener("blur", onBlur);
    window.addEventListener("focus", onFocus);
    return () => {
      document.removeEventListener("visibilitychange", onVisibility);
      window.removeEventListener("blur", onBlur);
      window.removeEventListener("focus", onFocus);
    };
  }, [attempt, result]);

  const answered = useMemo(
    () => Object.values(answers).filter((answer) => answer !== "" && answer !== undefined).length,
    [answers],
  );

  async function start() {
    setBusy(true);
    setError("");
    setDraftsReady(false);
    const supabase = createClient();
    const { data, error: startError } = await supabase.rpc("start_assessment", {
      p_assessment_id: assessment.id,
    });
    if (startError) {
      setError(startError.message);
      setDraftsReady(false);
    } else {
      const started = data as StartResult;
      setAnswers({});
      setFocusLeaves(started.focus_leave_count ?? 0);
      setAttempt(started);
    }
    setBusy(false);
  }

  async function submit(fromAuto = false, overrideAnswers?: Record<string, unknown>) {
    if (!attempt) return;
    setBusy(true);
    setError("");
    const payload = overrideAnswers ?? answersRef.current;
    const supabase = createClient();
    const { data, error: submitError } = await supabase.rpc("submit_assessment", {
      p_attempt_id: attempt.attempt_id,
      p_answers: payload,
    });
    if (submitError) {
      setError(submitError.message);
      if (fromAuto) autoSubmitted.current = false;
    } else {
      setResult(data as SubmitResult);
      window.localStorage.removeItem(storageKey(attempt.attempt_id));
    }
    setBusy(false);
  }

  if (result) {
    const isDaily = (result.assessment_type ?? assessment.assessment_type) === "daily_quiz";
    const unlockLabel = formatUnlock(result.unlock_at ?? assessment.due_at);
    return (
      <div className="production-result">
        <span><CheckCircle2 size={31} /></span>
        <p className="eyebrow">SUBMISSION RECEIVED</p>
        <h1>Answers submitted</h1>
        <p>
          {isDaily
            ? unlockLabel
              ? `Your score and explanations unlock after the quiz deadline (${unlockLabel}).`
              : "Your score and explanations unlock after the quiz deadline (7:00 pm IST next day if no due time is set)."
            : "This is a one-time test. Your teacher will release the result later."}
        </p>
        <p className="assessment-focus-summary">
          Tab / app leaves during this attempt: <strong>{focusLeaves}</strong>
        </p>
        <a className="button primary" href={`/student/progress/${result.attempt_id}`}>
          View submission status <ArrowRight size={16} />
        </a>
      </div>
    );
  }

  if (!attempt) {
    const needsAttendance =
      Boolean(attendance?.required) ||
      (assessment.assessment_type != null && assessment.assessment_type !== "daily_quiz");
    const attendanceStatus = attendance?.status ?? "none";
    const canBegin = attendance ? attendance.canStart : !needsAttendance;
    const waiting = attendanceStatus === "requested";
    const rejected = attendanceStatus === "rejected";

    if (needsAttendance && !canBegin) {
      return (
        <div className="assessment-start-card">
          {waiting && (
            <p className="attendance-status-banner">Status: awaiting teacher verification</p>
          )}
          <span className="metric-icon amber"><ShieldCheck size={24} /></span>
          <p className="eyebrow">ATTENDANCE REQUIRED</p>
          <h1>{assessment.title}</h1>
          <p>
            {waiting
              ? "Attendance is marked. Waiting for your teacher to verify it before this test can start."
              : rejected
                ? "Attendance was not approved. Mark attendance again when you are ready."
                : "Mark attendance first. Your teacher must verify it before you can see questions or begin."}
          </p>
          <div>
            <span><Clock3 size={16} /> {assessment.duration_minutes} minutes</span>
            <span>{assessment.total_marks} marks</span>
            <span>{(assessment.assessment_type ?? "weekly_test").replace(/_/g, " ")}</span>
          </div>
          {error && <p className="login-error">{error}</p>}
          {!waiting && (
            <form action={requestAssessmentAttendance}>
              <input type="hidden" name="assessmentId" value={assessment.id} />
              <button className="button primary" type="submit">
                {rejected ? "Mark attendance again" : "Mark attendance"} <ArrowRight size={16} />
              </button>
            </form>
          )}
        </div>
      );
    }

    return (
      <div className="assessment-start-card">
        <span className="metric-icon blue"><ShieldCheck size={24} /></span>
        <p className="eyebrow">SECURE ASSESSMENT</p>
        <h1>{assessment.title}</h1>
        <p>{assessment.instructions || "Answer every question carefully. The timer starts only when you press Begin."}</p>
        <div>
          <span><Clock3 size={16} /> {assessment.duration_minutes} minutes</span>
          <span>{questions.length} questions</span>
          <span>{assessment.total_marks} marks</span>
        </div>
        <p className="assessment-window-note">
          Leaving this tab or app is counted and saved for your teacher.
        </p>
        {(assessment.opens_at || assessment.due_at) && (
          <p className="assessment-window-note">
            {assessment.opens_at ? `Opens ${new Date(assessment.opens_at).toLocaleString("en-IN")} · ` : ""}
            {assessment.due_at ? `Due ${new Date(assessment.due_at).toLocaleString("en-IN")}` : "No due time"}
            {assessment.assessment_type === "daily_quiz"
              ? " · Answers unlock after the deadline"
              : " · Result released later by teacher"}
          </p>
        )}
        {error && <p className="login-error">{error}</p>}
        <button className="button primary" onClick={start} disabled={busy || questions.length === 0}>
          {busy ? "Starting…" : "Begin assessment"} <ArrowRight size={16} />
        </button>
      </div>
    );
  }

  const inGrace = remaining === 0 && graceLeft > 0;

  return (
    <div className="production-quiz-layout">
      <section>
        <div className="production-quiz-head">
          <div>
            <span className="eyebrow">IN PROGRESS</span>
            <h1>{assessment.title}</h1>
            <p>
              {answered} of {questions.length} answered ·{" "}
              {draftsReady ? "drafts saved on this device" : "restoring draft…"}
            </p>
          </div>
          <div className="assessment-head-meters">
            <div className={`focus-leave-meter ${focusLeaves > 0 ? "warn" : ""} ${isAway ? "away" : ""}`}>
              <EyeOff size={18} />
              <div>
                <span>{isAway ? "You left the paper" : "Tab / app leaves"}</span>
                <strong>{focusLeaves}</strong>
              </div>
            </div>
            <div className={remaining < 120 || inGrace ? "timer urgent" : "timer"}>
              <Clock3 size={18} />
              <div>
                <span>{inGrace ? "Grace period" : "Time remaining"}</span>
                <strong>
                  {inGrace
                    ? `${String(Math.floor(graceLeft / 60)).padStart(2, "0")}:${String(graceLeft % 60).padStart(2, "0")}`
                    : `${String(Math.floor(remaining / 60)).padStart(2, "0")}:${String(remaining % 60).padStart(2, "0")}`}
                </strong>
              </div>
            </div>
          </div>
        </div>
        {isAway && (
          <p className="login-error">Return to this tab to continue. Each leave is counted and stored.</p>
        )}
        {questions.map((question, index) => (
          <article className="question-card production-question" key={question.id}>
            <div className="question-meta">
              <span>QUESTION {index + 1} OF {questions.length}</span>
              <b>{question.topic} · {question.marks} marks</b>
            </div>
            <MathContent as="h3" content={question.prompt} />
            {question.question_type === "short_answer" ? (
              <textarea
                value={(answers[question.id] as string) ?? ""}
                onChange={(event) => setAnswers({ ...answers, [question.id]: event.target.value })}
                placeholder="Write your complete answer…"
              />
            ) : (
              <div className="option-list">
                {(question.options ?? []).map((option, optionIndex) => (
                  <button
                    type="button"
                    key={optionIndex}
                    className={answers[question.id] === optionIndex ? "selected" : ""}
                    onClick={() => setAnswers({ ...answers, [question.id]: optionIndex })}
                  >
                    <span className="option-letter">{String.fromCharCode(65 + optionIndex)}</span>
                    <MathContent as="span" content={option} />
                  </button>
                ))}
              </div>
            )}
          </article>
        ))}
        {error && <p className="login-error">{error}</p>}
        <button
          className="button primary submit-production-assessment"
          onClick={() => void submit(false)}
          disabled={busy || !draftsReady || graceLeft === 0}
        >
          {busy ? "Submitting…" : inGrace ? "Submit now (grace period)" : "Submit final answers"}{" "}
          <ArrowRight size={16} />
        </button>
      </section>
      <aside className="production-quiz-side">
        <h3>Attempt status</h3>
        <p><ShieldCheck size={15} /> Answers draft locally until you submit.</p>
        <div className={`focus-leave-side ${focusLeaves > 0 ? "warn" : ""}`}>
          <EyeOff size={15} />
          <div>
            <strong>{focusLeaves}</strong>
            <span>tab / app leaves recorded</span>
          </div>
        </div>
        <div className="question-nav">
          {questions.map((question, index) => (
            <span className={answers[question.id] !== undefined ? "answered" : ""} key={question.id}>
              {index + 1}
            </span>
          ))}
        </div>
        <small>
          Stay on this page. Switching tabs or apps increases the leave count, which is saved to your attempt
          record. Submission is final. A two-minute grace period remains after the timer hits zero; then the
          paper auto-submits.
        </small>
      </aside>
    </div>
  );
}
