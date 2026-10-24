"use client";

import {
  ArrowRight,
  BarChart3,
  Bell,
  BookMarked,
  BookOpen,
  CalendarDays,
  Check,
  CheckCircle2,
  ChevronDown,
  ChevronRight,
  CircleAlert,
  Clock3,
  Download,
  Eye,
  File,
  FileText,
  HelpCircle,
  Home,
  LayoutDashboard,
  Library,
  LockKeyhole,
  LogOut,
  Menu,
  MessageSquareText,
  NotebookPen,
  PanelLeftClose,
  PencilLine,
  Plus,
  Search,
  Settings,
  Sparkles,
  Target,
  TestTube2,
  TrendingUp,
  Upload,
  UserRound,
  UsersRound,
  X,
} from "lucide-react";
import Image from "next/image";
import { useMemo, useState } from "react";
import {
  dailyQuiz,
  physicsClasses,
  Role,
  scoreTrend,
  topicPerformance,
  weeklyTests,
} from "@/lib/data";

type View =
  | "overview"
  | "classes"
  | "notes"
  | "daily-quiz"
  | "weekly-tests"
  | "progress"
  | "content-studio"
  | "assessment-studio"
  | "review-queue"
  | "parent-summary";

const navByRole: Record<Role, { label: string; view: View; icon: typeof Home }[]> = {
  student: [
    { label: "Overview", view: "overview", icon: Home },
    { label: "My classes", view: "classes", icon: BookOpen },
    { label: "Notes & resources", view: "notes", icon: Library },
    { label: "Daily quizzes", view: "daily-quiz", icon: TestTube2 },
    { label: "Weekly tests", view: "weekly-tests", icon: CalendarDays },
    { label: "My progress", view: "progress", icon: BarChart3 },
  ],
  teacher: [
    { label: "Dashboard", view: "overview", icon: LayoutDashboard },
    { label: "Content studio", view: "content-studio", icon: BookOpen },
    { label: "Assessment studio", view: "assessment-studio", icon: PencilLine },
    { label: "Review queue", view: "review-queue", icon: MessageSquareText },
    { label: "Student progress", view: "progress", icon: TrendingUp },
  ],
  parent: [
    { label: "Weekly summary", view: "parent-summary", icon: Home },
    { label: "Learning progress", view: "progress", icon: TrendingUp },
    { label: "Tests & attendance", view: "weekly-tests", icon: CalendarDays },
  ],
};

const roleLabel: Record<Role, string> = {
  student: "Student",
  teacher: "Teacher",
  parent: "Parent",
};

function Brand() {
  return (
    <div className="brand">
      <span className="brand-mark"><Image src="/smxphysics-logo.png" alt="" width={36} height={36} /></span>
      <div>
        <strong>SMXPhysics</strong>
        <span>Excellence Programme</span>
      </div>
    </div>
  );
}

function Sidebar({
  role,
  view,
  setView,
  collapsed,
  setCollapsed,
  mobileOpen,
  setMobileOpen,
}: {
  role: Role;
  view: View;
  setView: (view: View) => void;
  collapsed: boolean;
  setCollapsed: (value: boolean) => void;
  mobileOpen: boolean;
  setMobileOpen: (value: boolean) => void;
}) {
  return (
    <>
      {mobileOpen && <button className="mobile-overlay" aria-label="Close navigation" onClick={() => setMobileOpen(false)} />}
      <aside className={`sidebar ${collapsed ? "collapsed" : ""} ${mobileOpen ? "mobile-open" : ""}`}>
        <div className="sidebar-top">
          <Brand />
          <button className="icon-button mobile-close" onClick={() => setMobileOpen(false)} aria-label="Close menu"><X size={19} /></button>
        </div>
        <div className="programme-pill">
          <span className="live-dot" />
          <div><strong>Class 11 Physics</strong><span>CBSE + NEET 2026–27</span></div>
        </div>
        <nav className="nav-list" aria-label={`${roleLabel[role]} navigation`}>
          <span className="nav-kicker">{roleLabel[role]} workspace</span>
          {navByRole[role].map((item) => {
            const Icon = item.icon;
            return (
              <button
                key={item.view}
                className={`nav-item ${view === item.view ? "active" : ""}`}
                onClick={() => { setView(item.view); setMobileOpen(false); }}
                title={collapsed ? item.label : undefined}
              >
                <Icon size={19} />
                <span>{item.label}</span>
                {item.view === "review-queue" && <b className="nav-count">4</b>}
              </button>
            );
          })}
        </nav>
        <div className="sidebar-bottom">
          <button className="nav-item"><Settings size={19} /><span>Settings</span></button>
          <div className="sidebar-user">
            <span className="avatar">AS</span>
            <div><strong>Arjun Sharma</strong><span>{roleLabel[role]} account</span></div>
            <LogOut size={17} />
          </div>
          <button className="collapse-button" onClick={() => setCollapsed(!collapsed)}>
            <PanelLeftClose size={17} />
            <span>Collapse menu</span>
          </button>
        </div>
      </aside>
    </>
  );
}

function Header({
  role,
  setRole,
  openMenu,
}: {
  role: Role;
  setRole: (role: Role) => void;
  openMenu: () => void;
}) {
  return (
    <header className="topbar">
      <button className="icon-button menu-button" onClick={openMenu} aria-label="Open navigation"><Menu size={20} /></button>
      <div className="search-box">
        <Search size={18} />
        <input aria-label="Search classes, notes, tests" placeholder="Search classes, notes, tests..." />
        <kbd>⌘ K</kbd>
      </div>
      <div className="topbar-actions">
        <span className="sync-status"><CheckCircle2 size={15} /> Synced</span>
        <button className="icon-button notification-button" aria-label="Notifications"><Bell size={19} /><span /></button>
        <label className="role-switcher">
          <span className={`role-icon ${role}`}><UserRound size={16} /></span>
          <select value={role} onChange={(event) => setRole(event.target.value as Role)} aria-label="Preview role">
            <option value="student">Student view</option>
            <option value="teacher">Teacher view</option>
            <option value="parent">Parent view</option>
          </select>
          <ChevronDown size={15} />
        </label>
      </div>
    </header>
  );
}

function PageTitle({ eyebrow, title, description, actions }: { eyebrow: string; title: string; description: string; actions?: React.ReactNode }) {
  return (
    <div className="page-title">
      <div><span className="eyebrow">{eyebrow}</span><h1>{title}</h1><p>{description}</p></div>
      {actions && <div className="page-actions">{actions}</div>}
    </div>
  );
}

function ProgressRing({ value, label }: { value: number; label?: string }) {
  return (
    <div className="progress-ring" style={{ "--value": `${value * 3.6}deg` } as React.CSSProperties}>
      <div><strong>{value}%</strong>{label && <span>{label}</span>}</div>
    </div>
  );
}

function TrendChart() {
  const max = 100;
  const points = scoreTrend.map((point, index) => {
    const x = 20 + index * 70;
    const y = 105 - (point.score / max) * 80;
    return `${x},${y}`;
  }).join(" ");
  return (
    <div className="trend-chart">
      <svg viewBox="0 0 320 120" role="img" aria-label="Quiz score trend increasing from 64 to 82 percent">
        <line x1="12" y1="25" x2="308" y2="25" />
        <line x1="12" y1="65" x2="308" y2="65" />
        <line x1="12" y1="105" x2="308" y2="105" />
        <polyline points={points} />
        {scoreTrend.map((point, index) => (
          <g key={point.label}>
            <circle cx={20 + index * 70} cy={105 - (point.score / max) * 80} r="4" />
            <text x={20 + index * 70} y="118" textAnchor="middle">{point.label}</text>
          </g>
        ))}
      </svg>
    </div>
  );
}

function StudentOverview({ setView, startQuiz }: { setView: (view: View) => void; startQuiz: () => void }) {
  const current = physicsClasses[4];
  return (
    <>
      <PageTitle eyebrow="Wednesday, 9 September" title="Good evening, Arjun" description="Your next step is ready. Keep the learning loop moving." />
      <section className="hero-grid">
        <article className="continue-card">
          <div className="continue-head"><span className="status-chip current"><span /> Continue learning</span><span>Lecture 05 · 68% complete</span></div>
          <div className="continue-body">
            <div className="class-number">05</div>
            <div>
              <span className="content-label">MATHEMATICAL TOOLS</span>
              <h2>{current.title}</h2>
              <p>Resume with geometric progression and graph basics. Your class notes and daily quiz are attached.</p>
              <div className="lesson-progress"><span style={{ width: "68%" }} /></div>
              <div className="continue-actions">
                <button className="button primary" onClick={() => setView("classes")}>Continue class <ArrowRight size={17} /></button>
                <button className="button secondary" onClick={() => setView("notes")}><FileText size={17} /> Open notes</button>
              </div>
            </div>
          </div>
        </article>
        <article className="quiz-callout">
          <div className="quiz-callout-icon"><TestTube2 size={23} /></div>
          <span className="small-label">TODAY&apos;S DAILY QUIZ</span>
          <h3>Phase, A.P., G.P. & graphs</h3>
          <p>5 questions · 12 minutes</p>
          <div className="due-row"><Clock3 size={15} /> Due today at 9:00 PM</div>
          <button className="button light full" onClick={startQuiz}>Start daily quiz <ArrowRight size={17} /></button>
        </article>
      </section>

      <section className="metric-grid">
        <article className="metric-card"><span className="metric-icon blue"><BookOpen size={20} /></span><div><span>Classes completed</span><strong>4 <small>/ 6</small></strong><p><TrendingUp size={14} /> 2 this week</p></div></article>
        <article className="metric-card"><span className="metric-icon green"><Target size={20} /></span><div><span>Average accuracy</span><strong>76%</strong><p><TrendingUp size={14} /> Up 8% this week</p></div></article>
        <article className="metric-card"><span className="metric-icon amber"><Clock3 size={20} /></span><div><span>Study time</span><strong>6h 40m</strong><p>Weekly goal: 8 hours</p></div></article>
        <article className="metric-card"><span className="metric-icon violet"><NotebookPen size={20} /></span><div><span>Errors to retry</span><strong>4</strong><p className="warn-text">2 due by tomorrow</p></div></article>
      </section>

      <section className="two-column">
        <article className="panel">
          <div className="panel-head"><div><span className="small-label">YOUR PATH</span><h3>Current learning journey</h3></div><button className="text-button" onClick={() => setView("classes")}>View all classes <ArrowRight size={15} /></button></div>
          <div className="path-list">
            {physicsClasses.slice(2, 6).map((item) => (
              <button className={`path-item ${item.status}`} key={item.id} onClick={() => item.status !== "upcoming" && setView("classes")}>
                <span className="path-state">{item.status === "mastered" || item.status === "practiced" ? <Check size={16} /> : item.status === "upcoming" ? <LockKeyhole size={15} /> : item.id}</span>
                <div><span>{item.lecture} · {item.date}</span><strong>{item.title}</strong></div>
                {item.score && <b>{item.score}%</b>}<ChevronRight size={17} />
              </button>
            ))}
          </div>
        </article>
        <article className="panel">
          <div className="panel-head"><div><span className="small-label">PERFORMANCE</span><h3>Quiz score trend</h3></div><span className="trend-pill"><TrendingUp size={14} /> +18%</span></div>
          <TrendChart />
          <div className="insight"><Sparkles size={17} /><p><strong>Good momentum.</strong> Phase is improving. Review maximum–minimum before the weekly test.</p></div>
        </article>
      </section>
    </>
  );
}

function ClassLibrary({ notify }: { notify: (message: string) => void }) {
  const [selected, setSelected] = useState(5);
  const selectedClass = physicsClasses.find((item) => item.id === selected) ?? physicsClasses[4];
  return (
    <>
      <PageTitle eyebrow="My classes" title="Classroom library" description="Every lecture, note and assignment—organized in the order it was taught." actions={<button className="button secondary"><CalendarDays size={17} /> Full schedule</button>} />
      <div className="class-layout">
        <aside className="class-list">
          <div className="class-list-head"><span>MATHEMATICAL TOOLS</span><b>5 of 6 classes</b></div>
          {physicsClasses.map((item) => (
            <button key={item.id} className={`class-list-item ${selected === item.id ? "selected" : ""}`} onClick={() => setSelected(item.id)}>
              <span className={`class-dot ${item.status}`}>{item.status === "mastered" ? <Check size={14} /> : item.id}</span>
              <div><span>{item.lecture} · {item.date}</span><strong>{item.title}</strong></div>
              <ChevronRight size={16} />
            </button>
          ))}
          <div className="locked-unit"><LockKeyhole size={17} /><div><strong>Units & Measurements</strong><span>Unlocks after Maths Tools</span></div></div>
        </aside>
        <section className="class-detail">
          <div className="class-detail-hero">
            <div className="class-hero-meta"><span className="status-chip current"><span /> {selectedClass.status.replace("-", " ")}</span><span>{selectedClass.lecture} · {selectedClass.date} · {selectedClass.duration}</span></div>
            <span className="content-label">{selectedClass.chapter}</span>
            <h2>{selectedClass.title}</h2>
            <p>{selectedClass.description}</p>
            <div className="tag-row">{selectedClass.tags.map((tag) => <span className="tag" key={tag}>{tag}</span>)}</div>
          </div>
          <div className="resource-section">
            <div className="section-head"><div><h3>Class materials</h3><p>Published by your teacher</p></div><span>{selectedClass.materials.length} resources</span></div>
            <div className="resource-list">
              {selectedClass.materials.map((material) => (
                <article className="resource-row" key={material.id}>
                  <span className={`file-icon ${material.kind === "Short notes" ? "note" : ""}`}><FileText size={21} /></span>
                  <div><span>{material.kind}{material.updated && <b className="updated-badge">UPDATED</b>}</span><strong>{material.title}</strong><small>{material.format}{material.size ? ` · ${material.size}` : " · Read in app"}</small></div>
                  <button className="icon-button" title="Preview" onClick={() => notify(`Opening ${material.title}`)}><Eye size={18} /></button>
                  {material.format !== "Note" && <button className="icon-button" title="Download" onClick={() => notify("Download will use the teacher-uploaded file after Supabase is connected.")}><Download size={18} /></button>}
                </article>
              ))}
            </div>
          </div>
          <div className="short-note">
            <div><span className="small-label">SHORT NOTES · LECTURE {String(selectedClass.id).padStart(2, "0")}</span><h3>Lock-lines for revision</h3></div>
            <div className="formula-grid">
              <p><span>01</span><b>Amplitude</b><strong>R = √(a² + b²)</strong></p>
              <p><span>02</span><b>Arithmetic progression</b><strong>Tₙ = a + (n−1)d</strong></p>
              <p><span>03</span><b>Infinite G.P.</b><strong>S∞ = a/(1−r), |r| &lt; 1</strong></p>
              <p><span>04</span><b>Graph rule</b><strong>y = constant → horizontal</strong></p>
            </div>
          </div>
        </section>
      </div>
    </>
  );
}

function QuizView({ onFinish }: { onFinish: (score: number) => void }) {
  const [answers, setAnswers] = useState<Record<number, number>>({});
  const [submitted, setSubmitted] = useState(false);
  const score = useMemo(() => dailyQuiz.questions.filter((question) => answers[question.id] === question.answer).length, [answers]);
  if (submitted) {
    const percentage = Math.round((score / dailyQuiz.questions.length) * 100);
    return (
      <div className="quiz-result">
        <span className="result-icon"><CheckCircle2 size={32} /></span>
        <span className="eyebrow">Attempt submitted</span>
        <h1>{percentage >= 70 ? "Daily goal achieved" : "Review before moving on"}</h1>
        <p>You answered {score} of {dailyQuiz.questions.length} questions correctly.</p>
        <ProgressRing value={percentage} label="accuracy" />
        <div className="result-summary">
          <div><span>Correct</span><strong>{score}</strong></div><div><span>Needs retry</span><strong>{dailyQuiz.questions.length - score}</strong></div><div><span>Time</span><strong>08:42</strong></div>
        </div>
        <div className="answer-review">
          {dailyQuiz.questions.map((question, index) => {
            const correct = answers[question.id] === question.answer;
            return <article key={question.id} className={correct ? "correct" : "incorrect"}><span>{correct ? <Check size={16} /> : <X size={16} />}</span><div><strong>Question {index + 1} · {question.topic}</strong><p>{question.explanation}</p></div></article>;
          })}
        </div>
        <button className="button primary" onClick={() => onFinish(percentage)}>Return to overview <ArrowRight size={17} /></button>
      </div>
    );
  }
  return (
    <>
      <PageTitle eyebrow="Daily assessment" title={dailyQuiz.title} description={`${dailyQuiz.subtitle} · ${dailyQuiz.duration} · Due ${dailyQuiz.due}`} />
      <div className="quiz-layout">
        <section className="quiz-paper">
          <div className="quiz-progress-head"><span>Question progress</span><strong>{Object.keys(answers).length} / {dailyQuiz.questions.length} answered</strong></div>
          <div className="quiz-progress"><span style={{ width: `${Object.keys(answers).length / dailyQuiz.questions.length * 100}%` }} /></div>
          {dailyQuiz.questions.map((question, index) => (
            <article className="question-card" key={question.id}>
              <div className="question-meta"><span>QUESTION {index + 1} OF {dailyQuiz.questions.length}</span><b>{question.topic}</b></div>
              <h3>{question.prompt}</h3>
              <div className="option-list">
                {question.options.map((option, optionIndex) => (
                  <button key={option} className={answers[question.id] === optionIndex ? "selected" : ""} onClick={() => setAnswers({ ...answers, [question.id]: optionIndex })}>
                    <span className="option-letter">{String.fromCharCode(65 + optionIndex)}</span>{option}{answers[question.id] === optionIndex && <Check size={17} />}
                  </button>
                ))}
              </div>
            </article>
          ))}
          <button className="button primary submit-quiz" disabled={Object.keys(answers).length !== dailyQuiz.questions.length} onClick={() => setSubmitted(true)}>Submit daily quiz <ArrowRight size={17} /></button>
        </section>
        <aside className="quiz-sidebar">
          <div className="timer"><Clock3 size={18} /><div><span>Time remaining</span><strong>11:42</strong></div></div>
          <h4>Question navigator</h4>
          <div className="question-nav">{dailyQuiz.questions.map((q, i) => <a href={`#q-${q.id}`} className={answers[q.id] !== undefined ? "answered" : ""} key={q.id}>{i + 1}</a>)}</div>
          <div className="quiz-info"><CircleAlert size={17} /><p>Your answers save automatically. Explanations appear after submission.</p></div>
        </aside>
      </div>
    </>
  );
}

function WeeklyTests({ role }: { role: Role }) {
  return (
    <>
      <PageTitle eyebrow={role === "parent" ? "Assessment record" : "Assessments"} title={role === "parent" ? "Tests and attendance" : "Weekly tests"} description={role === "parent" ? "Released results and completion records for your child." : "Cumulative CBSE + NEET assessments with earlier-topic retrieval."} />
      <div className="test-grid">
        {weeklyTests.map((test) => (
          <article className={`test-card ${test.status}`} key={test.id}>
            <div className="test-card-top"><span className="metric-icon violet"><CalendarDays size={20} /></span><span className={`status-chip ${test.status}`}>{test.status}</span></div>
            <span className="small-label">{test.scope}</span><h3>{test.title}</h3>
            <div className="test-facts"><span><CalendarDays size={15} />{test.date}</span><span><Clock3 size={15} />{test.duration}</span><span><FileText size={15} />{test.questions} questions · {test.marks} marks</span></div>
            {test.score ? <div className="released-score"><div><span>Released score</span><strong>{test.score}<small> / {test.marks}</small></strong></div><ProgressRing value={Math.round(test.score / test.marks * 100)} /></div> : <div className="scope-note"><BookMarked size={17} /><p>Includes NEET objective questions and CBSE written numericals.</p></div>}
            <button className={`button ${test.status === "scheduled" ? "primary" : "secondary"} full`}>{test.status === "scheduled" ? "View test instructions" : "Review result"} <ArrowRight size={16} /></button>
          </article>
        ))}
      </div>
    </>
  );
}

function ProgressView({ role }: { role: Role }) {
  return (
    <>
      <PageTitle eyebrow="Learning analytics" title={role === "parent" ? "Arjun’s learning progress" : role === "teacher" ? "Student progress" : "My progress"} description="Evidence from completed lessons, daily quizzes, retries and weekly tests." />
      <section className="progress-hero">
        <div><span className="small-label">MATHEMATICAL TOOLS</span><h2>Foundation mastery</h2><p>4 of 6 classes completed · 1 weekly assessment released</p></div>
        <ProgressRing value={74} label="mastery" />
        <div className="progress-summary"><div><span>Daily quiz average</span><strong>76%</strong></div><div><span>Completion</span><strong>92%</strong></div><div><span>Current streak</span><strong>5 days</strong></div></div>
      </section>
      <section className="two-column">
        <article className="panel">
          <div className="panel-head"><div><span className="small-label">TOPIC MASTERY</span><h3>Strengths and priorities</h3></div></div>
          <div className="topic-list">
            {topicPerformance.map((topic) => <div key={topic.topic}><div><strong>{topic.topic}</strong><span>{topic.status}</span></div><div className="topic-bar"><span style={{ width: `${topic.accuracy}%` }} /></div><b>{topic.accuracy}%</b></div>)}
          </div>
        </article>
        <article className="panel">
          <div className="panel-head"><div><span className="small-label">ERROR LOG</span><h3>Correction priorities</h3></div><span className="nav-count">4 open</span></div>
          <div className="error-list">
            <div><span className="error-icon formula">F</span><p><strong>Formula selection</strong><small>Maximum–minimum · retry tomorrow</small></p><b>2</b></div>
            <div><span className="error-icon concept">C</span><p><strong>Concept confusion</strong><small>Phase conversion · due today</small></p><b>1</b></div>
            <div><span className="error-icon reading">R</span><p><strong>Question reading</strong><small>G.P. convergence condition</small></p><b>1</b></div>
          </div>
          {role !== "parent" && <button className="button secondary full">Open 48-hour retry set <ArrowRight size={16} /></button>}
        </article>
      </section>
    </>
  );
}

function TeacherOverview({ setView }: { setView: (view: View) => void }) {
  return (
    <>
      <PageTitle eyebrow="Teacher workspace" title="Good evening, Dr. Sharma" description="Four actions need your attention before the next class." actions={<button className="button primary" onClick={() => setView("content-studio")}><Plus size={17} /> Create class</button>} />
      <div className="attention-banner"><span><CircleAlert size={19} /></span><div><strong>Next class preparation</strong><p>Lecture 06 slides are saved as a draft. Add short notes and publish before class.</p></div><button className="button light" onClick={() => setView("content-studio")}>Continue preparation <ArrowRight size={16} /></button></div>
      <section className="metric-grid">
        <article className="metric-card"><span className="metric-icon blue"><UsersRound size={20} /></span><div><span>Active students</span><strong>1</strong><p>100% active this week</p></div></article>
        <article className="metric-card"><span className="metric-icon green"><CheckCircle2 size={20} /></span><div><span>Daily quiz completion</span><strong>92%</strong><p>11 of 12 attempts</p></div></article>
        <article className="metric-card"><span className="metric-icon amber"><MessageSquareText size={20} /></span><div><span>Answers to review</span><strong>4</strong><p className="warn-text">Before next lecture</p></div></article>
        <article className="metric-card"><span className="metric-icon violet"><FileText size={20} /></span><div><span>Published resources</span><strong>17</strong><p>3 updated this week</p></div></article>
      </section>
      <section className="two-column teacher-columns">
        <article className="panel">
          <div className="panel-head"><div><span className="small-label">NEXT-DAY REVIEW QUEUE</span><h3>Start tomorrow’s class here</h3></div><button className="text-button" onClick={() => setView("review-queue")}>Open queue <ArrowRight size={15} /></button></div>
          <div className="review-list">
            <div><span className="priority high">HIGH</span><div><strong>Maximum of a sin θ + b cos θ</strong><p>Missed in 2 attempts · formula-selection error</p></div><b>58%</b></div>
            <div><span className="priority medium">RETRY</span><div><strong>Convert cosine before phase comparison</strong><p>48-hour correction due tomorrow</p></div><b>64%</b></div>
            <div><span className="priority low">CHECK</span><div><strong>G.P. convergence condition</strong><p>High-confidence wrong answer</p></div><b>72%</b></div>
          </div>
        </article>
        <article className="panel">
          <div className="panel-head"><div><span className="small-label">CLASS PERFORMANCE</span><h3>Daily quiz accuracy</h3></div><span className="trend-pill"><TrendingUp size={14} /> +8%</span></div>
          <TrendChart />
          <div className="teacher-quick-actions"><button onClick={() => setView("assessment-studio")}><Plus size={17} /> Create daily quiz</button><button onClick={() => setView("content-studio")}><Upload size={17} /> Upload notes</button></div>
        </article>
      </section>
    </>
  );
}

function ContentStudio({ notify }: { notify: (message: string) => void }) {
  const [files, setFiles] = useState<string[]>([]);
  const [published, setPublished] = useState(false);
  return (
    <>
      <PageTitle eyebrow="Teacher tools" title="Content studio" description="Build a complete class page, preview it as a student, then publish." actions={<><button className="button secondary"><Eye size={17} /> Student preview</button><button className="button primary" onClick={() => { setPublished(true); notify("Lecture 06 is now published to the student."); }}><Check size={17} /> Publish class</button></>} />
      {published && <div className="success-banner"><CheckCircle2 size={18} /><strong>Lecture 06 published.</strong> The student can now access its materials.</div>}
      <div className="studio-layout">
        <section className="studio-main">
          <article className="form-panel">
            <div className="form-panel-head"><span>1</span><div><h3>Class details</h3><p>Used throughout the student timeline and reports.</p></div><span className="status-chip draft">Draft</span></div>
            <div className="form-grid"><label>Class title<input defaultValue="Straight-line graphs" /></label><label>Class number<input defaultValue="Lecture 06" /></label><label>Chapter<select defaultValue="maths"><option value="maths">Mathematical Tools</option><option>Units & Measurements</option></select></label><label>Class date<input type="date" defaultValue="2026-09-10" /></label><label className="full-field">Description<textarea defaultValue="Coordinates, distance, slope, straight-line equations and Physics graph applications." /></label></div>
          </article>
          <article className="form-panel">
            <div className="form-panel-head"><span>2</span><div><h3>Upload class materials</h3><p>PDF, PPTX and images · maximum 50 MB per file.</p></div></div>
            <label className="upload-zone"><Upload size={28} /><strong>Drop teaching files here</strong><span>or click to browse from your computer</span><input type="file" multiple accept=".pdf,.ppt,.pptx,.png,.jpg,.jpeg" onChange={(event) => setFiles([...files, ...Array.from(event.target.files ?? []).map((file) => file.name)])} /></label>
            <div className="uploaded-files">
              <div><span className="file-icon"><File size={19} /></span><p><strong>Lecture_06_PW_Class05_Straight_Line_Graphs.pptx</strong><small>Slides · 10.8 MB · Ready</small></p><CheckCircle2 size={18} /></div>
              {files.map((file) => <div key={file}><span className="file-icon"><File size={19} /></span><p><strong>{file}</strong><small>New upload · Ready</small></p><CheckCircle2 size={18} /></div>)}
            </div>
          </article>
          <article className="form-panel">
            <div className="form-panel-head"><span>3</span><div><h3>Short notes</h3><p>Write the compact revision card students see below the class files.</p></div></div>
            <div className="editor-toolbar"><button><strong>B</strong></button><button><em>I</em></button><button>H₂</button><button>∑ Formula</button><button>NEET trap</button><button>CBSE</button></div>
            <textarea className="note-editor" defaultValue={"Slope = rise/run = Δy/Δx\n\nStraight-line equation: y = mx + c\n\nNEET trap: slope is tan θ, not θ."} />
          </article>
        </section>
        <aside className="studio-sidebar">
          <article className="publish-card"><h3>Publishing</h3><label>Status<select><option>Draft</option><option>Published</option></select></label><label>Release<select><option>Publish immediately</option><option>Schedule date and time</option></select></label><label className="check-row"><input type="checkbox" defaultChecked /> Notify student when published</label><hr /><div><span>Last saved</span><strong>Just now</strong></div><div><span>Created by</span><strong>Dr. Sharma</strong></div></article>
          <article className="checklist-card"><h3>Class readiness</h3><p><CheckCircle2 size={16} /> Class details complete</p><p><CheckCircle2 size={16} /> Lecture slides uploaded</p><p><CircleAlert size={16} /> Add class notes PDF</p><p><CheckCircle2 size={16} /> Short notes added</p><p><CircleAlert size={16} /> Link a daily quiz</p></article>
        </aside>
      </div>
    </>
  );
}

function AssessmentStudio({ notify }: { notify: (message: string) => void }) {
  const [questionCount, setQuestionCount] = useState(5);
  return (
    <>
      <PageTitle eyebrow="Teacher tools" title="Assessment studio" description="Create daily quizzes and weekly CBSE + NEET tests from one question bank." actions={<button className="button primary" onClick={() => notify("Assessment saved as a scheduled draft.")}><Check size={17} /> Save assessment</button>} />
      <div className="assessment-tabs"><button className="active">Daily quiz</button><button>Weekly test</button><button>Question bank <span>45</span></button></div>
      <div className="studio-layout">
        <section className="studio-main">
          <article className="form-panel">
            <div className="form-panel-head"><span>1</span><div><h3>Assessment settings</h3><p>Linked to Lecture 06 · Straight-line graphs.</p></div></div>
            <div className="form-grid"><label>Quiz title<input defaultValue="Lecture 06 daily quiz" /></label><label>Duration<select><option>12 minutes</option><option>15 minutes</option></select></label><label>Opens<input type="datetime-local" defaultValue="2026-09-10T18:00" /></label><label>Due<input type="datetime-local" defaultValue="2026-09-10T21:00" /></label></div>
            <div className="assessment-options"><label><input type="checkbox" defaultChecked /> Show explanations after submission</label><label><input type="checkbox" defaultChecked /> Add wrong answers to 48-hour retry</label><label><input type="checkbox" /> Randomize answer options</label></div>
          </article>
          <article className="form-panel">
            <div className="form-panel-head"><span>2</span><div><h3>Questions</h3><p>70% current lesson · 30% spaced retrieval.</p></div><button className="button secondary small" onClick={() => setQuestionCount(questionCount + 1)}><Plus size={15} /> Add question</button></div>
            <div className="question-builder-list">
              {Array.from({ length: questionCount }, (_, index) => (
                <div key={index}><span>{index + 1}</span><div><strong>{index < 3 ? ["What is the slope of a line making angle θ with +x-axis?", "Which equation represents a horizontal line?", "For y = 2x − 4, identify slope and y-intercept."][index] : index === 3 ? "A.P. retrieval: find T₁₀ when a = 2, d = 3." : index === 4 ? "Phase retrieval: convert cosine before comparing." : "New question — click to edit"}</strong><small>{index < 3 ? "Current lesson" : "Spaced retrieval"} · MCQ · {index < 3 ? "NEET" : "Foundation"}</small></div><PencilLine size={16} /></div>
              ))}
            </div>
          </article>
        </section>
        <aside className="studio-sidebar">
          <article className="publish-card"><h3>Quiz composition</h3><div className="composition-ring"><ProgressRing value={60} /></div><div><span>Current lesson</span><strong>60%</strong></div><div><span>Spaced retrieval</span><strong>40%</strong></div><hr /><div><span>Questions</span><strong>{questionCount}</strong></div><div><span>Total marks</span><strong>{questionCount * 4}</strong></div></article>
          <article className="scope-note"><HelpCircle size={18} /><p>Daily quizzes should diagnose learning, not block a struggling student. The chapter test controls the 70% mastery gate.</p></article>
        </aside>
      </div>
    </>
  );
}

function ReviewQueue() {
  return (
    <>
      <PageTitle eyebrow="Before the next class" title="Next-day review queue" description="Prioritized from wrong answers, slow correct responses and overdue retries." actions={<button className="button secondary"><Check size={17} /> Mark session reviewed</button>} />
      <div className="queue-banner"><div><strong>Recommended review time</strong><span>10 minutes at the start of Lecture 06</span></div><b>4 priority items</b></div>
      <section className="queue-list">
        {[
          ["01", "High priority", "Maximum value of a sin θ + b cos θ", "Formula-selection error repeated twice", "58%", "Review on board"],
          ["02", "48-hour retry", "Convert both functions before phase comparison", "Incorrect on L04 and L05 daily quizzes", "64%", "Retry due tomorrow"],
          ["03", "Confidence mismatch", "Condition for an infinite G.P. sum", "Student selected |r| > 1 with high confidence", "72%", "Ask for reasoning"],
          ["04", "Not attempted", "Graph: y = constant", "One daily-test item left unanswered", "—", "Assign quick check"],
        ].map((item) => <article key={item[0]}><span className="queue-number">{item[0]}</span><div className="queue-copy"><span>{item[1]}</span><h3>{item[2]}</h3><p>{item[3]}</p></div><strong className="queue-score">{item[4]}</strong><button className="button secondary small">{item[5]}</button></article>)}
      </section>
    </>
  );
}

function ParentSummary() {
  return (
    <>
      <PageTitle eyebrow="Week of 7–12 September" title="Arjun’s weekly summary" description="A concise view of consistency, learning progress and where support is needed." actions={<button className="button secondary"><Download size={17} /> Download report</button>} />
      <div className="parent-hero"><div><span className="status-chip current"><span /> On track</span><h2>A consistent week with improving accuracy</h2><p>Arjun completed all scheduled classes and is responding well to correction work. Maximum–minimum needs one focused review before Saturday’s test.</p></div><ProgressRing value={82} label="weekly score" /></div>
      <section className="metric-grid parent-metrics">
        <article className="metric-card"><span className="metric-icon green"><CheckCircle2 size={20} /></span><div><span>Attendance</span><strong>100%</strong><p>3 of 3 classes</p></div></article>
        <article className="metric-card"><span className="metric-icon blue"><TestTube2 size={20} /></span><div><span>Daily tests</span><strong>92%</strong><p>11 of 12 completed</p></div></article>
        <article className="metric-card"><span className="metric-icon amber"><BookOpen size={20} /></span><div><span>Homework</span><strong>83%</strong><p>5 of 6 submitted</p></div></article>
        <article className="metric-card"><span className="metric-icon violet"><Target size={20} /></span><div><span>Accuracy</span><strong>76%</strong><p><TrendingUp size={14} /> Up 8%</p></div></article>
      </section>
      <section className="two-column">
        <article className="panel"><div className="panel-head"><div><span className="small-label">TEACHER OBSERVATION</span><h3>This week’s feedback</h3></div></div><blockquote>“Arjun is confident with trigonometric identities and A.P./G.P. formulae. He should slow down when identifying the correct max–min formula. The assigned retry should resolve this before the weekly test.”</blockquote><div className="teacher-sign"><span className="avatar teacher">DS</span><div><strong>Dr. Sharma</strong><span>Physics teacher · updated today</span></div></div></article>
        <article className="panel"><div className="panel-head"><div><span className="small-label">THIS WEEK</span><h3>Strengths and support</h3></div></div><div className="support-list"><div className="positive"><Check size={17} /><p><strong>Strong areas</strong><span>Trigonometry · Radians · A.P. & G.P.</span></p></div><div className="attention"><CircleAlert size={17} /><p><strong>Needs attention</strong><span>Maximum–minimum · phase conversion</span></p></div><div className="neutral"><CalendarDays size={17} /><p><strong>Next milestone</strong><span>Weekly Test 01 · Saturday, 12 September</span></p></div></div></article>
      </section>
    </>
  );
}

export default function PhysicsPlatform() {
  const [role, setRole] = useState<Role>("student");
  const [view, setView] = useState<View>("overview");
  const [collapsed, setCollapsed] = useState(false);
  const [mobileOpen, setMobileOpen] = useState(false);
  const [toast, setToast] = useState("");

  const notify = (message: string) => {
    setToast(message);
    window.setTimeout(() => setToast(""), 3200);
  };
  const changeRole = (nextRole: Role) => {
    setRole(nextRole);
    setView(nextRole === "parent" ? "parent-summary" : "overview");
  };

  let content: React.ReactNode;
  if (view === "classes" || view === "notes") content = <ClassLibrary notify={notify} />;
  else if (view === "daily-quiz") content = <QuizView onFinish={() => setView("overview")} />;
  else if (view === "weekly-tests") content = <WeeklyTests role={role} />;
  else if (view === "progress") content = <ProgressView role={role} />;
  else if (view === "content-studio") content = <ContentStudio notify={notify} />;
  else if (view === "assessment-studio") content = <AssessmentStudio notify={notify} />;
  else if (view === "review-queue") content = <ReviewQueue />;
  else if (view === "parent-summary") content = <ParentSummary />;
  else if (role === "teacher") content = <TeacherOverview setView={setView} />;
  else if (role === "parent") content = <ParentSummary />;
  else content = <StudentOverview setView={setView} startQuiz={() => setView("daily-quiz")} />;

  return (
    <div className={`app-shell ${collapsed ? "sidebar-collapsed" : ""}`}>
      <Sidebar role={role} view={view} setView={setView} collapsed={collapsed} setCollapsed={setCollapsed} mobileOpen={mobileOpen} setMobileOpen={setMobileOpen} />
      <div className="app-main">
        <Header role={role} setRole={changeRole} openMenu={() => setMobileOpen(true)} />
        <main className="page-content">{content}</main>
      </div>
      {toast && <div className="toast"><CheckCircle2 size={18} /><span>{toast}</span></div>}
    </div>
  );
}
