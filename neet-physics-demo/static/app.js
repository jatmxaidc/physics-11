const app = document.getElementById("app");
const toastRoot = document.getElementById("toast-root");

const state = {
  user: null,
  route: "home",
  routeParam: null,
  dashboard: null,
  quizStartedAt: null,
  quizTimer: null,
  selectedAnswers: {},
  quizData: null,
  lesson: null,
};

const icons = {
  home: "⌂",
  roadmap: "◫",
  notebook: "▤",
  errors: "△",
  dashboard: "◉",
  student: "◎",
  logout: "↗",
  menu: "☰",
};

function escapeHtml(value = "") {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function formatDate(value) {
  if (!value) return "Not yet";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleDateString("en-IN", { day: "numeric", month: "short", year: "numeric" });
}

function percentClass(value) {
  if (value < 50) return "accuracy-low";
  if (value < 70) return "accuracy-mid";
  return "accuracy-high";
}

async function api(path, options = {}) {
  const response = await fetch(path, {
    credentials: "same-origin",
    headers: { "Content-Type": "application/json", ...(options.headers || {}) },
    ...options,
  });
  if (response.status === 401) {
    state.user = null;
    renderLogin();
    throw new Error("Please sign in again");
  }
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(data.detail || "Something went wrong");
  return data;
}

function toast(message, type = "success") {
  const node = document.createElement("div");
  node.className = `toast ${type}`;
  node.textContent = message;
  toastRoot.appendChild(node);
  setTimeout(() => node.remove(), 3200);
}

function todayText() {
  return new Date().toLocaleDateString("en-IN", { weekday: "long", day: "numeric", month: "long" });
}

function brand() {
  return `<div class="brand-lockup"><div class="brand-mark">O</div><span>Orbit Physics</span></div>`;
}

function renderLogin(role = "student", error = "") {
  const isStudent = role === "student";
  app.innerHTML = `
    <div class="login-page">
      <section class="login-visual">
        ${brand()}
        <div class="hero-copy">
          <p class="eyebrow">Class 11 · CBSE + NEET</p>
          <h1>Learn. Test. Correct. Continue.</h1>
          <p>A complete local learning system built around daily practice, next-day correction and teacher-side topic analysis.</p>
          <div class="flow-preview">
            <span class="flow-chip">138 dated lectures</span>
            <span class="flow-chip">Daily 5–10 questions</span>
            <span class="flow-chip">48-hour retry</span>
            <span class="flow-chip">Teacher analytics</span>
          </div>
        </div>
        <div class="login-meta">Local demo · Data stays on this computer · SQLite powered</div>
      </section>
      <section class="login-panel">
        <form class="login-card" id="login-form">
          <h2>Welcome back</h2>
          <p>Choose a role to explore the complete student and teacher experience.</p>
          <div class="role-switch">
            <button type="button" class="${isStudent ? "active" : ""}" data-role="student">Student</button>
            <button type="button" class="${!isStudent ? "active" : ""}" data-role="teacher">Teacher</button>
          </div>
          ${error ? `<div class="error-banner">${escapeHtml(error)}</div>` : ""}
          <div class="form-group">
            <label>Email</label>
            <input class="input" id="email" type="email" value="${isStudent ? "student@physics.local" : "teacher@physics.local"}" required />
          </div>
          <div class="form-group">
            <label>Password</label>
            <input class="input" id="password" type="password" value="${isStudent ? "student123" : "teacher123"}" required />
          </div>
          <button class="btn btn-primary btn-block" type="submit">Open ${isStudent ? "student workspace" : "teacher dashboard"}</button>
          <div class="demo-credentials">
            <strong>Demo access</strong><br />
            Student: student@physics.local / student123<br />
            Teacher: teacher@physics.local / teacher123
          </div>
        </form>
      </section>
    </div>`;

  document.querySelectorAll("[data-role]").forEach((button) => {
    button.addEventListener("click", () => renderLogin(button.dataset.role));
  });
  document.getElementById("login-form").addEventListener("submit", async (event) => {
    event.preventDefault();
    const button = event.submitter;
    button.disabled = true;
    button.textContent = "Signing in…";
    try {
      const result = await api("/api/login", {
        method: "POST",
        body: JSON.stringify({
          email: document.getElementById("email").value,
          password: document.getElementById("password").value,
        }),
      });
      state.user = result.user;
      state.route = result.user.role === "teacher" ? "teacher" : "home";
      renderShell();
    } catch (error) {
      renderLogin(role, error.message);
    }
  });
}

function studentNav() {
  return [
    ["home", "Home", icons.home],
    ["roadmap", "Learning path", icons.roadmap],
    ["notebook", "Notebook", icons.notebook],
    ["errors", "Error log", icons.errors],
  ];
}

function teacherNav() {
  return [
    ["teacher", "Overview", icons.dashboard],
    ["student-detail", "Student progress", icons.student],
  ];
}

function shellMarkup(content = `<div class="boot-screen"><div class="boot-line"></div></div>`) {
  const navItems = state.user.role === "teacher" ? teacherNav() : studentNav();
  return `
    <div class="app-shell">
      <aside class="sidebar" id="sidebar">
        ${brand()}
        <div class="nav-section-label">${state.user.role === "teacher" ? "Teaching" : "Study space"}</div>
        <nav class="nav">
          ${navItems.map(([route, label, icon]) => `
            <button class="nav-button ${state.route === route ? "active" : ""}" data-nav="${route}">
              <span class="nav-icon">${icon}</span><span>${label}</span>
            </button>`).join("")}
        </nav>
        <div class="sidebar-footer">
          <div class="user-mini">
            <div class="avatar">${escapeHtml(state.user.name[0])}</div>
            <div class="user-mini-copy">
              <strong>${escapeHtml(state.user.name)}</strong>
              <span>${escapeHtml(state.user.role)}</span>
            </div>
          </div>
          <button class="nav-button" id="logout-button" style="margin-top:10px">
            <span class="nav-icon">${icons.logout}</span><span>Sign out</span>
          </button>
        </div>
      </aside>
      <main class="main">
        <header class="topbar">
          <div style="display:flex;align-items:center;gap:12px">
            <button class="btn btn-ghost btn-small mobile-menu" id="mobile-menu">${icons.menu}</button>
            <div>
              <div class="topbar-title">${state.user.role === "teacher" ? "Teacher command centre" : "Class 11 Physics"}</div>
              <div class="topbar-date">${todayText()}</div>
            </div>
          </div>
          <div class="top-actions">
            <a class="btn btn-secondary btn-small" href="/static/lecture.html" target="_blank">Lecture board</a>
            <span class="tag ${state.user.role === "teacher" ? "tag-cbse" : "tag-neet"}">${state.user.role === "teacher" ? "Live analytics" : "CBSE + NEET"}</span>
          </div>
        </header>
        <div class="content" id="page-content">${content}</div>
      </main>
    </div>`;
}

function bindShell() {
  document.querySelectorAll("[data-nav]").forEach((button) => {
    button.addEventListener("click", () => navigate(button.dataset.nav));
  });
  document.getElementById("logout-button")?.addEventListener("click", async () => {
    await api("/api/logout", { method: "POST" });
    state.user = null;
    renderLogin();
  });
  document.getElementById("mobile-menu")?.addEventListener("click", () => {
    document.getElementById("sidebar").classList.toggle("open");
  });
}

async function renderShell() {
  app.innerHTML = shellMarkup();
  bindShell();
  await renderRoute();
}

async function navigate(route, param = null) {
  clearInterval(state.quizTimer);
  state.route = route;
  state.routeParam = param;
  app.innerHTML = shellMarkup();
  bindShell();
  await renderRoute();
  document.getElementById("sidebar")?.classList.remove("open");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

async function renderRoute() {
  try {
    if (state.route === "home") return renderStudentHome();
    if (state.route === "roadmap") return renderRoadmap();
    if (state.route === "lesson") return renderLesson(state.routeParam);
    if (state.route === "quiz") return renderQuiz(state.routeParam);
    if (state.route === "notebook") return renderNotebook();
    if (state.route === "errors") return renderErrorLog();
    if (state.route === "teacher") return renderTeacherDashboard();
    if (state.route === "student-detail") return renderStudentDetail(state.routeParam || 1);
    return state.user.role === "teacher" ? renderTeacherDashboard() : renderStudentHome();
  } catch (error) {
    document.getElementById("page-content").innerHTML = `<div class="error-banner">${escapeHtml(error.message)}</div>`;
  }
}

async function renderStudentHome() {
  const data = await api("/api/student/dashboard");
  state.dashboard = data;
  const s = data.stats;
  document.getElementById("page-content").innerHTML = `
    <div class="page-head">
      <div>
        <p class="eyebrow">Good afternoon, ${escapeHtml(data.user.name.split(" ")[0])}</p>
        <h1>Your learning orbit</h1>
        <p>Continue today's lesson, complete the daily test, and close yesterday's errors.</p>
      </div>
      <button class="btn btn-secondary" data-go="roadmap">View full-year path</button>
    </div>
    <div class="grid grid-4">
      ${statCard("Lessons mastered", `${s.lessonsMastered}/${s.totalAvailable}`, "Demo pathway", "✓")}
      ${statCard("Quiz average", `${s.averageScore}%`, "Across all attempts", "↗")}
      ${statCard("Unresolved errors", s.unresolvedErrors, "Retry within 48 hours", "!")}
      ${statCard("Study streak", `${s.streak} days`, "Consistent practice", "◆")}
    </div>
    <div class="grid dashboard-grid" style="margin-top:18px">
      <section class="hero-card card">
        <div class="tag-row">
          <span class="tag tag-common">Lecture ${data.nextLesson.lecture_no}</span>
          <span class="tag tag-neet">Daily test</span>
          <span class="tag">${escapeHtml(data.nextLesson.chapter)}</span>
        </div>
        <h2>${escapeHtml(data.nextLesson.title)}</h2>
        <p>${escapeHtml(data.nextLesson.summary)}</p>
        <div class="hero-actions">
          <button class="btn btn-primary" data-lesson="${data.nextLesson.id}">Continue lesson</button>
          <button class="btn btn-secondary" data-quiz="${data.nextLesson.id}">Open daily test</button>
        </div>
      </section>
      <section class="card">
        <div class="card-header"><h2>Daily focus</h2><span class="tag tag-neet">Due today</span></div>
        <div class="card-body">
          <div style="display:flex;justify-content:space-between;margin-bottom:9px;font-size:12px">
            <span style="color:var(--muted)">Lesson progress</span><strong>${data.nextLesson.mastery}%</strong>
          </div>
          <div class="progress"><span style="width:${data.nextLesson.mastery}%"></span></div>
          <div class="activity-list" style="margin-top:12px">
            ${focusItem("Read lesson", data.nextLesson.status !== "not_started")}
            ${focusItem("Complete 5-question test", data.nextLesson.status === "mastered")}
            ${focusItem("Review incorrect answers", s.unresolvedErrors === 0)}
          </div>
        </div>
      </section>
    </div>
    <div class="grid grid-2" style="margin-top:18px">
      <section class="card">
        <div class="card-header"><h2>Demo lessons</h2><button class="btn btn-ghost btn-small" data-go="roadmap">See roadmap</button></div>
        <div class="card-body lesson-list">
          ${data.lessons.slice(0, 6).map(lessonRow).join("")}
        </div>
      </section>
      <section class="card">
        <div class="card-header"><h2>Recent attempts</h2><span class="tag">${data.attempts.length} tracked</span></div>
        <div class="card-body activity-list">
          ${data.attempts.length ? data.attempts.map(attemptRow).join("") : `<div class="empty">Your quiz history will appear here.</div>`}
        </div>
      </section>
    </div>
    <div class="section-title"><h2>Upcoming programme schedule</h2><span class="tag">Imported from 138-lecture CSV</span></div>
    <section class="card">
      <div class="data-table-wrap">
        <table class="data-table">
          <thead><tr><th>Lecture</th><th>Date</th><th>Chapter</th><th>Topic</th><th>Learning output</th></tr></thead>
          <tbody>${data.upcoming.map(row => `
            <tr><td>#${row.lecture_no}</td><td>${formatDate(row.lecture_date)}</td><td>${escapeHtml(row.chapter)}</td><td>${escapeHtml(row.topic)}</td><td>${escapeHtml(row.homework || "")}</td></tr>
          `).join("")}</tbody>
        </table>
      </div>
    </section>`;
  bindStudentLinks();
}

function statCard(label, value, foot, icon) {
  return `<section class="card stat-card">
    <div class="stat-top"><span>${label}</span><span class="stat-icon">${icon}</span></div>
    <div class="stat-value">${value}</div><div class="stat-foot">${foot}</div>
  </section>`;
}

function focusItem(label, complete) {
  return `<div class="activity-row"><div class="lesson-number ${complete ? "mastered" : ""}">${complete ? "✓" : "·"}</div><div class="row-copy"><strong>${label}</strong><span>${complete ? "Complete" : "Ready for action"}</span></div></div>`;
}

function lessonRow(lesson) {
  return `<div class="lesson-row">
    <div class="lesson-number ${lesson.status === "mastered" ? "mastered" : ""}">${lesson.status === "mastered" ? "✓" : lesson.lecture_no}</div>
    <div class="row-copy"><strong>${escapeHtml(lesson.title)}</strong><span>${escapeHtml(lesson.chapter)} · ${formatDate(lesson.lesson_date)}</span></div>
    <span class="status ${lesson.status}">${lesson.status.replace("_", " ")}</span>
    <button class="btn btn-ghost btn-small" data-lesson="${lesson.id}">Open</button>
  </div>`;
}

function attemptRow(attempt) {
  const percentage = Math.round(attempt.score * 100 / attempt.total);
  return `<div class="activity-row">
    <div class="lesson-number ${percentage >= 70 ? "mastered" : ""}">${percentage}%</div>
    <div class="row-copy"><strong>${escapeHtml(attempt.lesson)}</strong><span>${formatDate(attempt.submitted_at)} · ${attempt.score}/${attempt.total} correct</span></div>
  </div>`;
}

function bindStudentLinks() {
  document.querySelectorAll("[data-lesson]").forEach(button => button.addEventListener("click", () => navigate("lesson", Number(button.dataset.lesson))));
  document.querySelectorAll("[data-quiz]").forEach(button => button.addEventListener("click", () => navigate("quiz", Number(button.dataset.quiz))));
  document.querySelectorAll("[data-go]").forEach(button => button.addEventListener("click", () => navigate(button.dataset.go)));
}

async function renderRoadmap() {
  const [road, dash] = await Promise.all([api("/api/roadmap"), api("/api/student/dashboard")]);
  document.getElementById("page-content").innerHTML = `
    <div class="page-head">
      <div><p class="eyebrow">24 Aug 2026 — 31 Jan 2027</p><h1>Full-year learning path</h1><p>The first two foundations are live. Later chapters remain visible as a clear locked roadmap.</p></div>
      <span class="tag tag-cbse">138 lectures · 23 weeks</span>
    </div>
    <section class="card" style="padding:20px;margin-bottom:18px">
      <div style="display:flex;justify-content:space-between;gap:15px;margin-bottom:11px">
        <strong>Demo pathway completion</strong><span style="color:var(--muted)">${dash.stats.lessonsMastered}/${dash.stats.totalAvailable} lessons</span>
      </div>
      <div class="progress"><span style="width:${Math.round(dash.stats.lessonsMastered * 100 / dash.stats.totalAvailable)}%"></span></div>
    </section>
    <div class="path-map">
      ${road.chapters.map((chapter, index) => `
        <article class="chapter-tile ${chapter.status === "locked" ? "locked" : ""}">
          <div class="chapter-index">${index === 0 ? "BRIDGE" : `CHAPTER ${index}`}</div>
          <strong>${escapeHtml(chapter.title)}</strong>
          <div class="chapter-foot">
            <span>${chapter.lesson_count ? `${chapter.mastered_count || 0}/${chapter.lesson_count} lessons` : "Planned"}</span>
            <span class="lock">${chapter.status === "locked" ? "Locked" : "Active"}</span>
          </div>
        </article>`).join("")}
    </div>
    <div class="section-title"><h2>Available lessons</h2><span class="tag">Maths Tools + Units</span></div>
    <section class="card"><div class="card-body lesson-list">${dash.lessons.map(lessonRow).join("")}</div></section>`;
  bindStudentLinks();
}

function blockTag(type) {
  const map = { common: ["Common", "tag-common"], neet: ["NEET emphasis", "tag-neet"], cbse: ["CBSE emphasis", "tag-cbse"], recall: ["Recall", "tag-common"], definition: ["Definition", ""], example: ["Worked example", ""] };
  return map[type] || ["Key concept", "tag-common"];
}

async function renderLesson(lessonId) {
  const lesson = await api(`/api/lessons/${lessonId}`);
  state.lesson = lesson;
  document.getElementById("page-content").innerHTML = `
    <button class="btn btn-ghost btn-small" id="back-home">← Back to pathway</button>
    <div class="lesson-layout" style="margin-top:18px">
      <article class="lesson-article">
        <div class="tag-row">
          <span class="tag">Lecture ${lesson.lecture_no}</span>
          <span class="tag tag-common">${escapeHtml(lesson.chapter)}</span>
          <span class="tag">${lesson.minutes} min</span>
        </div>
        <h1 class="lesson-title">${escapeHtml(lesson.title)}</h1>
        <p class="lesson-summary">${escapeHtml(lesson.summary)}</p>
        ${lesson.sections.map((section, index) => {
          const [label, klass] = blockTag(section.type);
          return `<section class="lesson-block type-${section.type}" data-block="${index + 1}">
            <div class="tag-row" style="margin-bottom:10px"><span class="tag ${klass}">${label}</span></div>
            <h3>${escapeHtml(section.title)}</h3>
            <p>${escapeHtml(section.body)}</p>
            ${section.formula ? `<div class="formula">${escapeHtml(section.formula)}</div>` : ""}
          </section>`;
        }).join("")}
        <section class="hero-card card" style="min-height:auto;margin-top:25px">
          <div class="tag-row"><span class="tag tag-neet">Daily test</span><span class="tag">5 questions</span><span class="tag">10–15 min</span></div>
          <h2>Turn reading into evidence.</h2>
          <p>Complete the daily test now. Incorrect questions will enter your error log and appear in the teacher's next-day review queue.</p>
          <div class="hero-actions"><button class="btn btn-primary" id="start-quiz">Start daily test</button></div>
        </section>
      </article>
      <aside class="lesson-aside">
        <section class="card">
          <div class="card-header"><h3>Lesson status</h3><span class="status ${lesson.progress.status}">${lesson.progress.status}</span></div>
          <div class="card-body">
            <div style="display:flex;justify-content:space-between;margin-bottom:9px;font-size:12px"><span style="color:var(--muted)">Mastery</span><strong>${lesson.progress.mastery}%</strong></div>
            <div class="progress"><span style="width:${lesson.progress.mastery}%"></span></div>
          </div>
        </section>
        <section class="card note-compose">
          <div class="card-header"><h3>Quick note</h3><button class="btn btn-ghost btn-small" id="capture-selection">Use selection</button></div>
          <div class="card-body">
            <div class="form-group"><label>Type</label><select class="input select" id="note-kind"><option>Key concept</option><option>Formula</option><option>Definition</option><option>CBSE derivation</option><option>NEET trap</option><option>Exam tip</option></select></div>
            <div class="form-group"><label>Selected text</label><input class="input" id="selected-text" placeholder="Select lesson text or type here" /></div>
            <div class="form-group"><label>Your note</label><textarea class="input" id="note-text" placeholder="Write a memory cue…"></textarea></div>
            <button class="btn btn-secondary btn-block" id="save-note">Save to notebook</button>
          </div>
        </section>
        <section class="card">
          <div class="card-header"><h3>Recent lesson notes</h3><span class="tag">${lesson.notes.length}</span></div>
          <div class="card-body">
            ${lesson.notes.length ? lesson.notes.slice(0, 4).map(note => `<div class="note-item"><small>${escapeHtml(note.kind)}</small><p>${escapeHtml(note.note)}</p></div>`).join("") : `<div class="empty">Capture a formula, definition or exam trap.</div>`}
          </div>
        </section>
      </aside>
    </div>`;
  document.getElementById("back-home").addEventListener("click", () => navigate("home"));
  document.getElementById("start-quiz").addEventListener("click", () => navigate("quiz", lessonId));
  document.getElementById("capture-selection").addEventListener("click", () => {
    const selection = window.getSelection().toString().trim();
    if (!selection) return toast("Select text in the lesson first", "error");
    document.getElementById("selected-text").value = selection.slice(0, 300);
    toast("Selection captured");
  });
  document.getElementById("save-note").addEventListener("click", async () => {
    const note = document.getElementById("note-text").value.trim();
    if (!note) return toast("Write a note first", "error");
    await api(`/api/lessons/${lessonId}/notes`, {
      method: "POST",
      body: JSON.stringify({
        kind: document.getElementById("note-kind").value,
        selectedText: document.getElementById("selected-text").value,
        note,
      }),
    });
    toast("Saved to notebook");
    renderLesson(lessonId);
  });
}

async function renderQuiz(lessonId) {
  const quiz = await api(`/api/lessons/${lessonId}/quiz`);
  state.quizData = quiz;
  state.selectedAnswers = {};
  state.quizStartedAt = Date.now();
  document.getElementById("page-content").innerHTML = `
    <div class="quiz-shell">
      <div class="quiz-top">
        <div><p class="eyebrow">Daily practice · 70% gate</p><h1 style="margin:0;letter-spacing:-.04em">${escapeHtml(quiz.lesson.title)}</h1><p style="color:var(--muted)">Answer all ${quiz.questions.length} questions. Results are tracked topic-wise for next-day review.</p></div>
        <div class="timer" id="quiz-timer">00:00</div>
      </div>
      <form id="quiz-form">
        ${quiz.questions.map((question, qIndex) => `
          <section class="question-card" data-question="${question.id}">
            <div class="question-meta"><span class="tag ${question.examTag === "NEET" ? "tag-neet" : question.examTag === "CBSE" ? "tag-cbse" : question.examTag === "Practical" ? "tag-practical" : "tag-common"}">${escapeHtml(question.examTag)}</span><span class="tag tag-${question.difficulty.toLowerCase()}">${escapeHtml(question.difficulty)}</span><span class="tag">${escapeHtml(question.topic)}</span></div>
            <h3>${qIndex + 1}. ${escapeHtml(question.prompt)}</h3>
            <div class="option-list">
              ${question.options.map((option, index) => `<label class="option"><input type="radio" name="q-${question.id}" value="${index}" /><span>${escapeHtml(option)}</span></label>`).join("")}
            </div>
          </section>`).join("")}
        <div style="display:flex;justify-content:space-between;gap:12px;margin-top:20px">
          <button type="button" class="btn btn-ghost" id="cancel-quiz">Return to lesson</button>
          <button type="submit" class="btn btn-primary">Submit daily test</button>
        </div>
      </form>
    </div>`;
  document.querySelectorAll(".option input").forEach(input => {
    input.addEventListener("change", () => {
      const card = input.closest(".question-card");
      card.querySelectorAll(".option").forEach(option => option.classList.remove("selected"));
      input.closest(".option").classList.add("selected");
      state.selectedAnswers[card.dataset.question] = Number(input.value);
    });
  });
  state.quizTimer = setInterval(() => {
    const seconds = Math.floor((Date.now() - state.quizStartedAt) / 1000);
    document.getElementById("quiz-timer").textContent = `${String(Math.floor(seconds / 60)).padStart(2, "0")}:${String(seconds % 60).padStart(2, "0")}`;
  }, 1000);
  document.getElementById("cancel-quiz").addEventListener("click", () => navigate("lesson", lessonId));
  document.getElementById("quiz-form").addEventListener("submit", async (event) => {
    event.preventDefault();
    if (Object.keys(state.selectedAnswers).length < quiz.questions.length && !confirm("Some questions are unanswered. Submit anyway?")) return;
    clearInterval(state.quizTimer);
    const result = await api(`/api/lessons/${lessonId}/quiz`, {
      method: "POST",
      body: JSON.stringify({ answers: state.selectedAnswers, durationSeconds: Math.floor((Date.now() - state.quizStartedAt) / 1000) }),
    });
    renderQuizResult(quiz, result, lessonId);
  });
}

function renderQuizResult(quiz, result, lessonId) {
  const passed = result.percentage >= 70;
  document.getElementById("page-content").innerHTML = `
    <div class="quiz-shell">
      <section class="card result-hero">
        <div class="result-score">${result.percentage}%</div>
        <p class="eyebrow">${passed ? "Mastery gate cleared" : "Remediation created"}</p>
        <h1 style="margin:0">${passed ? "Strong work." : "Good attempt. Now correct it."}</h1>
        <p style="color:var(--muted);line-height:1.6">${result.score}/${result.total} correct. ${passed ? "The lesson is marked mastered." : "Incorrect questions are now in your error log for a 48-hour retry."}</p>
        <div class="hero-actions" style="justify-content:center"><button class="btn btn-primary" id="result-home">Back to dashboard</button><button class="btn btn-secondary" id="result-errors">Open error log</button></div>
      </section>
      <div class="section-title"><h2>Question review</h2><span class="tag">${result.score}/${result.total} correct</span></div>
      ${quiz.questions.map((question, index) => {
        const evaluation = result.results.find(item => item.questionId === question.id);
        return `<section class="question-card">
          <div class="question-meta"><span class="tag ${evaluation.isCorrect ? "tag-cbse" : "tag-neet"}">${evaluation.isCorrect ? "Correct" : "Review"}</span><span class="tag">${escapeHtml(question.topic)}</span></div>
          <h3>${index + 1}. ${escapeHtml(question.prompt)}</h3>
          <div class="answer-review ${evaluation.isCorrect ? "correct" : ""}">
            <strong>${evaluation.isCorrect ? "Correct reasoning" : `Correct answer: ${escapeHtml(question.options[evaluation.correctIndex])}`}</strong>
            <p>${escapeHtml(evaluation.explanation)}</p>
          </div>
        </section>`;
      }).join("")}
    </div>`;
  document.getElementById("result-home").addEventListener("click", () => navigate("home"));
  document.getElementById("result-errors").addEventListener("click", () => navigate("errors"));
}

async function renderNotebook() {
  const data = await api("/api/student/notebook");
  const groups = data.notes.reduce((acc, note) => {
    acc[note.kind] = acc[note.kind] || [];
    acc[note.kind].push(note);
    return acc;
  }, {});
  document.getElementById("page-content").innerHTML = `
    <div class="page-head"><div><p class="eyebrow">Personal knowledge base</p><h1>Notebook</h1><p>Your formulas, definitions, concepts and exam traps in one place.</p></div><span class="tag tag-common">${data.notes.length} saved notes</span></div>
    ${data.notes.length ? Object.entries(groups).map(([kind, notes]) => `
      <div class="section-title"><h2>${escapeHtml(kind)}</h2><span class="tag">${notes.length}</span></div>
      <div class="grid grid-3">${notes.map(note => `
        <article class="card"><div class="card-body"><div class="tag-row"><span class="tag tag-common">${escapeHtml(note.chapter)}</span></div><h3 style="margin:14px 0 8px">${escapeHtml(note.lesson)}</h3>${note.selected_text ? `<p style="color:var(--gold);font-size:12px;line-height:1.5">“${escapeHtml(note.selected_text)}”</p>` : ""}<p style="color:var(--muted);font-size:13px;line-height:1.65">${escapeHtml(note.note)}</p></div></article>`).join("")}</div>
    `).join("") : `<section class="card empty">No notes yet. Open a lesson, select useful text and save a note.</section>`}`;
}

async function renderErrorLog() {
  const data = await api("/api/student/error-log");
  document.getElementById("page-content").innerHTML = `
    <div class="page-head"><div><p class="eyebrow">48-hour correction loop</p><h1>Error log</h1><p>Every wrong answer becomes a specific revision task—not a forgotten score.</p></div><span class="tag tag-neet">${data.errors.filter(e => !e.retry_resolved).length} unresolved</span></div>
    <div class="grid grid-2">
      ${data.errors.length ? data.errors.map(error => `
        <article class="card">
          <div class="card-header"><h3>${escapeHtml(error.topic)}</h3><span class="status ${error.retry_resolved ? "mastered" : "practiced"}">${error.retry_resolved ? "resolved" : "retry due"}</span></div>
          <div class="card-body">
            <div class="tag-row"><span class="tag">${escapeHtml(error.lesson)}</span><span class="tag">${formatDate(error.submitted_at)}</span></div>
            <h3 style="line-height:1.5;margin:14px 0">${escapeHtml(error.prompt)}</h3>
            <p style="color:var(--red);font-size:12px">Your answer: ${error.selected_index == null ? "Not answered" : escapeHtml(error.options[error.selected_index])}</p>
            <p style="color:var(--green);font-size:12px">Correct answer: ${escapeHtml(error.options[error.correct_index])}</p>
            <p style="color:var(--muted);font-size:12px;line-height:1.6">${escapeHtml(error.explanation)}</p>
          </div>
        </article>`).join("") : `<section class="card empty">No errors yet. Complete a daily test to build your correction log.</section>`}
    </div>`;
}

async function renderTeacherDashboard() {
  const data = await api("/api/teacher/dashboard");
  const student = data.students[0] || {};
  const unresolved = data.reviewQueue.filter(item => !item.reviewed).length;
  document.getElementById("page-content").innerHTML = `
    <div class="page-head">
      <div><p class="eyebrow">Next-day teaching intelligence</p><h1>Teacher dashboard</h1><p>See what was attempted, what went wrong, and what to review before today's new topic.</p></div>
      <button class="btn btn-primary" data-student="${student.id || 1}">Open student analysis</button>
    </div>
    <div class="grid grid-4">
      ${statCard("Active students", data.students.length, "Local demo cohort", "◎")}
      ${statCard("Average score", `${student.average_score || 0}%`, "All tracked attempts", "↗")}
      ${statCard("Lessons mastered", student.mastered || 0, "Programme progress", "✓")}
      ${statCard("Review items", unresolved, "Before new teaching", "!")}
    </div>
    <div class="grid dashboard-grid" style="margin-top:18px">
      <section class="card">
        <div class="card-header"><h2>Next-day review queue</h2><span class="tag tag-neet">Priority sorted</span></div>
        <div class="card-body review-list">
          ${data.reviewQueue.length ? data.reviewQueue.map(item => `
            <div class="review-row">
              <div class="lesson-number ${item.reviewed ? "mastered" : ""}">${item.reviewed ? "✓" : `${item.miss_rate}%`}</div>
              <div class="row-copy"><strong>${escapeHtml(item.prompt)}</strong><span>${escapeHtml(item.lesson)} · ${escapeHtml(item.topic)} · ${item.wrong}/${item.attempts} wrong</span></div>
              ${item.reviewed ? `<span class="status mastered">reviewed</span>` : `<button class="btn btn-secondary btn-small" data-review="${item.question_id}">Mark reviewed</button>`}
            </div>`).join("") : `<div class="empty">No review items yet.</div>`}
        </div>
      </section>
      <section class="card">
        <div class="card-header"><h2>Weakest topics</h2><span class="tag">Question level</span></div>
        <div class="card-body">
          ${data.weakTopics.length ? data.weakTopics.map(topic => `
            <div style="margin-bottom:17px">
              <div style="display:flex;justify-content:space-between;gap:12px;margin-bottom:7px;font-size:12px"><strong>${escapeHtml(topic.topic)}</strong><span class="${percentClass(topic.accuracy)}">${topic.accuracy}%</span></div>
              <div class="progress"><span style="width:${topic.accuracy}%;background:${topic.accuracy < 50 ? "var(--red)" : topic.accuracy < 70 ? "var(--gold)" : "var(--green)"}"></span></div>
            </div>`).join("") : `<div class="empty">Topic analytics appear after test attempts.</div>`}
        </div>
      </section>
    </div>
    <div class="grid grid-2" style="margin-top:18px">
      <section class="card">
        <div class="card-header"><h2>Students</h2><span class="tag">${data.students.length}</span></div>
        <div class="data-table-wrap">
          <table class="data-table"><thead><tr><th>Student</th><th>Attempts</th><th>Average</th><th>Mastered</th><th>Last active</th></tr></thead>
          <tbody>${data.students.map(row => `<tr data-student="${row.id}" style="cursor:pointer"><td><strong>${escapeHtml(row.name)}</strong><br><span style="color:var(--muted-2)">${escapeHtml(row.email)}</span></td><td>${row.attempts}</td><td class="${percentClass(row.average_score)}">${row.average_score}%</td><td>${row.mastered}</td><td>${formatDate(row.last_active)}</td></tr>`).join("")}</tbody></table>
        </div>
      </section>
      <section class="card">
        <div class="card-header"><h2>Recent test activity</h2><span class="tag tag-cbse">Live SQLite data</span></div>
        <div class="card-body activity-list">
          ${data.recentAttempts.map(row => `<div class="activity-row"><div class="lesson-number ${row.score / row.total >= .7 ? "mastered" : ""}">${Math.round(row.score * 100 / row.total)}%</div><div class="row-copy"><strong>${escapeHtml(row.student)} · ${escapeHtml(row.lesson)}</strong><span>${formatDate(row.submitted_at)} · ${row.score}/${row.total} correct</span></div></div>`).join("")}
        </div>
      </section>
    </div>`;
  document.querySelectorAll("[data-review]").forEach(button => button.addEventListener("click", async () => {
    await api(`/api/teacher/review/${button.dataset.review}`, { method: "POST" });
    toast("Question marked reviewed");
    renderTeacherDashboard();
  }));
  document.querySelectorAll("[data-student]").forEach(element => element.addEventListener("click", () => navigate("student-detail", Number(element.dataset.student))));
}

async function renderStudentDetail(studentId) {
  const data = await api(`/api/teacher/student/${studentId}`);
  const average = data.attempts.length ? Math.round(data.attempts.reduce((sum, row) => sum + row.score / row.total * 100, 0) / data.attempts.length) : 0;
  document.getElementById("page-content").innerHTML = `
    <div class="page-head">
      <div><p class="eyebrow">Individual learning record</p><h1>${escapeHtml(data.student.name)}</h1><p>${escapeHtml(data.student.email)} · Topic-wise progress and intervention evidence.</p></div>
      <button class="btn btn-secondary" id="back-teacher">Back to overview</button>
    </div>
    <div class="grid grid-4">
      ${statCard("Attempts", data.attempts.length, "Daily tests submitted", "▤")}
      ${statCard("Average", `${average}%`, "All tracked answers", "↗")}
      ${statCard("Lessons mastered", data.progress.filter(row => row.status === "mastered").length, "70% chapter gate", "✓")}
      ${statCard("Weak topics", data.topics.filter(row => row.accuracy < 70).length, "Need targeted support", "!")}
    </div>
    <div class="grid grid-2" style="margin-top:18px">
      <section class="card">
        <div class="card-header"><h2>Topic mastery</h2><span class="tag">Answer analytics</span></div>
        <div class="card-body">
          ${data.topics.map(topic => `<div style="margin-bottom:17px"><div style="display:flex;justify-content:space-between;margin-bottom:7px;font-size:12px"><strong>${escapeHtml(topic.topic)}</strong><span class="${percentClass(topic.accuracy)}">${topic.accuracy}%</span></div><div class="progress"><span style="width:${topic.accuracy}%;background:${topic.accuracy < 50 ? "var(--red)" : topic.accuracy < 70 ? "var(--gold)" : "var(--green)"}"></span></div></div>`).join("") || `<div class="empty">No topic data yet.</div>`}
        </div>
      </section>
      <section class="card">
        <div class="card-header"><h2>Lesson progression</h2><span class="tag">${data.progress.length} demo lessons</span></div>
        <div class="card-body lesson-list">
          ${data.progress.map((row, index) => `<div class="lesson-row"><div class="lesson-number ${row.status === "mastered" ? "mastered" : ""}">${row.status === "mastered" ? "✓" : index + 1}</div><div class="row-copy"><strong>${escapeHtml(row.lesson)}</strong><span>${escapeHtml(row.chapter)}</span></div><div class="meter"><div class="progress"><span style="width:${row.mastery}%"></span></div><span>${row.mastery}%</span></div></div>`).join("")}
        </div>
      </section>
    </div>
    <div class="section-title"><h2>Complete attempt history</h2><span class="tag">Read-only evidence</span></div>
    <section class="card"><div class="data-table-wrap"><table class="data-table"><thead><tr><th>Date</th><th>Chapter</th><th>Lesson</th><th>Score</th><th>Time</th><th>Action</th></tr></thead><tbody>${data.attempts.map(row => `<tr><td>${formatDate(row.submitted_at)}</td><td>${escapeHtml(row.chapter)}</td><td>${escapeHtml(row.lesson)}</td><td class="${percentClass(Math.round(row.score * 100 / row.total))}">${row.score}/${row.total} (${Math.round(row.score * 100 / row.total)}%)</td><td>${Math.floor(row.duration_seconds / 60)}m ${row.duration_seconds % 60}s</td><td>${row.score / row.total < .7 ? '<span class="status practiced">remediate</span>' : '<span class="status mastered">mastered</span>'}</td></tr>`).join("")}</tbody></table></div></section>`;
  document.getElementById("back-teacher").addEventListener("click", () => navigate("teacher"));
}

async function init() {
  try {
    state.user = await api("/api/me");
    state.route = state.user.role === "teacher" ? "teacher" : "home";
    await renderShell();
  } catch {
    renderLogin();
  }
}

init();
