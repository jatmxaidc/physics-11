import { createClient } from "@supabase/supabase-js";
import { randomUUID } from "node:crypto";

const url = process.env.NEXT_PUBLIC_SUPABASE_URL;
const anonKey = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY;
const email = process.env.TEACHER_EMAIL;
const password = process.env.TEACHER_PASSWORD;

if (!url || !anonKey || !email || !password) {
  throw new Error("Supabase URL, anonymous key, teacher email and teacher password are required.");
}

const supabase = createClient(url, anonKey, {
  auth: { persistSession: false, autoRefreshToken: false },
});

const questions = [
  {
    prompt: "If y = 2x + 4, find the change in y and the average rate of change when x changes from 1 to 3.",
    topic: "Rate of change",
    options: [
      "Change in y = 2; average rate = 1",
      "Change in y = 4; average rate = 2",
      "Change in y = 6; average rate = 3",
      "Change in y = 8; average rate = 4",
    ],
    correctIndex: 1,
    explanation: "y(1) = 6 and y(3) = 10. Therefore Δy = 4 and Δx = 2, so the average rate is Δy/Δx = 2.",
  },
  {
    prompt: "Find both roots of x² − 7x + 12 = 0 by factorisation.",
    topic: "Quadratic equations",
    options: ["x = 3 or x = 4", "x = −3 or x = −4", "x = 2 or x = 6", "x = −2 or x = −6"],
    correctIndex: 0,
    explanation: "x² − 7x + 12 = (x − 3)(x − 4). Hence x = 3 or x = 4.",
  },
  {
    prompt: "Find both roots of x² + 7x − 18 = 0 by factorisation.",
    topic: "Quadratic equations",
    options: ["x = 9 or x = −2", "x = 2 or x = −9", "x = 3 or x = −6", "x = −3 or x = 6"],
    correctIndex: 1,
    explanation: "x² + 7x − 18 = (x + 9)(x − 2). Hence x = 2 or x = −9.",
  },
  {
    prompt: "Find both roots of x² + 8x + 15 = 0 by factorisation.",
    topic: "Quadratic equations",
    options: ["x = 3 or x = 5", "x = −3 or x = −5", "x = −1 or x = −15", "x = 1 or x = 15"],
    correctIndex: 1,
    explanation: "x² + 8x + 15 = (x + 3)(x + 5). Hence x = −3 or x = −5.",
  },
  {
    prompt: "A ball is thrown vertically downward at 10 m s⁻¹ from a height of 75 m. Taking g = 10 m s⁻², find its time of flight.",
    topic: "Quadratic equations in physics",
    options: ["2 s", "3 s", "5 s", "7.5 s"],
    correctIndex: 1,
    explanation: "Use 75 = 10t + 5t². This gives (t + 5)(t − 3) = 0, so the physical time is 3 s.",
  },
  {
    prompt: "For y = 4x³ − 5x² + 2x − 7, find dy/dx.",
    topic: "Differentiation",
    options: ["12x² − 10x + 2", "12x² − 5x + 2", "4x² − 10x + 2", "12x³ − 10x² + 2"],
    correctIndex: 0,
    explanation: "Differentiate term by term. The derivative is 12x² − 10x + 2.",
  },
  {
    prompt: "Differentiate y = tan x · logₑx.",
    topic: "Product rule",
    options: ["sec²x/x", "tan x/x + logₑx · sec²x", "tan x · sec²x/x", "logₑx/x + tan x · sec²x"],
    correctIndex: 1,
    explanation: "Using (AB)′ = AB′ + BA′ gives tan x/x + logₑx · sec²x.",
  },
  {
    prompt: "If y = A/x³ − B/x⁴, find x for which dy/dx = 0. Assume A and B are non-zero constants.",
    topic: "Stationary points",
    options: ["x = 3A/(4B)", "x = 4A/(3B)", "x = 4B/(3A)", "x = 3B/(4A)"],
    correctIndex: 2,
    explanation: "dy/dx = −3Ax⁻⁴ + 4Bx⁻⁵. Setting this to zero gives −3Ax + 4B = 0, so x = 4B/(3A).",
  },
  {
    prompt: "Differentiate y = eˣ/x², where x ≠ 0.",
    topic: "Quotient rule",
    options: ["eˣ(x − 2)/x³", "eˣ(x + 2)/x³", "eˣ(x − 2)/x²", "eˣ/x³"],
    correctIndex: 0,
    explanation: "The quotient rule gives [x²eˣ − eˣ(2x)]/x⁴ = eˣ(x − 2)/x³.",
  },
  {
    prompt: "Differentiate y = (x² − 1)/(x² + 1).",
    topic: "Quotient rule",
    options: ["2x/(x² + 1)", "4x/(x² + 1)²", "4x²/(x² + 1)²", "−4x/(x² + 1)²"],
    correctIndex: 1,
    explanation: "The quotient-rule numerator simplifies to 4x, giving 4x/(x² + 1)².",
  },
  {
    prompt: "Differentiate y = cos x + 2/x³, where x ≠ 0.",
    topic: "Mixed differentiation",
    options: ["sin x + 6/x⁴", "−sin x + 6/x⁴", "−sin x − 6/x⁴", "cos x − 6/x³"],
    correctIndex: 2,
    explanation: "Write 2/x³ as 2x⁻³. Its derivative is −6x⁻⁴, so dy/dx = −sin x − 6/x⁴.",
  },
  {
    prompt: "Differentiate y = (√x + 1/√x)² for x > 0.",
    topic: "Simplification before differentiation",
    options: ["1 + 1/x²", "1 − 1/x²", "2 − 2/x²", "1 − 2/x"],
    correctIndex: 1,
    explanation: "Expand first: y = x + 2 + 1/x. Therefore dy/dx = 1 − 1/x².",
  },
];

function fail(error, context) {
  throw new Error(`${context}: ${error.message}`);
}

const { data: authData, error: authError } = await supabase.auth.signInWithPassword({ email, password });
if (authError) fail(authError, "Teacher authentication failed");
const teacherId = authData.user?.id;
if (!teacherId) throw new Error("Teacher authentication returned no user.");

const { data: profile, error: profileError } = await supabase
  .from("profiles")
  .select("role")
  .eq("id", teacherId)
  .single();
if (profileError) fail(profileError, "Could not verify teacher profile");
if (profile.role !== "teacher") throw new Error("The supplied account is not a teacher.");

const { data: existing, error: existingError } = await supabase
  .from("assessments")
  .select("id,status,total_marks,due_at,assessment_questions(count)")
  .ilike("title", "Lecture 10 daily quiz")
  .eq("assessment_type", "daily_quiz")
  .maybeSingle();
if (existingError) fail(existingError, "Could not check existing quiz");
if (existing) {
  console.log(JSON.stringify({ action: "already_exists", assessment: existing }));
  process.exit(0);
}

const { data: chapter, error: chapterError } = await supabase
  .from("chapters")
  .select("id")
  .eq("title", "Basic Maths and Calculus")
  .limit(1)
  .single();
if (chapterError) fail(chapterError, "Could not find chapter");

const { data: classRow, error: classError } = await supabase
  .from("classes")
  .select("id")
  .eq("chapter_id", chapter.id)
  .eq("lecture_number", 10)
  .limit(1)
  .maybeSingle();
if (classError) fail(classError, "Could not find Lecture 10 class");

const assessmentId = randomUUID();
const questionRows = questions.map((question) => ({
  id: randomUUID(),
  chapter_id: chapter.id,
  class_id: classRow?.id ?? null,
  prompt: question.prompt,
  question_type: "mcq",
  options: question.options,
  marks: 1,
  difficulty: "standard",
  exam_tags: ["CBSE", "NEET"],
  topic: question.topic,
  created_by: teacherId,
}));

const cleanup = async () => {
  await supabase.from("assessments").delete().eq("id", assessmentId);
  await supabase.from("questions").delete().in("id", questionRows.map((row) => row.id));
};

const { error: assessmentError } = await supabase.from("assessments").insert({
  id: assessmentId,
  class_id: classRow?.id ?? null,
  chapter_id: chapter.id,
  title: "Lecture 10 daily quiz",
  assessment_type: "daily_quiz",
  instructions: "Answer all 12 questions from Lecture 10. Each question carries one mark. You may resubmit before the deadline.",
  duration_minutes: 25,
  total_marks: 12,
  negative_marking: 0,
  attempt_limit: 1,
  randomize_questions: false,
  randomize_options: false,
  explanations_after_submission: true,
  opens_at: new Date().toISOString(),
  due_at: "2026-09-17T18:29:00.000Z",
  status: "draft",
  created_by: teacherId,
});
if (assessmentError) fail(assessmentError, "Could not create assessment");

const { error: questionsError } = await supabase.from("questions").insert(questionRows);
if (questionsError) {
  await cleanup();
  fail(questionsError, "Could not create questions");
}

const { error: keysError } = await supabase.from("question_keys").insert(
  questions.map((question, index) => ({
    question_id: questionRows[index].id,
    correct_answer: question.correctIndex,
    explanation: question.explanation,
  })),
);
if (keysError) {
  await cleanup();
  fail(keysError, "Could not create answer keys");
}

const { error: mappingError } = await supabase.from("assessment_questions").insert(
  questionRows.map((question, index) => ({
    assessment_id: assessmentId,
    question_id: question.id,
    sequence: index + 1,
    marks: 1,
  })),
);
if (mappingError) {
  await cleanup();
  fail(mappingError, "Could not attach questions");
}

const { data: published, error: publishError } = await supabase
  .from("assessments")
  .update({ status: "published" })
  .eq("id", assessmentId)
  .select("id,title,status,total_marks,opens_at,due_at")
  .single();
if (publishError) {
  await cleanup();
  fail(publishError, "Could not publish assessment");
}

const { count, error: countError } = await supabase
  .from("assessment_questions")
  .select("*", { count: "exact", head: true })
  .eq("assessment_id", assessmentId);
if (countError) fail(countError, "Could not verify question count");

console.log(JSON.stringify({ action: "created", assessment: published, questionCount: count }));
