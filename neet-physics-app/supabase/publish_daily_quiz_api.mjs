import { createClient } from "@supabase/supabase-js";
import { randomUUID } from "node:crypto";
import { readFile } from "node:fs/promises";

const configPath = process.argv[2];
if (!configPath) throw new Error("Pass the quiz JSON file path.");

const config = JSON.parse(await readFile(configPath, "utf8"));
const { NEXT_PUBLIC_SUPABASE_URL: url, NEXT_PUBLIC_SUPABASE_ANON_KEY: anonKey } = process.env;
const email = process.env.TEACHER_EMAIL;
const password = process.env.TEACHER_PASSWORD;
if (!url || !anonKey || !email || !password) throw new Error("API and teacher credentials are required.");

const supabase = createClient(url, anonKey, {
  auth: { persistSession: false, autoRefreshToken: false },
});
const fail = (error, context) => {
  throw new Error(`${context}: ${error.message}`);
};

const { data: auth, error: authError } = await supabase.auth.signInWithPassword({ email, password });
if (authError) fail(authError, "Teacher authentication failed");
const teacherId = auth.user?.id;
if (!teacherId) throw new Error("Teacher authentication returned no user.");

const { data: profile, error: profileError } = await supabase
  .from("profiles").select("role").eq("id", teacherId).single();
if (profileError) fail(profileError, "Could not verify profile");
if (profile.role !== "teacher") throw new Error("The supplied account is not a teacher.");

const { data: existing, error: existingError } = await supabase
  .from("assessments")
  .select("id,title,status,total_marks,due_at,assessment_questions(count)")
  .ilike("title", config.title)
  .eq("assessment_type", "daily_quiz")
  .maybeSingle();
if (existingError) fail(existingError, "Could not check existing quiz");
if (existing) {
  console.log(JSON.stringify({ action: "already_exists", assessment: existing }));
  process.exit(0);
}

const { data: chapter, error: chapterError } = await supabase
  .from("chapters").select("id").eq("title", "Basic Maths and Calculus").limit(1).single();
if (chapterError) fail(chapterError, "Could not find chapter");

const { data: classRow, error: classError } = await supabase
  .from("classes")
  .select("id")
  .eq("chapter_id", chapter.id)
  .eq("lecture_number", config.lectureNumber)
  .limit(1)
  .maybeSingle();
if (classError) fail(classError, "Could not find class");

const assessmentId = randomUUID();
const questionRows = config.questions.map((question) => ({
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
  title: config.title,
  assessment_type: "daily_quiz",
  instructions: config.instructions,
  duration_minutes: config.durationMinutes,
  total_marks: config.questions.length,
  negative_marking: 0,
  attempt_limit: 1,
  randomize_questions: false,
  randomize_options: false,
  explanations_after_submission: true,
  opens_at: new Date().toISOString(),
  due_at: config.dueAt,
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
  config.questions.map((question, index) => ({
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
