import { createClient } from "@supabase/supabase-js";
import { randomUUID } from "node:crypto";
import { readFile } from "node:fs/promises";

const configPath = process.argv[2];
if (!configPath) throw new Error("Pass the weekly-test JSON file path.");

const config = JSON.parse(await readFile(configPath, "utf8"));
const { NEXT_PUBLIC_SUPABASE_URL: url, NEXT_PUBLIC_SUPABASE_ANON_KEY: anonKey } = process.env;
const email = process.env.TEACHER_EMAIL;
const password = process.env.TEACHER_PASSWORD;
if (!url || !anonKey || !email || !password) {
  throw new Error("API and teacher credentials are required.");
}

const marksPerQuestion = config.marksPerQuestion ?? 2;
const supabase = createClient(url, anonKey, {
  auth: { persistSession: false, autoRefreshToken: false },
});
const fail = (error, context) => {
  throw new Error(`${context}: ${error.message}`);
};

const { data: auth, error: authError } = await supabase.auth.signInWithPassword({ email, password });
if (authError) fail(authError, "Teacher authentication failed");
const teacherId = auth.user?.id;

const { data: existing, error: existingError } = await supabase
  .from("assessments")
  .select("id")
  .eq("title", config.title)
  .eq("assessment_type", "weekly_test")
  .maybeSingle();
if (existingError) fail(existingError, "Could not find existing weekly test");
if (!existing) throw new Error("Weekly test not found to replace.");

const { count: attemptCount, error: attemptError } = await supabase
  .from("attempts")
  .select("*", { count: "exact", head: true })
  .eq("assessment_id", existing.id);
if (attemptError) fail(attemptError, "Could not check attempts");
if ((attemptCount ?? 0) > 0) {
  throw new Error(`Cannot replace: ${attemptCount} student attempt(s) already exist.`);
}

const { data: oldLinks, error: oldLinksError } = await supabase
  .from("assessment_questions")
  .select("question_id")
  .eq("assessment_id", existing.id);
if (oldLinksError) fail(oldLinksError, "Could not load old questions");
const oldQuestionIds = (oldLinks ?? []).map((row) => row.question_id);

const { data: chapter, error: chapterError } = await supabase
  .from("chapters")
  .select("id")
  .eq("title", "Basic Maths and Calculus")
  .limit(1)
  .single();
if (chapterError) fail(chapterError, "Could not find chapter");

let classId = null;
for (const lectureNumber of [config.lectureNumber, 9, 8]) {
  const { data: classRow } = await supabase
    .from("classes")
    .select("id")
    .eq("chapter_id", chapter.id)
    .eq("lecture_number", lectureNumber)
    .limit(1)
    .maybeSingle();
  if (classRow) {
    classId = classRow.id;
    break;
  }
}

const questionRows = config.questions.map((question) => ({
  id: randomUUID(),
  chapter_id: chapter.id,
  class_id: classId,
  prompt: question.prompt,
  question_type: "mcq",
  options: question.options,
  marks: marksPerQuestion,
  difficulty: "challenging",
  exam_tags: ["CBSE", "NEET"],
  topic: question.topic,
  created_by: teacherId,
}));

const { error: unlinkError } = await supabase
  .from("assessment_questions")
  .delete()
  .eq("assessment_id", existing.id);
if (unlinkError) fail(unlinkError, "Could not unlink old questions");

if (oldQuestionIds.length) {
  await supabase.from("question_keys").delete().in("question_id", oldQuestionIds);
  await supabase.from("questions").delete().in("id", oldQuestionIds);
}

const { error: questionsError } = await supabase.from("questions").insert(questionRows);
if (questionsError) fail(questionsError, "Could not create questions");

const { error: keysError } = await supabase.from("question_keys").insert(
  config.questions.map((question, index) => ({
    question_id: questionRows[index].id,
    correct_answer: question.correctIndex,
    explanation: question.explanation,
  })),
);
if (keysError) fail(keysError, "Could not create answer keys");

const { error: mappingError } = await supabase.from("assessment_questions").insert(
  questionRows.map((question, index) => ({
    assessment_id: existing.id,
    question_id: question.id,
    sequence: index + 1,
    marks: marksPerQuestion,
  })),
);
if (mappingError) fail(mappingError, "Could not attach questions");

const totalMarks = marksPerQuestion * config.questions.length;
const { data: updated, error: updateError } = await supabase
  .from("assessments")
  .update({
    instructions: config.instructions,
    duration_minutes: config.durationMinutes,
    total_marks: totalMarks,
    opens_at: new Date().toISOString(),
    due_at: config.dueAt,
    status: "published",
    class_id: classId,
  })
  .eq("id", existing.id)
  .select("id,title,status,total_marks,opens_at,due_at,duration_minutes")
  .single();
if (updateError) fail(updateError, "Could not update assessment");

console.log(JSON.stringify({
  action: "replaced_questions",
  assessment: updated,
  questionCount: config.questions.length,
}, null, 2));
