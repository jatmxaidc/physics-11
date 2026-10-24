"use server";

import { revalidatePath } from "next/cache";
import { redirect } from "next/navigation";
import { z } from "zod";
import { requireProfile } from "@/lib/auth";
import { createClient } from "@/lib/supabase/server";

const classSchema = z.object({
  chapterId: z.string().uuid(),
  lectureNumber: z.coerce.number().int().positive(),
  title: z.string().trim().min(3).max(160),
  description: z.string().trim().max(1000).optional(),
  classDate: z.string().date(),
  durationMinutes: z.coerce.number().int().min(15).max(360),
});

export async function createClass(formData: FormData) {
  const teacher = await requireProfile(["teacher"]);
  const parsed = classSchema.safeParse(Object.fromEntries(formData));
  if (!parsed.success) {
    redirect("/teacher/classes?error=Check+the+class+details");
  }

  const supabase = await createClient();
  const { error } = await supabase.from("classes").insert({
    chapter_id: parsed.data.chapterId,
    lecture_number: parsed.data.lectureNumber,
    title: parsed.data.title,
    description: parsed.data.description || null,
    class_date: parsed.data.classDate,
    duration_minutes: parsed.data.durationMinutes,
    created_by: teacher.id,
    status: "draft",
  });

  if (error) redirect(`/teacher/classes?error=${encodeURIComponent(error.message)}`);
  revalidatePath("/teacher/classes");
  redirect("/teacher/classes?success=Class+draft+created");
}

const publishSchema = z.object({ classId: z.string().uuid() });

export async function publishClass(formData: FormData) {
  await requireProfile(["teacher"]);
  const parsed = publishSchema.safeParse(Object.fromEntries(formData));
  if (!parsed.success) {
    redirect("/teacher/classes?error=Choose+a+valid+class+to+publish");
  }

  const supabase = await createClient();
  const publishedAt = new Date().toISOString();
  const { error } = await supabase
    .from("classes")
    .update({ status: "published", published_at: publishedAt })
    .eq("id", parsed.data.classId);
  if (error) redirect(`/teacher/classes?error=${encodeURIComponent(error.message)}`);

  await supabase
    .from("materials")
    .update({ status: "published", published_at: publishedAt })
    .eq("class_id", parsed.data.classId)
    .eq("status", "draft");

  revalidatePath("/teacher/classes");
  revalidatePath("/student/classes");
  redirect("/teacher/classes?success=Class+published");
}

export async function unpublishClass(formData: FormData) {
  await requireProfile(["teacher"]);
  const parsed = publishSchema.safeParse(Object.fromEntries(formData));
  if (!parsed.success) {
    redirect("/teacher/classes?error=Choose+a+valid+class+to+unpublish");
  }

  const supabase = await createClient();
  const { error } = await supabase
    .from("classes")
    .update({ status: "draft", published_at: null })
    .eq("id", parsed.data.classId);
  if (error) redirect(`/teacher/classes?error=${encodeURIComponent(error.message)}`);

  revalidatePath("/teacher/classes");
  revalidatePath("/student/classes");
  redirect("/teacher/classes?success=Class+unpublished");
}

const materialKindSchema = z.enum([
  "slides",
  "class_notes",
  "short_notes",
  "worksheet",
  "formula_sheet",
  "homework",
]);

const materialSchema = z.object({
  materialId: z.string().uuid(),
  lectureNumber: z.coerce.number().int().positive(),
});

const updateMaterialSchema = materialSchema.extend({
  title: z.string().trim().min(3).max(160),
  kind: materialKindSchema,
});

function materialRedirect(
  lectureNumber: number,
  type: "error" | "success",
  message: string,
) {
  redirect(
    `/teacher/classes?lecture=${lectureNumber}&${type}=${encodeURIComponent(message)}`,
  );
}

export async function updateMaterial(formData: FormData) {
  await requireProfile(["teacher"]);
  const parsed = updateMaterialSchema.safeParse(Object.fromEntries(formData));
  if (!parsed.success) {
    redirect("/teacher/classes?error=Check+the+material+details");
  }

  const supabase = await createClient();
  const { error } = await supabase
    .from("materials")
    .update({ title: parsed.data.title, kind: parsed.data.kind })
    .eq("id", parsed.data.materialId);
  if (error) {
    materialRedirect(parsed.data.lectureNumber, "error", error.message);
  }

  revalidatePath("/teacher/classes");
  revalidatePath("/student/classes");
  materialRedirect(parsed.data.lectureNumber, "success", "Material details updated");
}

export async function toggleMaterialPublication(formData: FormData) {
  await requireProfile(["teacher"]);
  const parsed = materialSchema.extend({
    nextStatus: z.enum(["draft", "published"]),
  }).safeParse(Object.fromEntries(formData));
  if (!parsed.success) {
    redirect("/teacher/classes?error=Choose+a+valid+material");
  }

  const supabase = await createClient();
  const published = parsed.data.nextStatus === "published";
  const { error } = await supabase
    .from("materials")
    .update({
      status: parsed.data.nextStatus,
      published_at: published ? new Date().toISOString() : null,
    })
    .eq("id", parsed.data.materialId);
  if (error) {
    materialRedirect(parsed.data.lectureNumber, "error", error.message);
  }

  revalidatePath("/teacher/classes");
  revalidatePath("/student/classes");
  materialRedirect(
    parsed.data.lectureNumber,
    "success",
    published ? "Material published" : "Material unpublished",
  );
}

export async function deleteMaterial(formData: FormData) {
  await requireProfile(["teacher"]);
  const parsed = materialSchema.safeParse(Object.fromEntries(formData));
  if (!parsed.success) {
    redirect("/teacher/classes?error=Choose+a+valid+material");
  }

  const supabase = await createClient();
  // Soft-delete so student progress, submissions and private notes are preserved.
  const { error } = await supabase
    .from("materials")
    .update({ status: "archived", published_at: null })
    .eq("id", parsed.data.materialId);
  if (error) {
    materialRedirect(parsed.data.lectureNumber, "error", error.message);
  }

  revalidatePath("/teacher/classes");
  revalidatePath("/student/classes");
  materialRedirect(parsed.data.lectureNumber, "success", "Material removed");
}

const shortNoteSchema = z.object({
  classId: z.string().uuid(),
  title: z.string().trim().min(3).max(160),
  content: z.string().trim().min(5).max(10000),
});

export async function publishShortNote(formData: FormData) {
  const teacher = await requireProfile(["teacher"]);
  const parsed = shortNoteSchema.safeParse(Object.fromEntries(formData));
  if (!parsed.success) redirect("/teacher/classes?error=Short+note+is+incomplete");

  const supabase = await createClient();
  const { data: classRow, error: classError } = await supabase
    .from("classes")
    .select("status")
    .eq("id", parsed.data.classId)
    .single();
  if (classError || !classRow) {
    redirect("/teacher/classes?error=Class+not+found");
  }

  const materialStatus = classRow.status === "published" ? "published" : "draft";
  const { error } = await supabase.from("materials").insert({
    class_id: parsed.data.classId,
    title: parsed.data.title,
    kind: "short_notes",
    rich_text: parsed.data.content,
    created_by: teacher.id,
    status: materialStatus,
    published_at: materialStatus === "published" ? new Date().toISOString() : null,
  });
  if (error) redirect(`/teacher/classes?error=${encodeURIComponent(error.message)}`);
  revalidatePath("/teacher/classes");
  revalidatePath("/student/classes");
  redirect(
    materialStatus === "published"
      ? "/teacher/classes?success=Short+notes+published"
      : "/teacher/classes?success=Short+notes+saved+as+draft.+Publish+the+class+to+share+with+students",
  );
}

const optionalLocalDate = z.string().optional().or(z.literal(""));

const assessmentSchema = z.object({
  chapterId: z.string().uuid(),
  classId: z.string().uuid().optional().or(z.literal("")),
  title: z.string().trim().min(3).max(160),
  assessmentType: z.enum(["daily_quiz", "weekly_test", "chapter_test"]),
  durationMinutes: z.coerce.number().int().min(5).max(240),
  totalMarks: z.coerce.number().positive().max(1000),
  instructions: z.string().trim().max(4000).optional().or(z.literal("")),
  opensAt: optionalLocalDate,
  dueAt: optionalLocalDate,
});

function parseWindow(opensAt?: string, dueAt?: string) {
  const opens = opensAt ? new Date(opensAt) : null;
  const due = dueAt ? new Date(dueAt) : null;
  if ((opens && Number.isNaN(opens.getTime())) || (due && Number.isNaN(due.getTime()))) {
    return { error: "Check+the+date+and+time" as const };
  }
  if (Boolean(opens) !== Boolean(due)) {
    return { error: "Set+both+opening+and+due+times,+or+leave+both+empty" as const };
  }
  if (opens && due && due <= opens) {
    return { error: "Due+time+must+be+after+opening+time" as const };
  }
  return {
    opens_at: opens ? opens.toISOString() : null,
    due_at: due ? due.toISOString() : null,
  };
}

export async function createAssessment(formData: FormData) {
  const teacher = await requireProfile(["teacher"]);
  const parsed = assessmentSchema.safeParse(Object.fromEntries(formData));
  if (!parsed.success) redirect("/teacher/assessments?error=Check+assessment+settings");
  const window = parseWindow(parsed.data.opensAt, parsed.data.dueAt);
  if ("error" in window) redirect(`/teacher/assessments?error=${window.error}`);

  const supabase = await createClient();
  const { error } = await supabase.from("assessments").insert({
    chapter_id: parsed.data.chapterId,
    class_id: parsed.data.classId || null,
    title: parsed.data.title,
    assessment_type: parsed.data.assessmentType,
    instructions: parsed.data.instructions || null,
    duration_minutes: parsed.data.durationMinutes,
    total_marks: parsed.data.totalMarks,
    opens_at: window.opens_at,
    due_at: window.due_at,
    created_by: teacher.id,
    status: "draft",
  });
  if (error) redirect(`/teacher/assessments?error=${encodeURIComponent(error.message)}`);
  revalidatePath("/teacher/assessments");
  redirect("/teacher/assessments?success=Assessment+draft+created");
}

const updateAssessmentSchema = assessmentSchema.extend({
  assessmentId: z.string().uuid(),
});

export async function updateAssessment(formData: FormData) {
  await requireProfile(["teacher"]);
  const parsed = updateAssessmentSchema.safeParse(Object.fromEntries(formData));
  if (!parsed.success) redirect("/teacher/assessments?error=Check+assessment+details");

  const supabase = await createClient();
  const { data: questionRows } = await supabase
    .from("assessment_questions")
    .select("marks")
    .eq("assessment_id", parsed.data.assessmentId);
  const totalMarks = (questionRows ?? []).length
    ? (questionRows ?? []).reduce((sum, row) => sum + Number(row.marks), 0)
    : parsed.data.totalMarks;

  const { error } = await supabase
    .from("assessments")
    .update({
      chapter_id: parsed.data.chapterId,
      class_id: parsed.data.classId || null,
      title: parsed.data.title,
      assessment_type: parsed.data.assessmentType,
      instructions: parsed.data.instructions || null,
      duration_minutes: parsed.data.durationMinutes,
      total_marks: totalMarks,
    })
    .eq("id", parsed.data.assessmentId);
  if (error) redirect(`/teacher/assessments?error=${encodeURIComponent(error.message)}`);
  revalidatePath("/teacher/assessments");
  redirect("/teacher/assessments?success=Assessment+details+updated");
}

const scheduleSchema = z.object({
  assessmentId: z.string().uuid(),
  opensAt: z.string().min(1),
  dueAt: z.string().min(1),
});

export async function scheduleAssessment(formData: FormData) {
  await requireProfile(["teacher"]);
  const parsed = scheduleSchema.safeParse(Object.fromEntries(formData));
  if (!parsed.success) redirect("/teacher/assessments?error=Set+both+opening+and+due+times");
  const window = parseWindow(parsed.data.opensAt, parsed.data.dueAt);
  if ("error" in window) redirect(`/teacher/assessments?error=${window.error}`);

  const supabase = await createClient();
  const { data: current, error: currentError } = await supabase
    .from("assessments")
    .select("status")
    .eq("id", parsed.data.assessmentId)
    .single();
  if (currentError) redirect(`/teacher/assessments?error=${encodeURIComponent(currentError.message)}`);

  const nextStatus = current.status === "published" ? "published" : "scheduled";
  const { error } = await supabase
    .from("assessments")
    .update({
      opens_at: window.opens_at,
      due_at: window.due_at,
      status: nextStatus,
    })
    .eq("id", parsed.data.assessmentId);
  if (error) redirect(`/teacher/assessments?error=${encodeURIComponent(error.message)}`);
  revalidatePath("/teacher/assessments");
  revalidatePath("/student/assessments");
  revalidatePath("/student/tests");
  redirect(
    current.status === "published"
      ? "/teacher/assessments?success=Assessment+rescheduled"
      : "/teacher/assessments?success=Times+saved.+Publish+when+students+should+see+it",
  );
}

export async function unpublishAssessment(formData: FormData) {
  await requireProfile(["teacher"]);
  const assessmentId = z.string().uuid().safeParse(formData.get("assessmentId"));
  if (!assessmentId.success) {
    redirect("/teacher/assessments?error=Choose+a+valid+assessment");
  }
  const supabase = await createClient();
  const { error } = await supabase
    .from("assessments")
    .update({ status: "draft" })
    .eq("id", assessmentId.data);
  if (error) redirect(`/teacher/assessments?error=${encodeURIComponent(error.message)}`);
  revalidatePath("/teacher/assessments");
  revalidatePath("/student/assessments");
  revalidatePath("/student/tests");
  redirect("/teacher/assessments?success=Assessment+withdrawn+from+students");
}

const questionSchema = z.object({
  assessmentId: z.string().uuid(),
  chapterId: z.string().uuid(),
  prompt: z.string().trim().min(5).max(2000),
  topic: z.string().trim().min(2).max(100),
  options: z.string().trim().min(7),
  correctIndex: z.coerce.number().int().min(0),
  explanation: z.string().trim().min(5).max(3000),
  marks: z.coerce.number().positive().max(100),
  examTag: z.enum(["CBSE", "NEET"]).optional(),
});

export async function addMcqQuestion(formData: FormData) {
  const teacher = await requireProfile(["teacher"]);
  const parsed = questionSchema.safeParse(Object.fromEntries(formData));
  if (!parsed.success) redirect("/teacher/assessments?error=Check+the+question+and+answer");
  const options = parsed.data.options.split(/\r?\n/).map((option) => option.trim()).filter(Boolean);
  if (options.length < 2 || parsed.data.correctIndex >= options.length) {
    redirect("/teacher/assessments?error=Provide+at+least+two+options+and+a+valid+answer+number");
  }

  const supabase = await createClient();
  const { data: question, error: questionError } = await supabase
    .from("questions")
    .insert({
      chapter_id: parsed.data.chapterId,
      prompt: parsed.data.prompt,
      question_type: "mcq",
      options,
      marks: parsed.data.marks,
      difficulty: "standard",
      exam_tags: [parsed.data.examTag ?? "NEET"],
      topic: parsed.data.topic,
      created_by: teacher.id,
    })
    .select("id")
    .single();
  if (questionError) redirect(`/teacher/assessments?error=${encodeURIComponent(questionError.message)}`);

  const { error: keyError } = await supabase.from("question_keys").insert({
    question_id: question.id,
    correct_answer: parsed.data.correctIndex,
    explanation: parsed.data.explanation,
  });
  if (keyError) {
    await supabase.from("questions").delete().eq("id", question.id);
    redirect(`/teacher/assessments?error=${encodeURIComponent(keyError.message)}`);
  }

  const { count } = await supabase
    .from("assessment_questions")
    .select("*", { count: "exact", head: true })
    .eq("assessment_id", parsed.data.assessmentId);
  const { error: mappingError } = await supabase.from("assessment_questions").insert({
    assessment_id: parsed.data.assessmentId,
    question_id: question.id,
    sequence: (count ?? 0) + 1,
    marks: parsed.data.marks,
  });
  if (mappingError) {
    await supabase.from("questions").delete().eq("id", question.id);
    redirect(`/teacher/assessments?error=${encodeURIComponent(mappingError.message)}`);
  }
  revalidatePath("/teacher/assessments");
  redirect("/teacher/assessments?success=Question+added");
}

const updateQuestionSchema = questionSchema.extend({
  questionId: z.string().uuid(),
});

export async function updateMcqQuestion(formData: FormData) {
  await requireProfile(["teacher"]);
  const parsed = updateQuestionSchema.safeParse(Object.fromEntries(formData));
  if (!parsed.success) redirect("/teacher/assessments?error=Check+the+question+and+answer");
  const options = parsed.data.options.split(/\r?\n/).map((option) => option.trim()).filter(Boolean);
  if (options.length < 2 || parsed.data.correctIndex >= options.length) {
    redirect("/teacher/assessments?error=Provide+at+least+two+options+and+a+valid+answer+number");
  }

  const supabase = await createClient();
  const { error: questionError } = await supabase
    .from("questions")
    .update({
      prompt: parsed.data.prompt,
      options,
      marks: parsed.data.marks,
      topic: parsed.data.topic,
      exam_tags: [parsed.data.examTag ?? "NEET"],
    })
    .eq("id", parsed.data.questionId);
  if (questionError) redirect(`/teacher/assessments?error=${encodeURIComponent(questionError.message)}`);

  const { error: keyError } = await supabase
    .from("question_keys")
    .update({
      correct_answer: parsed.data.correctIndex,
      explanation: parsed.data.explanation,
    })
    .eq("question_id", parsed.data.questionId);
  if (keyError) redirect(`/teacher/assessments?error=${encodeURIComponent(keyError.message)}`);

  const { error: mappingError } = await supabase
    .from("assessment_questions")
    .update({ marks: parsed.data.marks })
    .eq("assessment_id", parsed.data.assessmentId)
    .eq("question_id", parsed.data.questionId);
  if (mappingError) redirect(`/teacher/assessments?error=${encodeURIComponent(mappingError.message)}`);

  revalidatePath("/teacher/assessments");
  redirect("/teacher/assessments?success=Question+updated");
}

export async function deleteAssessmentQuestion(formData: FormData) {
  await requireProfile(["teacher"]);
  const assessmentId = z.string().uuid().safeParse(formData.get("assessmentId"));
  const questionId = z.string().uuid().safeParse(formData.get("questionId"));
  if (!assessmentId.success || !questionId.success) {
    redirect("/teacher/assessments?error=Choose+a+valid+question+to+remove");
  }

  const supabase = await createClient();
  const { error: mappingError } = await supabase
    .from("assessment_questions")
    .delete()
    .eq("assessment_id", assessmentId.data)
    .eq("question_id", questionId.data);
  if (mappingError) redirect(`/teacher/assessments?error=${encodeURIComponent(mappingError.message)}`);

  const { error: questionError } = await supabase.from("questions").delete().eq("id", questionId.data);
  if (questionError) redirect(`/teacher/assessments?error=${encodeURIComponent(questionError.message)}`);

  const { data: questionRows } = await supabase
    .from("assessment_questions")
    .select("marks")
    .eq("assessment_id", assessmentId.data);
  const totalMarks = (questionRows ?? []).reduce((sum, row) => sum + Number(row.marks), 0);
  await supabase.from("assessments").update({ total_marks: totalMarks }).eq("id", assessmentId.data);

  revalidatePath("/teacher/assessments");
  redirect("/teacher/assessments?success=Question+removed");
}

export async function publishAssessment(formData: FormData) {
  await requireProfile(["teacher"]);
  const assessmentId = z.string().uuid().safeParse(formData.get("assessmentId"));
  if (!assessmentId.success) {
    redirect("/teacher/assessments?error=Choose+a+valid+assessment");
  }
  const supabase = await createClient();
  const { data: assessment, error: assessmentError } = await supabase
    .from("assessments")
    .select("opens_at, due_at")
    .eq("id", assessmentId.data)
    .single();
  if (assessmentError) {
    redirect(`/teacher/assessments?error=${encodeURIComponent(assessmentError.message)}`);
  }
  if (!assessment.opens_at || !assessment.due_at) {
    redirect("/teacher/assessments?error=Set+opening+and+due+times+before+publishing");
  }

  const { data: questionRows, count } = await supabase
    .from("assessment_questions")
    .select("marks", { count: "exact" })
    .eq("assessment_id", assessmentId.data);
  if (!count) redirect("/teacher/assessments?error=Add+at+least+one+question+before+publishing");
  const totalMarks = (questionRows ?? []).reduce(
    (sum, row) => sum + Number(row.marks),
    0,
  );
  const { error } = await supabase
    .from("assessments")
    .update({ status: "published", total_marks: totalMarks })
    .eq("id", assessmentId.data);
  if (error) redirect(`/teacher/assessments?error=${encodeURIComponent(error.message)}`);
  revalidatePath("/teacher/assessments");
  revalidatePath("/student/assessments");
  revalidatePath("/student/tests");
  redirect("/teacher/assessments?success=Assessment+published");
}

const reviewSchema = z.object({
  responseId: z.string().uuid(),
  marksAwarded: z.coerce.number().min(0).max(100),
  feedback: z.string().trim().max(2000).optional(),
});

export async function reviewResponse(formData: FormData) {
  const teacher = await requireProfile(["teacher"]);
  const parsed = reviewSchema.safeParse(Object.fromEntries(formData));
  if (!parsed.success) redirect("/teacher/review?error=Check+marks+and+feedback");
  const supabase = await createClient();

  const { data: response, error: responseError } = await supabase
    .from("responses")
    .update({
      marks_awarded: parsed.data.marksAwarded,
      teacher_feedback: parsed.data.feedback || null,
      reviewed_by: teacher.id,
      reviewed_at: new Date().toISOString(),
    })
    .eq("id", parsed.data.responseId)
    .select("attempt_id")
    .single();
  if (responseError) redirect(`/teacher/review?error=${encodeURIComponent(responseError.message)}`);

  const { data: responses, error: listError } = await supabase
    .from("responses")
    .select("marks_awarded, questions(question_type)")
    .eq("attempt_id", response.attempt_id);
  if (listError) redirect(`/teacher/review?error=${encodeURIComponent(listError.message)}`);

  const shortResponses = (responses ?? []).filter((item) => {
    const question = Array.isArray(item.questions) ? item.questions[0] : item.questions;
    return question?.question_type === "short_answer";
  });
  const finished = shortResponses.every((item) => item.marks_awarded !== null);
  if (finished) {
    const { error: releaseError } = await supabase.rpc("release_attempt_results", {
      p_attempt_id: response.attempt_id,
    });
    if (releaseError) {
      redirect(`/teacher/review?error=${encodeURIComponent(releaseError.message)}`);
    }
  }
  revalidatePath("/teacher/review");
  revalidatePath("/teacher");
  revalidatePath("/student/progress");
  revalidatePath("/student/tests");
  revalidatePath("/student/assessments");
  redirect(
    finished
      ? "/teacher/review?success=Review+saved+and+results+updated"
      : "/teacher/review?success=Review+saved",
  );
}

const releaseSchema = z.object({
  attemptId: z.string().uuid(),
});

export async function releaseAttemptResults(formData: FormData) {
  await requireProfile(["teacher"]);
  const parsed = releaseSchema.safeParse(Object.fromEntries(formData));
  if (!parsed.success) redirect("/teacher/review?error=Choose+a+valid+attempt");

  const supabase = await createClient();
  const { data, error } = await supabase.rpc("release_attempt_results", {
    p_attempt_id: parsed.data.attemptId,
  });
  if (error) redirect(`/teacher/review?error=${encodeURIComponent(error.message)}`);

  revalidatePath("/teacher/review");
  revalidatePath("/teacher");
  revalidatePath("/student/progress");
  revalidatePath("/student/tests");
  revalidatePath("/student/assessments");

  const message =
    typeof data === "object" && data && "message" in data && typeof data.message === "string"
      ? data.message
      : "Results updated";
  redirect(`/teacher/review?success=${encodeURIComponent(message)}`);
}

const attendanceSchema = z.object({
  studentId: z.string().uuid(),
  classId: z.string().uuid(),
  status: z.enum(["present", "absent", "excused"]),
});

export async function markAttendance(formData: FormData) {
  const teacher = await requireProfile(["teacher"]);
  const parsed = attendanceSchema.safeParse(Object.fromEntries(formData));
  if (!parsed.success) redirect("/teacher/students?error=Choose+a+class+and+attendance+status");
  const supabase = await createClient();
  const { error } = await supabase.from("attendance").upsert({
    student_id: parsed.data.studentId,
    class_id: parsed.data.classId,
    status: parsed.data.status,
    marked_by: teacher.id,
    marked_at: new Date().toISOString(),
  });
  if (error) redirect(`/teacher/students?error=${encodeURIComponent(error.message)}`);
  revalidatePath("/teacher/students");
  revalidatePath("/parent");
  redirect("/teacher/students?success=Attendance+saved");
}

const assessmentAttendanceSchema = z.object({
  assessmentId: z.string().uuid(),
  studentId: z.string().uuid(),
  decision: z.enum(["verified", "rejected"]),
  note: z.string().trim().max(500).optional().or(z.literal("")),
});

export async function reviewAssessmentAttendance(formData: FormData) {
  const teacher = await requireProfile(["teacher"]);
  const parsed = assessmentAttendanceSchema.safeParse(Object.fromEntries(formData));
  if (!parsed.success) {
    redirect("/teacher/review?error=Choose+a+valid+attendance+request");
  }

  const supabase = await createClient();
  const { error } = await supabase
    .from("assessment_attendance")
    .update({
      status: parsed.data.decision,
      verified_at: new Date().toISOString(),
      verified_by: teacher.id,
      teacher_note: parsed.data.note || null,
    })
    .eq("assessment_id", parsed.data.assessmentId)
    .eq("student_id", parsed.data.studentId)
    .eq("status", "requested");
  if (error) redirect(`/teacher/review?error=${encodeURIComponent(error.message)}`);

  revalidatePath("/teacher/review");
  revalidatePath("/teacher");
  revalidatePath("/student/tests");
  revalidatePath("/student/assessments");
  revalidatePath(`/student/assessments/${parsed.data.assessmentId}`);
  redirect(
    parsed.data.decision === "verified"
      ? "/teacher/review?success=Attendance+verified.+Student+can+start+the+test"
      : "/teacher/review?success=Attendance+rejected",
  );
}

const plannerLectureSchema = z.object({
  lectureNumber: z.coerce.number().int().positive(),
});

export async function markPlannerLectureComplete(formData: FormData) {
  const teacher = await requireProfile(["teacher"]);
  const parsed = plannerLectureSchema.safeParse(Object.fromEntries(formData));
  if (!parsed.success) redirect("/teacher/planner?error=Choose+a+valid+lecture");

  const supabase = await createClient();
  const { error } = await supabase.from("lecture_planner_progress").upsert({
    lecture_number: parsed.data.lectureNumber,
    completed_at: new Date().toISOString(),
    completed_by: teacher.id,
  });
  if (error) redirect(`/teacher/planner?error=${encodeURIComponent(error.message)}`);
  revalidatePath("/teacher/planner");
  revalidatePath("/student/planner");
  redirect(`/teacher/planner?success=Lecture+${String(parsed.data.lectureNumber).padStart(2, "0")}+marked+complete`);
}

export async function markPlannerLectureIncomplete(formData: FormData) {
  await requireProfile(["teacher"]);
  const parsed = plannerLectureSchema.safeParse(Object.fromEntries(formData));
  if (!parsed.success) redirect("/teacher/planner?error=Choose+a+valid+lecture");

  const supabase = await createClient();
  const { error } = await supabase
    .from("lecture_planner_progress")
    .delete()
    .eq("lecture_number", parsed.data.lectureNumber);
  if (error) redirect(`/teacher/planner?error=${encodeURIComponent(error.message)}`);
  revalidatePath("/teacher/planner");
  revalidatePath("/student/planner");
  redirect(`/teacher/planner?success=Lecture+${String(parsed.data.lectureNumber).padStart(2, "0")}+marked+upcoming`);
}

export async function markPlannerCompletedThrough(formData: FormData) {
  const teacher = await requireProfile(["teacher"]);
  const parsed = plannerLectureSchema.safeParse(Object.fromEntries(formData));
  if (!parsed.success) redirect("/teacher/planner?error=Choose+a+valid+lecture");

  const through = parsed.data.lectureNumber;
  const supabase = await createClient();
  const { data: plannerRows, error: plannerError } = await supabase
    .from("lecture_planner")
    .select("lecture_number")
    .lte("lecture_number", through)
    .order("lecture_number");
  if (plannerError) redirect(`/teacher/planner?error=${encodeURIComponent(plannerError.message)}`);

  const now = new Date().toISOString();
  const rows = (plannerRows ?? []).map((row) => ({
    lecture_number: row.lecture_number,
    completed_at: now,
    completed_by: teacher.id,
  }));

  if (rows.length) {
    const { error } = await supabase.from("lecture_planner_progress").upsert(rows);
    if (error) redirect(`/teacher/planner?error=${encodeURIComponent(error.message)}`);
  }

  const { error: clearError } = await supabase
    .from("lecture_planner_progress")
    .delete()
    .gt("lecture_number", through);
  if (clearError) redirect(`/teacher/planner?error=${encodeURIComponent(clearError.message)}`);

  revalidatePath("/teacher/planner");
  revalidatePath("/student/planner");
  redirect(`/teacher/planner?success=Marked+complete+through+Lecture+${String(through).padStart(2, "0")}`);
}
