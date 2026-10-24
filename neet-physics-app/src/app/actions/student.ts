"use server";

import { revalidatePath } from "next/cache";
import { redirect } from "next/navigation";
import { z } from "zod";
import { requireProfile } from "@/lib/auth";
import { createClient } from "@/lib/supabase/server";

const materialIdSchema = z.string().uuid();

export async function markMaterialComplete(formData: FormData) {
  const student = await requireProfile(["student"]);
  const parsed = materialIdSchema.safeParse(formData.get("materialId"));
  if (!parsed.success) redirect("/student/classes?error=Invalid+material");
  const supabase = await createClient();
  const { error } = await supabase.from("material_progress").upsert({
    material_id: parsed.data,
    student_id: student.id,
    viewed_at: new Date().toISOString(),
    completed_at: new Date().toISOString(),
  });
  if (error) redirect(`/student/classes?error=${encodeURIComponent(error.message)}`);
  revalidatePath("/student/classes");
  redirect("/student/classes?success=Material+marked+complete");
}

const homeworkAllowedTypes = new Set([
  "application/pdf",
  "image/png",
  "image/jpeg",
]);

export async function submitHomework(formData: FormData) {
  const student = await requireProfile(["student"]);
  const parsed = materialIdSchema.safeParse(formData.get("materialId"));
  if (!parsed.success) redirect("/student/classes?error=Invalid+homework");

  const answerText = String(formData.get("answerText") ?? "").trim();
  const file = formData.get("file");
  const hasFile = file instanceof File && file.size > 0;
  if (!hasFile && answerText.length < 2) {
    redirect("/student/classes?error=Add+a+written+answer+or+upload+a+PDF%2FPNG%2FJPG");
  }
  if (hasFile) {
    if (file.size > 20 * 1024 * 1024 || !homeworkAllowedTypes.has(file.type)) {
      redirect("/student/classes?error=Use+a+PDF%2C+PNG+or+JPG+under+20+MB");
    }
  }

  const supabase = await createClient();
  let storagePath: string | null = null;
  let fileName: string | null = null;
  let mimeType: string | null = null;

  if (hasFile && file instanceof File) {
    const safeName = file.name.replace(/[^a-zA-Z0-9._-]+/g, "-");
    storagePath = `homework/${student.id}/${parsed.data}/${crypto.randomUUID()}-${safeName}`;
    fileName = file.name;
    mimeType = file.type;
    const { error: uploadError } = await supabase.storage
      .from("class-materials")
      .upload(storagePath, file, { contentType: file.type, upsert: false });
    if (uploadError) {
      redirect(`/student/classes?error=${encodeURIComponent(uploadError.message)}`);
    }
  }

  const { data: existing } = await supabase
    .from("homework_submissions")
    .select("storage_path, file_name, mime_type")
    .eq("material_id", parsed.data)
    .eq("student_id", student.id)
    .maybeSingle();

  const { error } = await supabase.from("homework_submissions").upsert({
    material_id: parsed.data,
    student_id: student.id,
    status: "submitted",
    submitted_at: new Date().toISOString(),
    answer_text: answerText || null,
    storage_path: storagePath ?? existing?.storage_path ?? null,
    file_name: fileName ?? existing?.file_name ?? null,
    mime_type: mimeType ?? existing?.mime_type ?? null,
  });
  if (error) {
    if (storagePath) await supabase.storage.from("class-materials").remove([storagePath]);
    redirect(`/student/classes?error=${encodeURIComponent(error.message)}`);
  }

  if (storagePath && existing?.storage_path && existing.storage_path !== storagePath) {
    await supabase.storage.from("class-materials").remove([existing.storage_path]);
  }

  revalidatePath("/student/classes");
  revalidatePath("/parent");
  redirect("/student/classes?success=Homework+submitted");
}

const noteSchema = z.object({
  classId: z.string().uuid(),
  materialId: z.string().uuid().optional().or(z.literal("")),
  selectedText: z.string().trim().max(2000).optional().or(z.literal("")),
  noteText: z.string().trim().min(2).max(4000),
  highlightType: z
    .enum(["formula", "definition", "key_concept", "cbse_derivation", "neet_trap", "exam_tip"])
    .optional()
    .or(z.literal("")),
  isFlashcard: z.string().optional(),
  noteId: z.string().uuid().optional().or(z.literal("")),
});

export async function savePrivateNote(formData: FormData) {
  await requireProfile(["student"]);
  const parsed = noteSchema.safeParse(Object.fromEntries(formData));
  if (!parsed.success) redirect("/student/notes?error=Write+a+note+and+choose+a+class");
  const supabase = await createClient();
  const { error } = await supabase.rpc("upsert_private_note", {
    p_class_id: parsed.data.classId,
    p_material_id: parsed.data.materialId || null,
    p_selected_text: parsed.data.selectedText || null,
    p_note_text: parsed.data.noteText,
    p_highlight_type: parsed.data.highlightType || null,
    p_is_flashcard: parsed.data.isFlashcard === "on",
    p_note_id: parsed.data.noteId || null,
  });
  if (error) redirect(`/student/notes?error=${encodeURIComponent(error.message)}`);
  revalidatePath("/student/notes");
  revalidatePath("/student/flashcards");
  revalidatePath("/student/classes");
  redirect("/student/notes?success=Private+note+saved");
}

export async function deletePrivateNote(formData: FormData) {
  await requireProfile(["student"]);
  const noteId = z.string().uuid().safeParse(formData.get("noteId"));
  if (!noteId.success) redirect("/student/notes?error=Invalid+note");
  const supabase = await createClient();
  const { error } = await supabase.rpc("delete_private_note", { p_note_id: noteId.data });
  if (error) redirect(`/student/notes?error=${encodeURIComponent(error.message)}`);
  revalidatePath("/student/notes");
  revalidatePath("/student/flashcards");
  redirect("/student/notes?success=Note+deleted");
}

export async function resolveErrorItem(formData: FormData) {
  const student = await requireProfile(["student"]);
  const errorId = z.string().uuid().safeParse(formData.get("errorId"));
  if (!errorId.success) redirect("/student/retry?error=Invalid+item");
  const supabase = await createClient();
  const { error } = await supabase
    .from("error_log")
    .update({ resolved_at: new Date().toISOString() })
    .eq("id", errorId.data)
    .eq("student_id", student.id);
  if (error) redirect(`/student/retry?error=${encodeURIComponent(error.message)}`);
  revalidatePath("/student/retry");
  revalidatePath("/student/progress");
  revalidatePath("/student");
  redirect("/student/retry?success=Marked+as+reviewed");
}

export async function requestAssessmentAttendance(formData: FormData) {
  const student = await requireProfile(["student"]);
  const assessmentId = z.string().uuid().safeParse(formData.get("assessmentId"));
  if (!assessmentId.success) {
    redirect("/student/tests?error=Choose+a+valid+test");
  }

  const supabase = await createClient();
  const { data: assessment, error: assessmentError } = await supabase
    .from("assessments")
    .select("id, assessment_type, status")
    .eq("id", assessmentId.data)
    .eq("status", "published")
    .single();
  if (assessmentError || !assessment) {
    redirect(`/student/assessments/${assessmentId.data}?error=${encodeURIComponent("Test not available")}`);
  }
  if (assessment.assessment_type === "daily_quiz") {
    redirect(`/student/assessments/${assessmentId.data}?error=${encodeURIComponent("Daily quizzes do not need attendance")}`);
  }

  const { data: existing } = await supabase
    .from("assessment_attendance")
    .select("status")
    .eq("assessment_id", assessmentId.data)
    .eq("student_id", student.id)
    .maybeSingle();

  if (existing?.status === "verified") {
    redirect(`/student/assessments/${assessmentId.data}?success=${encodeURIComponent("Attendance already verified")}`);
  }
  if (existing?.status === "requested") {
    redirect(`/student/assessments/${assessmentId.data}?success=${encodeURIComponent("Waiting for teacher verification")}`);
  }

  if (existing) {
    const { error } = await supabase
      .from("assessment_attendance")
      .update({
        status: "requested",
        requested_at: new Date().toISOString(),
        verified_at: null,
        verified_by: null,
        teacher_note: null,
      })
      .eq("assessment_id", assessmentId.data)
      .eq("student_id", student.id);
    if (error) {
      redirect(`/student/assessments/${assessmentId.data}?error=${encodeURIComponent(error.message)}`);
    }
  } else {
    const { error } = await supabase.from("assessment_attendance").insert({
      assessment_id: assessmentId.data,
      student_id: student.id,
      status: "requested",
      requested_at: new Date().toISOString(),
    });
    if (error) {
      redirect(`/student/assessments/${assessmentId.data}?error=${encodeURIComponent(error.message)}`);
    }
  }

  revalidatePath("/student/tests");
  revalidatePath("/student/assessments");
  revalidatePath(`/student/assessments/${assessmentId.data}`);
  revalidatePath("/teacher/review");
  revalidatePath("/teacher");
  redirect(`/student/assessments/${assessmentId.data}?success=${encodeURIComponent("Attendance marked. Waiting for teacher verification.")}`);
}
