import { createClient } from "@/lib/supabase/server";

export type ChapterRow = {
  id: string;
  title: string;
  sequence: number;
  is_locked: boolean;
};

export type MaterialRow = {
  id: string;
  title: string;
  kind: string;
  status: string;
  storage_path: string | null;
  rich_text: string | null;
  mime_type: string | null;
  file_size: number | null;
  updated_at: string;
};

export type ClassRow = {
  id: string;
  lecture_number: number;
  title: string;
  description: string | null;
  class_date: string | null;
  duration_minutes: number;
  status: string;
  tags: string[];
  chapter_id: string;
  chapters: { title: string } | null;
  materials: MaterialRow[];
};

export type AssessmentRow = {
  id: string;
  title: string;
  assessment_type: string;
  duration_minutes: number;
  total_marks: number;
  opens_at: string | null;
  due_at: string | null;
  status: string;
  class_id: string | null;
  chapter_id: string;
  instructions?: string | null;
  attempt_limit?: number;
};

export type StudentAssessmentCard = AssessmentRow & {
  attempt: {
    id: string;
    status: string;
    total_score: number | null;
    submitted_at: string | null;
  } | null;
  action:
    | "start"
    | "resume"
    | "retry"
    | "review"
    | "closed"
    | "mark_attendance"
    | "awaiting_attendance";
  attendanceStatus: "none" | "requested" | "verified" | "rejected" | null;
};

export type TeacherQuestionRow = {
  sequence: number;
  marks: number;
  question: {
    id: string;
    prompt: string;
    topic: string;
    options: string[];
    exam_tags: string[];
    correctIndex: number;
    explanation: string;
  };
};

export type TeacherAssessmentRow = AssessmentRow & {
  questions: TeacherQuestionRow[];
};

export type PrivateNoteRow = {
  id: string;
  class_id: string;
  material_id: string | null;
  selected_text: string | null;
  note_text: string | null;
  highlight_type: string | null;
  is_flashcard: boolean;
  updated_at: string;
  classes: { lecture_number: number; title: string } | null;
};

export type LecturePlannerRow = {
  lecture_number: number;
  planned_date: string | null;
  day_name: string | null;
  phase: string | null;
  chapter_title: string;
  topic: string;
  delivery_focus: string | null;
  assessment_or_homework: string | null;
};

export type PlannerBoardRow = LecturePlannerRow & {
  completion: "completed" | "upcoming";
  completedAt: string | null;
};

function isWithinOpenWindow(row: Pick<AssessmentRow, "status" | "opens_at" | "due_at">, now = new Date()) {
  if (row.status !== "published") return false;
  if (row.opens_at && new Date(row.opens_at) > now) return false;
  if (row.due_at && new Date(row.due_at) < now) return false;
  return true;
}

export async function getChapters() {
  const supabase = await createClient();
  const { data, error } = await supabase
    .from("chapters")
    .select("id, title, sequence, is_locked")
    .order("sequence");
  if (error) throw new Error(error.message);
  return (data ?? []) as ChapterRow[];
}

export async function getPublishedClasses() {
  const supabase = await createClient();
  const { data, error } = await supabase
    .from("classes")
    .select(`
      id, lecture_number, title, description, class_date, duration_minutes,
      status, tags, chapter_id, chapters(title),
      materials(id, title, kind, status, storage_path, rich_text, mime_type, file_size, updated_at)
    `)
    .eq("status", "published")
    .order("lecture_number");
  if (error) throw new Error(error.message);
  return ((data ?? []) as unknown as ClassRow[]).map((row) => ({
    ...row,
    materials: (row.materials ?? []).filter((material) => material.status === "published"),
  }));
}

export async function getTeacherClasses() {
  const supabase = await createClient();
  const { data, error } = await supabase
    .from("classes")
    .select(`
      id, lecture_number, title, description, class_date, duration_minutes,
      status, tags, chapter_id, chapters(title),
      materials(id, title, kind, status, storage_path, rich_text, mime_type, file_size, updated_at)
    `)
    .order("lecture_number");
  if (error) throw new Error(error.message);
  return ((data ?? []) as unknown as ClassRow[]).map((row) => ({
    ...row,
    materials: (row.materials ?? []).filter((material) => material.status !== "archived"),
  }));
}

export async function getOpenAssessments() {
  const supabase = await createClient();
  const { data, error } = await supabase
    .from("assessments")
    .select("id, title, assessment_type, duration_minutes, total_marks, opens_at, due_at, status, class_id, chapter_id, attempt_limit")
    .eq("status", "published")
    .order("due_at");
  if (error) throw new Error(error.message);
  return ((data ?? []) as AssessmentRow[]).filter((row) => isWithinOpenWindow(row));
}

export async function getStudentAssessmentCards(studentId: string): Promise<StudentAssessmentCard[]> {
  const supabase = await createClient();
  const [{ data: assessments, error }, { data: attempts, error: attemptError }, { data: attendanceRows, error: attendanceError }] =
    await Promise.all([
      supabase
        .from("assessments")
        .select("id, title, assessment_type, duration_minutes, total_marks, opens_at, due_at, status, class_id, chapter_id, attempt_limit")
        .eq("status", "published")
        .order("due_at"),
      supabase
        .from("attempts")
        .select("id, assessment_id, status, total_score, submitted_at, started_at")
        .eq("student_id", studentId)
        .order("started_at", { ascending: false }),
      supabase
        .from("assessment_attendance")
        .select("assessment_id, status")
        .eq("student_id", studentId),
    ]);
  if (error) throw new Error(error.message);
  if (attemptError) throw new Error(attemptError.message);
  if (attendanceError) throw new Error(attendanceError.message);

  const latestByAssessment = new Map<string, {
    id: string;
    status: string;
    total_score: number | null;
    submitted_at: string | null;
  }>();
  for (const attempt of attempts ?? []) {
    const assessmentId = attempt.assessment_id as string;
    if (!latestByAssessment.has(assessmentId)) {
      latestByAssessment.set(assessmentId, {
        id: attempt.id as string,
        status: attempt.status as string,
        total_score: attempt.total_score == null ? null : Number(attempt.total_score),
        submitted_at: (attempt.submitted_at as string | null) ?? null,
      });
    }
  }

  const attendanceByAssessment = new Map<string, string>();
  for (const row of attendanceRows ?? []) {
    attendanceByAssessment.set(row.assessment_id as string, row.status as string);
  }

  return ((assessments ?? []) as AssessmentRow[])
    .filter((row) => isWithinOpenWindow(row) || latestByAssessment.has(row.id))
    .map((row) => {
      const attempt = latestByAssessment.get(row.id) ?? null;
      const attendanceRaw = attendanceByAssessment.get(row.id) ?? null;
      const attendanceStatus = (attendanceRaw ?? "none") as StudentAssessmentCard["attendanceStatus"];
      const needsAttendanceGate = row.assessment_type !== "daily_quiz";
      let action: StudentAssessmentCard["action"] = "start";
      if (attempt?.status === "in_progress") action = "resume";
      else if (
        row.assessment_type === "daily_quiz"
        && attempt
        && isWithinOpenWindow(row)
      ) {
        action = "retry";
      }
      else if (attempt && ["submitted", "teacher_review", "released"].includes(attempt.status)) {
        action = "review";
      } else if (!isWithinOpenWindow(row)) {
        action = "closed";
      } else if (attempt && (row.attempt_limit ?? 1) <= 1) {
        action = "review";
      } else if (needsAttendanceGate && attendanceStatus !== "verified") {
        action = attendanceStatus === "requested" ? "awaiting_attendance" : "mark_attendance";
      }
      return { ...row, attempt, action, attendanceStatus };
    })
    .filter((row) => row.action !== "closed" || Boolean(row.attempt));
}

export async function getTeacherAssessments() {
  const supabase = await createClient();
  const { data, error } = await supabase
    .from("assessments")
    .select(`
      id, title, assessment_type, duration_minutes, total_marks, opens_at, due_at, status,
      class_id, chapter_id, instructions,
      assessment_questions(
        sequence, marks,
        questions(id, prompt, topic, options, exam_tags, question_keys(correct_answer, explanation))
      )
    `)
    .order("created_at", { ascending: false });
  if (error) throw new Error(error.message);

  return ((data ?? []) as unknown as Array<AssessmentRow & {
    assessment_questions: Array<{
      sequence: number;
      marks: number;
      questions: unknown;
    }>;
  }>).map((row) => ({
    id: row.id,
    title: row.title,
    assessment_type: row.assessment_type,
    duration_minutes: row.duration_minutes,
    total_marks: row.total_marks,
    opens_at: row.opens_at,
    due_at: row.due_at,
    status: row.status,
    class_id: row.class_id,
    chapter_id: row.chapter_id,
    instructions: row.instructions ?? null,
    questions: (row.assessment_questions ?? [])
      .slice()
      .sort((a, b) => a.sequence - b.sequence)
      .flatMap((mapping) => {
        const relation = mapping.questions;
        const question = Array.isArray(relation) ? relation[0] : relation;
        if (!question || typeof question !== "object") return [];
        const item = question as {
          id: string;
          prompt: string;
          topic: string;
          options: unknown;
          exam_tags: string[] | null;
          question_keys: unknown;
        };
        const keyRelation = item.question_keys;
        const key = Array.isArray(keyRelation) ? keyRelation[0] : keyRelation;
        const keyRow = (key ?? {}) as { correct_answer?: unknown; explanation?: string | null };
        return [{
          sequence: mapping.sequence,
          marks: Number(mapping.marks),
          question: {
            id: item.id,
            prompt: item.prompt,
            topic: item.topic,
            options: Array.isArray(item.options) ? item.options.map(String) : [],
            exam_tags: item.exam_tags ?? [],
            correctIndex: Number(keyRow.correct_answer ?? 0),
            explanation: keyRow.explanation ?? "",
          },
        }];
      }),
  })) as TeacherAssessmentRow[];
}

export async function getAssessmentForStudent(assessmentId: string) {
  const supabase = await createClient();
  const {
    data: { user },
  } = await supabase.auth.getUser();
  if (!user) throw new Error("Not signed in");

  const { data: assessment, error: assessmentError } = await supabase
    .from("assessments")
    .select("id, title, assessment_type, duration_minutes, total_marks, opens_at, due_at, status, instructions")
    .eq("id", assessmentId)
    .eq("status", "published")
    .single();
  if (assessmentError) throw new Error(assessmentError.message);

  let activeAttempt: {
    id: string;
    started_at: string;
    status: string;
    focus_leave_count?: number | null;
  } | null = null;

  const activeAttemptQuery = await supabase
    .from("attempts")
    .select("id, started_at, status, focus_leave_count")
    .eq("assessment_id", assessmentId)
    .eq("student_id", user.id)
    .eq("status", "in_progress")
    .order("started_at", { ascending: false })
    .limit(1)
    .maybeSingle();

  if (activeAttemptQuery.error?.message?.includes("focus_leave_count")) {
    const fallback = await supabase
      .from("attempts")
      .select("id, started_at, status")
      .eq("assessment_id", assessmentId)
      .eq("student_id", user.id)
      .eq("status", "in_progress")
      .order("started_at", { ascending: false })
      .limit(1)
      .maybeSingle();
    if (fallback.error) throw new Error(fallback.error.message);
    activeAttempt = fallback.data
      ? { ...(fallback.data as { id: string; started_at: string; status: string }), focus_leave_count: 0 }
      : null;
  } else if (activeAttemptQuery.error) {
    throw new Error(activeAttemptQuery.error.message);
  } else {
    activeAttempt = activeAttemptQuery.data as typeof activeAttempt;
  }

  const canResume = Boolean(activeAttempt);
  if (!canResume && !isWithinOpenWindow(assessment)) {
    const opensLater = assessment.opens_at && new Date(assessment.opens_at) > new Date();
    throw new Error(
      opensLater
        ? "This assessment is not open yet."
        : "This assessment is closed or past its due time.",
    );
  }

  const { data: attendance } = await supabase
    .from("assessment_attendance")
    .select("status, requested_at, verified_at")
    .eq("assessment_id", assessmentId)
    .eq("student_id", user.id)
    .maybeSingle();

  const attendanceStatus = (attendance?.status ?? "none") as
    | "none"
    | "requested"
    | "verified"
    | "rejected";
  const needsAttendanceGate = assessment.assessment_type !== "daily_quiz";
  const attendanceVerified = !needsAttendanceGate || attendanceStatus === "verified" || canResume;

  let mappings: unknown[] = [];
  if (attendanceVerified) {
    const { data, error: questionsError } = await supabase
      .from("assessment_questions")
      .select("sequence, marks, questions(id, prompt, question_type, options, topic, difficulty, exam_tags)")
      .eq("assessment_id", assessmentId)
      .order("sequence");
    if (questionsError) throw new Error(questionsError.message);
    mappings = data ?? [];
  }

  return {
    assessment,
    questions: mappings,
    activeAttempt: activeAttempt
      ? {
          id: activeAttempt.id,
          started_at: activeAttempt.started_at,
          focus_leave_count: Number(activeAttempt.focus_leave_count ?? 0),
        }
      : null,
    attendance: {
      required: needsAttendanceGate,
      status: attendanceStatus,
      requested_at: (attendance?.requested_at as string | null) ?? null,
      verified_at: (attendance?.verified_at as string | null) ?? null,
      canStart: attendanceVerified,
    },
  };
}

export async function getPendingAssessmentAttendance() {
  const supabase = await createClient();
  const { data, error } = await supabase
    .from("assessment_attendance")
    .select(`
      assessment_id, student_id, status, requested_at, verified_at, teacher_note,
      profiles!assessment_attendance_student_id_fkey(full_name),
      assessments(id, title, assessment_type, due_at, total_marks, duration_minutes)
    `)
    .eq("status", "requested")
    .order("requested_at");
  if (error) throw new Error(error.message);
  return data ?? [];
}

export async function getStudentAttempts(studentId: string) {
  const supabase = await createClient();
  const { data, error } = await supabase
    .from("attempts")
    .select(`
      id, assessment_id, status, started_at, submitted_at, released_at,
      objective_score, subjective_score, total_score, teacher_feedback, focus_leave_count,
      assessments(id, title, assessment_type, total_marks, due_at)
    `)
    .eq("student_id", studentId)
    .order("started_at", { ascending: false });
  if (error) throw new Error(error.message);
  return data ?? [];
}

export type StudentErrorRow = {
  id: string;
  error_type: string;
  retry_due_at: string;
  resolved_at: string | null;
  created_at: string;
  prompt: string;
  topic: string;
  explanation: string | null;
};

export async function getStudentOpenErrors(studentId: string): Promise<StudentErrorRow[]> {
  const supabase = await createClient();
  const { data, error } = await supabase
    .from("error_log")
    .select(`
      id, error_type, retry_due_at, resolved_at, created_at,
      responses(
        questions(prompt, topic, question_keys(explanation))
      )
    `)
    .eq("student_id", studentId)
    .is("resolved_at", null)
    .order("retry_due_at");
  if (error) throw new Error(error.message);

  return (data ?? []).map((row) => {
    const response = Array.isArray(row.responses) ? row.responses[0] : row.responses;
    const question = response && typeof response === "object"
      ? (Array.isArray((response as { questions?: unknown }).questions)
        ? (response as { questions: unknown[] }).questions[0]
        : (response as { questions?: unknown }).questions)
      : null;
    const q = (question ?? {}) as { prompt?: string; topic?: string; question_keys?: unknown };
    const key = Array.isArray(q.question_keys) ? q.question_keys[0] : q.question_keys;
    return {
      id: row.id as string,
      error_type: row.error_type as string,
      retry_due_at: row.retry_due_at as string,
      resolved_at: (row.resolved_at as string | null) ?? null,
      created_at: row.created_at as string,
      prompt: q.prompt ?? "Question",
      topic: q.topic ?? "Physics",
      explanation: (key as { explanation?: string | null } | null)?.explanation ?? null,
    };
  });
}

export async function getClassLearningState(studentId: string) {
  const supabase = await createClient();
  const [{ data: progress, error: progressError }, { data: homework, error: homeworkError }] =
    await Promise.all([
      supabase
        .from("material_progress")
        .select("material_id, completed_at")
        .eq("student_id", studentId),
      supabase
        .from("homework_submissions")
        .select("material_id, status, submitted_at, answer_text, storage_path, file_name, mime_type, teacher_feedback")
        .eq("student_id", studentId),
    ]);
  if (progressError) throw new Error(progressError.message);
  if (homeworkError) throw new Error(homeworkError.message);
  return {
    completedMaterialIds: new Set(
      (progress ?? [])
        .filter((row) => row.completed_at)
        .map((row) => row.material_id as string),
    ),
    homeworkByMaterial: new Map(
      (homework ?? []).map((row) => [
        row.material_id as string,
        {
          status: row.status as string,
          submitted_at: (row.submitted_at as string | null) ?? null,
          answer_text: (row.answer_text as string | null) ?? null,
          storage_path: (row.storage_path as string | null) ?? null,
          file_name: (row.file_name as string | null) ?? null,
          mime_type: (row.mime_type as string | null) ?? null,
          teacher_feedback: (row.teacher_feedback as string | null) ?? null,
        },
      ]),
    ),
  };
}

export type AttemptReviewItem = {
  question_id: string;
  prompt: string;
  topic: string;
  question_type: string;
  options: unknown;
  marks: number;
  response: unknown;
  is_correct: boolean | null;
  marks_awarded: number | null;
  teacher_feedback: string | null;
  correct_answer: unknown;
  explanation: string | null;
};

export type AttemptReview = {
  attempt: {
    id: string;
    status: string;
    started_at: string;
    submitted_at: string | null;
    released_at: string | null;
    objective_score: number | null;
    subjective_score: number | null;
    total_score: number | null;
    teacher_feedback: string | null;
  };
  assessment: {
    id: string;
    title: string;
    assessment_type: string;
    total_marks: number;
    due_at?: string | null;
    explanations_after_submission: boolean;
  };
  show_explanations: boolean;
  results_visible?: boolean;
  unlock_at?: string | null;
  items: AttemptReviewItem[];
};

export async function getAttemptReview(attemptId: string): Promise<AttemptReview> {
  const supabase = await createClient();
  const { data, error } = await supabase.rpc("get_attempt_review", {
    p_attempt_id: attemptId,
  });
  if (error) throw new Error(error.message);
  return data as AttemptReview;
}

export async function getTeacherReviewAttempts() {
  const supabase = await createClient();
  const { data, error } = await supabase
    .from("attempts")
    .select(`
      id, student_id, status, submitted_at, objective_score, total_score,
      profiles!attempts_student_id_fkey(full_name),
      assessments(title, total_marks, assessment_type),
      responses(id, question_id, response, marks_awarded, teacher_feedback, questions(prompt, topic, marks, question_type))
    `)
    .eq("status", "teacher_review")
    .order("submitted_at");
  if (error) throw new Error(error.message);
  return data ?? [];
}

export function splitTeacherReviewQueue(
  attempts: Awaited<ReturnType<typeof getTeacherReviewAttempts>>,
) {
  const shortAnswerQueue = attempts.flatMap((attempt) => {
    const student = Array.isArray(attempt.profiles) ? attempt.profiles[0] : attempt.profiles;
    const assessment = Array.isArray(attempt.assessments) ? attempt.assessments[0] : attempt.assessments;
    return (attempt.responses ?? []).flatMap((response) => {
      const question = Array.isArray(response.questions) ? response.questions[0] : response.questions;
      if (question?.question_type !== "short_answer" || response.marks_awarded !== null) return [];
      return [{ attempt, response, question, student, assessment }];
    });
  });

  const readyToRelease = attempts.filter((attempt) => {
    const pendingShort = (attempt.responses ?? []).some((response) => {
      const question = Array.isArray(response.questions) ? response.questions[0] : response.questions;
      return question?.question_type === "short_answer" && response.marks_awarded === null;
    });
    return !pendingShort;
  });

  return { shortAnswerQueue, readyToRelease };
}

export async function getLinkedStudent(adultId: string) {
  const students = await getLinkedStudents(adultId);
  return students[0] ?? null;
}

export async function getLinkedStudents(adultId: string) {
  const supabase = await createClient();
  const { data, error } = await supabase
    .from("student_links")
    .select("student_id, profiles!student_links_student_id_fkey(id, full_name, avatar_url)")
    .eq("adult_id", adultId)
    .order("created_at", { ascending: true });
  if (error) throw new Error(error.message);
  return data ?? [];
}

export async function getTeacherLinkedAttempts() {
  const supabase = await createClient();
  const {
    data: { user },
  } = await supabase.auth.getUser();
  if (!user) return [];

  const links = await getLinkedStudents(user.id);
  const studentIds = links
    .map((link) => {
      const relation = link.profiles as unknown;
      const profile = (Array.isArray(relation) ? relation[0] : relation) as { id?: string } | null;
      return profile?.id ?? link.student_id;
    })
    .filter(Boolean);
  if (studentIds.length === 0) return [];

  const { data, error } = await supabase
    .from("attempts")
    .select(`
      id, student_id, assessment_id, status, started_at, submitted_at, released_at,
      objective_score, subjective_score, total_score, teacher_feedback, focus_leave_count,
      profiles!attempts_student_id_fkey(id, full_name),
      assessments(id, title, assessment_type, total_marks, due_at)
    `)
    .in("student_id", studentIds)
    .neq("status", "in_progress")
    .order("submitted_at", { ascending: false });
  if (error) throw new Error(error.message);
  return data ?? [];
}

export async function getStudentSummaryData(studentId: string) {
  const supabase = await createClient();
  const { data, error } = await supabase.rpc("get_student_summary", {
    p_student_id: studentId,
  });
  if (error) throw new Error(error.message);
  return data as {
    attendance_total: number;
    attendance_present: number;
    attendance_percent: number;
    homework_total: number;
    homework_complete: number;
    homework_percent: number;
    attempt_total: number;
    attempt_released: number;
    average_accuracy: number;
    open_errors: number;
  };
}

export async function createMaterialSignedUrl(path: string) {
  const supabase = await createClient();
  const { data, error } = await supabase.storage
    .from("class-materials")
    .createSignedUrl(path, 60 * 10);
  if (error) throw new Error(error.message);
  return data.signedUrl;
}

export async function getPrivateNotes(studentId: string) {
  const supabase = await createClient();
  const { data, error } = await supabase
    .from("private_notes")
    .select("id, class_id, material_id, selected_text, note_text, highlight_type, is_flashcard, updated_at, classes(lecture_number, title)")
    .eq("student_id", studentId)
    .order("updated_at", { ascending: false });
  if (error) throw new Error(error.message);
  return (data ?? []) as unknown as PrivateNoteRow[];
}

export async function getLecturePlanner() {
  const supabase = await createClient();
  const { data, error } = await supabase
    .from("lecture_planner")
    .select("lecture_number, planned_date, day_name, phase, chapter_title, topic, delivery_focus, assessment_or_homework")
    .order("lecture_number");
  if (error) throw new Error(error.message);
  return (data ?? []) as LecturePlannerRow[];
}

export async function getPlannerBoard(): Promise<PlannerBoardRow[]> {
  const supabase = await createClient();
  const [planner, progressResult] = await Promise.all([
    getLecturePlanner(),
    supabase
      .from("lecture_planner_progress")
      .select("lecture_number, completed_at"),
  ]);

  if (progressResult.error) throw new Error(progressResult.error.message);

  const completedAtByLecture = new Map(
    (progressResult.data ?? []).map((row) => [row.lecture_number as number, row.completed_at as string]),
  );

  return planner.map((row) => {
    const completedAt = completedAtByLecture.get(row.lecture_number) ?? null;
    return {
      ...row,
      completion: completedAt ? "completed" : "upcoming",
      completedAt,
    };
  });
}
