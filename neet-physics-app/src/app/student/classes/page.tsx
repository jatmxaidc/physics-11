import { BookOpen, CheckCircle2, FileText, NotebookPen } from "lucide-react";
import { markMaterialComplete } from "@/app/actions/student";
import FocusLecture from "@/components/focus-lecture";
import HomeworkSubmitForm from "@/components/homework-submit-form";
import { PageHeading, EmptyState, Notice, formatDate } from "@/components/production-ui";
import { requireProfile } from "@/lib/auth";
import { getClassLearningState, getPublishedClasses } from "@/lib/queries";

const kindLabels: Record<string, string> = {
  slides: "Lecture slides",
  class_notes: "Class notes",
  short_notes: "Short notes",
  worksheet: "Worksheet",
  formula_sheet: "Formula sheet",
  homework: "Homework",
};

export default async function StudentClassesPage({
  searchParams,
}: {
  searchParams: Promise<{ lecture?: string; error?: string; success?: string }>;
}) {
  const profile = await requireProfile(["student"]);
  const params = await searchParams;
  const classes = await getPublishedClasses();
  let learning: Awaited<ReturnType<typeof getClassLearningState>> | null = null;
  let learningError = "";
  try {
    learning = await getClassLearningState(profile.id);
  } catch (error) {
    learningError = error instanceof Error ? error.message : "Could not load completion state.";
    learning = {
      completedMaterialIds: new Set(),
      homeworkByMaterial: new Map(),
    };
  }

  return (
    <>
      <FocusLecture lecture={params.lecture} />
      <PageHeading
        eyebrow="Learning library"
        title="My classes"
        description="Teacher-published materials arranged in the exact order taught."
      />
      {params.error && <Notice kind="error">{params.error}</Notice>}
      {params.success && <Notice kind="success">{params.success}</Notice>}
      {learningError && <Notice kind="error">{learningError}</Notice>}
      {classes.length === 0 ? (
        <EmptyState title="No published classes" description="Your teacher’s published classes will appear here." />
      ) : (
        <div className="production-class-library">
          {classes.map((item) => (
            <article
              className="production-class-card"
              key={item.id}
              id={`lecture-${item.lecture_number}`}
            >
              <header>
                <span className="class-number">{String(item.lecture_number).padStart(2, "0")}</span>
                <div>
                  <span className="small-label">
                    {item.chapters?.title ?? "Physics"} · {formatDate(item.class_date)}
                  </span>
                  <h2>{item.title}</h2>
                  <p>{item.description}</p>
                </div>
              </header>
              {item.materials.length === 0 ? (
                <p className="inline-empty">No materials have been attached to this class.</p>
              ) : (
                <div className="resource-list">
                  {item.materials.map((material) => {
                    const homework = learning.homeworkByMaterial.get(material.id);
                    const completed = learning.completedMaterialIds.has(material.id);
                    const isHomework = material.kind === "homework";
                    const done = isHomework
                      ? homework?.status === "submitted" || homework?.status === "reviewed"
                      : completed;
                    return (
                      <div className={isHomework ? "homework-block" : undefined} key={material.id}>
                        <div className="resource-row">
                          <span className={`file-icon ${material.kind === "short_notes" ? "note" : ""}`}>
                            {material.kind === "short_notes" ? <NotebookPen size={20} /> : <FileText size={20} />}
                          </span>
                          <div>
                            <span>{kindLabels[material.kind] ?? material.kind}</span>
                            <strong>{material.title}</strong>
                            <small>
                              {done
                                ? isHomework
                                  ? homework?.file_name
                                    ? `Submitted · ${homework.file_name}`
                                    : "Submitted"
                                  : "Completed"
                                : material.rich_text
                                  ? "Read in app"
                                  : material.mime_type}
                            </small>
                            {isHomework && homework?.answer_text && (
                              <p className="homework-answer-preview">{homework.answer_text}</p>
                            )}
                            {isHomework && homework?.teacher_feedback && (
                              <p className="homework-feedback">Teacher: {homework.teacher_feedback}</p>
                            )}
                          </div>
                          {material.storage_path || material.rich_text ? (
                            <a className="icon-button" href={`/materials/${material.id}`} title="Open in reader">
                              <BookOpen size={18} />
                            </a>
                          ) : (
                            <span />
                          )}
                          {isHomework ? (
                            done ? (
                              <button className="resource-complete-button done" type="button" disabled>
                                <CheckCircle2 size={12} /> Submitted
                              </button>
                            ) : (
                              <span />
                            )
                          ) : done ? (
                            <button className="resource-complete-button done" type="button" disabled>
                              <CheckCircle2 size={12} /> Completed
                            </button>
                          ) : (
                            <form action={markMaterialComplete}>
                              <input type="hidden" name="materialId" value={material.id} />
                              <button className="resource-complete-button">Complete</button>
                            </form>
                          )}
                        </div>
                        {isHomework && !done && <HomeworkSubmitForm materialId={material.id} />}
                      </div>
                    );
                  })}
                </div>
              )}
            </article>
          ))}
        </div>
      )}
    </>
  );
}
