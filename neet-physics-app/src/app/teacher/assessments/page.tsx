import { CalendarDays, CheckCircle2, Plus, Trash2 } from "lucide-react";
import {
  addMcqQuestion,
  createAssessment,
  deleteAssessmentQuestion,
  publishAssessment,
  scheduleAssessment,
  unpublishAssessment,
  updateAssessment,
  updateMcqQuestion,
} from "@/app/actions/teaching";
import { Notice, PageHeading, formatDate, toDatetimeLocal } from "@/components/production-ui";
import {
  getChapters,
  getTeacherAssessments,
  getTeacherClasses,
} from "@/lib/queries";

export default async function TeacherAssessmentsPage({
  searchParams,
}: {
  searchParams: Promise<{ error?: string; success?: string }>;
}) {
  const [chapters, classes, assessments, params] = await Promise.all([
    getChapters(),
    getTeacherClasses(),
    getTeacherAssessments(),
    searchParams,
  ]);
  return (
    <>
      <PageHeading eyebrow="Assessment management" title="Daily quizzes and weekly tests" description="Create drafts, edit questions, then set opening times yourself. Nothing reaches students until you publish." />
      {params.error && <Notice kind="error">{params.error}</Notice>}
      {params.success && <Notice kind="success">{params.success}</Notice>}
      <details className="production-create-panel">
        <summary><span><Plus size={17} /> Create an assessment</span><small>Saved as draft. Leave times empty until you are ready to schedule.</small></summary>
        <form action={createAssessment} className="form-grid">
          <label>Type<select name="assessmentType"><option value="daily_quiz">Daily quiz</option><option value="weekly_test">Weekly test</option><option value="chapter_test">Chapter test</option></select></label>
          <label>Title<input name="title" required placeholder="Lecture 06 daily quiz" /></label>
          <label>Chapter<select name="chapterId">{chapters.map((chapter) => <option value={chapter.id} key={chapter.id}>{chapter.title}</option>)}</select></label>
          <label>Linked class<select name="classId"><option value="">No specific class</option>{classes.map((item) => <option value={item.id} key={item.id}>Lecture {item.lecture_number} · {item.title}</option>)}</select></label>
          <label>Duration (minutes)<input name="durationMinutes" type="number" defaultValue="12" min="5" required /></label>
          <label>Total marks<input name="totalMarks" type="number" defaultValue="20" min="1" required /></label>
          <label>Opens at<input name="opensAt" type="datetime-local" /></label>
          <label>Due at<input name="dueAt" type="datetime-local" /></label>
          <label className="full-field">Instructions<textarea name="instructions" placeholder="Optional instructions shown before the student starts." /></label>
          <button className="button primary">Create assessment draft</button>
        </form>
      </details>
      <div className="teacher-class-stack">
        {assessments.map((assessment) => (
          <details className="teacher-class-editor" key={assessment.id}>
            <summary>
              <span className="metric-icon violet"><CheckCircle2 size={19} /></span>
              <div>
                <span className="small-label">{assessment.assessment_type.replace("_", " ")} · {assessment.duration_minutes} minutes · {assessment.questions.length} questions</span>
                <h2>{assessment.title}</h2>
                <p>{assessment.opens_at ? `Opens ${formatDate(assessment.opens_at)} · Due ${formatDate(assessment.due_at)}` : "Not scheduled"}</p>
              </div>
              <span className={`status-chip ${assessment.status}`}>{assessment.status}</span>
            </summary>
            <div className="teacher-class-editor-body">
              <section>
                <div className="section-head"><div><h3>Edit details</h3><p>Total marks sync from questions when questions exist. Times are not changed here — use Schedule below.</p></div></div>
                <form action={updateAssessment} className="form-grid">
                  <input type="hidden" name="assessmentId" value={assessment.id} />
                  <input type="hidden" name="totalMarks" value={assessment.total_marks} />
                  <label>Type<select name="assessmentType" defaultValue={assessment.assessment_type}><option value="daily_quiz">Daily quiz</option><option value="weekly_test">Weekly test</option><option value="chapter_test">Chapter test</option></select></label>
                  <label>Title<input name="title" required defaultValue={assessment.title} /></label>
                  <label>Chapter<select name="chapterId" defaultValue={assessment.chapter_id}>{chapters.map((chapter) => <option value={chapter.id} key={chapter.id}>{chapter.title}</option>)}</select></label>
                  <label>Linked class<select name="classId" defaultValue={assessment.class_id ?? ""}><option value="">No specific class</option>{classes.map((item) => <option value={item.id} key={item.id}>Lecture {item.lecture_number} · {item.title}</option>)}</select></label>
                  <label>Duration (minutes)<input name="durationMinutes" type="number" min="5" required defaultValue={assessment.duration_minutes} /></label>
                  <label>Total marks (from questions)<input value={assessment.questions.reduce((sum, item) => sum + Number(item.marks), 0) || assessment.total_marks} readOnly /></label>
                  <label className="full-field">Instructions<textarea name="instructions" defaultValue={assessment.instructions ?? ""} /></label>
                  <button className="button secondary">Save details</button>
                </form>
              </section>
              <section>
                <div className="section-head"><div><h3>{assessment.status === "published" ? "Reschedule" : "Schedule"}</h3><p>Students cannot start until the paper is published and the opening time is reached.</p></div></div>
                <form action={scheduleAssessment} className="form-grid">
                  <input type="hidden" name="assessmentId" value={assessment.id} />
                  <label>Opens at<input name="opensAt" type="datetime-local" required defaultValue={toDatetimeLocal(assessment.opens_at)} /></label>
                  <label>Due at<input name="dueAt" type="datetime-local" required defaultValue={toDatetimeLocal(assessment.due_at)} /></label>
                  <button className="button primary">{assessment.status === "published" ? "Reschedule" : "Save schedule"}</button>
                </form>
              </section>
              <section>
                <div className="section-head"><div><h3>Questions</h3><p>Answer keys stay in the teacher-only table.</p></div></div>
                {assessment.questions.length === 0 ? <p className="inline-empty">No questions yet.</p> : assessment.questions.map((item) => (
                  <details className="production-create-panel" key={item.question.id}>
                    <summary><span>Q{item.sequence}. {item.question.prompt.slice(0, 88)}</span><small>{item.marks} marks</small></summary>
                    <form action={updateMcqQuestion} className="form-grid">
                      <input type="hidden" name="assessmentId" value={assessment.id} />
                      <input type="hidden" name="questionId" value={item.question.id} />
                      <input type="hidden" name="chapterId" value={assessment.chapter_id} />
                      <label className="full-field">Question prompt<textarea name="prompt" required defaultValue={item.question.prompt} /></label>
                      <label>Topic<input name="topic" required defaultValue={item.question.topic} /></label>
                      <label>Marks<input name="marks" type="number" min="1" required defaultValue={item.marks} /></label>
                      <label>Paper<select name="examTag" defaultValue={item.question.exam_tags.includes("CBSE") ? "CBSE" : "NEET"}><option value="CBSE">CBSE</option><option value="NEET">NEET</option></select></label>
                      <label className="full-field">Options, one per line<textarea name="options" rows={5} required defaultValue={item.question.options.join("\n")} /></label>
                      <label>Correct option number<select name="correctIndex" defaultValue={String(item.question.correctIndex)}>{Array.from({ length: 8 }, (_, index) => <option value={index} key={index}>{index + 1}{index === 0 ? " — first option" : ""}</option>)}</select></label>
                      <label className="full-field">Explanation<textarea name="explanation" required defaultValue={item.question.explanation} /></label>
                      <button className="button secondary">Update question</button>
                    </form>
                    <form action={deleteAssessmentQuestion} className="publish-class-row">
                      <input type="hidden" name="assessmentId" value={assessment.id} />
                      <input type="hidden" name="questionId" value={item.question.id} />
                      <div><Trash2 size={18} /><p><strong>Remove this question</strong><span>It is deleted from this paper.</span></p></div>
                      <button className="button secondary">Remove</button>
                    </form>
                  </details>
                ))}
                <form action={addMcqQuestion} className="form-grid">
                  <input type="hidden" name="assessmentId" value={assessment.id} />
                  <input type="hidden" name="chapterId" value={assessment.chapter_id} />
                  <label className="full-field">Question prompt<textarea name="prompt" required /></label>
                  <label>Topic<input name="topic" required placeholder="Straight-line graphs" /></label>
                  <label>Marks<input name="marks" type="number" defaultValue="4" min="1" required /></label>
                  <label>Paper<select name="examTag"><option value="CBSE">CBSE</option><option value="NEET">NEET</option></select></label>
                  <label className="full-field">Options, one per line<textarea name="options" rows={5} placeholder={"Option A\nOption B\nOption C\nOption D"} required /></label>
                  <label>Correct option number<select name="correctIndex">{Array.from({ length: 8 }, (_, index) => <option value={index} key={index}>{index + 1}{index === 0 ? " — first option" : ""}</option>)}</select></label>
                  <label className="full-field">Explanation<textarea name="explanation" required placeholder="Explain why the answer is correct." /></label>
                  <button className="button secondary">Add question</button>
                </form>
              </section>
              {assessment.status !== "published" && (
                <form action={publishAssessment} className="publish-class-row">
                  <input type="hidden" name="assessmentId" value={assessment.id} />
                  <div><CheckCircle2 size={18} /><p><strong>Publish this assessment</strong><span>Requires a schedule. Students see it only after the opening time.</span></p></div>
                  <button className="button primary">Publish assessment</button>
                </form>
              )}
              {assessment.status === "published" && (
                <form action={unpublishAssessment} className="publish-class-row">
                  <input type="hidden" name="assessmentId" value={assessment.id} />
                  <div><CalendarDays size={18} /><p><strong>Unpublish</strong><span>Withdraw it from the student lists. Existing attempts are kept.</span></p></div>
                  <button className="button secondary">Unpublish</button>
                </form>
              )}
            </div>
          </details>
        ))}
      </div>
    </>
  );
}
