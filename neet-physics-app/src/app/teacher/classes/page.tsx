import { BookOpen, CheckCircle2, EyeOff, FileText, Pencil, Plus, Trash2 } from "lucide-react";
import {
  createClass,
  deleteMaterial,
  publishClass,
  publishShortNote,
  toggleMaterialPublication,
  unpublishClass,
  updateMaterial,
} from "@/app/actions/teaching";
import FocusLecture from "@/components/focus-lecture";
import { Notice, PageHeading, formatDate } from "@/components/production-ui";
import TeacherUploadForm from "@/components/teacher-upload-form";
import { getChapters, getTeacherClasses } from "@/lib/queries";

export default async function TeacherClassesPage({
  searchParams,
}: {
  searchParams: Promise<{ error?: string; success?: string; lecture?: string }>;
}) {
  const [chapters, classes, params] = await Promise.all([
    getChapters(),
    getTeacherClasses(),
    searchParams,
  ]);
  const focusLecture = params.lecture ? Number(params.lecture) : null;

  return (
    <>
      <FocusLecture lecture={params.lecture} />
      <PageHeading eyebrow="Content management" title="Classes and materials" description="Create each class, upload private files, add short notes and publish to students." />
      {params.error && <Notice kind="error">{params.error}</Notice>}
      {params.success && <Notice kind="success">{params.success}</Notice>}
      {chapters.length === 0 ? (
        <Notice kind="error">No chapters exist. Run the production seed in Supabase before creating classes.</Notice>
      ) : (
        <details className="production-create-panel">
          <summary><span><Plus size={17} /> Create a new class</span><small>Saved as draft</small></summary>
          <form action={createClass} className="form-grid">
            <label>Chapter<select name="chapterId" required>{chapters.map((chapter) => <option value={chapter.id} key={chapter.id}>{chapter.title}</option>)}</select></label>
            <label>Lecture number<input name="lectureNumber" type="number" min="1" required /></label>
            <label>Class title<input name="title" placeholder="Straight-line graphs" required /></label>
            <label>Class date<input name="classDate" type="date" required /></label>
            <label>Duration in minutes<input name="durationMinutes" type="number" defaultValue="120" min="15" max="360" required /></label>
            <label className="full-field">Description<textarea name="description" placeholder="What students will learn in this class" /></label>
            <button className="button primary" type="submit">Create draft class</button>
          </form>
        </details>
      )}
      <div className="teacher-class-stack">
        {classes.map((item) => (
          <details
            className="teacher-class-editor"
            key={item.id}
            id={`lecture-${item.lecture_number}`}
            open={focusLecture === item.lecture_number || undefined}
          >
            <summary>
              <span className="class-dot">{item.lecture_number}</span>
              <div>
                <span className="small-label">{item.chapters?.title} · {formatDate(item.class_date)}</span>
                <h2>{item.title}</h2>
                <p>{item.materials.length} materials</p>
              </div>
              <span className={`status-chip ${item.status}`}>{item.status}</span>
            </summary>
            <div className="teacher-class-editor-body">
              <section>
                <div className="section-head">
                  <div>
                    <h3>Teaching materials</h3>
                    <p>Edit, publish, unpublish or remove each item. Only published items are visible to students.</p>
                  </div>
                </div>
                {item.materials.length === 0 ? (
                  <p className="inline-empty">No materials attached.</p>
                ) : (
                  <>
                    <div className="resource-list">
                      {item.materials.map((material) => (
                        <div className="material-management-item" key={material.id}>
                          <div className="resource-row">
                            <span className="file-icon"><FileText size={19} /></span>
                            <div>
                              <span>{material.kind.replace("_", " ")}</span>
                              <strong>{material.title}</strong>
                              <small>{material.rich_text ? "Read in app" : material.mime_type ?? "File"}</small>
                            </div>
                            {material.storage_path || material.rich_text ? (
                              <a className="button secondary" href={`/materials/${material.id}`} title="Open in reader">
                                <BookOpen size={16} /> Read
                              </a>
                            ) : (
                              <span />
                            )}
                            <span className={`status-chip ${material.status === "published" ? "released" : material.status}`}>
                              {material.status}
                            </span>
                          </div>
                          <details className="production-create-panel material-edit-panel">
                            <summary>
                              <span><Pencil size={15} /> Manage material</span>
                              <small>Edit visibility or delete</small>
                            </summary>
                            <div className="material-edit-body">
                              <form action={updateMaterial} className="form-grid">
                                <input type="hidden" name="materialId" value={material.id} />
                                <input type="hidden" name="lectureNumber" value={item.lecture_number} />
                                <label>Material title<input name="title" defaultValue={material.title} required /></label>
                                <label>
                                  Material type
                                  <select name="kind" defaultValue={material.kind}>
                                    <option value="slides">Lecture slides</option>
                                    <option value="class_notes">Class notes</option>
                                    <option value="short_notes">Short notes</option>
                                    <option value="worksheet">Worksheet</option>
                                    <option value="formula_sheet">Formula sheet</option>
                                    <option value="homework">Homework</option>
                                  </select>
                                </label>
                                <button className="button secondary">Save changes</button>
                              </form>
                              <div className="material-action-row">
                                <form action={toggleMaterialPublication}>
                                  <input type="hidden" name="materialId" value={material.id} />
                                  <input type="hidden" name="lectureNumber" value={item.lecture_number} />
                                  <input
                                    type="hidden"
                                    name="nextStatus"
                                    value={material.status === "published" ? "draft" : "published"}
                                  />
                                  <button className="button secondary">
                                    <EyeOff size={15} />
                                    {material.status === "published" ? "Unpublish" : "Publish"}
                                  </button>
                                </form>
                                <form action={deleteMaterial}>
                                  <input type="hidden" name="materialId" value={material.id} />
                                  <input type="hidden" name="lectureNumber" value={item.lecture_number} />
                                  <button className="button danger">
                                    <Trash2 size={15} /> Remove material
                                  </button>
                                </form>
                              </div>
                            </div>
                          </details>
                        </div>
                      ))}
                    </div>
                    {item.materials.filter((material) => material.rich_text).map((material) => (
                      <section className="published-short-note" key={`${material.id}-text`}>
                        <span>{material.title}</span>
                        <p>{material.rich_text}</p>
                      </section>
                    ))}
                  </>
                )}
              </section>
              <section>
                <div className="section-head">
                  <div>
                    <h3>Upload a teaching file</h3>
                    <p>Uploads go directly to private Supabase Storage.</p>
                  </div>
                </div>
                <TeacherUploadForm classId={item.id} classStatus={item.status} />
              </section>
              <section>
                <div className="section-head">
                  <div>
                    <h3>Add short notes</h3>
                    <p>Concise formulas, definitions, traps and revision points.</p>
                  </div>
                </div>
                <form action={publishShortNote} className="form-grid">
                  <input type="hidden" name="classId" value={item.id} />
                  <label className="full-field">Title<input name="title" placeholder="Lecture lock-lines" required /></label>
                  <label className="full-field">Short-note content<textarea name="content" rows={6} placeholder="Use one clear idea per line…" required /></label>
                  <button className="button secondary" type="submit">Publish short notes</button>
                </form>
              </section>
              {item.status !== "published" && (
                <form action={publishClass} className="publish-class-row">
                  <input type="hidden" name="classId" value={item.id} />
                  <div>
                    <CheckCircle2 size={18} />
                    <p>
                      <strong>Ready for students?</strong>
                      <span>Publishing makes this class visible immediately.</span>
                    </p>
                  </div>
                  <button className="button primary">Publish class</button>
                </form>
              )}
              {item.status === "published" && (
                <form action={unpublishClass} className="publish-class-row unpublish-class-row">
                  <input type="hidden" name="classId" value={item.id} />
                  <div>
                    <EyeOff size={18} />
                    <p>
                      <strong>Hide this complete class?</strong>
                      <span>Unpublishing removes the class and all its materials from student view.</span>
                    </p>
                  </div>
                  <button className="button secondary">Unpublish class</button>
                </form>
              )}
            </div>
          </details>
        ))}
      </div>
    </>
  );
}
