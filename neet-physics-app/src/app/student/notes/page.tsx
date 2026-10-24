import { NotebookPen, Trash2 } from "lucide-react";
import { deletePrivateNote, savePrivateNote } from "@/app/actions/student";
import { EmptyState, Notice, PageHeading } from "@/components/production-ui";
import { requireProfile } from "@/lib/auth";
import { getPrivateNotes, getPublishedClasses } from "@/lib/queries";

export default async function StudentNotesPage({
  searchParams,
}: {
  searchParams: Promise<{ error?: string; success?: string }>;
}) {
  const profile = await requireProfile(["student"]);
  const [classes, notes, params] = await Promise.all([
    getPublishedClasses(),
    getPrivateNotes(profile.id),
    searchParams,
  ]);
  return (
    <>
      <PageHeading eyebrow="Private study" title="Highlights and notes" description="These notes stay on your account. Teachers and parents cannot read them." />
      {params.error && <Notice kind="error">{params.error}</Notice>}
      {params.success && <Notice kind="success">{params.success}</Notice>}
      <details className="production-create-panel" open>
        <summary><span><NotebookPen size={17} /> Add a private note</span><small>Attach it to a published class</small></summary>
        {classes.length === 0 ? (
          <p className="inline-empty">No published classes yet. Notes can be pinned once your teacher publishes a class.</p>
        ) : (
          <form action={savePrivateNote} className="form-grid">
            <label>Class<select name="classId">{classes.map((item) => <option value={item.id} key={item.id}>Lecture {item.lecture_number} · {item.title}</option>)}</select></label>
            <label>Highlight type<select name="highlightType"><option value="key_concept">Key concept</option><option value="formula">Formula</option><option value="definition">Definition</option><option value="cbse_derivation">CBSE derivation</option><option value="neet_trap">NEET trap</option><option value="exam_tip">Exam tip</option></select></label>
            <label className="full-field">Selected line (optional)<input name="selectedText" placeholder="Paste a formula or sentence you want to keep" /></label>
            <label className="full-field">Your note<textarea name="noteText" required placeholder="Write the explanation in your own words." /></label>
            <label><input name="isFlashcard" type="checkbox" /> Save as flashcard</label>
            <button className="button primary">Save private note</button>
          </form>
        )}
      </details>
      {notes.length === 0 ? (
        <EmptyState title="No private notes yet" description="Save formulas, definitions and exam traps as you revise." />
      ) : (
        <div className="teacher-class-stack">
          {notes.map((note) => (
            <article className="production-panel" key={note.id}>
              <div className="panel-head">
                <div>
                  <span className="small-label">{note.highlight_type?.replace("_", " ") ?? "note"}{note.is_flashcard ? " · flashcard" : ""}</span>
                  <h3>Lecture {note.classes?.lecture_number} · {note.classes?.title}</h3>
                </div>
                <form action={deletePrivateNote}>
                  <input type="hidden" name="noteId" value={note.id} />
                  <button className="icon-button" type="submit" title="Delete note"><Trash2 size={16} /></button>
                </form>
              </div>
              {note.selected_text && <blockquote>{note.selected_text}</blockquote>}
              <p>{note.note_text}</p>
            </article>
          ))}
        </div>
      )}
    </>
  );
}
