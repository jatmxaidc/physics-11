import Link from "next/link";
import FlashcardDeck from "@/components/flashcard-deck";
import { EmptyState, PageHeading } from "@/components/production-ui";
import { requireProfile } from "@/lib/auth";
import { getPrivateNotes } from "@/lib/queries";

export default async function StudentFlashcardsPage() {
  const profile = await requireProfile(["student"]);
  const notes = (await getPrivateNotes(profile.id)).filter((note) => note.is_flashcard);
  const cards = notes.map((note) => ({
    id: note.id,
    front: note.selected_text?.trim() || note.classes?.title || "Flashcard",
    back: note.note_text ?? "",
    meta: `Lecture ${note.classes?.lecture_number ?? "—"} · ${(note.highlight_type ?? "note").replace(/_/g, " ")}`,
  }));

  return (
    <>
      <PageHeading
        eyebrow="Active recall"
        title="Flashcards"
        description="Notes you marked as flashcards appear here for quick revision."
      />
      {cards.length === 0 ? (
        <EmptyState
          title="No flashcards yet"
          description={
            <>
              Open <Link href="/student/notes">My notes</Link>, tick “Save as flashcard”, and your deck will build here.
            </>
          }
        />
      ) : (
        <FlashcardDeck cards={cards} />
      )}
    </>
  );
}
