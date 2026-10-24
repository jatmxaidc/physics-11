import Link from "next/link";
import { notFound, redirect } from "next/navigation";
import PdfReader from "@/components/pdf-reader";
import { requireProfile } from "@/lib/auth";
import { createClient } from "@/lib/supabase/server";

export const dynamic = "force-dynamic";

export default async function MaterialReaderPage({
  params,
}: {
  params: Promise<{ id: string }>;
}) {
  const profile = await requireProfile(["student", "teacher"]);
  const { id } = await params;
  const supabase = await createClient();

  const { data: material, error } = await supabase
    .from("materials")
    .select("id, title, kind, storage_path, mime_type, rich_text, status, classes(status)")
    .eq("id", id)
    .single();

  if (error || !material || material.status !== "published") notFound();

  const classRelation = material.classes as unknown;
  const classRow = (Array.isArray(classRelation) ? classRelation[0] : classRelation) as
    | { status: string }
    | null;
  if (profile.role === "student" && classRow?.status !== "published") notFound();

  const backHref = profile.role === "teacher" ? "/teacher/classes" : "/student/classes";
  const homeHref = profile.role === "teacher" ? "/teacher" : "/student";
  const kindLabel = material.kind.replace(/_/g, " ");

  if (material.rich_text && !material.storage_path) {
    return (
      <div className="pdf-reader">
        <header className="pdf-reader-bar">
          <div className="pdf-reader-identity">
            <Link className="icon-button" href={backHref} aria-label="Back to classes">
              ←
            </Link>
            <Link className="icon-button" href={homeHref} aria-label="Workspace home">
              ⌂
            </Link>
            <div>
              <span className="small-label">{kindLabel} · focus reader</span>
              <h1>{material.title}</h1>
            </div>
          </div>
          <Link className="button secondary pdf-back-classes" href={backHref}>
            My classes
          </Link>
        </header>
        <div className="pdf-reader-stage">
          <article className="pdf-reader-text">
            <h2>{material.title}</h2>
            <p>{material.rich_text}</p>
          </article>
        </div>
      </div>
    );
  }

  if (!material.storage_path) notFound();

  const { error: signedError } = await supabase.storage
    .from("class-materials")
    .createSignedUrl(material.storage_path, 60);
  if (signedError) {
    redirect(`${backHref}?error=${encodeURIComponent("Unable to open class notes")}`);
  }

  return (
    <PdfReader
      materialId={material.id}
      title={material.title}
      kind={material.kind}
      backHref={backHref}
      homeHref={homeHref}
      initialUrl={`/api/materials/${material.id}?format=inline`}
      expiresAt={Date.now() + 60 * 10 * 1000}
    />
  );
}
