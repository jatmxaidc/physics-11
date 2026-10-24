import { NextResponse } from "next/server";
import { createClient } from "@/lib/supabase/server";

export async function GET(
  request: Request,
  { params }: { params: Promise<{ id: string }> },
) {
  const { id } = await params;
  const format = new URL(request.url).searchParams.get("format");
  const supabase = await createClient();
  const {
    data: { user },
  } = await supabase.auth.getUser();
  if (!user) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const { data: material, error } = await supabase
    .from("materials")
    .select("title, storage_path, mime_type, status, classes(status)")
    .eq("id", id)
    .eq("status", "published")
    .single();
  if (error || !material?.storage_path) {
    return NextResponse.json({ error: "Material not found" }, { status: 404 });
  }

  const { data: profile } = await supabase
    .from("profiles")
    .select("role")
    .eq("id", user.id)
    .single();
  const classRelation = material.classes as unknown;
  const classRow = (Array.isArray(classRelation) ? classRelation[0] : classRelation) as
    | { status: string }
    | null;
  if (profile?.role === "student" && classRow?.status !== "published") {
    return NextResponse.json({ error: "Material not found" }, { status: 404 });
  }

  if (format === "json") {
    const { data, error: signedUrlError } = await supabase.storage
      .from("class-materials")
      .createSignedUrl(material.storage_path, 60 * 10);
    if (signedUrlError || !data?.signedUrl) {
      return NextResponse.json({ error: "Unable to open material" }, { status: 500 });
    }
    return NextResponse.json({
      url: `/api/materials/${id}?format=inline&t=${Date.now()}`,
      signedUrl: data.signedUrl,
      title: material.title,
      mimeType: material.mime_type,
      expiresIn: 600,
    });
  }

  if (format === "inline" || format === null) {
    const { data: file, error: downloadError } = await supabase.storage
      .from("class-materials")
      .download(material.storage_path);
    if (downloadError || !file) {
      return NextResponse.json({ error: "Unable to open material" }, { status: 500 });
    }

    const filename = `${material.title.replace(/[^\w.\- ]+/g, "").trim() || "notes"}.pdf`;
    return new NextResponse(file.stream(), {
      headers: {
        "Content-Type": material.mime_type || "application/pdf",
        "Content-Disposition": `inline; filename="${filename}"`,
        "Cache-Control": "private, max-age=60",
        "X-Content-Type-Options": "nosniff",
      },
    });
  }

  const { data, error: signedUrlError } = await supabase.storage
    .from("class-materials")
    .createSignedUrl(material.storage_path, 60 * 10);
  if (signedUrlError || !data?.signedUrl) {
    return NextResponse.json({ error: "Unable to open material" }, { status: 500 });
  }
  return NextResponse.redirect(data.signedUrl);
}
