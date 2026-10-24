"use client";

import { CheckCircle2, Upload } from "lucide-react";
import { useRouter } from "next/navigation";
import { useState } from "react";
import { createClient } from "@/lib/supabase/client";

const allowedTypes = new Set([
  "application/pdf",
  "application/vnd.ms-powerpoint",
  "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  "image/png",
  "image/jpeg",
]);

export default function TeacherUploadForm({
  classId,
  classStatus,
}: {
  classId: string;
  classStatus: string;
}) {
  const router = useRouter();
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState("");
  const publishNow = classStatus === "published";

  async function upload(formData: FormData) {
    const file = formData.get("file");
    const title = String(formData.get("title") ?? "").trim();
    const kind = String(formData.get("kind") ?? "");
    if (!(file instanceof File) || !file.size || !title) return;
    if (file.size > 50 * 1024 * 1024 || !allowedTypes.has(file.type)) {
      setMessage("Use a PDF, PPTX, PNG or JPG file under 50 MB.");
      return;
    }

    setBusy(true);
    setMessage("");
    const supabase = createClient();
    const { data: authData } = await supabase.auth.getUser();
    if (!authData.user) {
      setMessage("Your session expired. Sign in again.");
      setBusy(false);
      return;
    }
    const safeName = file.name.replace(/[^a-zA-Z0-9._-]+/g, "-");
    const path = `${classId}/${crypto.randomUUID()}-${safeName}`;
    const { error: uploadError } = await supabase.storage
      .from("class-materials")
      .upload(path, file, { contentType: file.type });
    if (uploadError) {
      setMessage(uploadError.message);
      setBusy(false);
      return;
    }
    const { error: insertError } = await supabase.from("materials").insert({
      class_id: classId,
      title,
      kind,
      storage_path: path,
      mime_type: file.type,
      file_size: file.size,
      created_by: authData.user.id,
      status: publishNow ? "published" : "draft",
      published_at: publishNow ? new Date().toISOString() : null,
    });
    if (insertError) {
      await supabase.storage.from("class-materials").remove([path]);
      setMessage(insertError.message);
    } else {
      setMessage(
        publishNow
          ? "Material uploaded and published."
          : "Material saved as draft. Publish the class to share it with students.",
      );
      router.refresh();
    }
    setBusy(false);
  }

  return (
    <form action={upload} className="production-upload-form">
      <div className="form-grid">
        <label>
          Material title
          <input name="title" placeholder="Lecture slides or class notes" required />
        </label>
        <label>
          Material type
          <select name="kind" defaultValue="slides">
            <option value="slides">Lecture slides</option>
            <option value="class_notes">Class notes</option>
            <option value="worksheet">Worksheet</option>
            <option value="formula_sheet">Formula sheet</option>
            <option value="homework">Homework</option>
          </select>
        </label>
        <label className="full-field file-picker">
          <Upload size={17} />
          <span>Select PDF, PPTX or image</span>
          <input name="file" type="file" accept=".pdf,.ppt,.pptx,.png,.jpg,.jpeg" required />
        </label>
      </div>
      {message && (
        <p className={message.includes("uploaded") || message.includes("draft") ? "upload-success" : "login-error"}>
          {(message.includes("uploaded") || message.includes("draft")) && <CheckCircle2 size={15} />}
          {message}
        </p>
      )}
      <button className="button primary" disabled={busy}>
        {busy ? "Uploading securely…" : publishNow ? "Upload and publish" : "Upload as draft"}
      </button>
    </form>
  );
}
