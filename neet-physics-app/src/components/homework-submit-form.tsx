"use client";

import { Upload } from "lucide-react";
import { useFormStatus } from "react-dom";
import { submitHomework } from "@/app/actions/student";

function SubmitButton() {
  const { pending } = useFormStatus();
  return (
    <button className="button primary" type="submit" disabled={pending}>
      {pending ? "Submitting…" : "Submit homework"}
    </button>
  );
}

export default function HomeworkSubmitForm({ materialId }: { materialId: string }) {
  return (
    <details className="homework-submit">
      <summary>Submit homework</summary>
      <form action={submitHomework} className="homework-submit-form" encType="multipart/form-data">
        <input type="hidden" name="materialId" value={materialId} />
        <label>
          Written answer
          <textarea
            name="answerText"
            rows={3}
            placeholder="Type your answer, or leave blank if you are only uploading a file."
          />
        </label>
        <label className="file-picker">
          <Upload size={16} />
          <span>Upload PDF, PNG or JPG (optional if you wrote an answer)</span>
          <input name="file" type="file" accept=".pdf,.png,.jpg,.jpeg,application/pdf,image/png,image/jpeg" />
        </label>
        <SubmitButton />
      </form>
    </details>
  );
}
