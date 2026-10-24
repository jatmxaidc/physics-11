"use client";

import { useEffect } from "react";

export default function FocusLecture({ lecture }: { lecture?: string }) {
  useEffect(() => {
    if (!lecture) return;
    const target = document.getElementById(`lecture-${lecture}`);
    if (!target) return;
    target.scrollIntoView({ behavior: "smooth", block: "start" });
    target.classList.add("lecture-focus");
    const timer = window.setTimeout(() => target.classList.remove("lecture-focus"), 2600);
    return () => window.clearTimeout(timer);
  }, [lecture]);

  return null;
}
