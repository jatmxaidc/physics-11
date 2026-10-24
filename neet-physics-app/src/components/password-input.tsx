"use client";

import { Eye, EyeOff } from "lucide-react";
import { useState } from "react";

export default function PasswordInput({
  name,
  placeholder,
  autoComplete,
  minLength,
}: {
  name: string;
  placeholder?: string;
  autoComplete: "current-password" | "new-password";
  minLength: number;
}) {
  const [visible, setVisible] = useState(false);
  return (
    <span className="password-input">
      <input
        name={name}
        type={visible ? "text" : "password"}
        autoComplete={autoComplete}
        placeholder={placeholder}
        minLength={minLength}
        required
      />
      <button
        type="button"
        onClick={() => setVisible((current) => !current)}
        aria-label={visible ? "Hide password" : "Show password"}
        title={visible ? "Hide password" : "Show password"}
      >
        {visible ? <EyeOff size={17} /> : <Eye size={17} />}
      </button>
    </span>
  );
}
