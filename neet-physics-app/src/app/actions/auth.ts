"use server";

import { redirect } from "next/navigation";
import { z } from "zod";
import { createClient } from "@/lib/supabase/server";

const loginSchema = z.object({
  email: z.string().trim().email(),
  password: z.string().min(6),
  next: z.string().optional(),
});

export async function login(formData: FormData) {
  const parsed = loginSchema.safeParse(Object.fromEntries(formData));
  if (!parsed.success) {
    redirect("/login?error=Enter+a+valid+email+and+password");
  }

  const supabase = await createClient();
  const { error } = await supabase.auth.signInWithPassword({
    email: parsed.data.email,
    password: parsed.data.password,
  });

  if (error) {
    redirect(`/login?error=${encodeURIComponent(error.message)}`);
  }

  const destination =
    parsed.data.next?.startsWith("/") &&
    !parsed.data.next.startsWith("//")
      ? parsed.data.next
      : "/dashboard";
  redirect(destination);
}

export async function logout() {
  const supabase = await createClient();
  await supabase.auth.signOut();
  redirect("/login");
}

const passwordSchema = z
  .object({
    password: z.string().min(6),
    confirmPassword: z.string(),
  })
  .refine((data) => data.password === data.confirmPassword);

export async function changeInitialPassword(formData: FormData) {
  const parsed = passwordSchema.safeParse(Object.fromEntries(formData));
  if (!parsed.success) {
    redirect("/account/password?error=Use+at+least+6+characters+and+match+both+fields");
  }
  const supabase = await createClient();
  const {
    data: { user },
  } = await supabase.auth.getUser();
  if (!user) redirect("/login");

  const { error } = await supabase.auth.updateUser({
    password: parsed.data.password,
    data: { ...user.user_metadata, must_change_password: false },
  });
  if (error) redirect(`/account/password?error=${encodeURIComponent(error.message)}`);
  redirect("/dashboard");
}

export async function changeAccountPassword(formData: FormData) {
  const parsed = passwordSchema.safeParse(Object.fromEntries(formData));
  if (!parsed.success) {
    redirect("/student/account?error=Use+at+least+6+characters+and+match+both+fields");
  }
  const supabase = await createClient();
  const {
    data: { user },
  } = await supabase.auth.getUser();
  if (!user) redirect("/login");

  const { error } = await supabase.auth.updateUser({
    password: parsed.data.password,
    data: { ...user.user_metadata, must_change_password: false },
  });
  if (error) redirect(`/student/account?error=${encodeURIComponent(error.message)}`);
  redirect("/student/account?success=Password+updated");
}
