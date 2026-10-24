import { cache } from "react";
import { redirect } from "next/navigation";
import { createClient } from "@/lib/supabase/server";
import type { Role } from "@/lib/data";

export type Profile = {
  id: string;
  full_name: string;
  role: Role;
  avatar_url: string | null;
};

export const getCurrentProfile = cache(async (): Promise<Profile | null> => {
  if (
    !process.env.NEXT_PUBLIC_SUPABASE_URL ||
    !process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY
  ) {
    return null;
  }

  const supabase = await createClient();
  const {
    data: { user },
  } = await supabase.auth.getUser();
  if (!user) return null;

  const { data, error } = await supabase
    .from("profiles")
    .select("id, full_name, role, avatar_url")
    .eq("id", user.id)
    .single();

  if (error) throw new Error(`Unable to load user profile: ${error.message}`);
  return data as Profile;
});

export async function requireProfile(allowed?: Role[]) {
  const profile = await getCurrentProfile();
  if (!profile) redirect("/login");

  if (allowed && !allowed.includes(profile.role)) {
    redirect(`/${profile.role}`);
  }
  return profile;
}
