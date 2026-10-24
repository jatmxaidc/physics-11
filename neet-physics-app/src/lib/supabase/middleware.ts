import { createServerClient } from "@supabase/ssr";
import { NextResponse, type NextRequest } from "next/server";

const roleHomes = {
  student: "/student",
  teacher: "/teacher",
  parent: "/parent",
} as const;

type Role = keyof typeof roleHomes;

function roleFromPath(pathname: string): Role | null {
  if (pathname.startsWith("/student")) return "student";
  if (pathname.startsWith("/teacher")) return "teacher";
  if (pathname.startsWith("/parent")) return "parent";
  return null;
}

export async function updateSession(request: NextRequest) {
  const url = process.env.NEXT_PUBLIC_SUPABASE_URL;
  const key = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY;

  if (!url || !key) {
    return NextResponse.next({ request });
  }

  let response = NextResponse.next({ request });
  const supabase = createServerClient(url, key, {
    cookies: {
      getAll: () => request.cookies.getAll(),
      setAll(cookiesToSet) {
        cookiesToSet.forEach(({ name, value }) =>
          request.cookies.set(name, value),
        );
        response = NextResponse.next({ request });
        cookiesToSet.forEach(({ name, value, options }) =>
          response.cookies.set(name, value, options),
        );
      },
    },
  });

  const {
    data: { user },
  } = await supabase.auth.getUser();
  const pathname = request.nextUrl.pathname;
  const isProtected = ["/student", "/teacher", "/parent", "/dashboard", "/account", "/materials"].some(
    (path) => pathname.startsWith(path),
  );

  if (!user && isProtected) {
    const loginUrl = request.nextUrl.clone();
    loginUrl.pathname = "/login";
    loginUrl.searchParams.set("next", pathname);
    return NextResponse.redirect(loginUrl);
  }

  if (
    user?.user_metadata?.must_change_password === true &&
    pathname !== "/account/password"
  ) {
    const passwordUrl = request.nextUrl.clone();
    passwordUrl.pathname = "/account/password";
    passwordUrl.search = "";
    return NextResponse.redirect(passwordUrl);
  }

  if (user && isProtected) {
    const { data: profile } = await supabase
      .from("profiles")
      .select("role")
      .eq("id", user.id)
      .maybeSingle();
    const role = profile?.role as Role | undefined;
    if (role && role in roleHomes) {
      const pathRole = roleFromPath(pathname);
      if (pathRole && pathRole !== role) {
        const home = request.nextUrl.clone();
        home.pathname = roleHomes[role];
        home.search = "";
        return NextResponse.redirect(home);
      }
      if (pathname === "/dashboard" || pathname.startsWith("/dashboard/")) {
        const home = request.nextUrl.clone();
        home.pathname = roleHomes[role];
        home.search = "";
        return NextResponse.redirect(home);
      }
    }
  }

  return response;
}
