# SMXPhysics

Professional teaching and learning platform for the Class 11 CBSE + NEET
Physics Excellence Programme.

## Included

- Responsive student, teacher and parent workspaces
- Class-wise slides, class notes, short notes and worksheets
- Secure timed daily quizzes with server-side marking
- Weekly test schedule and released results
- Student mastery, score trends and 48-hour error log
- Teacher class creation, direct-to-storage uploads and publishing
- Teacher assessment creation, protected answer keys and manual marking
- Parent weekly summary with attendance and teacher feedback
- Supabase schema, row-level security and private file storage
- Current teaching sequence through Lecture 06

The production entry point requires a real authenticated account. The original
design sample remains available at `/preview` only for visual review.

## Run locally

```bash
npm run dev
```

Open [http://localhost:3000](http://localhost:3000).

## Connect Supabase

1. Create a free Supabase project.
2. Run `supabase/schema.sql` and then `supabase/seed.sql` in its SQL editor.
3. Copy `.env.example` to `.env.local` and add the project URL and anon key.
4. In Authentication, create the student, teacher and parent users.
5. New users receive a student profile automatically. Promote the teacher and
   parent with the administrator-only SQL examples at the end of `seed.sql`.
6. Insert the teacher/student and parent/student relationships shown in
   `seed.sql`.

Production pages do not use hard-coded learning records. Classes, files,
assessments, attempts, attendance and summaries are queried from Supabase.
Files upload directly from the teacher’s browser to private Supabase Storage,
avoiding serverless request-size limits.

Security properties:

- Server-verified sessions and role redirects
- Row-level security on every application table
- Parent access limited to linked summary and released scores
- Student private notes hidden from teacher and parent
- Answer keys isolated from student-readable questions
- Timed attempts started and graded in security-definer database functions
- Signed, ten-minute material download URLs

## Deploy

Live production URL: [https://smxphysics.vercel.app](https://smxphysics.vercel.app)

Vercel project: `manish-jats-projects/smxphysics` (Mumbai / bom1 preferred via
`vercel.json`). Supabase env vars are set for Production, Preview and
Development.

If Auth redirects fail, open Supabase Authentication → URL Configuration and set:

- **Site URL:** `https://smxphysics.vercel.app`
- **Redirect URLs:** `https://smxphysics.vercel.app/**`

Initial Supabase accounts:

- Teacher — Manish Jat: `jatmanis1.mail@gmail.com`
- Teacher — Sapna Mehrat: `sapnamehrat@gmail.com`
- Student: `Sanjaykrishnan363@gmail.com`
- Parent: to be added later

All initial accounts are marked `must_change_password`. After first login the
middleware forces each person to create a private password of at least 6
characters before any dashboard can be opened.

## Important files

- `src/app/student`, `teacher`, `parent` — authenticated production workspaces
- `src/app/actions` — validated role-protected mutations
- `src/components/physics-platform.tsx` — visual preview only
- `src/lib/supabase/` — browser and server Supabase clients
- `supabase/schema.sql` — database, RLS, secure grading and storage
- `supabase/seed.sql` — curriculum and account-link setup
