-- Planner progress is teacher-marked and independent of live classes.
create table if not exists public.lecture_planner_progress (
  lecture_number integer primary key references public.lecture_planner(lecture_number) on delete cascade,
  completed_at timestamptz not null default now(),
  completed_by uuid not null references public.profiles(id) on delete restrict,
  note text
);

alter table public.lecture_planner_progress enable row level security;

drop policy if exists "authenticated read planner progress" on public.lecture_planner_progress;
create policy "authenticated read planner progress" on public.lecture_planner_progress for select
to authenticated
using (true);

drop policy if exists "teachers manage planner progress" on public.lecture_planner_progress;
create policy "teachers manage planner progress" on public.lecture_planner_progress for all
using (public.current_role() = 'teacher')
with check (public.current_role() = 'teacher');
