-- Orbit Physics: production schema for Supabase PostgreSQL
-- Run in the Supabase SQL editor after creating a project.

create extension if not exists "pgcrypto";

create type public.user_role as enum ('student', 'teacher', 'parent');
create type public.publish_status as enum ('draft', 'scheduled', 'published', 'archived');
create type public.assessment_type as enum ('daily_quiz', 'weekly_test', 'chapter_test');
create type public.question_type as enum ('mcq', 'multi_select', 'numerical', 'assertion_reason', 'short_answer');
create type public.attempt_status as enum ('in_progress', 'submitted', 'teacher_review', 'released');

create table public.profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  full_name text not null,
  role public.user_role not null,
  avatar_url text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table public.student_links (
  student_id uuid not null references public.profiles(id) on delete cascade,
  adult_id uuid not null references public.profiles(id) on delete cascade,
  relationship public.user_role not null check (relationship in ('teacher', 'parent')),
  created_at timestamptz not null default now(),
  primary key (student_id, adult_id)
);

create table public.chapters (
  id uuid primary key default gen_random_uuid(),
  title text not null,
  sequence integer not null unique,
  description text,
  is_locked boolean not null default true,
  created_at timestamptz not null default now()
);

create table public.classes (
  id uuid primary key default gen_random_uuid(),
  chapter_id uuid not null references public.chapters(id) on delete restrict,
  lecture_number integer not null,
  title text not null,
  description text,
  class_date date,
  duration_minutes integer not null default 120,
  tags text[] not null default '{}',
  status public.publish_status not null default 'draft',
  published_at timestamptz,
  created_by uuid not null references public.profiles(id) on delete restrict,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (chapter_id, lecture_number)
);

create table public.materials (
  id uuid primary key default gen_random_uuid(),
  class_id uuid not null references public.classes(id) on delete cascade,
  title text not null,
  kind text not null check (kind in ('slides', 'class_notes', 'short_notes', 'worksheet', 'formula_sheet', 'homework')),
  storage_path text,
  rich_text text,
  mime_type text,
  file_size bigint,
  version integer not null default 1,
  status public.publish_status not null default 'draft',
  published_at timestamptz,
  created_by uuid not null references public.profiles(id) on delete restrict,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  check (storage_path is not null or rich_text is not null)
);

create table public.material_progress (
  material_id uuid not null references public.materials(id) on delete cascade,
  student_id uuid not null references public.profiles(id) on delete cascade,
  viewed_at timestamptz,
  completed_at timestamptz,
  bookmarked boolean not null default false,
  primary key (material_id, student_id)
);

create table public.attendance (
  class_id uuid not null references public.classes(id) on delete cascade,
  student_id uuid not null references public.profiles(id) on delete cascade,
  status text not null check (status in ('present', 'absent', 'excused')),
  marked_by uuid not null references public.profiles(id) on delete restrict,
  marked_at timestamptz not null default now(),
  primary key (class_id, student_id)
);

create table public.homework_submissions (
  material_id uuid not null references public.materials(id) on delete cascade,
  student_id uuid not null references public.profiles(id) on delete cascade,
  status text not null check (status in ('assigned', 'submitted', 'reviewed', 'late')),
  submitted_at timestamptz,
  teacher_feedback text,
  primary key (material_id, student_id)
);

create table public.assessments (
  id uuid primary key default gen_random_uuid(),
  class_id uuid references public.classes(id) on delete set null,
  chapter_id uuid not null references public.chapters(id) on delete restrict,
  title text not null,
  assessment_type public.assessment_type not null,
  instructions text,
  duration_minutes integer not null,
  total_marks numeric(7,2) not null,
  negative_marking numeric(5,2) not null default 0,
  attempt_limit integer not null default 1,
  randomize_questions boolean not null default false,
  randomize_options boolean not null default false,
  explanations_after_submission boolean not null default true,
  opens_at timestamptz,
  due_at timestamptz,
  status public.publish_status not null default 'draft',
  created_by uuid not null references public.profiles(id) on delete restrict,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table public.questions (
  id uuid primary key default gen_random_uuid(),
  chapter_id uuid not null references public.chapters(id) on delete restrict,
  class_id uuid references public.classes(id) on delete set null,
  prompt text not null,
  question_type public.question_type not null,
  options jsonb,
  marks numeric(5,2) not null default 4,
  difficulty text not null check (difficulty in ('foundation', 'standard', 'challenging')),
  exam_tags text[] not null default '{}',
  topic text not null,
  misconception_tag text,
  created_by uuid not null references public.profiles(id) on delete restrict,
  created_at timestamptz not null default now()
);

-- Answer keys are isolated from student-readable question rows.
create table public.question_keys (
  question_id uuid primary key references public.questions(id) on delete cascade,
  correct_answer jsonb not null,
  explanation text,
  created_at timestamptz not null default now()
);

create table public.assessment_questions (
  assessment_id uuid not null references public.assessments(id) on delete cascade,
  question_id uuid not null references public.questions(id) on delete restrict,
  sequence integer not null,
  marks numeric(5,2) not null,
  primary key (assessment_id, question_id),
  unique (assessment_id, sequence)
);

create table public.attempts (
  id uuid primary key default gen_random_uuid(),
  assessment_id uuid not null references public.assessments(id) on delete cascade,
  student_id uuid not null references public.profiles(id) on delete cascade,
  attempt_number integer not null default 1,
  status public.attempt_status not null default 'in_progress',
  started_at timestamptz not null default now(),
  submitted_at timestamptz,
  released_at timestamptz,
  objective_score numeric(7,2),
  subjective_score numeric(7,2),
  total_score numeric(7,2),
  teacher_feedback text,
  focus_leave_count integer not null default 0,
  last_focus_left_at timestamptz,
  unique (assessment_id, student_id, attempt_number)
);

create table public.responses (
  id uuid primary key default gen_random_uuid(),
  attempt_id uuid not null references public.attempts(id) on delete cascade,
  question_id uuid not null references public.questions(id) on delete restrict,
  response jsonb,
  is_correct boolean,
  marks_awarded numeric(5,2),
  time_seconds integer,
  confidence smallint check (confidence between 1 and 5),
  teacher_feedback text,
  reviewed_by uuid references public.profiles(id) on delete set null,
  reviewed_at timestamptz,
  unique (attempt_id, question_id)
);

create table public.error_log (
  id uuid primary key default gen_random_uuid(),
  student_id uuid not null references public.profiles(id) on delete cascade,
  response_id uuid not null references public.responses(id) on delete cascade,
  error_type text not null check (error_type in ('concept', 'formula', 'calculation', 'reading', 'time_management')),
  retry_due_at timestamptz not null,
  resolved_at timestamptz,
  teacher_note text,
  created_at timestamptz not null default now()
);

create table public.private_notes (
  id uuid primary key default gen_random_uuid(),
  student_id uuid not null references public.profiles(id) on delete cascade,
  class_id uuid not null references public.classes(id) on delete cascade,
  material_id uuid references public.materials(id) on delete cascade,
  selected_text text,
  note_text text,
  highlight_type text check (highlight_type in ('formula', 'definition', 'key_concept', 'cbse_derivation', 'neet_trap', 'exam_tip')),
  is_flashcard boolean not null default false,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table public.lecture_planner (
  lecture_number integer primary key,
  planned_date date,
  day_name text,
  phase text,
  chapter_title text not null,
  topic text not null,
  delivery_focus text,
  assessment_or_homework text,
  created_at timestamptz not null default now()
);

alter table public.profiles enable row level security;
alter table public.student_links enable row level security;
alter table public.chapters enable row level security;
alter table public.classes enable row level security;
alter table public.materials enable row level security;
alter table public.material_progress enable row level security;
alter table public.attendance enable row level security;
alter table public.homework_submissions enable row level security;
alter table public.assessments enable row level security;
alter table public.questions enable row level security;
alter table public.question_keys enable row level security;
alter table public.assessment_questions enable row level security;
alter table public.attempts enable row level security;
alter table public.responses enable row level security;
alter table public.error_log enable row level security;
alter table public.private_notes enable row level security;
alter table public.lecture_planner enable row level security;

create or replace function public.current_role()
returns public.user_role language sql stable security definer
set search_path = public
as $$ select role from public.profiles where id = auth.uid() $$;

create or replace function public.is_linked_to_student(target_student uuid)
returns boolean language sql stable security definer
set search_path = public
as $$
  select exists (
    select 1 from public.student_links
    where student_id = target_student and adult_id = auth.uid()
  )
$$;

create policy "profiles own or linked read" on public.profiles for select
using (id = auth.uid() or public.is_linked_to_student(id));

create policy "links visible to participants" on public.student_links for select
using (student_id = auth.uid() or adult_id = auth.uid());
create policy "teachers manage student links" on public.student_links for all
using (public.current_role() = 'teacher')
with check (public.current_role() = 'teacher');

create policy "published curriculum readable" on public.chapters for select using (true);
create policy "teachers manage chapters" on public.chapters for all
using (public.current_role() = 'teacher') with check (public.current_role() = 'teacher');

create policy "published classes readable" on public.classes for select
using (status = 'published' or created_by = auth.uid());
create policy "teachers manage classes" on public.classes for all
using (public.current_role() = 'teacher') with check (public.current_role() = 'teacher');

create policy "published materials readable" on public.materials for select
using (
  created_by = auth.uid()
  or (
    status = 'published'
    and public.current_role() = 'teacher'
  )
  or (
    status = 'published'
    and public.current_role() = 'student'
    and exists (
      select 1 from public.classes c
      where c.id = materials.class_id
        and c.status = 'published'
    )
  )
);
create policy "teachers manage materials" on public.materials for all
using (public.current_role() = 'teacher') with check (public.current_role() = 'teacher');

create policy "students manage material progress" on public.material_progress for all
using (student_id = auth.uid()) with check (student_id = auth.uid());
create policy "linked adults read material progress" on public.material_progress for select
using (public.is_linked_to_student(student_id));

create policy "students read own attendance" on public.attendance for select
using (student_id = auth.uid());
create policy "linked adults read attendance" on public.attendance for select
using (public.is_linked_to_student(student_id));
create policy "teachers manage linked attendance" on public.attendance for all
using (public.current_role() = 'teacher' and public.is_linked_to_student(student_id))
with check (public.current_role() = 'teacher' and public.is_linked_to_student(student_id));

create policy "students read own homework status" on public.homework_submissions for select
using (student_id = auth.uid());
create policy "students submit own homework" on public.homework_submissions for insert
with check (student_id = auth.uid());
create policy "students update own homework submission" on public.homework_submissions for update
using (student_id = auth.uid());
create policy "linked adults read homework status" on public.homework_submissions for select
using (public.is_linked_to_student(student_id));
create policy "teachers manage linked homework" on public.homework_submissions for all
using (public.current_role() = 'teacher' and public.is_linked_to_student(student_id))
with check (public.current_role() = 'teacher' and public.is_linked_to_student(student_id));

create policy "published assessments readable" on public.assessments for select
using (status = 'published' or created_by = auth.uid());
create policy "teachers manage assessments" on public.assessments for all
using (public.current_role() = 'teacher') with check (public.current_role() = 'teacher');

create policy "published questions readable" on public.questions for select
using (
  public.current_role() = 'teacher'
  or (
    public.current_role() = 'student'
    and exists (
    select 1 from public.assessment_questions aq
    join public.assessments a on a.id = aq.assessment_id
    where aq.question_id = questions.id and a.status = 'published'
    )
  )
);
create policy "teachers manage questions" on public.questions for all
using (public.current_role() = 'teacher') with check (public.current_role() = 'teacher');

create policy "teachers manage answer keys" on public.question_keys for all
using (public.current_role() = 'teacher')
with check (public.current_role() = 'teacher');

create policy "assessment mapping readable" on public.assessment_questions for select using (true);
create policy "teachers manage assessment mapping" on public.assessment_questions for all
using (public.current_role() = 'teacher') with check (public.current_role() = 'teacher');

create policy "students read own attempts" on public.attempts for select
using (student_id = auth.uid());
create policy "linked adults read attempts" on public.attempts for select
using (public.is_linked_to_student(student_id));
create policy "teachers update linked attempts" on public.attempts for update
using (public.current_role() = 'teacher' and public.is_linked_to_student(student_id));

create policy "students read own responses after submission" on public.responses for select
using (
  exists (
    select 1 from public.attempts a
    where a.id = attempt_id
      and a.student_id = auth.uid()
      and a.status in ('submitted', 'released')
  )
);
create policy "teachers review linked responses" on public.responses for select
using (
  public.current_role() = 'teacher'
  and exists (
    select 1 from public.attempts a
    where a.id = attempt_id and public.is_linked_to_student(a.student_id)
  )
);
create policy "teachers mark linked responses" on public.responses for update
using (
  public.current_role() = 'teacher'
  and exists (
    select 1 from public.attempts a
    where a.id = attempt_id and public.is_linked_to_student(a.student_id)
  )
);

create policy "students manage own errors" on public.error_log for all
using (student_id = auth.uid()) with check (student_id = auth.uid());
create policy "linked teachers read errors" on public.error_log for select
using (public.current_role() = 'teacher' and public.is_linked_to_student(student_id));

-- Private note text is deliberately hidden from teacher and parent roles.
create policy "private notes owner only" on public.private_notes for all
using (student_id = auth.uid()) with check (student_id = auth.uid());
create policy "teachers manage lecture planner" on public.lecture_planner for all
using (public.current_role() = 'teacher') with check (public.current_role() = 'teacher');
create policy "authenticated read lecture planner" on public.lecture_planner for select
to authenticated using (true);

create table if not exists public.lecture_planner_progress (
  lecture_number integer primary key references public.lecture_planner(lecture_number) on delete cascade,
  completed_at timestamptz not null default now(),
  completed_by uuid not null references public.profiles(id) on delete restrict,
  note text
);
alter table public.lecture_planner_progress enable row level security;
create policy "authenticated read planner progress" on public.lecture_planner_progress for select
to authenticated using (true);
create policy "teachers manage planner progress" on public.lecture_planner_progress for all
using (public.current_role() = 'teacher') with check (public.current_role() = 'teacher');

create or replace function public.handle_new_user()
returns trigger language plpgsql security definer
set search_path = public
as $$
begin
  insert into public.profiles (id, full_name, role)
  values (
    new.id,
    coalesce(nullif(new.raw_user_meta_data ->> 'full_name', ''), split_part(new.email, '@', 1)),
    'student'
  );
  return new;
end;
$$;

create trigger on_auth_user_created
after insert on auth.users
for each row execute procedure public.handle_new_user();

create or replace function public.touch_updated_at()
returns trigger language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

create trigger profiles_updated_at before update on public.profiles
for each row execute procedure public.touch_updated_at();
create trigger classes_updated_at before update on public.classes
for each row execute procedure public.touch_updated_at();
create trigger materials_updated_at before update on public.materials
for each row execute procedure public.touch_updated_at();
create trigger assessments_updated_at before update on public.assessments
for each row execute procedure public.touch_updated_at();
create trigger private_notes_updated_at before update on public.private_notes
for each row execute procedure public.touch_updated_at();

create or replace function public.start_assessment(p_assessment_id uuid)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_assessment public.assessments%rowtype;
  v_attempt_id uuid;
  v_attempt_started timestamptz;
  v_attempt_number integer;
  v_started_at timestamptz := now();
begin
  if auth.uid() is null or public.current_role() <> 'student' then
    raise exception 'Only an authenticated student can start an assessment';
  end if;

  select * into v_assessment
  from public.assessments
  where id = p_assessment_id
    and status = 'published'
  for share;
  if not found then raise exception 'Assessment is not open'; end if;

  -- Always resume an existing in-progress attempt, even after due_at.
  select id, started_at into v_attempt_id, v_attempt_started
  from public.attempts
  where assessment_id = p_assessment_id
    and student_id = auth.uid()
    and status = 'in_progress'
  order by started_at desc
  limit 1;

  if v_attempt_id is not null then
    return jsonb_build_object(
      'attempt_id', v_attempt_id,
      'started_at', v_attempt_started,
      'duration_minutes', v_assessment.duration_minutes
    );
  end if;

  -- New attempts must be inside the published open window.
  if (v_assessment.opens_at is not null and v_assessment.opens_at > now())
     or (v_assessment.due_at is not null and v_assessment.due_at < now()) then
    raise exception 'Assessment is not open';
  end if;

  select count(*) + 1 into v_attempt_number
  from public.attempts
  where assessment_id = p_assessment_id and student_id = auth.uid();
  if v_attempt_number > v_assessment.attempt_limit then
    raise exception 'Attempt limit reached';
  end if;

  insert into public.attempts (
    assessment_id, student_id, attempt_number, status, started_at
  ) values (
    p_assessment_id, auth.uid(), v_attempt_number, 'in_progress', v_started_at
  ) returning id into v_attempt_id;

  return jsonb_build_object(
    'attempt_id', v_attempt_id,
    'started_at', v_started_at,
    'duration_minutes', v_assessment.duration_minutes
  );
end;
$$;

create or replace function public.submit_assessment(
  p_attempt_id uuid,
  p_answers jsonb
)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_assessment public.assessments%rowtype;
  v_attempt public.attempts%rowtype;
  v_score numeric(7,2) := 0;
  v_objective_count integer := 0;
  v_correct_count integer := 0;
  v_needs_review boolean := false;
  v_answer jsonb;
  v_correct boolean;
  v_row record;
  v_hard_end timestamptz;
begin
  if auth.uid() is null or public.current_role() <> 'student' then
    raise exception 'Only an authenticated student can submit an assessment';
  end if;

  select * into v_attempt
  from public.attempts
  where id = p_attempt_id
    and student_id = auth.uid()
    and status = 'in_progress'
  for update;
  if not found then raise exception 'Active attempt not found'; end if;

  select * into v_assessment
  from public.assessments
  where id = v_attempt.assessment_id
    and status = 'published'
  for share;
  if not found then raise exception 'Assessment is unavailable'; end if;

  -- Once started, finish against the attempt clock (duration + 2 min grace),
  -- not the assessment due_at. Allow a short recovery window after grace so
  -- drafts can still submit if auto-submit raced.
  v_hard_end := v_attempt.started_at
    + make_interval(mins => v_assessment.duration_minutes)
    + interval '2 minutes'
    + interval '30 seconds';
  if now() > v_hard_end then
    raise exception 'Assessment time has expired';
  end if;

  -- Clear partial rows from a previous failed submit of the same attempt.
  delete from public.error_log
  where response_id in (
    select id from public.responses where attempt_id = p_attempt_id
  );
  delete from public.responses where attempt_id = p_attempt_id;

  for v_row in
    select q.*, aq.marks as assigned_marks, qk.correct_answer
    from public.assessment_questions aq
    join public.questions q on q.id = aq.question_id
    left join public.question_keys qk on qk.question_id = q.id
    where aq.assessment_id = v_attempt.assessment_id
    order by aq.sequence
  loop
    v_answer := p_answers -> v_row.id::text;

    if v_row.question_type = 'short_answer' then
      v_correct := null;
      v_needs_review := true;
    else
      v_objective_count := v_objective_count + 1;
      v_correct := v_answer is not null and v_answer = v_row.correct_answer;
      if v_correct then
        v_correct_count := v_correct_count + 1;
        v_score := v_score + v_row.assigned_marks;
      elsif v_assessment.negative_marking > 0 and v_answer is not null then
        v_score := v_score - v_assessment.negative_marking;
      end if;
    end if;

    insert into public.responses (
      attempt_id, question_id, response, is_correct, marks_awarded
    ) values (
      p_attempt_id,
      v_row.id,
      v_answer,
      v_correct,
      case
        when v_correct then v_row.assigned_marks
        when v_correct = false and v_answer is not null then -v_assessment.negative_marking
        else null
      end
    );

    if v_correct = false then
      insert into public.error_log (
        student_id, response_id, error_type, retry_due_at
      )
      select auth.uid(), r.id, 'concept', now() + interval '48 hours'
      from public.responses r
      where r.attempt_id = p_attempt_id and r.question_id = v_row.id;
    end if;
  end loop;

  update public.attempts
  set
    status = (case when v_needs_review then 'teacher_review' else 'released' end)::public.attempt_status,
    submitted_at = now(),
    objective_score = greatest(v_score, 0),
    total_score = case when v_needs_review then null else greatest(v_score, 0) end,
    released_at = case when v_needs_review then null else now() end
  where id = p_attempt_id;

  return jsonb_build_object(
    'attempt_id', p_attempt_id,
    'status', case when v_needs_review then 'teacher_review' else 'released' end,
    'objective_score', greatest(v_score, 0),
    'correct_count', v_correct_count,
    'objective_count', v_objective_count
  );
end;
$$;

revoke all on function public.start_assessment(uuid) from public;
grant execute on function public.start_assessment(uuid) to authenticated;
revoke all on function public.submit_assessment(uuid, jsonb) from public;
grant execute on function public.submit_assessment(uuid, jsonb) to authenticated;

create or replace function public.get_student_summary(p_student_id uuid)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_attendance_total integer;
  v_attendance_present integer;
  v_homework_total integer;
  v_homework_complete integer;
  v_attempt_total integer;
  v_attempt_released integer;
  v_average_accuracy numeric;
  v_open_errors integer;
begin
  if auth.uid() <> p_student_id and not public.is_linked_to_student(p_student_id) then
    raise exception 'Not authorized to view this summary';
  end if;

  select count(*), count(*) filter (where status = 'present')
  into v_attendance_total, v_attendance_present
  from public.attendance where student_id = p_student_id;

  select count(*), count(*) filter (where status in ('submitted', 'reviewed'))
  into v_homework_total, v_homework_complete
  from public.homework_submissions where student_id = p_student_id;

  select
    count(*),
    count(*) filter (where a.status = 'released'),
    avg(case when a.status = 'released' and s.total_marks > 0
      then (a.total_score / s.total_marks) * 100 end)
  into v_attempt_total, v_attempt_released, v_average_accuracy
  from public.attempts a
  join public.assessments s on s.id = a.assessment_id
  where a.student_id = p_student_id;

  select count(*) into v_open_errors
  from public.error_log
  where student_id = p_student_id and resolved_at is null;

  return jsonb_build_object(
    'attendance_total', v_attendance_total,
    'attendance_present', v_attendance_present,
    'attendance_percent', case when v_attendance_total = 0 then 0 else round(v_attendance_present * 100.0 / v_attendance_total) end,
    'homework_total', v_homework_total,
    'homework_complete', v_homework_complete,
    'homework_percent', case when v_homework_total = 0 then 0 else round(v_homework_complete * 100.0 / v_homework_total) end,
    'attempt_total', v_attempt_total,
    'attempt_released', v_attempt_released,
    'average_accuracy', coalesce(round(v_average_accuracy), 0),
    'open_errors', v_open_errors
  );
end;
$$;

revoke all on function public.get_student_summary(uuid) from public;
grant execute on function public.get_student_summary(uuid) to authenticated;

alter function public.touch_updated_at() set search_path = public;

revoke all on function public.current_role() from public, anon;
revoke all on function public.is_linked_to_student(uuid) from public, anon;
revoke all on function public.handle_new_user() from public, anon, authenticated;
revoke all on function public.touch_updated_at() from public, anon, authenticated;
revoke all on function public.start_assessment(uuid) from public, anon;
revoke all on function public.submit_assessment(uuid, jsonb) from public, anon;
revoke all on function public.get_student_summary(uuid) from public, anon;

grant execute on function public.current_role() to authenticated;
grant execute on function public.is_linked_to_student(uuid) to authenticated;
grant execute on function public.start_assessment(uuid) to authenticated;
grant execute on function public.submit_assessment(uuid, jsonb) to authenticated;
grant execute on function public.get_student_summary(uuid) to authenticated;

insert into storage.buckets (id, name, public, file_size_limit, allowed_mime_types)
values (
  'class-materials',
  'class-materials',
  false,
  52428800,
  array[
    'application/pdf',
    'application/vnd.ms-powerpoint',
    'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    'image/png',
    'image/jpeg'
  ]
) on conflict (id) do nothing;

create policy "teachers upload class materials" on storage.objects for insert
to authenticated with check (
  bucket_id = 'class-materials' and public.current_role() = 'teacher'
);
create policy "teachers update class materials" on storage.objects for update
to authenticated using (
  bucket_id = 'class-materials' and public.current_role() = 'teacher'
);
create policy "authenticated users read class materials" on storage.objects for select
to authenticated using (
  bucket_id = 'class-materials'
  and public.current_role() in ('student', 'teacher')
);
