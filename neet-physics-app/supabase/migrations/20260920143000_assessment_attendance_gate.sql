-- Weekly/chapter tests require student attendance request + teacher verification
-- before the paper (and start_assessment) can begin.

create table if not exists public.assessment_attendance (
  assessment_id uuid not null references public.assessments(id) on delete cascade,
  student_id uuid not null references public.profiles(id) on delete cascade,
  status text not null check (status in ('requested', 'verified', 'rejected')),
  requested_at timestamptz not null default now(),
  verified_at timestamptz,
  verified_by uuid references public.profiles(id) on delete set null,
  teacher_note text,
  primary key (assessment_id, student_id)
);

create index if not exists assessment_attendance_status_idx
  on public.assessment_attendance (status, requested_at desc);

alter table public.assessment_attendance enable row level security;

drop policy if exists "students read own assessment attendance" on public.assessment_attendance;
create policy "students read own assessment attendance" on public.assessment_attendance for select
using (student_id = auth.uid());

drop policy if exists "students request assessment attendance" on public.assessment_attendance;
create policy "students request assessment attendance" on public.assessment_attendance for insert
with check (
  student_id = auth.uid()
  and public.current_role() = 'student'
  and status = 'requested'
);

drop policy if exists "students update own requested attendance" on public.assessment_attendance;
create policy "students update own requested attendance" on public.assessment_attendance for update
using (
  student_id = auth.uid()
  and public.current_role() = 'student'
  and status in ('requested', 'rejected')
)
with check (
  student_id = auth.uid()
  and status = 'requested'
);

drop policy if exists "linked teachers read assessment attendance" on public.assessment_attendance;
create policy "linked teachers read assessment attendance" on public.assessment_attendance for select
using (public.current_role() = 'teacher' and public.is_linked_to_student(student_id));

drop policy if exists "linked teachers manage assessment attendance" on public.assessment_attendance;
create policy "linked teachers manage assessment attendance" on public.assessment_attendance for update
using (public.current_role() = 'teacher' and public.is_linked_to_student(student_id))
with check (public.current_role() = 'teacher' and public.is_linked_to_student(student_id));

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
  v_focus_leave_count integer := 0;
  v_started_at timestamptz := now();
  v_attendance_status text;
begin
  if auth.uid() is null or public.current_role() <> 'student' then
    raise exception 'Only an authenticated student can start an assessment';
  end if;

  select * into v_assessment
  from public.assessments
  where id = p_assessment_id
    and status = 'published'
  for share;
  if not found then
    raise exception 'Assessment is not open';
  end if;

  select id, started_at, focus_leave_count
  into v_attempt_id, v_attempt_started, v_focus_leave_count
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
      'duration_minutes', v_assessment.duration_minutes,
      'focus_leave_count', coalesce(v_focus_leave_count, 0)
    );
  end if;

  if (v_assessment.opens_at is not null and v_assessment.opens_at > now())
     or (v_assessment.due_at is not null and v_assessment.due_at < now()) then
    raise exception 'Assessment is not open';
  end if;

  -- Weekly and chapter tests require teacher-verified attendance first.
  if v_assessment.assessment_type <> 'daily_quiz' then
    select status into v_attendance_status
    from public.assessment_attendance
    where assessment_id = p_assessment_id
      and student_id = auth.uid();

    if v_attendance_status is null then
      raise exception 'Mark your attendance first. Your teacher must verify it before you can start.';
    end if;
    if v_attendance_status = 'requested' then
      raise exception 'Waiting for your teacher to verify attendance.';
    end if;
    if v_attendance_status = 'rejected' then
      raise exception 'Attendance was not approved. Mark attendance again or contact your teacher.';
    end if;
    if v_attendance_status <> 'verified' then
      raise exception 'Attendance is not verified yet.';
    end if;
  end if;

  select count(*) + 1
  into v_attempt_number
  from public.attempts
  where assessment_id = p_assessment_id
    and student_id = auth.uid();

  if v_assessment.assessment_type <> 'daily_quiz'
     and v_attempt_number > v_assessment.attempt_limit then
    raise exception 'Attempt limit reached';
  end if;

  insert into public.attempts (
    assessment_id, student_id, attempt_number, status, started_at, focus_leave_count
  ) values (
    p_assessment_id, auth.uid(), v_attempt_number, 'in_progress', v_started_at, 0
  )
  returning id into v_attempt_id;

  return jsonb_build_object(
    'attempt_id', v_attempt_id,
    'started_at', v_started_at,
    'duration_minutes', v_assessment.duration_minutes,
    'focus_leave_count', 0
  );
end;
$$;

revoke all on function public.start_assessment(uuid) from public, anon;
grant execute on function public.start_assessment(uuid) to authenticated;
