-- Track how many times a student leaves the assessment tab/app during an attempt.

alter table public.attempts
  add column if not exists focus_leave_count integer not null default 0;

alter table public.attempts
  add column if not exists last_focus_left_at timestamptz;

comment on column public.attempts.focus_leave_count is
  'How many times the student left the assessment tab/window/app during this attempt.';

create or replace function public.record_attempt_focus_leave(p_attempt_id uuid)
returns integer
language plpgsql
security definer
set search_path = public
as $$
declare
  v_count integer;
begin
  if auth.uid() is null or public.current_role() <> 'student' then
    raise exception 'Only the student can record a focus leave';
  end if;

  update public.attempts
  set
    focus_leave_count = focus_leave_count + 1,
    last_focus_left_at = now()
  where id = p_attempt_id
    and student_id = auth.uid()
    and status = 'in_progress'
  returning focus_leave_count into v_count;

  if not found then
    raise exception 'Active attempt not found';
  end if;

  return v_count;
end;
$$;

revoke all on function public.record_attempt_focus_leave(uuid) from public, anon;
grant execute on function public.record_attempt_focus_leave(uuid) to authenticated;

-- Return existing leave count when resuming an attempt.
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
