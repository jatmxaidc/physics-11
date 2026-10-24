-- Teacher release console + safe error_log writes (security definer).
-- Daily quizzes: mark as submitted until unlock_at, then release.
-- Weekly/chapter: release immediately once short answers (if any) are marked.

create or replace function public.release_attempt_results(p_attempt_id uuid)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_attempt public.attempts%rowtype;
  v_assessment public.assessments%rowtype;
  v_unlock_at timestamptz;
  v_pending integer;
  v_subjective numeric(7,2) := 0;
  v_total numeric(7,2);
  v_new_status public.attempt_status;
begin
  if auth.uid() is null or public.current_role() <> 'teacher' then
    raise exception 'Only a teacher can release results';
  end if;

  select * into v_attempt
  from public.attempts
  where id = p_attempt_id
  for update;
  if not found then
    raise exception 'Attempt not found';
  end if;

  if not public.is_linked_to_student(v_attempt.student_id) then
    raise exception 'You can only release results for linked students';
  end if;

  if v_attempt.status not in ('teacher_review', 'submitted') then
    raise exception 'This attempt is not waiting for release';
  end if;

  select * into v_assessment
  from public.assessments
  where id = v_attempt.assessment_id;
  if not found then
    raise exception 'Assessment not found';
  end if;

  select count(*) into v_pending
  from public.responses r
  join public.questions q on q.id = r.question_id
  where r.attempt_id = p_attempt_id
    and q.question_type = 'short_answer'
    and r.marks_awarded is null;
  if v_pending > 0 then
    raise exception 'Mark all short answers before releasing results';
  end if;

  select coalesce(sum(r.marks_awarded), 0) into v_subjective
  from public.responses r
  join public.questions q on q.id = r.question_id
  where r.attempt_id = p_attempt_id
    and q.question_type = 'short_answer';

  v_total := coalesce(v_attempt.objective_score, 0) + v_subjective;
  v_unlock_at := public.assessment_results_unlock_at(v_assessment, v_attempt);

  if v_assessment.assessment_type = 'daily_quiz'
     and v_unlock_at is not null
     and now() < v_unlock_at then
    v_new_status := 'submitted'::public.attempt_status;
    update public.attempts
    set
      status = v_new_status,
      subjective_score = v_subjective,
      total_score = null,
      released_at = null
    where id = p_attempt_id
    returning * into v_attempt;

    return jsonb_build_object(
      'attempt_id', p_attempt_id,
      'status', v_new_status,
      'results_visible', false,
      'unlock_at', v_unlock_at,
      'message', 'Marked complete. Results unlock automatically after the daily deadline.'
    );
  end if;

  v_new_status := 'released'::public.attempt_status;
  update public.attempts
  set
    status = v_new_status,
    subjective_score = v_subjective,
    total_score = v_total,
    released_at = coalesce(released_at, now())
  where id = p_attempt_id
  returning * into v_attempt;

  insert into public.error_log (student_id, response_id, error_type, retry_due_at)
  select v_attempt.student_id, r.id, 'concept', now() + interval '48 hours'
  from public.responses r
  where r.attempt_id = v_attempt.id
    and r.is_correct = false
    and not exists (
      select 1 from public.error_log e where e.response_id = r.id
    );

  return jsonb_build_object(
    'attempt_id', p_attempt_id,
    'status', v_new_status,
    'results_visible', true,
    'unlock_at', v_unlock_at,
    'total_score', v_total,
    'message', 'Results released to the student.'
  );
end;
$$;

revoke all on function public.release_attempt_results(uuid) from public, anon;
grant execute on function public.release_attempt_results(uuid) to authenticated;

-- Students may only read published materials on published classes.
drop policy if exists "published materials readable" on public.materials;
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
