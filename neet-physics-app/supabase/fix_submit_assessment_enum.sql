-- Fix: CASE returns text; attempts.status is attempt_status enum.
-- Also clear any partial responses from a failed prior submit so retry works.

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
    and (due_at is null or due_at >= now())
    and now() <= v_attempt.started_at
      + make_interval(mins => duration_minutes)
      + interval '2 minutes'
  for share;

  if not found then
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
