-- Delayed result release: daily quizzes unlock after due_at (else 7pm IST next day);
-- weekly/chapter tests wait for teacher release.

create or replace function public.assessment_results_unlock_at(
  p_assessment public.assessments,
  p_attempt public.attempts
)
returns timestamptz
language sql
stable
as $$
  select case
    when p_assessment.assessment_type = 'daily_quiz' then
      coalesce(
        p_assessment.due_at,
        (
          (timezone('Asia/Kolkata', coalesce(p_attempt.submitted_at, p_attempt.started_at)))::date
          + 1
          + time '19:00'
        ) at time zone 'Asia/Kolkata'
      )
    else null
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
  v_status public.attempt_status;
  v_unlock_at timestamptz;
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

  v_hard_end := v_attempt.started_at
    + make_interval(mins => v_assessment.duration_minutes)
    + interval '2 minutes'
    + interval '30 seconds';
  if now() > v_hard_end then
    raise exception 'Assessment time has expired';
  end if;

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
  end loop;

  -- Daily quiz: hold results until deadline (or 7pm IST next day).
  -- Weekly/chapter: hold until teacher releases.
  if v_assessment.assessment_type = 'daily_quiz' then
    v_status := case when v_needs_review then 'teacher_review'::public.attempt_status
                     else 'submitted'::public.attempt_status end;
  else
    v_status := 'teacher_review'::public.attempt_status;
  end if;

  update public.attempts
  set
    status = v_status,
    submitted_at = now(),
    objective_score = greatest(v_score, 0),
    total_score = null,
    released_at = null
  where id = p_attempt_id;

  select * into v_attempt from public.attempts where id = p_attempt_id;
  v_unlock_at := public.assessment_results_unlock_at(v_assessment, v_attempt);

  return jsonb_build_object(
    'attempt_id', p_attempt_id,
    'status', v_status,
    'results_visible', false,
    'unlock_at', v_unlock_at,
    'assessment_type', v_assessment.assessment_type,
    'objective_score', null,
    'correct_count', null,
    'objective_count', v_objective_count
  );
end;
$$;

create or replace function public.get_attempt_review(p_attempt_id uuid)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_attempt public.attempts%rowtype;
  v_assessment public.assessments%rowtype;
  v_show_keys boolean := false;
  v_items jsonb := '[]'::jsonb;
  v_unlock_at timestamptz;
  v_results_ready boolean := false;
begin
  if auth.uid() is null then
    raise exception 'Not authenticated';
  end if;

  select * into v_attempt
  from public.attempts
  where id = p_attempt_id;

  if not found then
    raise exception 'Attempt not found';
  end if;

  if v_attempt.student_id <> auth.uid()
     and not public.is_linked_to_student(v_attempt.student_id) then
    raise exception 'Not allowed to view this attempt';
  end if;

  if v_attempt.status = 'in_progress' then
    raise exception 'This attempt has not been submitted yet';
  end if;

  select * into v_assessment
  from public.assessments
  where id = v_attempt.assessment_id;

  v_unlock_at := public.assessment_results_unlock_at(v_assessment, v_attempt);

  -- Auto-release daily quizzes once the unlock time has passed (and no pending teacher review).
  if v_assessment.assessment_type = 'daily_quiz'
     and v_attempt.status = 'submitted'
     and v_unlock_at is not null
     and now() >= v_unlock_at then
    update public.attempts
    set
      status = 'released',
      released_at = coalesce(released_at, now()),
      total_score = coalesce(total_score, objective_score)
    where id = v_attempt.id
    returning * into v_attempt;

    insert into public.error_log (student_id, response_id, error_type, retry_due_at)
    select v_attempt.student_id, r.id, 'concept', now() + interval '48 hours'
    from public.responses r
    where r.attempt_id = v_attempt.id
      and r.is_correct = false
      and not exists (
        select 1 from public.error_log e where e.response_id = r.id
      );
  end if;

  v_results_ready :=
    v_attempt.status = 'released'
    or (
      v_assessment.assessment_type = 'daily_quiz'
      and v_attempt.status in ('submitted', 'released')
      and v_unlock_at is not null
      and now() >= v_unlock_at
    );

  v_show_keys :=
    v_results_ready
    and coalesce(v_assessment.explanations_after_submission, true)
    and (
      v_attempt.student_id = auth.uid()
      or public.current_role() = 'teacher'
    );

  -- Teachers can always see keys while reviewing.
  if public.current_role() = 'teacher' and public.is_linked_to_student(v_attempt.student_id) then
    v_show_keys := true;
    v_results_ready := true;
  end if;

  select coalesce(jsonb_agg(item order by seq), '[]'::jsonb)
  into v_items
  from (
    select
      aq.sequence as seq,
      jsonb_build_object(
        'question_id', q.id,
        'prompt', q.prompt,
        'topic', q.topic,
        'question_type', q.question_type,
        'options', q.options,
        'marks', aq.marks,
        'response', case when v_results_ready or v_attempt.student_id = auth.uid() then r.response else null end,
        'is_correct', case when v_show_keys then r.is_correct else null end,
        'marks_awarded', case when v_show_keys then r.marks_awarded else null end,
        'teacher_feedback', case when v_show_keys then r.teacher_feedback else null end,
        'correct_answer', case when v_show_keys then qk.correct_answer else null end,
        'explanation', case when v_show_keys then qk.explanation else null end
      ) as item
    from public.assessment_questions aq
    join public.questions q on q.id = aq.question_id
    left join public.responses r
      on r.attempt_id = v_attempt.id and r.question_id = q.id
    left join public.question_keys qk on qk.question_id = q.id
    where aq.assessment_id = v_attempt.assessment_id
  ) ranked;

  return jsonb_build_object(
    'attempt', jsonb_build_object(
      'id', v_attempt.id,
      'status', v_attempt.status,
      'started_at', v_attempt.started_at,
      'submitted_at', v_attempt.submitted_at,
      'released_at', v_attempt.released_at,
      'objective_score', case when v_results_ready then v_attempt.objective_score else null end,
      'subjective_score', case when v_results_ready then v_attempt.subjective_score else null end,
      'total_score', case when v_results_ready then v_attempt.total_score else null end,
      'teacher_feedback', case when v_results_ready then v_attempt.teacher_feedback else null end
    ),
    'assessment', jsonb_build_object(
      'id', v_assessment.id,
      'title', v_assessment.title,
      'assessment_type', v_assessment.assessment_type,
      'total_marks', v_assessment.total_marks,
      'due_at', v_assessment.due_at,
      'explanations_after_submission', v_assessment.explanations_after_submission
    ),
    'show_explanations', v_show_keys,
    'results_visible', v_results_ready,
    'unlock_at', v_unlock_at,
    'items', v_items
  );
end;
$$;

revoke all on function public.assessment_results_unlock_at(public.assessments, public.attempts) from public, anon;
grant execute on function public.submit_assessment(uuid, jsonb) to authenticated;
grant execute on function public.get_attempt_review(uuid) to authenticated;
