-- Student missing features: attempt review RPC, homework upload columns, storage policies.
-- Safe to re-run.

alter table public.homework_submissions
  add column if not exists answer_text text,
  add column if not exists storage_path text,
  add column if not exists file_name text,
  add column if not exists mime_type text;

drop policy if exists "students read own responses after submission" on public.responses;
create policy "students read own responses after submission" on public.responses for select
using (
  exists (
    select 1 from public.attempts a
    where a.id = attempt_id
      and a.student_id = auth.uid()
      and a.status in ('submitted', 'teacher_review', 'released')
  )
);

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

  v_show_keys :=
    v_attempt.status = 'released'
    and coalesce(v_assessment.explanations_after_submission, true)
    and (
      v_attempt.student_id = auth.uid()
      or public.current_role() = 'teacher'
    );

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
        'response', r.response,
        'is_correct', r.is_correct,
        'marks_awarded', r.marks_awarded,
        'teacher_feedback', r.teacher_feedback,
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
      'objective_score', v_attempt.objective_score,
      'subjective_score', v_attempt.subjective_score,
      'total_score', v_attempt.total_score,
      'teacher_feedback', v_attempt.teacher_feedback
    ),
    'assessment', jsonb_build_object(
      'id', v_assessment.id,
      'title', v_assessment.title,
      'assessment_type', v_assessment.assessment_type,
      'total_marks', v_assessment.total_marks,
      'explanations_after_submission', v_assessment.explanations_after_submission
    ),
    'show_explanations', v_show_keys,
    'items', v_items
  );
end;
$$;

revoke all on function public.get_attempt_review(uuid) from public, anon;
grant execute on function public.get_attempt_review(uuid) to authenticated;

drop policy if exists "students upload own homework files" on storage.objects;
create policy "students upload own homework files" on storage.objects for insert
to authenticated
with check (
  bucket_id = 'class-materials'
  and public.current_role() = 'student'
  and (storage.foldername(name))[1] = 'homework'
  and (storage.foldername(name))[2] = auth.uid()::text
);

drop policy if exists "students update own homework files" on storage.objects;
create policy "students update own homework files" on storage.objects for update
to authenticated
using (
  bucket_id = 'class-materials'
  and public.current_role() = 'student'
  and (storage.foldername(name))[1] = 'homework'
  and (storage.foldername(name))[2] = auth.uid()::text
)
with check (
  bucket_id = 'class-materials'
  and public.current_role() = 'student'
  and (storage.foldername(name))[1] = 'homework'
  and (storage.foldername(name))[2] = auth.uid()::text
);

drop policy if exists "students delete own homework files" on storage.objects;
create policy "students delete own homework files" on storage.objects for delete
to authenticated
using (
  bucket_id = 'class-materials'
  and public.current_role() = 'student'
  and (storage.foldername(name))[1] = 'homework'
  and (storage.foldername(name))[2] = auth.uid()::text
);

notify pgrst, 'reload schema';
