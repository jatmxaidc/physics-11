-- Additive backend completion for an already-provisioned Orbit Physics project.
-- Safe to re-run.

create table if not exists public.lecture_planner (
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

alter table public.lecture_planner enable row level security;

drop policy if exists "teachers manage lecture planner" on public.lecture_planner;
create policy "teachers manage lecture planner" on public.lecture_planner for all
using (public.current_role() = 'teacher')
with check (public.current_role() = 'teacher');

drop policy if exists "authenticated read lecture planner" on public.lecture_planner;
create policy "authenticated read lecture planner" on public.lecture_planner for select
to authenticated
using (true);

create or replace function public.is_assessment_open(p_assessment public.assessments)
returns boolean
language sql
immutable
as $$
  select
    p_assessment.status = 'published'
    and (p_assessment.opens_at is null or p_assessment.opens_at <= now())
    and (p_assessment.due_at is null or p_assessment.due_at >= now());
$$;

revoke all on function public.is_assessment_open(public.assessments) from public, anon;
grant execute on function public.is_assessment_open(public.assessments) to authenticated;

create or replace function public.upsert_private_note(
  p_class_id uuid,
  p_material_id uuid,
  p_selected_text text,
  p_note_text text,
  p_highlight_type text,
  p_is_flashcard boolean,
  p_note_id uuid default null
)
returns uuid
language plpgsql
security definer
set search_path = public
as $$
declare
  v_id uuid;
begin
  if auth.uid() is null or public.current_role() <> 'student' then
    raise exception 'Only an authenticated student can save private notes';
  end if;
  if p_note_text is null or length(btrim(p_note_text)) < 2 then
    raise exception 'Note text is required';
  end if;
  if p_highlight_type is not null and p_highlight_type not in (
    'formula', 'definition', 'key_concept', 'cbse_derivation', 'neet_trap', 'exam_tip'
  ) then
    raise exception 'Invalid highlight type';
  end if;

  if p_note_id is not null then
    update public.private_notes
    set
      class_id = p_class_id,
      material_id = p_material_id,
      selected_text = nullif(btrim(coalesce(p_selected_text, '')), ''),
      note_text = btrim(p_note_text),
      highlight_type = p_highlight_type,
      is_flashcard = coalesce(p_is_flashcard, false)
    where id = p_note_id and student_id = auth.uid()
    returning id into v_id;
    if v_id is null then
      raise exception 'Note not found';
    end if;
    return v_id;
  end if;

  insert into public.private_notes (
    student_id, class_id, material_id, selected_text, note_text, highlight_type, is_flashcard
  ) values (
    auth.uid(),
    p_class_id,
    p_material_id,
    nullif(btrim(coalesce(p_selected_text, '')), ''),
    btrim(p_note_text),
    p_highlight_type,
    coalesce(p_is_flashcard, false)
  ) returning id into v_id;
  return v_id;
end;
$$;

create or replace function public.delete_private_note(p_note_id uuid)
returns void
language plpgsql
security definer
set search_path = public
as $$
begin
  if auth.uid() is null or public.current_role() <> 'student' then
    raise exception 'Only an authenticated student can delete private notes';
  end if;
  delete from public.private_notes where id = p_note_id and student_id = auth.uid();
end;
$$;

revoke all on function public.upsert_private_note(uuid, uuid, text, text, text, boolean, uuid) from public, anon;
revoke all on function public.delete_private_note(uuid) from public, anon;
grant execute on function public.upsert_private_note(uuid, uuid, text, text, text, boolean, uuid) to authenticated;
grant execute on function public.delete_private_note(uuid) to authenticated;
