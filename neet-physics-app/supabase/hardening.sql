-- Permission hardening for an already-created Orbit Physics database.
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
revoke all on function public.upsert_private_note(uuid, uuid, text, text, text, boolean, uuid) from public, anon;
revoke all on function public.delete_private_note(uuid) from public, anon;
grant execute on function public.upsert_private_note(uuid, uuid, text, text, text, boolean, uuid) to authenticated;
grant execute on function public.delete_private_note(uuid) to authenticated;

do $$
begin
  if to_regprocedure('public.rls_auto_enable()') is not null then
    execute 'revoke execute on function public.rls_auto_enable() from public, anon, authenticated';
  end if;
end;
$$;
