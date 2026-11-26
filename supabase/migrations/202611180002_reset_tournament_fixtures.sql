create or replace function public.reset_tournament_fixtures(p_category_id uuid)
returns void
language plpgsql
security invoker
set search_path = public, pg_temp
as $$
begin
  if coalesce(auth.jwt() -> 'app_metadata' ->> 'role', '') <> 'admin' then
    raise exception 'Administrator access is required to reset fixtures';
  end if;

  if not exists (select 1 from public.tournament_categories where id = p_category_id) then
    raise exception 'Tournament category was not found';
  end if;

  delete from public.tournament_results where category_id = p_category_id;
  delete from public.points_table where category_id = p_category_id;
  delete from public.matches where category_id = p_category_id;
end;
$$;

revoke all on function public.reset_tournament_fixtures(uuid) from public, anon;
grant execute on function public.reset_tournament_fixtures(uuid) to authenticated;

notify pgrst, 'reload schema';