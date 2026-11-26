create or replace function public.mark_match_incomplete(p_match_id uuid)
returns void
language plpgsql
security invoker
set search_path = public, pg_temp
as $$
declare
  v_category_id uuid;
  v_match_status text;
  v_winner_id uuid;
  v_next_match_id uuid;
  v_next_match_slot smallint;
  v_next_match_status text;
begin
  if coalesce(auth.jwt() -> 'app_metadata' ->> 'role', '') <> 'admin' then
    raise exception 'Administrator access is required to reopen matches';
  end if;

  select category_id, match_status, winner_participant_id, next_match_id, next_match_slot
  into v_category_id, v_match_status, v_winner_id, v_next_match_id, v_next_match_slot
  from public.matches
  where id = p_match_id
  for update;

  if not found then
    raise exception 'Match was not found';
  end if;
  if v_match_status <> 'completed' or v_winner_id is null then
    raise exception 'Only completed matches can be marked incomplete';
  end if;

  if v_next_match_id is not null then
    select match_status into v_next_match_status
    from public.matches
    where id = v_next_match_id
    for update;

    if v_next_match_status = 'completed' then
      raise exception 'This result cannot be changed after the next match is completed';
    end if;

    if v_next_match_slot = 1 then
      update public.matches set participant1_id = null
      where id = v_next_match_id and participant1_id = v_winner_id;
    else
      update public.matches set participant2_id = null
      where id = v_next_match_id and participant2_id = v_winner_id;
    end if;
  else
    delete from public.tournament_results where category_id = v_category_id;
    delete from public.points_table where category_id = v_category_id;
  end if;

  update public.matches
  set winner_participant_id = null, match_status = 'upcoming'
  where id = p_match_id;
end;
$$;

revoke all on function public.mark_match_incomplete(uuid) from public, anon;
grant execute on function public.mark_match_incomplete(uuid) to authenticated;

notify pgrst, 'reload schema';