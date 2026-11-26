create or replace function public.save_tournament_result(
  p_category_id uuid,
  p_winner_team_id uuid,
  p_runner_up_team_id uuid,
  p_player_of_tournament_id uuid default null,
  p_best_batsman_id uuid default null,
  p_best_bowler_id uuid default null
)
returns void
language plpgsql
security invoker
set search_path = public, pg_temp
as $$
declare
  v_sport_id uuid;
begin
  if coalesce(auth.jwt() -> 'app_metadata' ->> 'role', '') <> 'admin' then
    raise exception 'Administrator access is required to save tournament results';
  end if;
  if p_winner_team_id is null or p_runner_up_team_id is null then
    raise exception 'Winner and runner-up are required';
  end if;

  select sport_id into strict v_sport_id
  from public.tournament_categories
  where id = p_category_id;

  insert into public.tournament_results (
    category_id,
    winner_team_id,
    runner_up_team_id,
    player_of_tournament_id,
    best_batsman_id,
    best_bowler_id
  ) values (
    p_category_id,
    p_winner_team_id,
    p_runner_up_team_id,
    p_player_of_tournament_id,
    p_best_batsman_id,
    p_best_bowler_id
  )
  on conflict (category_id) do update set
    winner_team_id = excluded.winner_team_id,
    runner_up_team_id = excluded.runner_up_team_id,
    player_of_tournament_id = excluded.player_of_tournament_id,
    best_batsman_id = excluded.best_batsman_id,
    best_bowler_id = excluded.best_bowler_id,
    updated_at = now();

  delete from public.points_table where category_id = p_category_id;

  insert into public.points_table (team_id, sport_id, category_id, winner_points, runner_up_points)
  select team_id, v_sport_id, p_category_id, sum(winner_points)::integer, sum(runner_up_points)::integer
  from (values
    (p_winner_team_id, 4, 0),
    (p_runner_up_team_id, 0, 2)
  ) as awards(team_id, winner_points, runner_up_points)
  group by team_id;
end;
$$;

revoke all on function public.save_tournament_result(uuid, uuid, uuid, uuid, uuid, uuid) from public, anon;
grant execute on function public.save_tournament_result(uuid, uuid, uuid, uuid, uuid, uuid) to authenticated;

notify pgrst, 'reload schema';