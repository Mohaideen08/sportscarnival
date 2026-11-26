create extension if not exists pgcrypto;

create table if not exists public.teams (
  id uuid primary key default gen_random_uuid(),
  team_name text not null check (length(trim(team_name)) between 1 and 80),
  department text,
  logo_url text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);
create unique index if not exists teams_team_name_lower_uidx on public.teams (lower(team_name));

create table if not exists public.players (
  id uuid primary key default gen_random_uuid(),
  team_id uuid not null references public.teams(id) on delete cascade,
  player_name text not null check (length(trim(player_name)) between 1 and 80),
  gender text not null check (gender in ('Male', 'Female')),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);
create index if not exists players_team_id_idx on public.players(team_id);

create table if not exists public.sports (
  id uuid primary key default gen_random_uuid(),
  sport_name text not null unique,
  created_at timestamptz not null default now()
);

create table if not exists public.tournament_categories (
  id uuid primary key default gen_random_uuid(),
  sport_id uuid not null references public.sports(id) on delete cascade,
  category_name text not null,
  competition_type text not null check (competition_type in ('singles', 'doubles', 'team')),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (sport_id, category_name)
);

create table if not exists public.tournament_participants (
  id uuid primary key default gen_random_uuid(),
  category_id uuid not null references public.tournament_categories(id) on delete cascade,
  team_id uuid not null references public.teams(id) on delete restrict,
  participant_name text not null check (length(trim(participant_name)) between 1 and 120),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (category_id, participant_name)
);
create index if not exists tournament_participants_category_idx on public.tournament_participants(category_id);

create table if not exists public.participant_players (
  participant_id uuid not null references public.tournament_participants(id) on delete cascade,
  player_id uuid not null references public.players(id) on delete restrict,
  created_at timestamptz not null default now(),
  primary key (participant_id, player_id)
);

create table if not exists public.matches (
  id uuid primary key default gen_random_uuid(),
  category_id uuid not null references public.tournament_categories(id) on delete cascade,
  round_number integer not null check (round_number > 0),
  match_number integer not null check (match_number > 0),
  participant1_id uuid references public.tournament_participants(id) on delete set null,
  participant2_id uuid references public.tournament_participants(id) on delete set null,
  winner_participant_id uuid references public.tournament_participants(id) on delete set null,
  match_status text not null default 'upcoming' check (match_status in ('upcoming', 'completed', 'bye')),
  next_match_id uuid references public.matches(id) on delete set null,
  next_match_slot smallint check (next_match_slot in (1, 2)),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (category_id, round_number, match_number)
);
create index if not exists matches_category_round_idx on public.matches(category_id, round_number, match_number);

create table if not exists public.tournament_results (
  id uuid primary key default gen_random_uuid(),
  category_id uuid not null unique references public.tournament_categories(id) on delete cascade,
  winner_team_id uuid references public.teams(id) on delete restrict,
  runner_up_team_id uuid references public.teams(id) on delete restrict,
  player_of_tournament_id uuid references public.players(id) on delete set null,
  best_batsman_id uuid references public.players(id) on delete set null,
  best_bowler_id uuid references public.players(id) on delete set null,
  updated_at timestamptz not null default now(),
  check (winner_team_id is null or runner_up_team_id is null or winner_team_id <> runner_up_team_id)
);

create table if not exists public.points_table (
  id uuid primary key default gen_random_uuid(),
  team_id uuid not null references public.teams(id) on delete cascade,
  sport_id uuid not null references public.sports(id) on delete cascade,
  category_id uuid not null references public.tournament_categories(id) on delete cascade,
  winner_points integer not null default 0 check (winner_points >= 0),
  runner_up_points integer not null default 0 check (runner_up_points >= 0),
  total_points integer generated always as (winner_points + runner_up_points) stored,
  updated_at timestamptz not null default now(),
  unique (team_id, category_id)
);
create index if not exists points_table_sport_category_idx on public.points_table(sport_id, category_id);

create table if not exists public.tournament_settings (
  id uuid primary key default gen_random_uuid(),
  tournament_name text not null,
  tournament_status text not null default 'active' check (tournament_status in ('draft', 'active', 'completed')),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create or replace function public.set_updated_at()
returns trigger language plpgsql as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

do $$
declare table_name text;
begin
  foreach table_name in array array['teams', 'players', 'tournament_categories', 'tournament_participants', 'matches', 'tournament_results', 'points_table', 'tournament_settings']
  loop
    execute format('drop trigger if exists set_updated_at on public.%I', table_name);
    execute format('create trigger set_updated_at before update on public.%I for each row execute function public.set_updated_at()', table_name);
  end loop;
end;
$$;

alter table public.teams enable row level security;
alter table public.players enable row level security;
alter table public.sports enable row level security;
alter table public.tournament_categories enable row level security;
alter table public.tournament_participants enable row level security;
alter table public.participant_players enable row level security;
alter table public.matches enable row level security;
alter table public.tournament_results enable row level security;
alter table public.points_table enable row level security;
alter table public.tournament_settings enable row level security;

do $$
declare table_name text;
begin
  foreach table_name in array array['teams', 'players', 'sports', 'tournament_categories', 'tournament_participants', 'participant_players', 'matches', 'tournament_results', 'points_table', 'tournament_settings']
  loop
    execute format('drop policy if exists "Public can view %s" on public.%I', table_name, table_name);
    execute format('create policy "Public can view %s" on public.%I for select to anon, authenticated using (true)', table_name, table_name);
    execute format('drop policy if exists "Admins can manage %s" on public.%I', table_name, table_name);
    execute format($policy$create policy "Admins can manage %s" on public.%I for all to authenticated using ((auth.jwt() -> 'app_metadata' ->> 'role') = 'admin') with check ((auth.jwt() -> 'app_metadata' ->> 'role') = 'admin')$policy$, table_name, table_name);
  end loop;
end;
$$;