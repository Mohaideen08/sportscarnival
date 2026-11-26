insert into public.sports (sport_name) values
  ('Badminton'),
  ('Table Tennis'),
  ('Carrom'),
  ('Chess'),
  ('Volleyball'),
  ('Football'),
  ('Cricket')
on conflict (sport_name) do nothing;

insert into public.tournament_categories (sport_id, category_name, competition_type)
select sports.id, categories.category_name, categories.competition_type
from (values
  ('Badminton', 'Men''s Doubles', 'doubles'),
  ('Badminton', 'Women''s Doubles', 'doubles'),
  ('Table Tennis', 'Men''s Doubles', 'doubles'),
  ('Table Tennis', 'Women''s Doubles', 'doubles'),
  ('Table Tennis', 'Men''s Singles', 'singles'),
  ('Table Tennis', 'Women''s Singles', 'singles'),
  ('Carrom', 'Men''s Doubles', 'doubles'),
  ('Carrom', 'Women''s Doubles', 'doubles'),
  ('Chess', 'Men''s', 'singles'),
  ('Chess', 'Women''s', 'singles'),
  ('Volleyball', 'Open', 'team'),
  ('Football', 'Open', 'team'),
  ('Cricket', 'Open', 'team')
) as categories(sport_name, category_name, competition_type)
join public.sports on sports.sport_name = categories.sport_name
on conflict (sport_id, category_name) do nothing;