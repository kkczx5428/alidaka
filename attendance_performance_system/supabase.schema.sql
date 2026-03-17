create extension if not exists pgcrypto;

create or replace function public.set_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = timezone('utc', now());
  return new;
end;
$$;

create table if not exists public.app_states (
  user_id uuid primary key references auth.users(id) on delete cascade,
  payload jsonb not null default '{"teams":[],"employees":[],"attendance":{},"performance":{}}'::jsonb,
  updated_at timestamptz not null default timezone('utc', now())
);

alter table public.app_states enable row level security;

drop policy if exists "Users can select own app state" on public.app_states;
create policy "Users can select own app state"
on public.app_states
for select
using (auth.uid() = user_id);

drop policy if exists "Users can insert own app state" on public.app_states;
create policy "Users can insert own app state"
on public.app_states
for insert
with check (auth.uid() = user_id);

drop policy if exists "Users can update own app state" on public.app_states;
create policy "Users can update own app state"
on public.app_states
for update
using (auth.uid() = user_id)
with check (auth.uid() = user_id);

drop trigger if exists set_app_states_updated_at on public.app_states;
create trigger set_app_states_updated_at
before update on public.app_states
for each row
execute function public.set_updated_at();

do $$
begin
  if not exists (
    select 1
    from pg_publication_tables
    where pubname = 'supabase_realtime'
      and schemaname = 'public'
      and tablename = 'app_states'
  ) then
    alter publication supabase_realtime add table public.app_states;
  end if;
end;
$$;
