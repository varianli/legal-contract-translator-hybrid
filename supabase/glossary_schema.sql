-- Supabase schema for project-level glossary sync.
-- Run this in the Supabase SQL editor.
-- Desktop clients should use the anon key plus a project access token.
-- Do NOT put the service_role key into the desktop app.

create extension if not exists pgcrypto;

create table if not exists public.projects (
  id uuid primary key default gen_random_uuid(),
  slug text not null unique,
  name text not null,
  access_token_hash text not null,
  status text not null default 'active' check (status in ('active', 'archived')),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.glossary_terms (
  id uuid primary key default gen_random_uuid(),
  project_id uuid not null references public.projects(id) on delete cascade,
  source_text text not null,
  target_text text not null,
  term_type text not null default 'legal_term' check (
    term_type in ('company', 'fund', 'person', 'legal_term', 'defined_term', 'document_term', 'other')
  ),
  strategy text not null default 'prompt_constraint' check (
    strategy in ('placeholder_lock', 'prompt_constraint', 'postcheck')
  ),
  priority integer not null default 100,
  enabled boolean not null default true,
  notes text not null default '',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (project_id, source_text)
);

create index if not exists idx_projects_slug on public.projects (slug);
create index if not exists idx_glossary_terms_project_enabled_priority
  on public.glossary_terms (project_id, enabled, priority desc);
create index if not exists idx_glossary_terms_source_length
  on public.glossary_terms (project_id, length(source_text) desc);

create or replace function public.set_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists trg_projects_updated_at on public.projects;
create trigger trg_projects_updated_at
before update on public.projects
for each row execute function public.set_updated_at();

drop trigger if exists trg_glossary_terms_updated_at on public.glossary_terms;
create trigger trg_glossary_terms_updated_at
before update on public.glossary_terms
for each row execute function public.set_updated_at();

alter table public.projects enable row level security;
alter table public.glossary_terms enable row level security;

revoke all on public.projects from anon, authenticated;
revoke all on public.glossary_terms from anon, authenticated;

create or replace function public.hash_project_access_token(raw_token text)
returns text
language sql
immutable
as $$
  select encode(digest(coalesce(raw_token, ''), 'sha256'), 'hex');
$$;

create or replace function public.get_project_glossary(
  p_project_slug text,
  p_access_token text
)
returns table (
  source_text text,
  target_text text,
  term_type text,
  strategy text,
  priority integer,
  notes text,
  updated_at timestamptz
)
language sql
security definer
set search_path = public
as $$
  select
    gt.source_text,
    gt.target_text,
    gt.term_type,
    gt.strategy,
    gt.priority,
    gt.notes,
    gt.updated_at
  from public.projects p
  join public.glossary_terms gt on gt.project_id = p.id
  where p.slug = p_project_slug
    and p.status = 'active'
    and p.access_token_hash = public.hash_project_access_token(p_access_token)
    and gt.enabled = true
  order by gt.priority desc, length(gt.source_text) desc, gt.source_text asc;
$$;

revoke all on function public.get_project_glossary(text, text) from public;
grant execute on function public.get_project_glossary(text, text) to anon, authenticated;

-- Example setup:
-- 1. Choose a private project token, e.g. 'step-star-preipo-token'.
-- 2. Store only its hash in the database:
--
-- insert into public.projects (slug, name, access_token_hash)
-- values (
--   'step-star-preipo',
--   'Step Star Pre-IPO',
--   public.hash_project_access_token('step-star-preipo-token')
-- );
--
-- insert into public.glossary_terms (project_id, source_text, target_text, term_type, strategy, priority)
-- select id, '上海阶跃星辰智能科技股份有限公司', 'Shanghai Step Star Intelligent Technology Co., Ltd.', 'company', 'placeholder_lock', 1000
-- from public.projects where slug = 'step-star-preipo';
--
-- insert into public.glossary_terms (project_id, source_text, target_text, term_type, strategy, priority)
-- select id, '交割', 'Closing', 'defined_term', 'prompt_constraint', 900
-- from public.projects where slug = 'step-star-preipo';
