-- Sincronização por senha administrativa (Planilha Financeira Pessoal)
-- Rode UMA VEZ no SQL Editor do seu projeto Supabase.
-- Os dados ficam sob um identificador derivado da senha (SHA-256); a tabela
-- não é exposta pela API — só as duas funções abaixo (security definer) a acessam.

create table if not exists public.financas_sync (
  space text primary key,
  doc jsonb not null,
  updated_ms bigint not null default 0,
  updated_at timestamptz not null default now()
);
alter table public.financas_sync enable row level security;
-- sem policies: a tabela só é acessada pelas funções abaixo

create or replace function public.sync_pull(p_space text)
returns table(doc jsonb, updated_ms bigint)
language sql security definer set search_path = public as $$
  select doc, updated_ms from public.financas_sync where space = p_space;
$$;

create or replace function public.sync_push(p_space text, p_doc jsonb, p_ms bigint)
returns void
language sql security definer set search_path = public as $$
  insert into public.financas_sync(space, doc, updated_ms, updated_at)
  values (p_space, p_doc, p_ms, now())
  on conflict (space) do update
    set doc = excluded.doc, updated_ms = excluded.updated_ms, updated_at = now()
    where public.financas_sync.updated_ms <= excluded.updated_ms;
$$;

grant execute on function public.sync_pull(text) to anon, authenticated;
grant execute on function public.sync_push(text, jsonb, bigint) to anon, authenticated;
