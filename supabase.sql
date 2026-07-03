-- Execute este código no SQL Editor do seu projeto Supabase.
-- Cria a tabela de sincronização do painel financeiro com RLS
-- (cada usuário só lê e grava a própria linha).

create table if not exists public.financas (
  user_id uuid primary key references auth.users(id) on delete cascade,
  doc jsonb not null,
  updated_ms bigint not null default 0,
  updated_at timestamptz not null default now()
);

alter table public.financas enable row level security;

create policy "sel_own" on public.financas for select using (auth.uid() = user_id);
create policy "ins_own" on public.financas for insert with check (auth.uid() = user_id);
create policy "upd_own" on public.financas for update using (auth.uid() = user_id);
