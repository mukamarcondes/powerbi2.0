-- Banco do sistema Conectel BI Operacional
-- Rode este arquivo no SQL Editor do Supabase.
-- Estrutura pensada para guardar clientes, vendas, pagos mensais, cancelamentos/downgrades,
-- perfis de acesso e historico de alteracoes.

create extension if not exists pgcrypto;

-- Tipos principais
do $$
begin
  create type public.app_role as enum ('admin', 'visualizador');
exception when duplicate_object then null;
end $$;
do $$
begin
  create type public.client_status as enum ('ativo', 'bloqueado', 'cancelado');
exception when duplicate_object then null;
end $$;
do $$
begin
  create type public.finance_status as enum ('pago', 'pendente', 'atraso', 'receber');
exception when duplicate_object then null;
end $$;
do $$
begin
  create type public.sale_type as enum ('venda', 'aditivo');
exception when duplicate_object then null;
end $$;
do $$
begin
  create type public.cancel_type as enum ('cancelamento', 'downgrade');
exception when duplicate_object then null;
end $$;

-- Usuarios do sistema. O login fica no Supabase Auth; esta tabela guarda permissao do app.
create table if not exists public.usuarios_sistema (
  id uuid primary key default gen_random_uuid(),
  user_id uuid unique references auth.users(id) on delete cascade,
  email text not null unique,
  nome text not null,
  perfil public.app_role not null default 'visualizador',
  ativo boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

-- Cadastro unico do cliente. Mantem a identidade do cliente mesmo quando muda valor/status.
create table if not exists public.clientes (
  id uuid primary key default gen_random_uuid(),
  empresa text not null,
  documento text,
  filial text,
  sistema text,
  servico text,
  entrada date,
  vencimento_dia integer check (vencimento_dia between 1 and 31),
  status public.client_status not null default 'ativo',
  valor_mensal numeric(14,2) not null default 0,
  observacao text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  constraint clientes_identidade_unique unique nulls not distinct (empresa, filial, sistema, servico)
);

-- Foto da carteira em cada mes. Ex.: junho tinha 350, julho tinha 356.
-- Competencia deve ser sempre o primeiro dia do mes: 2026-06-01, 2026-07-01 etc.
create table if not exists public.clientes_mensais (
  id uuid primary key default gen_random_uuid(),
  cliente_id uuid not null references public.clientes(id) on delete cascade,
  competencia date not null,
  empresa text not null,
  filial text,
  sistema text,
  servico text,
  status public.client_status not null default 'ativo',
  financeiro public.finance_status not null default 'receber',
  valor_mensal numeric(14,2) not null default 0,
  vencimento date,
  entrada date,
  observacao text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (cliente_id, competencia)
);

-- Vendas novas e aditivos. Nao mistura com pagos mensais.
create table if not exists public.vendas (
  id uuid primary key default gen_random_uuid(),
  cliente_id uuid references public.clientes(id) on delete set null,
  competencia date not null,
  data_venda date,
  tipo public.sale_type not null default 'venda',
  empresa text not null,
  sistema text,
  servico text,
  consultor text,
  valor numeric(14,2) not null default 0,
  observacao text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

-- Mensalidades por mes. Aqui entra todos os clientes ativos do mes, com status de pagamento.
create table if not exists public.pagos_mensais (
  id uuid primary key default gen_random_uuid(),
  cliente_id uuid not null references public.clientes(id) on delete cascade,
  competencia date not null,
  empresa text not null,
  sistema text,
  servico text,
  valor_original numeric(14,2) not null default 0,
  valor_downgrade numeric(14,2) not null default 0,
  valor_final numeric(14,2) generated always as (greatest(valor_original - valor_downgrade, 0)) stored,
  status public.finance_status not null default 'pendente',
  vencimento date,
  pago_em date,
  observacao text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (cliente_id, competencia)
);

-- Cancelamentos e downgrades. Downgrade reduz mensalidade; cancelamento encerra meses seguintes.
create table if not exists public.cancelamentos (
  id uuid primary key default gen_random_uuid(),
  cliente_id uuid references public.clientes(id) on delete set null,
  competencia date not null,
  data_evento date,
  tipo public.cancel_type not null default 'cancelamento',
  empresa text not null,
  sistema text,
  servico text,
  consultor text,
  motivo text,
  valor numeric(14,2) not null default 0,
  observacao text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

-- Controle opcional de importacoes, util para saber quando e por quem subiu planilha.
create table if not exists public.importacoes (
  id uuid primary key default gen_random_uuid(),
  tipo text not null check (tipo in ('clientes', 'financeiro', 'vendas', 'pagos_mensais', 'cancelamentos')),
  nome_arquivo text,
  total_linhas integer not null default 0,
  total_importadas integer not null default 0,
  usuario_id uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default now()
);

-- Historico de alteracoes feitas no CRUD.
create table if not exists public.auditoria (
  id uuid primary key default gen_random_uuid(),
  tabela text not null,
  registro_id uuid,
  acao text not null check (acao in ('insert', 'update', 'delete', 'import')),
  antes jsonb,
  depois jsonb,
  usuario_id uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default now()
);

-- Indices para filtros dos dashboards
create index if not exists idx_clientes_mensais_competencia on public.clientes_mensais (competencia);
create index if not exists idx_clientes_mensais_empresa on public.clientes_mensais using gin (to_tsvector('portuguese', empresa));
create index if not exists idx_clientes_mensais_status on public.clientes_mensais (status);
create index if not exists idx_clientes_mensais_sistema on public.clientes_mensais (sistema);
create index if not exists idx_clientes_mensais_servico on public.clientes_mensais (servico);
create index if not exists idx_vendas_competencia on public.vendas (competencia);
create index if not exists idx_vendas_consultor on public.vendas (consultor);
create index if not exists idx_pagos_competencia on public.pagos_mensais (competencia);
create index if not exists idx_pagos_status on public.pagos_mensais (status);
create index if not exists idx_cancelamentos_competencia on public.cancelamentos (competencia);
create index if not exists idx_cancelamentos_tipo on public.cancelamentos (tipo);

-- updated_at automatico
drop trigger if exists set_usuarios_sistema_updated_at on public.usuarios_sistema;
drop trigger if exists set_clientes_updated_at on public.clientes;
drop trigger if exists set_clientes_mensais_updated_at on public.clientes_mensais;
drop trigger if exists set_vendas_updated_at on public.vendas;
drop trigger if exists set_pagos_mensais_updated_at on public.pagos_mensais;
drop trigger if exists set_cancelamentos_updated_at on public.cancelamentos;
create or replace function public.set_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

create trigger set_usuarios_sistema_updated_at before update on public.usuarios_sistema for each row execute function public.set_updated_at();
create trigger set_clientes_updated_at before update on public.clientes for each row execute function public.set_updated_at();
create trigger set_clientes_mensais_updated_at before update on public.clientes_mensais for each row execute function public.set_updated_at();
create trigger set_vendas_updated_at before update on public.vendas for each row execute function public.set_updated_at();
create trigger set_pagos_mensais_updated_at before update on public.pagos_mensais for each row execute function public.set_updated_at();
create trigger set_cancelamentos_updated_at before update on public.cancelamentos for each row execute function public.set_updated_at();

-- Funcao para descobrir permissao do usuario logado
create or replace function public.current_app_role()
returns public.app_role
language sql
stable
security definer
set search_path = public
as $$
  select perfil
  from public.usuarios_sistema
  where user_id = auth.uid()
    and ativo = true
  limit 1;
$$;

create or replace function public.is_admin()
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select public.current_app_role() = 'admin';
$$;

create or replace function public.can_view()
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select public.current_app_role() in ('admin', 'visualizador');
$$;

-- RLS
alter table public.usuarios_sistema enable row level security;
alter table public.clientes enable row level security;
alter table public.clientes_mensais enable row level security;
alter table public.vendas enable row level security;
alter table public.pagos_mensais enable row level security;
alter table public.cancelamentos enable row level security;
alter table public.importacoes enable row level security;
alter table public.auditoria enable row level security;

drop policy if exists "usuarios_select" on public.usuarios_sistema;
drop policy if exists "usuarios_insert_admin" on public.usuarios_sistema;
drop policy if exists "usuarios_update_admin" on public.usuarios_sistema;
drop policy if exists "usuarios_delete_admin" on public.usuarios_sistema;
drop policy if exists "clientes_select" on public.clientes;
drop policy if exists "clientes_insert_admin" on public.clientes;
drop policy if exists "clientes_update_admin" on public.clientes;
drop policy if exists "clientes_delete_admin" on public.clientes;
drop policy if exists "clientes_mensais_select" on public.clientes_mensais;
drop policy if exists "clientes_mensais_insert_admin" on public.clientes_mensais;
drop policy if exists "clientes_mensais_update_admin" on public.clientes_mensais;
drop policy if exists "clientes_mensais_delete_admin" on public.clientes_mensais;
drop policy if exists "vendas_select" on public.vendas;
drop policy if exists "vendas_insert_admin" on public.vendas;
drop policy if exists "vendas_update_admin" on public.vendas;
drop policy if exists "vendas_delete_admin" on public.vendas;
drop policy if exists "pagos_select" on public.pagos_mensais;
drop policy if exists "pagos_insert_admin" on public.pagos_mensais;
drop policy if exists "pagos_update_admin" on public.pagos_mensais;
drop policy if exists "pagos_delete_admin" on public.pagos_mensais;
drop policy if exists "cancelamentos_select" on public.cancelamentos;
drop policy if exists "cancelamentos_insert_admin" on public.cancelamentos;
drop policy if exists "cancelamentos_update_admin" on public.cancelamentos;
drop policy if exists "cancelamentos_delete_admin" on public.cancelamentos;
drop policy if exists "importacoes_select" on public.importacoes;
drop policy if exists "importacoes_insert_admin" on public.importacoes;
drop policy if exists "importacoes_update_admin" on public.importacoes;
drop policy if exists "importacoes_delete_admin" on public.importacoes;
drop policy if exists "auditoria_select" on public.auditoria;
drop policy if exists "auditoria_insert_admin" on public.auditoria;

-- Usuarios: cada um ve o proprio perfil; admin gerencia todos.
create policy "usuarios_select" on public.usuarios_sistema for select using (user_id = auth.uid() or public.is_admin());
create policy "usuarios_insert_admin" on public.usuarios_sistema for insert with check (public.is_admin());
create policy "usuarios_update_admin" on public.usuarios_sistema for update using (public.is_admin()) with check (public.is_admin());
create policy "usuarios_delete_admin" on public.usuarios_sistema for delete using (public.is_admin());

-- Visualizador le tudo operacional. Admin faz CRUD.
create policy "clientes_select" on public.clientes for select using (public.can_view());
create policy "clientes_insert_admin" on public.clientes for insert with check (public.is_admin());
create policy "clientes_update_admin" on public.clientes for update using (public.is_admin()) with check (public.is_admin());
create policy "clientes_delete_admin" on public.clientes for delete using (public.is_admin());

create policy "clientes_mensais_select" on public.clientes_mensais for select using (public.can_view());
create policy "clientes_mensais_insert_admin" on public.clientes_mensais for insert with check (public.is_admin());
create policy "clientes_mensais_update_admin" on public.clientes_mensais for update using (public.is_admin()) with check (public.is_admin());
create policy "clientes_mensais_delete_admin" on public.clientes_mensais for delete using (public.is_admin());

create policy "vendas_select" on public.vendas for select using (public.can_view());
create policy "vendas_insert_admin" on public.vendas for insert with check (public.is_admin());
create policy "vendas_update_admin" on public.vendas for update using (public.is_admin()) with check (public.is_admin());
create policy "vendas_delete_admin" on public.vendas for delete using (public.is_admin());

create policy "pagos_select" on public.pagos_mensais for select using (public.can_view());
create policy "pagos_insert_admin" on public.pagos_mensais for insert with check (public.is_admin());
create policy "pagos_update_admin" on public.pagos_mensais for update using (public.is_admin()) with check (public.is_admin());
create policy "pagos_delete_admin" on public.pagos_mensais for delete using (public.is_admin());

create policy "cancelamentos_select" on public.cancelamentos for select using (public.can_view());
create policy "cancelamentos_insert_admin" on public.cancelamentos for insert with check (public.is_admin());
create policy "cancelamentos_update_admin" on public.cancelamentos for update using (public.is_admin()) with check (public.is_admin());
create policy "cancelamentos_delete_admin" on public.cancelamentos for delete using (public.is_admin());

create policy "importacoes_select" on public.importacoes for select using (public.can_view());
create policy "importacoes_insert_admin" on public.importacoes for insert with check (public.is_admin());
create policy "importacoes_update_admin" on public.importacoes for update using (public.is_admin()) with check (public.is_admin());
create policy "importacoes_delete_admin" on public.importacoes for delete using (public.is_admin());

create policy "auditoria_select" on public.auditoria for select using (public.is_admin());
create policy "auditoria_insert_admin" on public.auditoria for insert with check (public.is_admin());

-- View pronta para dashboards principais
create or replace view public.dashboard_resumo_mensal as
select
  cm.competencia,
  count(*) filter (where cm.status = 'ativo') as clientes_ativos,
  count(*) filter (where cm.status = 'bloqueado') as clientes_bloqueados,
  coalesce(sum(cm.valor_mensal), 0) as receita_carteira,
  coalesce((select sum(v.valor) from public.vendas v where v.competencia = cm.competencia), 0) as vendas_mes,
  coalesce((select sum(c.valor) from public.cancelamentos c where c.competencia = cm.competencia and c.tipo = 'cancelamento'), 0) as cancelamentos_mes,
  coalesce((select sum(c.valor) from public.cancelamentos c where c.competencia = cm.competencia and c.tipo = 'downgrade'), 0) as downgrades_mes,
  coalesce((select sum(p.valor_final) from public.pagos_mensais p where p.competencia = cm.competencia and p.status = 'pago'), 0) as pagos_mes,
  coalesce((select sum(p.valor_final) from public.pagos_mensais p where p.competencia = cm.competencia and p.status in ('pendente', 'atraso', 'receber')), 0) as aberto_mes
from public.clientes_mensais cm
group by cm.competencia;

-- Modelo para inserir o primeiro admin depois de criar o usuario no Supabase Auth:
-- insert into public.usuarios_sistema (user_id, email, nome, perfil, ativo)
-- values ('UUID_DO_AUTH_USER', 'admin@email.com', 'Nome Admin', 'admin', true)
-- on conflict (user_id) do update set
--   email = excluded.email,
--   nome = excluded.nome,
--   perfil = excluded.perfil,
--   ativo = excluded.ativo,
--   updated_at = now();


