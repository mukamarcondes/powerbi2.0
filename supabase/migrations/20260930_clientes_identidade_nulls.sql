-- Corrige clientes duplicáveis quando filial, sistema ou servico são NULL.
-- Execute este arquivo no SQL Editor do Supabase, uma vez por projeto existente.
-- A migração não apaga nem escolhe registros: se houver duplicados atuais, aborta
-- e informa a consulta de diagnóstico para revisão antes de aplicar a constraint.

begin;

do $$
declare
  duplicate_groups bigint;
begin
  select count(*)
    into duplicate_groups
  from (
    select 1
    from public.clientes
    group by empresa, filial, sistema, servico
    having count(*) > 1
  ) duplicates;

  if duplicate_groups > 0 then
    raise exception 'Migração cancelada: existem % grupos de clientes duplicados.', duplicate_groups
      using hint = 'Revise os registros com: SELECT empresa, filial, sistema, servico, count(*) AS quantidade, array_agg(id) AS ids FROM public.clientes GROUP BY empresa, filial, sistema, servico HAVING count(*) > 1;';
  end if;
end $$;

alter table public.clientes
  drop constraint if exists clientes_empresa_filial_sistema_servico_key;

do $$
begin
  if not exists (
    select 1
    from pg_constraint
    where conrelid = 'public.clientes'::regclass
      and conname = 'clientes_identidade_unique'
  ) then
    alter table public.clientes
      add constraint clientes_identidade_unique
      unique nulls not distinct (empresa, filial, sistema, servico);
  end if;
end $$;

commit;