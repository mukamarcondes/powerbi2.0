-- Reconstrui somente a carteira de novembro de 2026 a partir de Pagos Mensais.
-- Outubro, agosto, setembro e as demais competencias permanecem intactos.
-- A transacao aborta se Pagos Mensais de novembro nao tiver exatamente 335 linhas.

begin;

do $$
declare
  source_count integer;
begin
  select count(*) into source_count
  from public.pagos_mensais
  where competencia = date '2026-11-01';

  if source_count <> 335 then
    raise exception 'Esperadas 335 linhas em pagos_mensais para novembro; encontradas %.', source_count;
  end if;
end $$;

delete from public.clientes_mensais
where competencia = date '2026-11-01';

insert into public.clientes_mensais (
  cliente_id,
  competencia,
  empresa,
  filial,
  sistema,
  servico,
  status,
  financeiro,
  valor_mensal,
  vencimento,
  entrada,
  observacao
)
select
  p.cliente_id,
  p.competencia,
  p.empresa,
  c.filial,
  p.sistema,
  p.servico,
  c.status,
  p.status,
  p.valor_final,
  p.vencimento,
  c.entrada,
  p.observacao
from public.pagos_mensais p
join public.clientes c on c.id = p.cliente_id
where p.competencia = date '2026-11-01';

do $$
declare
  wallet_count integer;
begin
  select count(*) into wallet_count
  from public.clientes_mensais
  where competencia = date '2026-11-01';

  if wallet_count <> 335 then
    raise exception 'A carteira reconstruida ficou com % linhas, esperadas 335.', wallet_count;
  end if;
end $$;

commit;
