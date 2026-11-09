-- Diagnostico somente leitura: identifica possiveis duplicatas em Pagos Mensais.
-- Verifica outubro e novembro de 2026, pois a carga de outubro pode ter sido
-- salva com competencia de novembro.
-- Execute primeiro no SQL Editor do Supabase e revise os grupos antes de excluir.

with candidatos as (
  select
    date_trunc('month', p.competencia)::date as competencia,
    lower(trim(p.empresa)) as empresa_chave,
    lower(trim(coalesce(p.sistema, ''))) as sistema_chave,
    lower(trim(coalesce(p.servico, ''))) as servico_chave,
    count(*) as quantidade,
    array_agg(p.id order by p.updated_at desc nulls last, p.created_at desc, p.id) as ids,
    array_agg(p.cliente_id order by p.updated_at desc nulls last, p.created_at desc, p.id) as cliente_ids,
    array_agg(p.status::text order by p.updated_at desc nulls last, p.created_at desc, p.id) as status,
    array_agg(p.valor_original order by p.updated_at desc nulls last, p.created_at desc, p.id) as valores
  from public.pagos_mensais p
  where p.competencia >= date '2026-10-01'
    and p.competencia < date '2026-12-01'
  group by 1, 2, 3, 4
  having count(*) > 1
)
select *
from candidatos
order by competencia, empresa_chave, sistema_chave, servico_chave;
