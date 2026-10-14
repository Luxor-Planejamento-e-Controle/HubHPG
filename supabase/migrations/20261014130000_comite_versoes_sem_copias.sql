-- Limpeza única das cópias automáticas do histórico do comitê.
--
-- Até 14/10/2026 todo pedido de atualização cravava TODOS os meses com conteúdo,
-- e o rastro gravava versão sem autor. Sobraram versões com o mesmo conteúdo da
-- anterior: em agosto, v7–v10 iguais à v6 (e a v10 virou a cravada, que não se
-- remove pela tela). A 20261014120000 parou de criar cópia; esta apaga as que
-- ficaram. Sai só versão automática (sem autor ou "atualização automática") igual
-- à imediatamente anterior do mesmo mês — versão feita por pessoa fica, mesmo
-- repetida. Em 14/10: jun v3–v5, jul v4–v6, ago v2, v3, v7–v10.

with ordenadas as (
  select id, conteudo, criado_por,
         lag(conteudo) over (partition by mes order by versao) as anterior
    from comite_conteudo_versao
)
delete from comite_conteudo_versao x
 using ordenadas o
 where x.id = o.id
   and o.anterior = o.conteudo
   and coalesce(o.criado_por, 'atualização automática') = 'atualização automática';

-- a cravada passa pra última que sobrou em cada mês
update comite_conteudo_versao x
   set cravada = (x.versao = m.ult)
  from (select mes, max(versao) ult from comite_conteudo_versao group by mes) m
 where x.mes = m.mes and x.cravada is distinct from (x.versao = m.ult);
