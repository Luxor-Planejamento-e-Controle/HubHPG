-- Aba 'auditoria' no RBAC.
--
-- A policy do bucket usa o prefixo do nome do arquivo (auditoria.html ->
-- hub_can('auditoria')), entao o valor precisa ser aceito pelo CHECK. E o CHECK
-- de `create table if not exists` NAO altera tabela existente — por isso a
-- constraint e recriada aqui, mesmo padrao que o P&C precisou ao adicionar
-- 'vendas'. Sem isto, liberar a auditoria pra alguem volta como violacao.
alter table user_dashboard_access
  drop constraint if exists user_dashboard_access_dashboard_check;
alter table user_dashboard_access
  add constraint user_dashboard_access_dashboard_check
  check (dashboard in ('semanal','comite','auditoria','plantel'));
