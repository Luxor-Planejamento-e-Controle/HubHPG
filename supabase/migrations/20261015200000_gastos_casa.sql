-- Aba Gastos Casa — o painel de Gastos da Casa da Fazenda Pao Grande (orçado × realizado
-- por natureza), irmão do painel de Gastos do Haras.
--
-- Quem publica é a rotina `publicar_gastos_casa` do Function App (repo Automacoes,
-- pasta "Controle de Gastos da Casa"): sobe `gastos_casa.html` no bucket, toda segunda.
-- A policy de leitura do bucket é hub_can(split_part(name,'.',1)), então quem baixa o
-- arquivo é quem tem hub_can('gastos_casa') — admin, ou uma linha em
-- user_dashboard_access com esse id. E-mail real NÃO entra aqui.
--
-- A check só vale na criação da tabela: aba nova recria a lista inteira (mesma receita
-- da migration 20261005150000_sgpg_exposicoes).
alter table user_dashboard_access drop constraint if exists user_dashboard_access_dashboard_check;
alter table user_dashboard_access add constraint user_dashboard_access_dashboard_check
  check (dashboard in ('semanal','comite','auditoria','plantel','gastos','sgpg','gastos_casa'));
