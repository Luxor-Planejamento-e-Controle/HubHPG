-- Botão "Atualizar dados" nas abas Gastos (Haras) e Gastos Casa.
--
-- Os pedidos entram na MESMA fila do semanal (hub_job): o tipo é o id do painel, então a
-- policy que já existe (hub_can(tipo)) resolve quem vê e quem pede, e o índice "um pedido
-- ativo por tipo" impede clique duplo. O trigger hub_job_avisa_azure já avisa a rota para
-- qualquer tipo; a rota hpg_hub_job (repo de deploy) manda os de Gastos para a fila do
-- painel, e quem executa é a rotina do repo Automacoes (publicar_gastos_haras /
-- publicar_gastos_casa), que grava ok/erro + log no pedido como o executor daqui.
--
-- A check nasceu na criação da tabela (20260923120000): aba nova recria a lista inteira.
alter table hub_job drop constraint if exists hub_job_tipo_check;
alter table hub_job add constraint hub_job_tipo_check
  check (tipo in ('semanal', 'comite', 'gastos', 'gastos_casa'));
