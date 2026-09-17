-- Plantel / Movimentação — poder REINICIAR a conciliação de um mês aberto.
--
-- A migration original decidiu o contrário: "Apagar não é operação de tela
-- (sumiria o rastro de quem decidiu): corrigir = gravar outra classe por cima".
-- Isso cobre corrigir UMA linha, e não cobre o caso real de quem está rodando o
-- fechamento pela primeira vez: classificar meia dúzia testando, perceber que o
-- critério estava errado, e querer recomeçar o mês do zero. Sem apagar, a única
-- saída era reclassificar item por item para um valor que também não vale.
--
-- Fica restrito ao mesmo cerco de sempre: só quem tem a aba, e só em mês ABERTO.
-- Mês fechado continua imexível — lá o rastro é o que importa, e o argumento
-- original vale inteiro.

drop policy if exists pmc_delete on plantel_mov_classificacao;
create policy pmc_delete on plantel_mov_classificacao
  for delete to authenticated
  using ( public.hub_can('plantel') and not public.plantel_mes_fechado(mes) );
