-- Plantel / Movimentação — o CHECK de `classe` estava num vocabulário que a tela
-- nunca usou.
--
-- A tabela nasceu com o vocabulário do protótipo de terminal
-- (tools/build_plantel_mov.py: producao, transferencia, manter, baixar,
-- corrigir_planilha). A aba do hub (assets/plantel/plantel.js, CLASSES_MOV)
-- oferece outro: compra, embriao, venda, morte, doacao, reavaliacao, renome,
-- sem_efeito. Resultado: registrar `embriao`, `reavaliacao` ou `renome` batia no
-- CHECK e o insert era recusado. A tela avisa (alert com a mensagem do banco),
-- mas a decisão já entrava no estado local ANTES do upsert: a linha aparecia
-- registrada até alguém recarregar a aba, e aí sumia. O desfazer do estado
-- local entrou junto, em registra() (assets/plantel/plantel.js).
--
-- A tela é a fonte: quem classifica é quem está fechando o mês, e é lá que ele
-- escolhe. Mexeu em CLASSES_MOV, mexe aqui.

-- ---------------------------------------------------------------------
-- 1) Renomes sem ambiguidade do vocabulário antigo. Os outros valores antigos
--    (transferencia, manter, baixar, corrigir_planilha) não têm equivalente na
--    tela e ficam como estão — linha gravada é rastro de decisão humana, não se
--    reescreve no chute.
-- ---------------------------------------------------------------------
update plantel_mov_classificacao set classe = 'embriao'     where classe = 'producao';
update plantel_mov_classificacao set classe = 'reavaliacao' where classe = 'reaval';

-- ---------------------------------------------------------------------
-- 2) O CHECK passa a ser o da tela. NOT VALID de propósito: vale pra tudo que
--    for gravado daqui pra frente e não derruba a migração se sobrou linha do
--    vocabulário antigo — mês fechado (<= 2026-07) não se reclassifica mesmo,
--    então não há como consertar aquelas linhas pela tela.
-- ---------------------------------------------------------------------
alter table plantel_mov_classificacao
  drop constraint if exists plantel_mov_classificacao_classe_check;

alter table plantel_mov_classificacao
  drop constraint if exists plantel_mov_classe_ck;

alter table plantel_mov_classificacao
  add constraint plantel_mov_classe_ck check (classe in (
      'compra', 'embriao', 'venda', 'morte', 'doacao', 'reavaliacao',
      'renome', 'sem_efeito')) not valid;

comment on constraint plantel_mov_classe_ck on plantel_mov_classificacao is
  'Mesma lista de CLASSES_MOV em assets/plantel/plantel.js — mudou lá, muda aqui.';
