-- Plantel / Movimentação: remover movimentação que não condiz com a realidade.
--
-- O diff entre as planilhas do haras às vezes cria movimentação que não aconteceu no
-- mês: o LEGADO DA PAO GRANDE aparece trocando de local em ago/2026 porque a cópia de
-- agosto (AGO_26_v3) já vinha editada com a saída de 05/09. Quem fecha o mês remove na
-- ficha (assets/plantel/plantel.js, CLASSE_REMOVIDA): sai da fila, não entra no resumo
-- e fica listada na Conciliação com o motivo (coluna `nota`), para restaurar.
--
-- É gravada como mais uma linha de plantel_mov_classificacao, com classe 'removida':
-- mesmas policies (escrita só em mês aberto) e restaurar é apagar a linha (pmc_delete).

alter table plantel_mov_classificacao
  drop constraint if exists plantel_mov_classe_ck;

alter table plantel_mov_classificacao
  add constraint plantel_mov_classe_ck check (classe in (
      'compra', 'embriao', 'venda', 'morte', 'doacao', 'reavaliacao',
      'renome', 'sem_efeito', 'removida')) not valid;

comment on constraint plantel_mov_classe_ck on plantel_mov_classificacao is
  'CLASSES_MOV + CLASSE_REMOVIDA em assets/plantel/plantel.js — mudou lá, muda aqui.';
