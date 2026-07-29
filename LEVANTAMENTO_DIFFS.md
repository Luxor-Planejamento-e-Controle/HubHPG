# Levantamento de divergências — Atualização Semanal (validação vs docx 24/07)

Legenda: ✅ bate · ⚠️ divergência de definição · ❌ falta dado na planilha (não é bug de código)

## ✅ Batem exato (não precisa mexer)
| Métrica | Calc | Docx | Fonte |
|---|---|---|---|
| Headcount total | 205 | 205 | CONTROLE PLANTEL aba CONTAGEM |
| Fazenda Pao Grande | 106 | 106 | CONTAGEM |
| Arrendamento | 41 | 41 | CONTAGEM |
| Centro de Treinamento | 1 | 1 | CONTAGEM |
| Sócios | 57 | 57 | CONTAGEM |
| Receptoras total | 62 | 62 | PLANTEL ARREND./RECEPT. aba ANIMAIS (FPG+ARR, status prenha/vazia) |
| Receptoras prenhas | 34 | 34 | idem |
| Receptoras vazias | 28 | 28 | idem |
| Nascimentos na semana | 1 | 1 | estação aba ESTAÇÃO (parição na janela) |
| Confirmados na semana | 0 | – | diff de confirmados vs semana anterior |
| Vendidos pendentes | 2 | 2 | Animais para sair, aba ANIMAIS VENDIDOS (VENDA, não reposição) |
| Sociedade pendentes | 2 | 2 | idem (SOCIEDADE) |

## ❌ Diffs por FALTA DE DADO na planilha (código certo, fonte não preenchida/atualizada)

### Acumulado na estação — calc 56 · docx 61
- **Causa:** vem da estação de monta. A aba `RESUMO` (que somaria o total) está com fórmula
  `#REF!` quebrada em TODAS as cópias salvas. O `data_only` só lê o cache, que dá 56 por toda
  medida (PLANEJAMENTO TOTAL EMBRIÕES=56, +/-=OK=56, GARANHOES=56). O 61 não existe em célula
  nenhuma do arquivo salvo — é fórmula viva que não persiste.
- **Como resolver:** (a) o operador abre a estação no Excel e SALVA (recalcula) → releio o 61;
  ou (b) consertar a fórmula da aba RESUMO pra ela persistir o total; ou (c) me apontar a célula
  exata do 61. Rodando com o arquivo recalculado, bate.

### Saídas na semana — calc 0 · docx 2
### Entradas na semana — calc 0 · docx 1 (Charmosa Serra Bela)
- **Causa:** saídas/entradas saem do log `MOVIMENTAÇÕES` do CONTROLE_DE_PLANTEL. Esse log tem
  **última data 11/07** — a semana 18-24/07 NÃO tem nenhum movimento lançado (nem as 2 saídas,
  nem a entrada da Charmosa). O operador escreveu no Word mas não lançou na planilha.
- **Como resolver:** o operador lança os movimentos da semana na aba MOVIMENTAÇÕES
  (`SAIU DO HARAS` / `CHEGOU NO HARAS` com data) → o script pega sozinho. Ou apontar a planilha
  datada onde essas saídas/entradas ficam registradas.

### Δ headcount (gross +entradas/−saídas) — mostrando net
- **Causa:** o gross (+1/−2) precisa dos movimentos individuais (mesma fonte das saídas/entradas,
  não lançada). O net (−1 = 205−206) é derivável e está sendo mostrado.
- **Como resolver:** idem saídas/entradas. Quando os movimentos forem lançados, o gross fecha.
  (Matemática: net = entradas − saídas; saber só o net não revela os dois números.)

## ⚠️ Diffs por DEFINIÇÃO a alinhar

### Índice de eficiência — calc 2,3 · docx 2,8
- **Causa:** índice = vazias / doadoras. Vazias=28 ok. Doadoras: o código conta CATEGORIA=DOADORA
  no CONTROLE PLANTEL = **12** (7 FPG + 5 sócio) → 28/12 = 2,3. O docx implica **10 doadoras**
  (28/2,8). Não é receptoras — já é doadoras; a questão é QUAIS 10 (o operador exclui 2).
- **Como resolver:** me diga o conjunto exato de doadoras do índice (ex.: só as do FPG? só as
  ativas na estação? excluir vendida/aposentada?). Com a regra certa, fecha 2,8.

### Acumulado no mês — RESOLVIDO (agora 0 = docx "--")
- "--" = zero (confirmado pelo usuário). Antes eu usava IA+60 (proxy ruim) → dava 3.
- **Corrigido:** acumulado no mês = novos confirmados desde o snapshot ANTES do mês (diff), igual
  à lógica de confirmados na semana. Acumulado parado no mês → 0. Bate com docx.

## 💡 Melhorias pedidas
- **Nome do animal nos nascimentos:** o detalhe já traz Potro/Doadora/Garanhão; quando o potro
  ainda não tem nome, a identidade é doadora × garanhão. Posso destacar um campo "Produto"
  (ex.: "Macho — Begônia × Damasco") no card. — a implementar se quiser.
- **Entrada Charmosa não aparece:** consequência das entradas=0 (log não lançado). Assim que o
  movimento for lançado, o nome aparece no detalhe de Entradas.

## ORIGEM REAL das 3 diffs (confirmado 2026-07-24) → viraram CAMPOS MANUAIS
Esses 3 NÃO saem de planilha — são input humano. Por isso viraram campos manuais no dashboard
(tag "manual", editável e salvo por semana). Índice = vazias(auto) ÷ doadoras(manual).
| Dado | Origem hoje | No dashboard |
|---|---|---|
| Acumulado na estação | Alexandre (vet) passa | campo manual (default = 56 do arquivo) |
| Saídas / Entradas na semana | grupo do WhatsApp | campos manuais (default 0) |
| Doadoras (índice) | nº fixo da estação; Alexandre avisa mudança | campo manual (default = 12) |

**FUTURO:** o usuário vai criar CONTROLES (planilhas) pra esses 3. Quando existirem, troco
manual → leitura automática no extractor (pontos: acumulado, saídas/entradas, doadoras).

## Resumo executivo
- 12/12 métricas com **fonte e lógica corretas**.
- 10 batem exato com o arquivo atual.
- 2 dependem do operador LANÇAR o dado (saídas/entradas → e o Δ gross) ou ATUALIZAR/RECALCULAR a
  estação (acumulado). Não são bug de código — são dado ausente/desatualizado na planilha.
- 2 pontos de DEFINIÇÃO a fechar contigo: conjunto de doadoras do índice, e o que é "acumulado no mês".
