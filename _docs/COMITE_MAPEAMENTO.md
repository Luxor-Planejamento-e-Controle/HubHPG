# Comitê Mensal HPG — mapeamento de dados

De onde sai cada slide do *Relatório Mensal de Desempenho Estratégico*. Base:
o deck de **junho/2026** (52 slides) + o `Guia_Operacional_Relatorio_Mensal_PG.pdf`
(que descreve 41 — a diferença são os slides de foto, que multiplicaram).

Legenda de status:

- **auto** — dá pra ler direto de uma base que já existe, sem intervenção
- **semi** — a base existe, mas o número final depende de regra a fechar ou de campo que o haras ainda não preenche
- **manual** — é input humano (texto, print, foto); não tem fonte de dado
- **falta fonte** — o slide existe, mas o arquivo de origem não está no repo nem foi localizado

---

## Resumo por seção

| Seção | Slides | auto | semi | manual | falta fonte |
|---|---|---|---|---|---|
| Capa / Agenda / Divisores | 01, 02, 03, 15, 22, 28, 36, 52 | 8 | – | – | – |
| 01 Financeiro | 04–14 | 4 | 3 | 4 | – |
| 02 Estação de Monta | 16–21 | 4 | 1 | – | 1 |
| 03 Exposições | 23–27 | – | – | 5 | – |
| 04 Vendas | 29–35 | 5 | 1 | 1 | – |
| 05 Decisões e Manejo | 37–51 | 1 | – | 14 | – |

---

## Capa, agenda e divisores

| Slide | Conteúdo | Fonte | Status |
|---|---|---|---|
| 01 | Capa — mês/ano | parâmetro do run | auto |
| 02 | Agenda — 5 seções | fixo (estrutura do deck) | auto |
| 03, 15, 22, 28, 36 | Divisores de seção | fixo | auto |
| 52 | Encerramento | fixo | auto |

---

## Seção 01 — Financeiro (S04–S14)

| Slide | Conteúdo | Fonte | Aba / campos | Status |
|---|---|---|---|---|
| **04** | Resumo financeiro Haras — competência, orçado × realizado do mês | `DRE_2026_HPG_HARAS.xlsx` | aba `Real x Orçado (Comp)` — Orçado, Realizado, Δk, Δ% | auto |
| **05** | Análise de **custos** do mês, aberta por natureza | `DRE_2026_HPG_HARAS.xlsx` | aba `DRE-Compet`, col 30 = Orçado, col 31 = Realizado do mês | auto |
| **06** | Análise de **despesas** do mês, aberta por natureza | idem S05 | idem | auto |
| **07** | Haras competência — acumulado YTD | `DRE_2026_HPG_HARAS.xlsx` | aba `Real x Orçado (Comp)`, colunas de acumulado | auto |
| **08** | Comentários das variações YTD | `COMENTARIOS_DRE_HARAS.docx` | texto livre, escrito por pessoa | manual |
| **09** | Investimentos — compra de animais e produtos, mês a mês com descrição | `DRE_2026_HPG_HARAS.xlsx` aba `Investimentos` | descrição + valor por mês | semi — o valor é auto; a descrição de cada compra é escrita à mão |
| **10** | Haras **caixa** — orçado × realizado do mês | `DRE_2026_HPG_HARAS.xlsx` | aba `DRE-Caixa`, col 30–32 (mês), 33–35 (YTD) | semi — hoje o orçado do caixa é preenchido à mão |
| **11** | Estoque em equinos — headcount e patrimônio por categoria | `bases/base_bi.parquet` (PGBaseBI.py, este repo) | filtro `status_plantel = PLANTEL` e `sufixo_grupo` EXATO `DA PAO GRANDE` ou `OUTRO`; patrimônio = soma `patrimonio_proporcional`; valor médio = média `valor_100` | auto |
| **12** | Resumo da movimentação do plantel — saldo mensal e cascata | `LuxorMonthlyP-CRoutines/PlantelHPG/mov_cascata.parquet` | saldo_ini, compra, producao, venda, morte, doacao, reaval, saldo_fim | auto |
| **13** | Resumo financeiro Casa/FPG — orçado × realizado do mês | `DRE_2026_FPG_CASA.xlsx` | aba `Real x Orçado` | auto |
| **14** | Casa/FPG — acumulado YTD | idem S13 | idem, colunas de acumulado | auto |

> **Slide oculto — programação da casa.** Existe no arquivo mas fica escondido porque
> a casa não tem meta. Assim que houver orçado da casa, é o mesmo tratamento de S13/S14.

**Regras críticas (do guia, e valem à risca):**

- Competência é a aba `Real x Orçado (Comp)`; caixa é `DRE-Caixa`. Não misturar.
- S11 usa sufixo **exato** — `DA PAO GRANDE - E 50%` e variantes com percentual ficam de fora,
  senão o animal conta duas vezes.
- S09 é só compra de animais e produtos; obra e infraestrutura não entram.

**Onde isso encosta no que já existe:** o `LuxorP&CHub` já lê
`DRE Data/DRE_Historico.xlsx` do Drive com filtro de Centro de Custo (HPG/FPG) e
cenário Orçado/Realizado. Antes de criar leitura nova, conferir se o
`DRE_2026_HPG_HARAS.xlsx` é recorte desse mesmo histórico — se for, a fonte do
comitê é a que o hub do P&C já usa.

---

## Seção 02 — Estação de Monta (S16–S21)

| Slide | Conteúdo | Fonte | Aba / campos | Status |
|---|---|---|---|---|
| **16** | Embriões e prenhezes — funil (tentativas → lavados+ → 15/30/45/60d → abortos → confirmados) | `ESTACAO_DE_MONTA.xlsx` | aba `ESTAÇÃO`: K=lavado, M=15d, N=30d, O=45d, P=60d, Q=aborto, AJ=estação | auto |
| **17** | Garanhões — lavados, confirmados e índice, por tipo de sêmen | `ESTACAO_DE_MONTA.xlsx` | aba `GARANHOES`: garanhão, tipo de sêmen, total de lavados, lavados positivos, %, embriões confirmados, prenhez confirmada, aborto/absorção | auto |
| **18** | Comparativo com estações anteriores — embriões confirmados por mês | `ESTACAO_DE_MONTA.xlsx` aba `COMPARATIVO` | mês × estação | semi — a aba do arquivo em cache está em 20/21…23/24 e o deck mostra 22/23…25/26; confirmar se ela é atualizada ou se o número vem de outro lugar |
| **19** | Doadoras **Time A** — meta × realizado por doadora | `ESTACAO_DE_MONTA.xlsx` | aba `REC. EMBR.` (doadora, TIME, nº de embriões, lavados+) + aba `PLANEJAMENTO` col 7=Meta, 8=Real, 6=Time | auto |
| **20** | Doadoras **Time B** — mesma estrutura de S19 | idem S19, `TIME = B` | idem | auto |
| **21** | Coberturas disponíveis de garanhões de fora | `COBERTURAS_CAVALOS_FORA.xlsx`, aba `Planilha2` | garanhão + saldo; excluir *Trilho da Zizica* e *Quantum de Alcateia* | **falta fonte** — arquivo não está no repo nem no `_cache`; a `Planilha2` do ESTACAO é outra coisa (legenda) |

**Definições que o guia fixa e o dashboard tem que respeitar:**

- **Absorção** — perda antes da confirmação de 60 dias.
- **Aborto** — embrião já confirmado (>60d) que não nasceu.
- **Óbito** — nasceu e morreu.
- Contagem de confirmados: aba `ESTAÇÃO`, filtro estação `[ANO-1]/[ANO-2]`,
  `lavado = "+"`, `15d = "+"`, `30d/45d/60d = "+"` ou vazio, **menos** `aborto = "SIM"`.

Isso bate com o que o `PGSemanalReport.py` já faz no fechamento semanal
(`confirmado = coluna +/- == OK`) — vale alinhar as duas contagens antes de
publicar as duas telas lado a lado.

**Faltando definir (você marcou como "preciso revisar"):** de onde saem os
comparativos de anos anteriores nas visões doadoras e garanhões (S18, e o
histórico por trás de S19/S20).

---

## Seção 03 — Exposições (S23–S27)

| Slide | Conteúdo | Fonte | Status |
|---|---|---|---|
| **23** | Programação 2026 — evento, data, local, status | digitado | manual |
| **24–27** | Resultados por exposição — animais, títulos, colocações | digitado após cada evento | manual |

Não há planilha por trás. Se virar controle (uma aba de eventos com resultado),
passa a auto — hoje é texto.

---

## Seção 04 — Vendas (S29–S35)

| Slide | Conteúdo | Fonte | Aba / campos | Status |
|---|---|---|---|---|
| **29** | KPI de vendas — mês, YTD, meta anual, saldo para meta | `PG_MAPA_VENDAS.xlsx` + meta anual | aba `MAPA VENDAS`, col 7=valor, 14=vendedor, 21=ano, 22=mês | semi — a **meta anual** (R$ 4,5M) é parâmetro, não sai de planilha |
| **30** | Detalhamento de vendas por mês e evento/origem | idem S29 | idem | auto |
| **31** | Inadimplências e recebíveis | print do dashboard de cobrança (`controle-de-inadimplencia`) | imagem | manual hoje — o dashboard é HTML gerado por `ControleInadimplencia.py`; dá pra embutir em vez de printar |
| **32** | Embriões vendidos a fazer — quitado / pagando | `EMBRIOES_ENTREGAR_RECEBER.xlsx` | aba `ENTREGAR`, col 11=status pgto, 12=status embrião; filtro `status_embrião="A fazer"` **e** pgto `Pagando`/`Quitado` | auto |
| **33** | Embriões vendidos a fazer — pgto pausado / após confirmação | idem | `status_embrião="A fazer"` **e** pgto `Pgto pausado`/`Pgto após conf` | auto |
| **34** | Embriões de direito / reposição | idem | `status_embrião="Reposição"` **ou** (`"A fazer"` **e** pgto `Direito`/`Troca/Direito`) | auto |
| **35** | Embriões comprados a receber | idem, aba `RECEBER` | filtrar `status_embrião="A fazer"` | auto |

**Regras críticas:** vendas filtram `coluna 14 (VENDEDOR) = CARLA` e excluem status
`CANCELADO`. Um embrião só pode aparecer em **um** dos slides 32/33/34 — os três
filtros são mutuamente exclusivos e é onde o processo manual mais erra.

---

## Seção 05 — Decisões e Manejo (S37–S51)

| Slide | Conteúdo | Fonte | Status |
|---|---|---|---|
| **37** | Plantel — Pao Grande / Arrendamento / Sócios, com total | `CONTROLE PLANTEL.xlsx` aba `CONTAGEM` (o `PGSemanalReport.py` já lê isso) | auto |
| **38** | Manejo — histórico de intervenções e decisões, mês a mês | texto | manual |
| **39–51** | Fotos e registros do mês | fotos | manual |

---

## Planilhas de controle do HPG — o que cada uma é

| Planilha | O que controla | Alimenta |
|---|---|---|
| **Controle de entregas de embriões** (`EMBRIOES A ENTREGAR - A RECEBER.xlsx`, aba ENTREGAR) | embriões prontos e a fazer, por contrato de venda | S32, S33, S34 |
| **Controle de embriões a receber** (mesma pasta, aba RECEBER) | mesma ideia, do lado das compras | S35 |
| **Mapa de vendas** (`PG_Mapa Vendas.xlsx`) | controle geral de vendas | S29, S30 |
| **Mapa de compras** | idem, do lado das compras | S09 (investimentos), S35 |
| **Estação de monta** (`ESTACAO DE MONTA.xlsx`) | ciclo de fertilidade e reprodução — inseminação das doadoras, lavado (coleta do embrião), confirmação em 60d. **Referência da planilha = doadora** | S16–S20 |
| **Controle de nascimentos** | dados parecidos com a estação, em planilha separada | (a definir onde entra) |
| **Planejamento da estação de monta** (aba `PLANEJAMENTO`) | metas de embriões por doadora + informações dos garanhões | S19, S20 |
| **Recuperação embrionária** (aba `REC. EMBR.`) | embriões e lavados+ por doadora, com o Time (A/B) | S19, S20 |
| **Plantel de receptoras** (`PLANTEL ARRENDAMENTOS E RECEPTORAS.xlsx`) | disponibilidade das receptoras | fechamento semanal (não entra no comitê hoje) |
| **Controle de plantel** (`CONTROLE_DE_PLANTEL_PAO_GRANDE_*.xlsx`) | roster, movimentações, confirmações/abortos/mortes | S11, S12, S37 |
| **Coberturas de cavalos de fora** | saldo de coberturas compradas | S21 — **arquivo não localizado** |

---

## Pendências antes de automatizar

1. **`COBERTURAS_CAVALOS_FORA.xlsx`** — localizar. Sem ele, S21 fica manual.
2. **S18 (comparativo entre estações)** — a aba `COMPARATIVO` do arquivo em cache
   está desatualizada. Definir se ela é a fonte ou se o histórico vem de outro lugar.
3. **`DRE_2026_HPG_HARAS.xlsx` vs `DRE_Historico.xlsx`** — confirmar se é a mesma base.
   Se for, uma fonte só serve os dois hubs.
4. **Meta anual de vendas** e **meta de embriões** — hoje são parâmetro digitado.
   Definir onde ficam (planilha de metas ou campo do próprio hub).
5. **Alinhar a contagem de embriões confirmados** entre o comitê (filtro
   lavado/15d/30-45-60d menos aborto) e o fechamento semanal (`+/- == OK`).
6. **Inadimplência (S31)** — trocar o print por embed do HTML que o
   `ControleInadimplencia.py` já gera.
