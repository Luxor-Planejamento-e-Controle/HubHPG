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
| 01 Financeiro | 04–14 | 9 | 1 | 1 | – |
| 02 Estação de Monta | 16–21 | 6 | – | – | – |
| 03 Exposições | 23–27 | – | – | 5 | – |
| 04 Vendas | 29–35 | 6 | 1 | – | – |
| 05 Decisões e Manejo | 37–51 | 1 | – | 14 | – |

No build de 06/2026: **42 de 47 slides com dado**. Os 5 que sobram são os que
não têm base — comentários do DRE (S08), exposições (S23/S24), manejo (S38) e
fotos (S39).

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
| **04** | Resumo financeiro Haras — competência, orçado × realizado do mês | `DRE_Historico.xlsx` | `Base DRE Geral`, CC=HPG, Modelo=Competência, só `É Subtotal`; desde ago/26 abre Venda de Produtos em Embriões/Coberturas/Óvulos/Animais **quando a linha tem valor** (orçado ou realizado). Conferido 21/21 contra a face (arredondamento do Excel: meio pra longe do zero) | auto |
| **05** | Análise de **custos** do mês, aberta por natureza | `DRE_Historico.xlsx` | `Base DRE Geral`, `Grupo = CUSTOS E DESPESAS OPERACIONAIS`; linha zerada no mês omitida | auto |
| **06** | Análise de **despesas** do mês, aberta por natureza | `DRE_Historico.xlsx` | `Base DRE Geral`, `Grupo = DESPESAS` | auto |
| **07** | Haras competência — acumulado YTD | `DRE_Historico.xlsx` | `Base YTD`, faixa `NN-Jan a <Mês>`; mesmas linhas do S04. **Apresentado** desde ago/26 (em jul/26 era oculto). Ago/26: 23/23 com as colunas YTD da face | auto |
| **08** | Comentários das variações do mês | **Trello** — quadro `Fluxo de Caixa`, card `DRE Haras - <Mês> <ano>`, comentário `COMENTÁRIOS DRE (HPG) – <MÊS>/<AA>` da controladoria (arquivado em `_cache/trello/`) | uma faixa por categoria, na ordem do DRE; texto = cada natureza com o valor e a explicação dela; ∆ da categoria = face **Real x Orçado (Caixa)** (o comentário é sobre o caixa); quantos slides o texto pedir | auto — sem comentário no Trello, vale o escrito no hub (`comite_conteudo.comentarios`); o slide do Trello não abre o editor |
| **09** | Investimentos do mês — todos os blocos | `DRE 2026 HPG - HARAS v3.xlsx` aba `Investimentos` | seção `INVESTIMENTOS - <MÊS>/<AA>`: blocos INFRAESTRUTURA / FORMAÇÃO DE PASTAGEM / MÁQUINAS E EQUIPAMENTOS / COMPRA DE ANIMAIS E PRODUTOS (A=favorecido, B=descrição, C=valor) + total da seção; lançamento repetido vira uma linha (2×); ago/26 bate com o relatório | auto — é o total da aba, que difere do bloco Investimentos da face |
| **10** | Haras **caixa** — orçado × realizado do mês | `DRE_Historico.xlsx` | `Base DRE Geral`, CC=HPG, `Modelo = Caixa`; abre produtos e os blocos de investimento (Máquinas, Infraestrutura, Animais e Produtos) quando têm valor. Ago/26: 15/15 com a face | auto |
| **11** | Estoque em equinos — headcount e patrimônio por categoria | `bases/base_bi.parquet` (PGBaseBI.py, este repo) + aba `Resumo Contabil` do mapa de movimentações | filtro `status_plantel = PLANTEL` e `sufixo_grupo` EXATO `DA PAO GRANDE` ou `OUTRO` (o `E 100%` do Eduardo NÃO entra — vale só pro semanal); **patrimônio = saldo final do Resumo Contábil do mês** (o mesmo número do S12; jul/26 R$ 15.970.552,61), valor médio = soma `valor_100` ÷ **avaliados** desde ago/26 ("N animais avaliados"; até jul/26 ÷ todos, como foi apresentado); todas as categorias abertas (coluna CATEGORIA da aba PLANTEL — a divisão do relatório de ago/26 tem reclassificação manual, sem regra na planilha) | auto |
| **12** | Resumo da movimentação do plantel — saldo mensal | **aba Plantel do hub** — `tools/resumo_plantel_hub.js` roda o motor dela (`assets/plantel/plantel.js`) sobre `plantel.<AAAA-MM>.json` do bucket + `plantel_mov_classificacao` (arquivado em `_cache/plantel_hub/`) | Saldo inicial, (+) Compras, (+) Produção embriões, (−) Baixa vendas, (−) Baixa mortes e doações, (+/−) Reavaliações, Saldo final — só mês fechado (ou 100% classificado) no hub; jan–jul/26 idênticos ao Resumo Contábil divulgado | auto — mês que o hub não fechou sai da aba `Resumo Contabil` do mapa, e o subtítulo diz |
| **13** | Resumo financeiro Casa/FPG — orçado × realizado do mês | `DRE_Historico.xlsx` | `Base DRE Geral`, **CC=FPG**, `Modelo = Caixa`; + Receitas com Locação. Ago/26: 17/17 com a face | auto |
| **14** | Casa/FPG — acumulado YTD | `DRE_Historico.xlsx` | `Base YTD`, CC=FPG. Apresentado desde ago/26; `Deduções` = DEDUÇÕES E IMPOSTOS da base (saía zerada) | auto |

> **Slide oculto — programação da casa.** Existe no arquivo mas fica escondido porque
> a casa não tem meta. Assim que houver orçado da casa, é o mesmo tratamento de S13/S14.

**Regras críticas (do guia, e valem à risca):**

- Competência e caixa são o campo `Modelo` do histórico. Não misturar.
- A hierarquia vem de `Grupo`/`Subgrupo`/`É Subtotal`: nível 0 é a linha do grupo,
  nível 1 o subgrupo, nível 2 a natureza folha. Tratar tudo que é `É Subtotal`
  como nível 0 deixa o slide inteiro em negrito e sem leitura.
- Nome de subgrupo se repete entre grupos (Sanidade e Reprodução aparecem em
  Custos **e** em Despesas) — desambiguar com o grupo entre parênteses.
- S11 usa sufixo **exato** — `DA PAO GRANDE - E 50%`, `- E 25%` e as demais fatias
  parciais ficam de fora, senão o animal dividido conta duas vezes (ele já entra
  pela cota).
- **Exceção desde 04/08/2026:** `DA PAO GRANDE - E 100%` **entra**. É o que é
  inteiro do Eduardo (cota 1,0, sem divisão), e o headcount passou a contar
  Carla + Eduardo. Efeito em jun/26: 186 → **190 animais**, patrimônio
  R$ 15,8M → **R$ 18,3M**. Quase toda a diferença é um bicho só — Loucura da Pao
  Grande, égua de pista avaliada em R$ 2,0M; Loteria da Pao Grande entra com
  R$ 500k e os dois embriões, sem valor.
- S09 é só compra de animais e produtos; obra e infraestrutura não entram.

**Fonte única com o P&C:** toda a seção financeira lê o mesmo
`DRE Data/DRE_Historico.xlsx` que o `LuxorP&CHub` já usa. Os arquivos
`DRE 2026 HPG - HARAS.xlsx` / `DRE 2026 FPG - CASA.xlsx` **não** são fonte do
comitê (ver pendência 3). A exceção é o S09, que continua na aba `Investimentos`
do arquivo do ano, porque a descrição de cada compra não existe no histórico.

---

## Seção 02 — Estação de Monta (S16–S21)

| Slide | Conteúdo | Fonte | Aba / campos | Status |
|---|---|---|---|---|
| **16** | Embriões e prenhezes — funil (tentativas → lavados+ → 15/30/45/60d → abortos → confirmados) | `ESTACAO_DE_MONTA.xlsx` | aba `ESTAÇÃO`: K=lavado, M=15d, N=30d, O=45d, P=60d, Q=aborto, AJ=estação | auto |
| **17** | Garanhões — lavados, confirmados e índice, por tipo de sêmen | `ESTACAO_DE_MONTA.xlsx` | aba `GARANHOES` (lavados, positivos, total confirmados, tipo); garanhão com tentativa na safra que a aba esquece entra com a conta da aba `ESTAÇÃO` (como o relatório fez com Latino/Esteio/Invencível). Cartões = positivos ÷ lavados por tipo; "ref." é referência fixa (60% / 50-60% / 70%) | auto |
| **18** | Comparativo entre estações — as 4 últimas safras, confirmados por mês | `ESTACAO_DE_MONTA.xlsx` aba **`ESTAÇÃO`**; safra anterior a 23/24 sai do master da pasta dela (aba `CONTROLE_DOADORAS`) | mês do embrião = **mês da IA/cobrição** (24/25 bate mês a mês com o relatório); "doad" = doadoras com embrião confirmado; "Meta" = confirmados ÷ META TOTAL do PLANEJAMENTO da safra — safras fechadas sem PLANEJAMENTO usam o % publicado em jul/26 (`META_PCT_OFICIAL`) | auto — a aba `COMPARATIVO` está congelada em 20/21–23/24 |
| **19** | Doadoras **Time A** — meta × realizado por doadora | `ESTACAO_DE_MONTA.xlsx` | aba `PLANEJAMENTO`, colunas **pelo cabeçalho** (NOME, TIME, META TOTAL, TOTAL EMBRIÕES — em 26/27 não há TIME e sai um slide só); time vazio usa o da `REC. EMBR.`; meta do time soma só quem tem o time no PLANEJAMENTO; Rec. Embrionária = lavados+ ÷ tentativas e Prenhez = 15d ÷ lavados+, pela aba `ESTAÇÃO` | auto |
| **20** | Doadoras **Time B** — mesma estrutura de S19 | idem S19, `TIME = B` | idem | auto |
| **21** | Coberturas disponíveis de garanhões de fora | **`REPRODUÇÃO/COBERTURAS - CAVALOS DE FORA NÃO USADAS.xlsx`** no Drive | aba `Planilha2`: só **saldo > 0**, do maior pro menor; a leitura para em `ARQUIVO MORTO` (abaixo dele a aba repete o cabeçalho e guarda os encerrados); excluir *Trilho da Zizica* e *Quantum de Alcateia* | auto |

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
| **23** | Programação — evento, data, local, status | `_docs/comite_conteudo.json` → `exposicoes.programacao` | manual estruturado |
| **24–27** | Resultados por exposição — animal e prêmios | `_docs/comite_conteudo.json` → `exposicoes.resultados` | manual estruturado |

Não há planilha por trás, mas também não é mais placeholder: o conteúdo do deck
de junho foi extraído para o JSON e o slide é montado a partir dele. Todo mês
novo é uma chave `AAAA-MM` no arquivo. Mês sem conteúdo escrito volta a aparecer
como slide em aberto, dizendo exatamente isso — nunca repete o mês anterior.

---

## Seção 04 — Vendas (S29–S35)

| Slide | Conteúdo | Fonte | Aba / campos | Status |
|---|---|---|---|---|
| **29** | KPI de vendas — mês, YTD, **média mensal** (YTD ÷ meses decorridos), meta anual, saldo para meta, colunas por mês | `PG_Mapa Vendas.xlsx` **do fechamento do mês** (1ª versão gerada depois do fim do mês) + meta anual | aba `MAPA VENDAS`, col 8=valor da venda, 15=vendedor (contém CARLA), 19=status (sem CANCELADO), 22=ano, 23=mês | semi — a **meta anual** (R$ 4,5M) é parâmetro |
| **30** | Detalhamento de vendas por mês e evento/origem — **um slide só** | idem S29 | origem = `Venda Direta` ou o **NOME** do evento (col 12), espaço duplo normalizado (a Semana vinha em duas grafias) | auto |
| **31** | Inadimplências e recebíveis — o painel de cobrança | `controle-de-inadimplencia` → `output_pbi/historico/fato_titulos_<fim do mês>.parquet` (e o do mês anterior, pra variação) | carteira **CAR inteira** desde ago/26 — é o que o dash do hub P&C mostra e o que o ControleInadimplencia grava em `indicadores_kpi_historico.xlsx` (o build confere e avisa se divergir); até jul/26, carteira **Carla** (o print da Ana, que bate na vírgula). Regra do `dashboard_conferencia.html` (vencido = status vencido/parcial e > 7 dias; AJ/NE/Inadimplentes = quebra dos vencidos); 6 cartões + rosca + índice por ano de emissão | auto — **agregados, sem PII**; sem foto do fim do mês o slide fica pendente |
| **32** | Embriões vendidos a fazer — quitado / pagando | `EMBRIOES_ENTREGAR_RECEBER.xlsx` | aba `ENTREGAR`, col 11=status pgto, 12=status embrião; filtro `status_embrião="A fazer"` **e** pgto `Pagando`/`Quitado`; **ordem da planilha**, venda posterior ao mês do deck fica fora, altura de linha padrão (16 por slide) | auto |
| **33** | Embriões vendidos a fazer — pgto pausado / após confirmação | idem | `status_embrião="A fazer"` **e** pgto `Pgto pausado`/`Pgto após conf` | auto |
| **34** | Embriões de direito / reposição | idem | `status_embrião="Reposição"` **ou** (`"A fazer"` **e** pgto `Direito`/`Troca/Direito`); coluna PGTO = status do pagamento (ago/26), cor por status como nos outros três | auto |
| **35** | Embriões comprados a receber | idem, aba `RECEBER` | filtrar `status_embrião="A fazer"` | auto |

**Regras críticas:** vendas filtram `coluna 14 (VENDEDOR) = CARLA` e excluem status
`CANCELADO`. Um embrião só pode aparecer em **um** dos slides 32/33/34 — os três
filtros são mutuamente exclusivos e é onde o processo manual mais erra.

---

## Seção 05 — Decisões e Manejo (S37–S51)

| Slide | Conteúdo | Fonte | Status |
|---|---|---|---|
| **37** | Plantel — Pao Grande / Arrendamento / Sócios, com total | `CONTROLE_DE_PLANTEL` do fechamento do mês (`headcount_de`) | auto — **oculto**: o relatório de jul/26 não tem esse slide; ele fica no arquivo, fora da apresentação |
| **38** | Manejo — histórico de intervenções e decisões, **um slide por semestre** | `comite_conteudo.manejo` de **todos** os meses até o do deck (o texto mais recente de cada mês vale) | manual estruturado — o deck ao vivo remonta com a mesma regra |
| **39+** | Fotos e registros do mês | `_docs/comite_conteudo.json` → `fotos` + `hub/assets/comite/fotos/` | manual — 6 fotos por slide, quantos slides forem precisos |

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

1. ~~**`COBERTURAS_CAVALOS_FORA.xlsx`** — localizar.~~ **Resolvido:** é o
   `REPRODUÇÃO/COBERTURAS - CAVALOS DE FORA NÃO USADAS.xlsx` do Drive, só com
   outro nome. `Planilha1` é o log de compra; `Planilha2` é o consolidado por
   garanhão, que é o que o slide usa. Hoje: 13 garanhões com saldo, 35 coberturas
   a usar.
2. ~~**S18 (comparativo entre estações)**~~ **Resolvido:** deixou de ler a aba
   `COMPARATIVO` (congelada em 23/24) e passou a calcular da aba `ESTAÇÃO`,
   agrupando por safra. Hoje sai 23/24 = 78 · 24/25 = 87 · 25/26 = 56 confirmados,
   com abortos, absorções, lavados e tentativas na mesma tabela.
3. ~~**`DRE_2026_HPG_HARAS.xlsx`** — localizar.~~ **Resolvido, e a fonte é outra.**
   Os arquivos do ano existem em `G:/…/Ambiente de testes/DRE Data/`, mas **não
   servem**: a aba de resumo mostra só o mês em que o operador deixou o arquivo
   (em 03/08/2026 estava em *fevereiro*) e a coluna Orçado estava zerada.
   A fonte certa é o **`DRE_Historico.xlsx`, aba `Base DRE Geral`** (mês) e
   **`Base YTD`** (acumulado) — a base consolidada que o `LxDREdataExtractor`
   gera e que o `LuxorP&CHub` já lê. Ela traz os 12 meses, HPG e FPG,
   Competência e Caixa, com Orçado e Realizado de verdade, além de
   `Grupo`/`Subgrupo`/`É Subtotal`/`Ordem` para a hierarquia.
   Conferido: jun/26 Receita Bruta orçado R$ 301.670 / realizado R$ 561.254,68 —
   **idêntico ao deck oficial de junho**. Uma fonte serve os dois hubs.
4. **Meta anual de vendas** e **meta de embriões** — hoje são parâmetro digitado.
   Definir onde ficam (planilha de metas ou campo do próprio hub).
5. **Alinhar a contagem de embriões confirmados** entre o comitê (funil
   encadeado lavado → 15d → 30/45/60d, menos aborto) e o fechamento semanal
   (`+/- == OK`). O funil calculado bate exato com o deck oficial de junho
   (196 tentativas · 88 lavados · 66/62/59/59 · 3 abortos · **56 confirmados**).
7. **Aba `GARANHOES` × aba `ESTAÇÃO`** — somando a tabela de garanhões dá
   157 lavados / 74 positivos / 55 confirmados; o funil da aba `ESTAÇÃO` dá
   196 / 88 / 56. As duas abas são mantidas em separado e divergem. Definir qual
   manda (ou corrigir a de garanhões).
6. ~~**Inadimplência (S31)** — trocar o print.~~ **Resolvido:** o slide passou a
   sair dos agregados (`indicadores_kpi.xlsx` + `resumo_por_faixa.xlsx`), que não
   têm nome de devedor. Embutir o `dashboard_conferencia.html` foi descartado: são
   2,7 MB de HTML que não viram slide de PPTX e carregariam PII pra dentro do deck.

## Formato (desde 24/09/2026)

O desenho de cada slide é o do relatório da Ana (`RELATORIO_MENSAL_PG_JULHO26`),
na geometria dela — posições, corpos de fonte e cores tirados do próprio PPTX —,
em `assets/comite/layout.js`. A mesma lista de primitivas vira o HTML do hub e o
PPTX exportado; o PDF é a impressão do HTML. Sem rodapé. Três slides saem
**ocultos** (existem no arquivo, fora da apresentação, como no dela): Haras YTD,
Casa YTD e a contagem por local (S37); os comentários ficam ocultos só quando
estão vazios.

- **S03 Pendências da apresentação anterior** — novo, conteúdo manual em
  `comite_conteudo.pendencias` (migration 20260924150000), editável pelo hub.
- **Resumo (S04), Caixa (S10), Casa (S13)** — as linhas do relatório dela (19, 6
  e 15), com os três pesos de linha; o valor vem da face oficial (`Real x Orçado`).
  Na Casa, Marketing e Hospedagem Família entram quando têm valor (senão os
  detalhes não fecham com Despesas Gerais).
- **Análise de custos e de despesas** — duas páginas cada, com a lista de naturezas
  que ela acompanha (`ANALISE_CUSTOS` / `ANALISE_DESPESAS` no build), TOTAL no topo.
