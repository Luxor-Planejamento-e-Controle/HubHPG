# HubHPG — Automações do Haras Pao Grande

Automação dos relatórios do Haras Pao Grande. Lê as planilhas de controle do Drive
e gera as saídas (dashboard semanal, base do comitê mensal).

> **Só código aqui.** `bases/`, `dashboards/`, `_cache/` e `_docs/` estão no
> `.gitignore` — dado do plantel não é versionado. As fontes continuam no Drive
> (`G:\.shortcut-targets-by-id\...`, resolvido em `scripts/_pg_common.py`) e as
> pastas de saída são criadas pelo próprio pipeline no primeiro run.
>
> Produção roda em `Luxor Controladoria\Ambiente de testes\Projeto automações HPG`
> (git corrompe dentro do Google Drive): o repo é a fonte da verdade do código,
> a pasta do Drive é onde executa.

> **`_cache/` não é só cache — guarda o estado do pipeline.** Dois arquivos ali
> são memória, não arquivo descartável:
>
> - `_cache/semanal_snapshots.json` — a foto congelada de cada semana. É o que
>   alimenta o seletor de semanas do dashboard e a base de comparação.
> - `_cache/headcount_history.json` — o total por semana, de onde sai o Δ.
>
> Eles ficam fora do Git junto com o resto do `_cache/` (é dado do plantel, e a
> regra é dado só no Drive). Consequência prática: **rodar de um clone limpo
> começa o histórico do zero** — sem Δ de headcount, sem saídas/entradas por
> diferença de roster e com uma única semana no dashboard, até acumular a
> segunda captura. Por isso o fechamento roda da pasta do Drive, onde esse
> estado vive e é preservado pelo backup do próprio Drive.

```bash
pip install -r requirements.txt
```

## Estrutura

```text
.
├── PGSemanal.py            # ⭐ ORQUESTRADOR do fechamento semanal (único .py na raiz — entry point)
│
├── scripts/                # todos os scripts do pipeline
│   ├── PGSemanalDocx.py        # parseia relatórios oficiais (Word) → bases/semanal_docx.json  [validação]
│   ├── PGSemanalReport.py      # calcula a semana das planilhas → bases/semanal_data.json (+ snapshot)
│   ├── PGSemanalDashboard.py   # gera dashboards/dashboard_semanal.html
│   ├── PGSemanalValidar.py     # compara calculado vs docx (placar)
│   ├── PGDataExtractor.py      # [comitê mensal] extrai 1 mês do CONTROLE_DE_PLANTEL
│   ├── PGBaseBI.py             # [comitê mensal] consolida → bases/base_bi.xlsx
│   ├── PGRunHistorico.py       # [comitê mensal] bootstrap histórico
│   └── _pg_common.py           # helpers compartilhados (paths do Drive, cache, MesRef)
│
├── bases/                  # DADOS (.xlsx .parquet .json): base_bi, semanal_data, semanal_docx
├── dashboards/             # SAÍDAS visuais: dashboard_semanal.html, ComiteHPG.pbix
├── _cache/                 # cópia local das planilhas do Drive (read-only) + snapshots
└── _docs/                  # modelos de referência + relatórios antigos
```

> O orquestrador (`PGSemanal.py`) fica na raiz e adiciona `scripts/` ao path. Os scripts
> também rodam soltos: `python scripts/PGSemanalReport.py ...`, `python scripts/PGSemanalValidar.py`.

## Fechamento semanal — como rodar

Um comando faz tudo (parseia relatórios → calcula → gera dashboard → abre → valida):

```bash
python PGSemanal.py 17/07/2026      # semana de referência = essa data
python PGSemanal.py                  # semana de referência = hoje
python PGSemanal.py 17/07/2026 --no-open --no-docx
```

O dashboard mostra **só o que o extractor calcula** das planilhas. Os relatórios Word (docx)
servem **só de validação** (bater os números), nunca viram dado no dashboard.

### Fontes no Drive (layout 2026-07)

| Seção | Arquivo | Aba |
|---|---|---|
| Headcount | `ATUALIZACAO SEMANAL/CONTROLE PLANTEL.xlsx` | `CONTAGEM` |
| Receptoras | `PLANTEL/Estação 2025-2026/{YYMMDD}_PLANTEL ARRENDAMENTOS E RECEPTORAS.xlsx` | `ATUALIZAÇÃO SEMANAL` |
| Produção/estação | `REPRODUÇÃO/ESTAÇÃO DE MONTA/Estação 2025-2026/{YYMMDD} ESTACAO DE MONTA.xlsx` | `ESTAÇÃO`, `PLANEJAMENTO` |
| Movimentação | `PLANTEL/Estação 2025-2026/{YYMMDD}_CONTROLE_DE_PLANTEL_*.xlsx` | `MOVIMENTAÇÕES` |
| Pendentes | `ATUALIZACAO SEMANAL/CONTROLE PLANTEL.xlsx` | `PLANTEL` |
| Vendas (cota) | `VENDAS/MAPAS DE VENDAS/Estação 2025-2026/{YYMMDD}_PG_Mapa Vendas.xlsx` | `MAPA VENDAS` |

## Lógica das métricas (semanal)

| Métrica | Regra / fonte |
|---|---|
| Headcount (total + por local) | aba `CONTAGEM` (pré-agregada) |
| Acumulado na estação | `SUM(PLANEJAMENTO 'TOTAL EMBRIÕES' REAL)` |
| Embrião confirmado | coluna `+/-` == `OK` na aba ESTAÇÃO |
| Confirmados na semana | diff do conjunto confirmado vs semana anterior (forward) |
| Receptoras (total/prenhas/vazias) | aba ATUALIZAÇÃO SEMANAL, colunas FPG + ARRENDAMENTO |
| Índice eficiência | vazias ÷ doadoras (CATEGORIA=DOADORA no plantel) |
| Saídas/entradas na semana | diff do roster do plantel vs semana anterior (forward) |
| Nascimentos/abortos | data de parição/aborto na janela (aba ESTAÇÃO) |
| Vendidos pendentes | STATUS = `PENDENTE PARA SAÍDA` na aba PLANTEL |

**Snapshot / forward:** cada run congela o snapshot da semana em `_cache/semanal_snapshots.json`.
As métricas de *diff* (saídas/entradas, confirmados) só têm valor a partir da 2ª semana capturada;
na 1ª (bootstrap) são semeadas do relatório oficial. As métricas de *estado* (acumulado, receptoras,
headcount) só batem o docx quando a planilha-fonte está fresca (rodar na data do fechamento).

## Comitê mensal (inalterado)

```bash
python PGDataExtractor.py     # prompt MM/AAAA
python PGBaseBI.py            # reconstrói base_bi.xlsx
```
