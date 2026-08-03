# Hub HPG

Site único dos painéis do Haras Pao Grande. Mesmo desenho do `LuxorP&CHub`
(sidebar + rotas por hash + ECharts vendorizado, sem build step), com a
identidade do haras e **acesso próprio** — outras pessoas, outra allowlist,
nada compartilhado com o hub do P&C.

> **Fase atual: DEMO OFFLINE.** Abre o `index.html` direto no navegador
> (`file://`), sem login e sem rede. Os dados vêm dos `assets/data/*.js`
> gerados localmente. Supabase, Netlify e RBAC entram só quando for a gold.

## Rodar

```bash
python hub/tools/build_data.py      # bases -> assets/data/{comite,plantel}.js
python hub/tools/build_semanal.py   # dashboard do pipeline -> assets/semanal/dashboard.html
```

Depois é só abrir `hub/index.html`. Sem os builds, a aba correspondente nem
aparece na navegação (`temDado()` em `assets/app.js`) — melhor sumir do que
abrir painel vazio.

## Abas

| Aba | Fonte | Como entra |
|---|---|---|
| **Atualização Semanal** | `dashboards/dashboard_semanal.html` (PGSemanalDashboard.py, este repo) | iframe. Já nasce com a paleta do haras — o build só confere que segue autocontido e esconde o header interno, que duplicaria o título da aba |
| **Comitê Mensal** | `bases/base_bi.parquet` (PGBaseBI.py, este repo) | ECharts nativo. 27 meses (fev/24 → abr/26), plantel + patrimônio proporcional |
| **Plantel / Movimentação** | `mov_cascata` / `mov_detalhe` do `LuxorMonthlyP-CRoutines/PlantelHPG` (LxMovimentacao.py) | ECharts nativo. Cascata de valor por mês (59 meses, ago/21 → jun/26) |

Comitê e Plantel **não tinham painel web** — o comitê saía do `ComiteHPG.pbix` e
o plantel de e-mail com xlsx anexo. O que está aqui é primeira versão, feita a
partir do que as bases já entregam; nenhum número é recalculado no navegador.

## Estrutura

```text
hub/
├── index.html
├── assets/
│   ├── theme.css              # tema Haras Pao Grande (mesmos tokens do dashboard semanal)
│   ├── app.js                 # rotas, gráficos, tabelas
│   ├── config.js              # offline hoje; SUPABASE_URL + anon key na fase gold
│   ├── pg-logo.png
│   ├── vendor/echarts.min.js
│   ├── data/                  # GERADO — gitignored (dado do plantel)
│   └── semanal/               # GERADO — gitignored (dado do plantel)
└── tools/
    ├── build_data.py
    └── build_semanal.py
```

O `hub/` é autocontido de propósito: se a decisão for tirá-lo daqui e deixar
este repo só com o pipeline, é mover a pasta e nada mais.

## Tema

Os tokens do `theme.css` são os **mesmos** do `dashboard_semanal.html`
(navy `#04223B`, dourado `#CA9703`, azul `#7FA8C4`) — por isso a aba embutida
não precisa de re-skin e a costura entre casca e iframe não aparece. Gráficos
sem linha de grade.

## Para ir a gold (pendente)

1. **Repo/deploy** — o Netlify exige repo público pro deploy contínuo, e este
   repo é privado e versiona dado do plantel (`bases/*.json`,
   `_cache/*snapshots*`, `dashboards/dashboard_semanal.html`). Ou o `hub/` sai
   para um repo público próprio, ou o deploy passa a ser por CLI.
   **Atenção:** tornar este repo público expõe o histórico inteiro, não só o
   estado atual.
2. **Supabase próprio** — projeto novo, `allowed_users` + `user_dashboard_access`
   com a lista do haras. Copiar o `sql/hub_schema.sql` e o `assets/auth.js` do
   `LuxorP&CHub`, trocando os nomes dos dashboards para `semanal/comite/plantel`.
3. **Publish** — `tools/publish_hub.py` subindo os `.json` (já gerados ao lado
   do `.js`) e o `dashboard.html` do semanal para o bucket privado.
4. **Netlify** — site próprio. O plano é por conta, não por site: os créditos
   são os mesmos do hub do P&C.
