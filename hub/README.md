# Hub HPG

Site único dos painéis do Haras Pao Grande. Mesmo desenho do `LuxorP&CHub`
(sidebar + rotas por hash, sem build step), com a identidade do haras e
**acesso próprio** — outras pessoas, outra allowlist, nada compartilhado com o
hub do P&C.

> **Fase atual: DEMO OFFLINE.** Abre o `index.html` direto no navegador
> (`file://`), sem login e sem rede. Supabase, Netlify e RBAC entram só quando
> for a gold.

## Rodar

```bash
python hub/tools/build_semanal.py   # dashboard do pipeline -> assets/semanal/dashboard.html
```

Depois é só abrir `hub/index.html`.

## Abas

**Atualização Semanal** (no ar) — fonte: `dashboards/dashboard_semanal.html`
(PGSemanalDashboard.py, este repo). Entra por iframe: já nasce com a paleta do
haras, então o build só confere que segue autocontido e esconde o header
interno, que duplicaria o título da aba.

**Comitê Mensal** e **Plantel / Movimentação** (em breve) — entradas
*placeholder*: aparecem na navegação com o selo "em breve" e uma tela que diz
qual base vai alimentar e de onde o dado sai hoje. Nenhum gráfico, nenhum número
de exemplo — não existe painel web dos dois (o comitê é montado no
`ComiteHPG.pbix`, o plantel sai por e-mail com xlsx anexo via
`LuxorMonthlyP-CRoutines/PlantelHPG/LxEmailHPGPlantel.py`) e, sem referência do
que mostrar, a casca não inventa layout.

Para promover um placeholder: em `assets/app.js`, trocar `soon:true` pela função
de render; se a tela tiver gráfico, copiar o `assets/vendor/echarts.min.js` do
`LuxorP&CHub` e voltar a tag `<script>` no `index.html`.

## Estrutura

```text
hub/
├── index.html
├── assets/
│   ├── theme.css              # tema Haras Pao Grande (mesmos tokens do dashboard semanal)
│   ├── app.js                 # rotas + nav + embed da aba
│   ├── config.js              # offline hoje; SUPABASE_URL + anon key na fase gold
│   ├── pg-logo.png            # GERADO por tools/build_logo.py
│   └── semanal/               # GERADO — gitignored (dado do plantel)
└── tools/
    ├── build_semanal.py
    └── build_logo.py          # só roda de novo se o logo mudar
```

O `hub/` é autocontido de propósito: se a decisão for tirá-lo daqui e deixar
este repo só com o pipeline, é mover a pasta e nada mais.

## Tema

Os tokens do `theme.css` são os **mesmos** do `dashboard_semanal.html`
(navy `#04223B`, dourado `#CA9703`, azul `#7FA8C4`) — por isso a aba embutida
não precisa de re-skin e a costura entre casca e iframe não aparece.

O `theme.css` traz o kit completo da casca (KPI, tabela, toolbar, multi-select,
gráfico), herdado do P&C Hub. Parte está sem uso enquanto só existe a aba
embutida; é o molde das próximas.

## Para ir a gold (pendente)

1. **Repo/deploy** — o Netlify exige repo público pro deploy contínuo, e este
   repo é privado e versiona dado do plantel (`bases/*.json`,
   `_cache/*snapshots*`, `dashboards/dashboard_semanal.html`). Ou o `hub/` sai
   para um repo público próprio, ou o deploy passa a ser por CLI.
   **Atenção:** tornar este repo público expõe o histórico inteiro, não só o
   estado atual.
2. **Supabase próprio** — projeto novo, `allowed_users` + `user_dashboard_access`
   com a lista do haras. Copiar o `sql/hub_schema.sql` e o `assets/auth.js` do
   `LuxorP&CHub`, trocando os nomes dos dashboards.
3. **Publish** — `tools/publish_hub.py` subindo o `dashboard.html` do semanal
   para o bucket privado (é dado do plantel, não pode virar arquivo público).
4. **Netlify** — site próprio. O plano é por conta, não por site: os créditos
   são os mesmos do hub do P&C.
