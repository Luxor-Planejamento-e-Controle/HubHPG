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
python tools/build_semanal.py   # dashboard do pipeline -> assets/semanal/dashboard.html
python tools/build_comite.py    # bases -> assets/comite/spec.js (todos os meses fechados)
python tools/build_comite.py 06/2026   # só um mês
```

Depois é só abrir `hub/index.html`.

## Abas

**Atualização Semanal** (no ar) — fonte: `dashboards/dashboard_semanal.html`
(PGSemanalDashboard.py, este repo). Entra por iframe: já nasce com a paleta do
haras, então o build só confere que segue autocontido e esconde o header
interno, que duplicaria o título da aba.

**Comitê Mensal** (no ar) — o *Relatório Mensal de Desempenho Estratégico*, que
hoje é um PPTX montado à mão. Vira deck HTML em `comite.html`: slide de
1280×720 (16:9, mesma proporção do PPTX), navegação por seta, modo apresentação
(tecla **P**, **Esc** sai) e **Exportar PPTX**.

O desenho é *um spec, duas saídas*: `tools/build_comite.py` lê as bases e
grava `assets/comite/spec.js` — uma lista de `{t: <tipo de slide>, ...}`. O HTML
e o PPTX renderizam **o mesmo spec**, então não existe conteúdo que só viva num
dos dois. Slide sem fonte vira tipo `pendente` e diz na tela qual base vai
alimentar e por que ainda não tem — nada de número de exemplo.

Mapa de slide × fonte: [`_docs/COMITE_MAPEAMENTO.md`](../_docs/COMITE_MAPEAMENTO.md).

O que **não** sai de planilha — comentários do DRE, exposições, manejo e fotos —
mora em [`_docs/comite_conteudo.json`](../_docs/comite_conteudo.json), uma chave
por mês (`AAAA-MM`). Foi semeado do último deck aprovado por
`tools/extrair_conteudo.py`; daí em diante é editar o JSON. Mês sem conteúdo
escrito mostra o slide em aberto dizendo isso — nunca repete o mês anterior.

**Plantel / Movimentação** (em breve) — entrada *placeholder*: aparece na
navegação com o selo "em breve" e uma tela que diz qual base vai alimentar. Hoje
sai por e-mail com xlsx anexo (`LuxorMonthlyP-CRoutines/PlantelHPG/LxEmailHPGPlantel.py`);
sem referência do que mostrar, a casca não inventa layout.

Para promover um placeholder: em `assets/app.js`, trocar `soon:true` pela função
de render; se a tela tiver gráfico, copiar o `assets/vendor/echarts.min.js` do
`LuxorP&CHub` e voltar a tag `<script>` no `index.html`.

## Estrutura

```text
hub/
├── index.html                 # casca do hub
├── comite.html                # deck do comitê (roda dentro do hub, em iframe)
├── assets/
│   ├── theme.css              # tema Haras Pao Grande (mesmos tokens do dashboard semanal)
│   ├── app.js                 # rotas + nav + embed das abas
│   ├── config.js              # offline hoje; SUPABASE_URL + anon key na fase gold
│   ├── pg-logo.png            # GERADO por tools/build_logo.py
│   ├── comite/
│   │   ├── deck.css · deck.js # renderer do deck + export PPTX
│   │   ├── spec.js            # GERADO — gitignored (número de DRE)
│   │   └── fotos/             # GERADO — gitignored (12 MB)
│   ├── semanal/               # GERADO — gitignored (dado do plantel)
│   └── vendor/pptxgen.bundle.js
└── tools/
    ├── build_semanal.py
    ├── build_comite.py
    ├── extrair_conteudo.py    # semeia o conteúdo manual do último PPTX aprovado
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
