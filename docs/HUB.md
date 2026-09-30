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
│   │   ├── layout.js          # desenho dos slides (geometria do relatório da Ana) → HTML e PPTX
│   │   ├── deck.css · deck.js # navegação, conteúdo ao vivo, editor, versões, exportações
│   │   ├── logo-ouro.png · logo-navy.png  # logo nas duas cores do relatório
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

## Atualizar os dados pelo hub (botão + agente)

O botão **Atualizar dados** no hub não executa nada: ele grava um pedido na
tabela `hub_job`, e quem executa é um de dois executores, com os mesmos passos
(`tools/pipelines_hub.py`) e tomada atômica do pedido — cada pedido roda uma vez
só, mesmo com os dois de pé.

**Azure (desde 30/09/2026).** Um gatilho no `hub_job` (migration
`20260930170000`) avisa a rota `hpg_hub_job` do `luxor-planejamento-functions`,
que enfileira; o consumidor da fila roda `tools/roda_pedido.py <id> --nuvem`
numa cópia do repo: restaura o estado do bucket (inclusive
`estado_fontes.zip`, `estado_pdf.zip` e `estado_confirmados_extra.json`) e
espelha do Drive só o que o pipeline abre (`tools/sync_drive.py`, ~95 MB contra
~5,7 GB das pastas inteiras), com a conta de serviço do HPG
(`HPG_GOOGLE_SERVICE_ACCOUNT_JSON`). A App Setting `HPG_HUB_JOB_TIPOS` diz que
tipos a Azure atende; o resto fica para o notebook. Uma varredura a cada 10 min
reenfileira pedido parado e fecha como erro o que ficou preso em "rodando". A
sexta das 09:00 só cria um pedido de semanal. Endereço e chave da rota ficam no
Vault do Supabase (`hub_job_define_rota`, só service_role).

O comitê na nuvem lê do Blob (`LuxorControlDatabase/`, pela identidade do
Function App) o que no notebook está em outros repos: a base e os DRE anuais,
que o `LxDREdataExtractor.py` sobe ao gravar, e a foto da carteira, que o
`ControleInadimplencia.py` sobe. O executor baixa e aponta `HPG_DRE_DIR`,
`HPG_DRE_ANUAL_DIR` e `HPG_INAD_DIR` para a cópia. Do bucket vem a memória do
comitê (`estado_comite_*`: base_bi e parquets mensais, resumo da aba Plantel,
safras encerradas, cache do Trello), que o pedido de comitê republica — e só ela,
nunca a do semanal. O resumo do plantel roda o mesmo `tools/resumo_plantel_hub.js`
com o node que o `deploy.yml` põe no pacote (`HPG_NODE`). Atualizar/Gerar
reextraem o plantel do mês e reconsolidam o `base_bi` antes do deck.

Conferir o caminho da nuvem sem mexer em nada: numa CÓPIA do repo,
`python tools/roda_pedido.py <id> --nuvem --ensaio` (não toma o pedido, não
grava status, não publica).

**Notebook.** O pipeline também roda com o Google Drive montado em `G:`, pelo
agente desta máquina — é o executor dos tipos que ainda não foram para a Azure.

```bash
python tools/agente_hub.py            # processa a fila e sai
python tools/agente_hub.py --loop     # fica de pé, olhando a fila a cada 5 s
python tools/agente_hub.py --status   # só mostra a fila
```

Agendado em modo `--loop`, o pedido começa em até 5 s. O agendador dispara a
cada minuto só como vigia: a tarefa não inicia segunda instância, então os
disparos são ignorados enquanto o loop vive e, se ele cair (reboot, rede), o
minuto seguinte o levanta.

```bat
schtasks /create /tn "HPG - Agente do hub" /sc minute /mo 1 ^
  /tr "pythonw C:\Users\Arthur\repos\HubHPG\tools\agente_hub.py --loop"
```

Depois de criar, tirar o limite de duração da tarefa (`ExecutionTimeLimit =
"PT0S"`), senão o Windows mata o loop no meio de um pedido. Mudou o
`agente_hub.py`? Encerrar o `pythonw` do loop; o vigia sobe o código novo em
até 1 min.

O que o agente roda:

| pedido | sequência |
|---|---|
| `semanal` | `PGSemanal.py --no-open` → `tools/build_semanal.py` → `tools/build_auditoria.py` → `tools/publish_hub.py semanal auditoria estado` |
| `comite` — botão **Atualizar ‹mês›** | crava versão de cada mês → `tools/build_comite.py --atualizar` → `tools/publish_hub.py comite` |
| `comite` + `detalhe.acao = novo` — botão **Gerar ‹mês seguinte›** | crava versão de cada mês → `tools/build_comite.py --novo` → `tools/publish_hub.py comite` |

Os dois pedidos do comitê remontam **um** mês só e mesclam com o `comite.json`
publicado — os outros meses ficam exatamente como estão no ar. Atualizar remonta
o mês padrão do deck; Gerar cria o seguinte (que vira o padrão) e recusa enquanto
ele não acabou. Mês sem DRE fechado sai com os slides financeiros pendentes; o
Atualizar depois do fechamento (e do extractor do DRE) os preenche.

Duas garantias que o botão sozinho não daria:

- **Semanal não retroage.** O agente chama `PGSemanal.py` sem data e sem
  `--forcar`. A janela sai do último snapshot congelado (`_janela()`) e o próprio
  script aborta se existir semana posterior já fechada. Recusa do pipeline não é
  sucesso: o agente procura `ABORTADO` na saída e marca o pedido como erro, para
  o hub mostrar em vermelho em vez de dizer "atualizado".
- **Comitê não perde texto.** `build_comite.py` só lê `comite_conteudo`, mas o
  agente ainda crava uma versão de cada mês antes de rodar — rede para o caso de
  alguém pedir atualização no meio de uma edição.

## Versões do conteúdo do comitê

`comite_conteudo` tem uma linha por mês e o editor grava por upsert. Com várias
pessoas escrevendo o mesmo mês, quem salvasse por último sobrescrevia o anterior
sem deixar rastro. Agora:

- **toda gravação** guarda o estado anterior em `comite_conteudo_versao`, por
  trigger no banco (`comite_conteudo_versiona`) — não depende de o navegador
  lembrar. Update que não muda nada não gera versão.
- **⏱ Versões**, na barra do deck, lista o histórico, permite **cravar** o estado
  atual com um rótulo (`comite_cravar_versao`) e **restaurar** qualquer versão
  (`comite_restaurar_versao`). Restaurar não perde o atual: ele vira versão antes
  de ser substituído.

Os dois botões aparecem só para quem está em `comite_editores` (ou é admin).
