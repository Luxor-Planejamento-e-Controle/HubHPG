# Deploy do Hub HPG

Mesmos conceitos do `LuxorP&CHub`: repo público no Netlify, dado no bucket
privado do Supabase, acesso por allowlist invite-only com magic-link.

O que muda em relação ao P&C: **projeto Supabase próprio**. Outra allowlist,
outras pessoas, outro bucket (`hpg-data`). Erro de policy num hub não alcança o
outro.

---

## Regra número um: crédito do Netlify

O plano Free dá **300 créditos/mês e cobra 15 por deploy de PRODUÇÃO** — 20
deploys por ciclo, e o ciclo vira dia 20. No P&C isso estourou em 28/07/2026 e o
site ficou **três semanas congelado**: todo push passou a falhar com
`Skipped due to account credit usage exceeded`.

Duas defesas, as duas já no repo:

1. **`netlify.toml` tem `ignore`**: commit que só toca `tools/`, `scripts/`,
   `docs/`, `sql/`, `*.md`, `.githooks/` ou `PGSemanal.py` **não** dispara build.
   No P&C, 17 dos 28 deploys do primeiro ciclo republicaram um site idêntico —
   255 créditos à toa.
2. **Deploy preview não consome crédito.** Só produção.

**Publicar dado não é deploy.** Rodar o pipeline e o `publish_hub.py` atualiza o
hub sem tocar no Netlify — o site lê o bucket em tempo de execução. Só mexer em
HTML/CSS/JS exige deploy novo.

---

## Ordem de operações (primeira vez)

### 1. Supabase

1. Criar projeto novo, só do HPG.
2. `Authentication > Providers > Email`: **desligar Signups**. O hub é
   invite-only; magic-link não cria conta.
3. `Authentication > URL Configuration`: Site URL e Redirect URLs apontando pro
   domínio do Netlify. Redirect não liberado devolve `validation_failed` no link.
4. Rodar `sql/hub_schema.sql` no SQL editor. Cria allowlist, RBAC, RLS, o bucket
   privado `hpg-data` e o log de acesso.
5. Semear a allowlist **no Supabase, nunca no Git** — modelo em
   `sql/seed_allowlist.example.sql`. O primeiro registro tem de ser um `admin`,
   senão ninguém entra pra liberar os outros.

### 2. Configuração pública

Preencher `assets/config.js` com a URL do projeto e a **anon key**. Ela é pública
por design: quem protege é a RLS. A `service_role` **nunca** entra aí — vai só no
`.env` local, que é gitignored, e serve ao `publish_hub.py`.

### 3. Netlify

1. Conectar o repo. `publish = "."`, sem build command — é site estático.
2. Conferir que o `netlify.toml` veio com o `ignore`.
3. **Não** subir para produção antes de o Supabase estar de pé: sem allowlist,
   quem abrir o site vê a tela de login e não passa.

### 4. Publicar o dado

```bash
python PGSemanal.py                  # fecha a semana e gera o dashboard
python tools/build_comite.py         # monta o deck do mês
python tools/publish_hub.py          # semanal + comite -> bucket privado
python tools/publish_hub.py estado   # backup da memória do pipeline
```

---

## O que nunca entra no repo

O repo é **público**. Barreira em três camadas:

| camada | onde | obrigatória? |
|---|---|---|
| `.gitignore` | raiz | não — só evita o `git add` distraído |
| hook `pre-commit` / `pre-push` | `.githooks/` | não — só em quem rodou `python tools/install_hooks.py` |
| Action `guarda` | `.github/workflows/` | **sim** — roda em todo PR e push na `main` |

A regra mora num lugar só: `tools/scan_segredos.sh`. O hook e a Action chamam o
mesmo script, então não existem duas versões podendo divergir.

**Rodar `python tools/install_hooks.py` em toda máquina.** Hook não vem no clone.

O que o scanner recusa, e por quê:

- `_cache/`, `bases/`, `dashboards/` — nome de animal, comprador, headcount
- `assets/semanal/`, `assets/comite/spec.*` — saída dos builds, dado do plantel
- `assets/comite/fotos/` — imagem da fazenda. **As fotos não ficam no repo**:
  entram reduzidas, como data URI, dentro do `spec.json`, que sai pelo bucket
  privado. O `deck.js` consome direto do spec, sem caminho de arquivo
- `*.xlsx`, `*.parquet`, `*.pptx`, `_docs/comite_conteudo.json`, `ComitêHPG/`
- `.env`, `*.local.sql`, chave privada, JWT com `role: service_role`
- e-mail `@luxor.com.br` real

---

## A memória do pipeline saiu do Git

`_cache/semanal_snapshots.json`, `headcount_history.json`, `paricoes_extra.json`
e `acumulado_piso.json` **não se reconstroem**: as planilhas do Drive são
sobrescritas toda semana, então snapshot perdido é dado perdido.

Enquanto o repo era privado, eles eram versionados. Agora não. O backup passou a
ser o bucket, sob prefixo `estado`, com policy de admin:

```bash
python tools/publish_hub.py estado
```

**Rodar isso toda semana, junto do fechamento.** Sem isso, formatar a máquina
apaga o histórico do plantel.

---

## Por que o dado não pode ser arquivo estático

O site é público. Qualquer coisa em `assets/` fica numa URL adivinhável, sem
login. Por isso o dado do plantel só existe:

1. no bucket **privado** `hpg-data`;
2. em memória no navegador de quem passou pela allowlist.

O `assets/auth.js` baixa o snapshot já autenticado e injeta no iframe:

- **semanal**: HTML entra por `srcdoc`
- **comitê**: `comite.html` é iframe de mesmo origin e lê o spec por
  `window.parent.HUB.comiteSpec`

O comitê usa `src` de verdade, e não `srcdoc`, porque com `srcdoc` o caminho fica
opaco e o `<script src="assets/comite/deck.js">` de dentro não resolve.

---

## Proteção da branch — depende de o repo ser público

Branch protection e ruleset **não funcionam em repo privado** sem GitHub Pro: a
API devolve `403 Upgrade to GitHub Pro or make this repository public`. Isso
força a ordem:

```bash
# 1. história já purgada localmente (git filter-repo)
git push --force-with-lease origin main

# 2. conferir na aba de commits do GitHub que xlsx/parquet/fotos sumiram

# 3. Settings > General > Change visibility  ->  público

# 4. só então:
python tools/proteger_main.py
```

O script aplica o mesmo par do P&C:

- **proteção clássica**: exige o check `guarda`, 1 aprovação, dismiss de review
  velha, review de code owner, sem force-push, sem deleção
- **ruleset `protege main`**: bloqueia deleção e non-fast-forward, exige PR com
  1 aprovação

Os dois com **bypass de admin de propósito**, igual ao P&C (`enforce_admins:false`
e `RepositoryRole 5`). Consequência assumida: push direto do admin não passa por
PR nem pela Action `guarda` — nesse caminho os hooks locais são a única barreira.
Mais uma razão para rodar `python tools/install_hooks.py` em toda máquina.

O script se recusa a rodar enquanto o repo for privado, explicando o motivo, em
vez de estourar um 403 cru.
