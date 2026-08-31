"""Página de auditoria do comitê para publicar como Artifact.

A versão anterior era escrita à mão e envelheceu: dizia "Build em 18/08" e listava
data de arquivo que já tinha mudado. O texto que NÃO envelhece — o histórico do que
estava errado e por quê — fica aqui como prosa fixa; tudo que é número ou data sai do
`assets/comite/spec.json`, que o build do comitê grava.

Divisão, então:
  - narrativa (o que estava errado, o que limita o deck): escrita, estável;
  - fontes, caminhos, datas, contagem de slides e pendências: geradas.

Uso: python tools/build_artifact_comite.py [destino]
"""
from __future__ import annotations

import html
import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SPEC = ROOT / "assets" / "comite" / "spec.json"
CONTEUDO = ROOT / "_docs" / "comite_conteudo.json"

TITULO = "Auditoria de Fontes do Comitê"

# Fonte mensal parada há mais de isto não descreve o mês que o deck publica.
DIAS_VELHA = 45

# Onde cada fonte mora, em uma palavra — o caminho completo já vai na coluna ao lado,
# e "Controladoria" vs "Haras" é o que responde de quem é o arquivo.
DONO = {
    "DRE histórico": "Controladoria",
    "inadimplência (KPI)": "Controladoria",
    "inadimplência (faixas)": "Controladoria",
}

# Slide -> (rótulo da fonte, aba, regra). Metadado ESTÁVEL, chaveado pelo número do
# slide, que não muda de mês para mês — a posição muda, porque tabela longa vira duas
# ou três páginas conforme o volume.
POR_SLIDE = {
    4:  ("DRE histórico", "Base DRE Geral",
         "CC = HPG, modelo Competência, mês da referência. Só linhas marcadas 'É Subtotal', que é o resumo que o comitê discute."),
    5:  ("DRE histórico", "Base DRE Geral",
         "Grupo = CUSTOS E DESPESAS OPERACIONAIS, aberto por natureza. Natureza zerada no mês fica de fora — encheria o slide de linha sem informação."),
    6:  ("DRE histórico", "Base DRE Geral",
         "Grupo = DESPESAS, mesma regra do slide de custos."),
    7:  ("DRE histórico", "Base YTD",
         "Acumulado do ano até o mês da referência, só subtotais. É o número que o comitê trimestral olha — por isso este slide fica nos dois decks."),
    8:  ("_docs/comite_conteudo.json", "comentarios",
         "Escrito à mão por mês. Não sai de planilha: é a leitura de quem fechou o mês sobre as variações do YTD."),
    9:  ("DRE anual (Haras)", "Investimentos",
         "Única parte do financeiro fora do histórico: a descrição de cada compra só existe no arquivo do ano. Lia a cópia em 'Ambiente de testes', parada em 18/03/2026; passou a ler o original em Relatórios Gerenciais."),
    10: ("DRE histórico", "Base DRE Geral",
         "CC = HPG, modelo Caixa, mês da referência. Regime de caixa, não competência."),
    11: ("base_bi.parquet", "fato_plantel",
         "Consolidado dos controles mensais do plantel: uma linha por animal por mês, com cota e avaliação."),
    12: ("mov_cascata.parquet", "—",
         "Movimentação do ano em cascata. Vem de outro repositório (LuxorMonthlyP-CRoutines): se aquele não rodar, este slide congela."),
    13: ("DRE histórico", "Base DRE Geral",
         "Organização FPG, modelo Caixa. Só linhas com valor no mês."),
    14: ("DRE histórico", "Base YTD",
         "Casa/FPG acumulado no ano. Também fica no deck trimestral."),
    16: ("estacao de monta", "ESTAÇÃO",
         "Funil da safra: coberturas, confirmados aos 60 dias, absorções e abortos. Absorção é perda antes dos 60d; aborto é embrião já confirmado."),
    17: ("estacao de monta", "GARANHOES",
         "Embriões por garanhão na safra corrente."),
    18: ("estacao de monta", "ESTAÇÃO",
         "Mesma contagem aplicada às safras anteriores, para comparar o ritmo."),
    19: ("estacao de monta", "PLANEJAMENTO",
         "Doadoras do time A: meta contra realizado. O time vem da coluna TIME da própria planilha."),
    20: ("estacao de monta", "PLANEJAMENTO", "Doadoras do time B, mesma regra."),
    21: ("coberturas de fora", "Planilha2",
         "Saldo de cobertura comprada ou de direito, por garanhão de fora."),
    23: ("_docs/comite_conteudo.json", "exposicoes.programacao",
         "Escrito à mão. A fonte declarada no próprio slide é o grupo da equipe mais o site da ABCCMM."),
    24: ("_docs/comite_conteudo.json", "exposicoes.resultados",
         "Escrito à mão, uma tabela por exposição."),
    29: ("mapa de vendas", "MAPA VENDAS",
         "Filtro do guia: vendedor CARLA, sem contrato cancelado. Meta anual de R$ 4,5M."),
    30: ("mapa de vendas", "MAPA VENDAS",
         "Mesmo filtro, aberto por mês e evento."),
    31: ("inadimplência (KPI)", "—",
         "Só agregados — nenhum nome de devedor entra no deck. A saída vem do repositório controle-de-inadimplencia, o mesmo que o hub do P&C lê; havia uma cópia no Drive e o build escolhe a mais recente."),
    32: ("embriões a entregar", "ENTREGAR",
         "Embrião vendido e ainda não gestado, com pagamento quitado ou em curso."),
    33: ("embriões a entregar", "ENTREGAR",
         "Mesma aba, recorte de pagamento pausado ou após confirmação."),
    34: ("embriões a entregar", "ENTREGAR",
         "Embrião de direito, troca ou reposição — não é venda."),
    35: ("embriões a entregar", "RECEBER",
         "Embrião que a PG comprou e ainda vai receber."),
    37: ("snapshot local", "semanal_snapshots.json",
         "Último fechamento semanal DO MÊS do deck. Era a aba CONTAGEM, que não tem dimensão de mês e trazia a contagem de hoje para qualquer deck."),
    38: ("_docs/comite_conteudo.json", "manejo", "Escrito à mão."),
    39: ("fotos do mês", "ATA & APRESENTACOES MENSAIS/<ano>/FOTOS",
         "12 por mês (2 slides), embutidas no spec. Mês vem da data no nome do arquivo do WhatsApp — que é a data do encaminhamento, não da foto: nenhum dos 80 arquivos tem EXIF e 9 pares são byte-idênticos com nomes de dias diferentes. Dedup por hash; com mais de 12 fotos, escolhe um dia por vez para o slide cobrir o mês. Antes vinha de 28 imagens extraídas do PPTX de junho e declaradas à mão, e TODO mês herdava as de junho."),
}

# O que cada fonte alimenta. Metadado estável: muda quando a origem muda.
ALIMENTA = {
    "DRE histórico": "Haras competência, custos, despesas, YTD, caixa e Casa/FPG",
    "DRE anual (Haras)": "Investimentos do ano, com a descrição de cada compra",
    "fotos do mês": "Registros de manejo do mês",
    "estacao de monta": "Embriões e prenhezes, garanhões, comparativo, doadoras A e B",
    "coberturas de fora": "Coberturas disponíveis por garanhão de fora",
    "mapa de vendas": "Resultado acumulado e detalhamento por mês e evento",
    "inadimplência (KPI)": "Inadimplências e recebíveis — agregados",
    "inadimplência (faixas)": "Inadimplências por faixa de atraso",
    "embriões a entregar": "Embriões vendidos a fazer, de direito e a receber",
}


def _idade(iso: str | None):
    if not iso:
        return None, "—"
    dias = (datetime.now() - datetime.fromisoformat(iso)).days
    return dias, f"{dias} dia{'s' if dias != 1 else ''}"


def _linha_fonte(rotulo: str, f: dict) -> str:
    dias, txt = _idade(f.get("modificado"))
    marca = (f'<span class="chip warn">{html.escape(txt)}</span>'
             if dias is not None and dias > DIAS_VELHA else html.escape(txt))
    quando = (f.get("modificado") or "—").replace("T", " ")
    return f"""        <tr>
          <td>{html.escape(ALIMENTA.get(rotulo, "—"))}</td>
          <td class="file">{html.escape(f.get("arquivo") or "—")}</td>
          <td class="tight">{html.escape(DONO.get(rotulo, "Haras"))}</td>
          <td class="file">{html.escape(f.get("caminho") or "—")}</td>
          <td class="num tight">{html.escape(quando)}</td>
          <td class="num tight">{marca}</td>
        </tr>"""


def _linha_slide(sl: dict, fontes: dict) -> str:
    """Uma linha por slide com dado: de onde vem, por qual regra, e se saiu.

    Mesmo formato da auditoria semanal — lá a chave é o indicador, aqui é o número
    do slide."""
    n = sl.get("n")
    rotulo, aba, regra = POR_SLIDE.get(n, ("—", "—", ""))
    pendente = sl.get("t") == "pendente"
    if pendente:
        situacao = '<span class="chip bad">pendente</span>'
        regra = f'<b>{html.escape(sl.get("motivo") or "")}</b><br>{html.escape(regra)}'
    else:
        situacao = '<span class="chip ok">com dado</span>'
        regra = html.escape(regra)
    # caminho e data vêm do que o build LEU; fonte escrita à mão não tem registro
    f = fontes.get(rotulo) or {}
    caminho = f.get("caminho") or rotulo
    dias, idade = _idade(f.get("modificado"))
    if dias is not None and dias > DIAS_VELHA:
        idade = f'<span class="chip warn">{html.escape(idade)}</span>'
    else:
        idade = html.escape(idade)
    return f"""        <tr>
          <td class="num">{n}</td>
          <td>{html.escape((sl.get("titulo") or "").split(" (")[0])}</td>
          <td>{situacao}</td>
          <td class="file">{html.escape(rotulo)}<br>
              <span class="aba">{html.escape(aba)}</span><br>
              <span class="cam">{html.escape(caminho)}</span></td>
          <td class="num tight">{idade}</td>
          <td class="obs">{regra}</td>
        </tr>"""


def _cartoes_limite(fontes: dict, pendentes: list, meses_conteudo: list) -> str:
    cartoes = []
    velhas = [(r, f) for r, f in fontes.items()
              if (_idade(f.get("modificado"))[0] or 0) > DIAS_VELHA]
    if velhas:
        itens = " · ".join(f"{html.escape(r)} {_idade(f['modificado'])[1]}"
                           for r, f in velhas)
        cartoes.append(f"""    <div class="card">
      <h3>Fonte parada há mais de {DIAS_VELHA} dias</h3>
      <p>O deck publica o mês corrente, mas estas fontes não foram atualizadas desde
      então — o número sai, e sai velho. Nenhum código conserta isso: alguém precisa
      rodar a rotina de origem.</p>
      <div class="figures">{itens}</div>
    </div>""")
    if pendentes:
        cartoes.append(f"""    <div class="card">
      <h3>Conteúdo manual do mês não escrito</h3>
      <p>Comentários do YTD, exposições, decisões de manejo e fotos não saem de
      planilha: são escritos a cada mês em <code>_docs/comite_conteudo.json</code>.
      Sem isso o slide vira pendência explícita — que é melhor que herdar o texto de
      outro mês, mas não preenche o deck.</p>
      <div class="figures">{len(pendentes)} pendente(s) · meses com conteúdo: {
          ", ".join(meses_conteudo) or "nenhum"}</div>
    </div>""")
    cartoes.append("""    <div class="card">
      <h3>Duas cópias do mesmo derivado</h3>
      <p>O <code>DRE_Historico.xlsx</code> é derivado, e o extractor grava a saída ao
      lado de si mesmo — duas cópias do extractor, duas saídas. O build escolhe a mais
      recente e diz no log qual usou, mas nada garante que as duas andem juntas. Uma
      saída única resolveria de vez.</p>
      <div class="figures">o log do build diz qual cópia venceu</div>
    </div>""")
    return "\n".join(cartoes)


# Prosa que não envelhece: o que já estava errado e foi corrigido. Fica escrita porque
# é história — nenhum build sabe disso.
HISTORICO = """<section>
  <h2>O que já estava errado</h2>
  <p class="lede">O deck monta sozinho e reportava 54 slides sem nenhuma pendência,
  então nada denunciava os problemas abaixo. Os quatro foram corrigidos; ficam
  registrados porque explicam por que o build hoje avisa em vez de calar.</p>

  <div class="cards">
    <div class="card">
      <h3>O deck lia a cópia velha do DRE</h3>
      <p>Duas cópias do extractor, duas saídas: a do repo de rotinas tinha julho
      fechado, a do Drive estava em 15/07 com julho zerado — e era a do Drive que o
      comitê lia. O build passou a escolher a mais recente e a dizer no log qual usou.</p>
      <div class="figures">Drive 15/07 → 18/08 · deck 06/2026 → 07/2026</div>
    </div>
    <div class="card">
      <h3>Um slide dizia junho e mostrava agosto</h3>
      <p>Lia a aba <code>CONTAGEM</code> do CONTROLE PLANTEL, que é retrato ao vivo e
      não tem dimensão de mês: o deck de junho exibia a contagem de 14/08. Agora usa o
      snapshot datado do fechamento semanal, e o subtítulo diz de qual data o número
      veio.</p>
      <div class="figures">antes 203 (100/44/1/58) · agora 206 (104/43/1/58)</div>
    </div>
    <div class="card">
      <h3>Julho não estava na base do plantel</h3>
      <p>O controle de julho estava no Drive desde 13/08 sem parquet extraído, então o
      <code>base_bi</code> parava em junho.</p>
      <div class="figures">29 → 30 meses · 391 animais em julho</div>
    </div>
    <div class="card">
      <h3>Rótulo de fonte adivinhado</h3>
      <p>O resolvedor decidia o rótulo pelo padrão do nome do arquivo: o mapa de vendas
      era registrado como "controle mensal". Não mudava número — só o comitê chama
      aquela função —, mas ia direto para esta página, que existe para dizer de onde
      veio o dado. Agora quem chama passa o rótulo.</p>
      <div class="figures">9 fontes com rótulo duplicado → 7 corretas</div>
    </div>
  </div>
</section>"""


def build(destino: Path | None = None) -> Path:
    if not SPEC.exists():
        raise SystemExit("rode tools/build_comite.py primeiro — falta spec.json")
    spec = json.loads(SPEC.read_text(encoding="utf-8"))
    mes = spec.get("padrao")
    rotulo_mes = (spec.get("labels") or {}).get(mes, mes or "")
    deck = (spec.get("decks") or {}).get(mes) or []
    fontes = spec.get("fontes") or {}
    pendentes = [s for s in deck if s.get("t") == "pendente"]
    com_fonte = [s for s in deck
                 if s.get("t") not in ("capa", "agenda", "divisor", "encerramento")]
    meses = spec.get("meses") or []

    # um slide por número: tabela longa vira '(1/3)', '(cont. 2/3)'… e todas as
    # partes têm a mesma fonte e a mesma regra
    por_slide, vistos = [], set()
    for x in deck:
        n = x.get("n")
        if x.get("t") in ("capa", "agenda", "divisor", "encerramento") or n is None:
            continue
        if n in vistos:
            continue
        vistos.add(n)
        por_slide.append(x)

    conteudo = []
    if CONTEUDO.exists():
        try:
            conteudo = [k for k in json.loads(CONTEUDO.read_text(encoding="utf-8"))
                        if re.fullmatch(r"\d{4}-\d{2}", k)]
        except Exception:
            conteudo = []

    gerado = datetime.now().strftime("%d/%m/%Y %H:%M")
    css = (ROOT / "tools" / "_artifact_comite.css")
    estilo = css.read_text(encoding="utf-8") if css.exists() else ESTILO

    corpo = f"""<title>{TITULO}</title>
<style>{estilo}</style>

<div class="wrap">

<header class="page">
  <span class="eyebrow">Haras Pao Grande &middot; comitê mensal</span>
  <h1>Auditoria de Fontes do Comitê</h1>
  <div class="runstamp">
    <span>Deck de <b>{html.escape(rotulo_mes)}</b></span>
    <span>{len(deck)} slides · {len(pendentes)} pendentes</span>
    <span>Meses no deck <b>{html.escape((meses[0] if meses else "") + " – " + (meses[-1] if meses else ""))}</b></span>
    <span>Gerada em <b>{gerado}</b></span>
  </div>
</header>

<section>
  <h2>De onde vem cada número</h2>
  <p class="lede">As {len(fontes)} fontes que o build abriu nesta rodada, com o caminho
  do arquivo e a data dele. Não é o que está escrito em documento: é o que o código
  leu. O caminho importa porque o mesmo nome de arquivo existe em mais de uma pasta —
  a estação de monta, por exemplo, troca de pasta quando a safra vira.</p>

  <div class="tiles">
    <div class="tile is-ok"><span class="num">{len(com_fonte) - len(pendentes)}</span><span class="lab">slides com dado, de {len(com_fonte)} com fonte</span></div>
    <div class="tile is-bad"><span class="num">{len(pendentes)}</span><span class="lab">pendentes — falta conteúdo ou fonte</span></div>
    <div class="tile is-warn"><span class="num">{len(fontes)}</span><span class="lab">arquivos lidos no build</span></div>
  </div>

  <div class="scroll">
    <table class="main">
      <thead>
        <tr><th>Alimenta</th><th>Arquivo</th><th>De quem</th><th>Caminho</th>
        <th class="num">Modificado</th><th class="num">Idade</th></tr>
      </thead>
      <tbody>
{chr(10).join(_linha_fonte(r, fontes[r]) for r in sorted(fontes))}
      </tbody>
      <tfoot>
        <tr><td colspan="6">Capa, agenda, divisores e encerramento não têm fonte de dado — são estrutura do deck.</td></tr>
      </tfoot>
    </table>
  </div>
</section>

<section>
  <h2>Slide a slide</h2>
  <p class="lede">Cada slide que carrega número: a fonte, a aba, o caminho do arquivo
  que o build abriu, a idade dele e a regra que produz o conteúdo. Capa, agenda,
  divisores e encerramento ficam de fora — são estrutura do deck, não têm dado.</p>
  <div class="scroll">
    <table class="main">
      <thead><tr><th class="num">#</th><th>Slide</th><th>Situação</th>
      <th>Fonte &middot; aba &middot; caminho</th><th class="num">Idade</th><th>Regra</th></tr></thead>
      <tbody>
{chr(10).join(_linha_slide(x, fontes) for x in por_slide)}
      </tbody>
    </table>
  </div>
</section>

<section>
  <h2>O que limita o deck hoje</h2>
  <p class="lede">Nenhuma é problema de código. São fontes que atrasam ou dependem de
  alguém escrever, cada uma travando uma parte diferente.</p>
  <div class="cards">
{_cartoes_limite(fontes, pendentes, conteudo)}
  </div>
</section>

{HISTORICO}

<footer class="page">
  Gerada por tools/build_artifact_comite.py a partir de assets/comite/spec.json — os
  números e as datas não são digitados.
</footer>

</div>
"""

    destino = destino or Path(os.getenv("TEMP") or "/tmp")
    destino.mkdir(parents=True, exist_ok=True)
    alvo = destino / "artifact_auditoria_comite.html"
    alvo.write_text(corpo, encoding="utf-8")
    print(f"[artifact] comitê: {alvo.stat().st_size // 1024} KB · {len(fontes)} fontes · "
          f"{len(pendentes)} pendentes -> {alvo}")
    return alvo


# Mesmo sistema visual da página do hub: o leitor reconhece as duas como a mesma
# auditoria, e não há motivo para inventar uma segunda identidade.
ESTILO = """
:root{--ground:#F4F6F3;--surface:#FFFFFF;--surface-alt:#ECEFEA;--line:#D5DBD3;--line-soft:#E4E8E2;
--ink:#1B211D;--ink-soft:#4E5852;--ink-mute:#77827B;--accent:#8C4A2F;--accent-soft:#F0E2DA;
--ok:#2E6F4E;--ok-bg:#E2EFE7;--warn:#8A6112;--warn-bg:#F4EBD6;--bad:#9C352F;--bad-bg:#F5E1DF;
--zebra:#FAFBF9;
--serif:"Iowan Old Style","Palatino Linotype",Palatino,"Book Antiqua",Georgia,serif;
--sans:system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;
--mono:ui-monospace,"SF Mono","Cascadia Mono",Menlo,Consolas,monospace}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){--ground:#121614;--surface:#1A1F1C;
--surface-alt:#222824;--line:#333B36;--line-soft:#272E2A;--ink:#E6EAE5;--ink-soft:#AFB8B1;
--ink-mute:#838D86;--accent:#CE8462;--accent-soft:#35231B;--ok:#6FBE8F;--ok-bg:#1B2E23;
--warn:#D6A947;--warn-bg:#322913;--bad:#E08078;--bad-bg:#33201E;--zebra:#1D2320}}
:root[data-theme="dark"]{--ground:#121614;--surface:#1A1F1C;--surface-alt:#222824;--line:#333B36;
--line-soft:#272E2A;--ink:#E6EAE5;--ink-soft:#AFB8B1;--ink-mute:#838D86;--accent:#CE8462;
--accent-soft:#35231B;--ok:#6FBE8F;--ok-bg:#1B2E23;--warn:#D6A947;--warn-bg:#322913;
--bad:#E08078;--bad-bg:#33201E;--zebra:#1D2320}
*{box-sizing:border-box}
body{background:var(--ground);color:var(--ink);font-family:var(--sans);font-size:16px;
line-height:1.55;margin:0;padding:0 20px 72px;-webkit-font-smoothing:antialiased}
.wrap{max-width:1240px;margin:0 auto}
header.page{display:flex;flex-direction:column;gap:12px;padding:52px 0 24px;
border-bottom:2px solid var(--ink)}
.eyebrow{font-family:var(--mono);font-size:11px;letter-spacing:.16em;text-transform:uppercase;
color:var(--accent)}
h1{font-family:var(--serif);font-size:clamp(2rem,5vw,2.9rem);line-height:1.08;font-weight:600;
margin:0;letter-spacing:-.01em;text-wrap:balance}
.runstamp{display:flex;flex-wrap:wrap;gap:6px 22px;font-family:var(--mono);font-size:12px;
color:var(--ink-mute)}
.runstamp b{color:var(--ink-soft);font-weight:600}
section{padding-top:44px}
h2{font-family:var(--serif);font-size:1.5rem;font-weight:600;margin:0 0 4px;
letter-spacing:-.01em;text-wrap:balance}
.lede{color:var(--ink-soft);max-width:72ch;margin:0 0 20px;font-size:.93rem}
.tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;
margin-bottom:22px}
.tile{background:var(--surface);border:1px solid var(--line);
border-top:3px solid var(--tile-hue,var(--ink-mute));padding:14px 16px 16px;display:flex;
flex-direction:column;gap:2px}
.tile .num{font-family:var(--mono);font-size:1.9rem;font-weight:600;line-height:1.1;
color:var(--tile-hue,var(--ink));font-variant-numeric:tabular-nums}
.tile .lab{font-size:.8rem;color:var(--ink-soft);line-height:1.4}
.tile.is-ok{--tile-hue:var(--ok)}.tile.is-warn{--tile-hue:var(--warn)}.tile.is-bad{--tile-hue:var(--bad)}
.scroll{overflow-x:auto;-webkit-overflow-scrolling:touch;border:1px solid var(--line);background:var(--surface)}
table{border-collapse:collapse;width:100%;font-size:.86rem}
table.main{min-width:1040px}table.legend{min-width:760px}
thead th{text-align:left;font-family:var(--mono);font-size:10.5px;letter-spacing:.1em;
text-transform:uppercase;color:var(--ink-mute);font-weight:600;padding:10px 13px;
background:var(--surface-alt);border-bottom:1px solid var(--line);white-space:nowrap}
tbody td{padding:9px 13px;border-bottom:1px solid var(--line-soft);vertical-align:top}
tbody tr:nth-child(even){background:var(--zebra)}
tbody tr:last-child td{border-bottom:none}
td.num,th.num{font-family:var(--mono);font-variant-numeric:tabular-nums;white-space:nowrap}
td.file{font-family:var(--mono);font-size:.76rem;line-height:1.45;word-break:break-word;
max-width:38ch}
td.tight{white-space:nowrap}
td.file .aba{color:var(--ink-mute)}
td.file .cam{display:block;color:var(--ink-mute);font-size:.68rem;margin-top:2px;
word-break:break-word}
td.obs{color:var(--ink-soft);font-size:.84rem;min-width:240px}
tfoot td{padding:9px 13px;color:var(--ink-mute);font-size:.8rem;background:var(--surface-alt)}
.chip{display:inline-block;font-family:var(--mono);font-size:10px;letter-spacing:.06em;
text-transform:uppercase;padding:3px 7px;border-radius:2px;white-space:nowrap;font-weight:600}
.chip.ok{color:var(--ok);background:var(--ok-bg)}
.chip.warn{color:var(--warn);background:var(--warn-bg)}
.chip.bad{color:var(--bad);background:var(--bad-bg)}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(270px,1fr));gap:14px}
.card{background:var(--surface);border:1px solid var(--line);padding:16px 18px 18px;
display:flex;flex-direction:column;gap:8px}
.card h3{margin:0;font-size:.95rem;font-weight:650;letter-spacing:-.005em}
.card p{margin:0;font-size:.87rem;color:var(--ink-soft)}
.card .figures{font-family:var(--mono);font-size:.8rem;background:var(--surface-alt);
padding:7px 10px;font-variant-numeric:tabular-nums}
code{font-family:var(--mono);font-size:.88em;background:var(--surface-alt);padding:1px 4px;
border-radius:2px}
td code,.card code{background:transparent;padding:0}
footer.page{margin-top:52px;padding-top:18px;border-top:1px solid var(--line);
font-family:var(--mono);font-size:11.5px;color:var(--ink-mute)}
@media (max-width:620px){
header.page{padding-top:32px}
section{padding-top:34px}
body{padding:0 12px 48px}
.tiles{gap:8px;margin-bottom:16px}
.tile{padding:11px 12px 13px}
.tile .num{font-size:1.6rem}
.cards{gap:10px}
.card{padding:13px 14px 15px}
footer.page{margin-top:36px;padding-top:14px}
}
"""


if __name__ == "__main__":
    build(Path(sys.argv[1]) if len(sys.argv) > 1 else None)
