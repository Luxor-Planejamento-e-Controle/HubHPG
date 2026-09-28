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
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SPEC = ROOT / "assets" / "comite" / "spec.json"

TITULO = "Auditoria de Fontes do Comitê"

# Fonte mensal parada há mais de isto não descreve o mês que o deck publica.
DIAS_VELHA = 45

# De quem é cada fonte — o caminho completo já vai na coluna ao lado; esta coluna
# responde a quem pedir quando o número sai errado.
DONO = {
    "DRE histórico": "P&C (extrator)",
    "DRE anual (Haras)": "Controladoria",
    "DRE anual (Casa)": "Controladoria",
    "Comentários do DRE (Trello)": "Controladoria",
    "resumo contábil (mapa)": "Controladoria",
    "resumo do plantel (aba Plantel do hub)": "Hub HPG (aba Plantel)",
    "plantel consolidado (base_bi)": "P&C (extrator)",
    "inadimplência (foto do mês)": "P&C (controle de inadimplência)",
    "inadimplência (KPI do dash)": "P&C (controle de inadimplência)",
    "conteúdo do hub (comite_conteudo)": "Haras (escrito no hub)",
    "fotos do mês": "Haras (escrito no hub)",
}

# O que cada fonte alimenta. Metadado estável: muda quando a origem muda.
ALIMENTA = {
    "DRE histórico": "Valores do resumo, das análises, do YTD, do caixa e da Casa/FPG",
    "DRE anual (Haras)": "Rótulos e ordem das linhas (abas Real x Orçado) e a aba Investimentos",
    "DRE anual (Casa)": "Rótulos e ordem das linhas da Casa/FPG",
    "Comentários do DRE (Trello)": "Comentários do mês, por categoria",
    "resumo do plantel (aba Plantel do hub)": "Movimentação do plantel e patrimônio do estoque",
    "resumo contábil (mapa)": "Reserva da movimentação para mês que o hub não fechou",
    "plantel consolidado (base_bi)": "Estoque: animais por categoria e valor médio",
    "estacao de monta": "Embriões e prenhezes, garanhões, comparativo e doadoras",
    "coberturas de fora": "Coberturas disponíveis por garanhão de fora",
    "mapa de vendas": "Resultado acumulado, média mensal e detalhamento por evento",
    "inadimplência (foto do mês)": "Painel de inadimplência: cartões, rosca e índice por ano",
    "inadimplência (KPI do dash)": "Conferência dos cartões contra o dash do hub P&C",
    "embriões a entregar": "Embriões vendidos a fazer, de direito e a receber",
    "contagem do fechamento": "Plantel por local (slide oculto)",
    "receptoras do fechamento": "Receptoras por local (slide oculto)",
    "conteúdo do hub (comite_conteudo)": "Exposições, manejo, pendências e comentários sem Trello",
    "fotos do mês": "Fotos e registros de manejo, por tema",
}


def _historica(rotulo: str) -> bool:
    """Master de safra encerrada: não muda mais, idade grande ali não é atraso."""
    return rotulo.startswith("estação ")


# Regra de cada slide, pelo TIPO (e, no DRE e nos contratos, pelo número que o
# build dá a cada recorte). Chavear pela posição não serve: tabela longa vira
# duas páginas, slide vazio some e o que vem depois muda de lugar.
DRE_POR_N = {
    4: ("DRE histórico", "Base DRE Geral + face Real x Orçado (Comp)",
        "Rótulo e ordem da face oficial, valor da base (CC HPG, Competência, mês), arredondado como o Excel. Desde agosto/2026 entram Outras Receitas com o Haras, Embriões Produzidos e Reavaliação de Plantel (abertura da Variação Patrimonial) e a abertura de Investimentos em máquinas com infraestrutura e animais e produtos."),
    5: ("DRE histórico", "Base DRE Geral · conferência na DRE-Compet",
        "CUSTOS TOTAIS no topo de cada página e os cinco subgrupos fixos. Desde agosto/2026, embaixo de cada subgrupo entram as naturezas que variaram pelo menos R$ 2 mil contra o orçado, para cima ou para baixo; até julho, a lista fixa do relatório."),
    6: ("DRE histórico", "Base DRE Geral · conferência na DRE-Compet",
        "Mesma regra, para as despesas: DESPESAS TOTAIS (Despesas e os dois arrendamentos), cinco subgrupos fixos e os dois arrendamentos em linhas próprias. D. Lúdia abre nas naturezas dele e Vassouras nos blocos dele (Volumes e Concentrados, Pessoal, Reprodução, Sanidade). Fecha com Resultado Operacional."),
    7: ("DRE histórico", "Base YTD + face Real x Orçado (Comp)",
        "Acumulado do ano, com as mesmas linhas do resumo e sem subtítulo. Oculto na apresentação, como no relatório; fica no deck trimestral."),
    10: ("DRE histórico", "Base DRE Geral + face Real x Orçado (Caixa)",
         "As 6 linhas do caixa do relatório (CC HPG, modelo Caixa)."),
    13: ("DRE histórico", "Base DRE Geral + face Real x Orçado (Casa)",
         "CC FPG, modelo Caixa. Marketing e Hospedagem Família entram quando têm valor; sem elas os detalhes não fecham com Despesas Gerais."),
    14: ("DRE histórico", "Base YTD + face Real x Orçado (Casa)",
         "Acumulado da Casa/FPG. Oculto, como no relatório."),
}
CONTRATOS_POR_N = {
    32: "Vendido e ainda a fazer, com pagamento quitado ou em curso.",
    33: "Vendido e ainda a fazer, com pagamento pausado, após confirmação ou a pagar.",
    34: "Reposição, ou a fazer com direito ou troca. A coluna PGTO mostra o pagamento.",
    35: "Aba RECEBER: embrião que a PG comprou e ainda vai receber.",
}
POR_TIPO = {
    "pendencias": ("conteúdo do hub (comite_conteudo)", "pendencias",
                   "O que ficou combinado na apresentação anterior. Só existe quando há item escrito; para escrever, o Editar da agenda abre o editor."),
    "lista_mes": ("DRE anual (Haras)", "Investimentos",
                  "Só o bloco Compra de Animais e Produtos, de janeiro até o mês do deck, um grupo por mês com o total do bloco. É a aba, que pode diferir da face do DRE."),
    "estoque": ("plantel consolidado (base_bi)", "fato_plantel",
                "Status PLANTEL, sufixo exato Da PG ou Outro. Categoria pela coluna CATEGORIA do controle, todas abertas. Valor médio sobre os animais avaliados desde agosto/2026. Patrimônio é o saldo final da movimentação do hub."),
    "matriz": ("resumo contábil (mapa)", "Resumo Contábil",
               "Reserva: o hub não tem nenhum mês do ano fechado, e a movimentação sai do mapa da controladoria."),
    "movimentacao": ("resumo do plantel (aba Plantel do hub)", "Resumo contábil",
                     "O resumo contábil da aba Plantel do hub HPG: controle do mês importado na aba e movimentações classificadas, com a mesma conta da tela. Só entra mês fechado no hub; o que faltar sai do Resumo Contábil do mapa."),
    "funil": ("estacao de monta", "ESTAÇÃO",
              "Funil da safra do mês do deck: tentativas, lavados positivos, prenhez aos 15, 30, 45 e 60 dias, abortos e confirmados."),
    "garanhoes": ("estacao de monta", "GARANHOES + ESTAÇÃO",
                  "Lavados e confirmados por garanhão. Quem tem tentativa na safra e não está na aba entra pela conta da ESTAÇÃO."),
    "comparativo": ("estacao de monta", "ESTAÇÃO + masters das safras antigas",
                    "Confirmados por mês da IA nas quatro últimas safras; a meta é confirmados sobre a META TOTAL do PLANEJAMENTO."),
    "doadoras": ("estacao de monta", "PLANEJAMENTO + REC. EMBR.",
                 "Meta contra realizado por doadora, com as colunas lidas pelo cabeçalho. Sem coluna TIME na safra, sai um slide só."),
    "coberturas": ("coberturas de fora", "Planilha2",
                   "Só saldo maior que zero, do maior para o menor. A leitura para em ARQUIVO MORTO."),
    "tabela": ("conteúdo do hub (comite_conteudo)", "exposicoes.programacao",
               "Escrito no hub. Evento com resultado no mês que sumiu da programação volta com a linha do mês anterior."),
    "resultados": ("conteúdo do hub (comite_conteudo)", "exposicoes.resultados",
                   "Escrito no hub, um slide por exposição. O subtítulo traz a data da programação."),
    "vendas_acum": ("mapa de vendas", "MAPA VENDAS",
                    "Mapa do fechamento do mês (a primeira versão gerada depois do fim do mês), vendedor CARLA, sem cancelado. Média mensal é o acumulado dividido pelos meses decorridos."),
    "vendas_mes": ("mapa de vendas", "MAPA VENDAS", "Mesmo filtro, aberto por mês e evento."),
    "inadimplencia": ("inadimplência (foto do mês)", "fato_titulos do fim do mês",
                      "Carteira CAR inteira desde agosto/2026, como o dash do hub P&C mostra (até julho, carteira Carla). O build confere os cartões contra o histórico de KPIs do dash. Só agregados."),
    "kpis_tabela": ("contagem do fechamento", "PLANTEL",
                    "Plantel por local no fechamento do mês. Oculto: o relatório não tem esse slide."),
    "manejo": ("conteúdo do hub (comite_conteudo)", "manejo",
               "Um slide por semestre, juntando o manejo escrito em todos os meses. Para cada mês vale o texto mais recente."),
    "fotos": ("fotos do mês", "fotos (bucket do hub)",
              "Por tema, quantos slides forem precisos. Tema digitado em caixa alta sai como título."),
}


def _meta(sl: dict):
    t, n = sl.get("t"), sl.get("n")
    if t == "dre":
        return DRE_POR_N.get(n, ("DRE histórico", "Base DRE Geral", ""))
    if t == "contratos":
        return ("embriões a entregar", "RECEBER" if n == 35 else "ENTREGAR", CONTRATOS_POR_N.get(n, ""))
    if t == "pendente":
        return (sl.get("fonte") or "—", "—", "")
    if t == "comentarios":
        if sl.get("origem") == "trello":
            return ("Comentários do DRE (Trello)", "card DRE Haras - <mês>",
                    "Comentário da controladoria no card do mês (quadro Fluxo de Caixa), uma faixa por categoria. O Δ de cada categoria sai da face Real x Orçado (Caixa), porque o comentário é sobre o caixa.")
        return ("conteúdo do hub (comite_conteudo)", "comentarios",
                "Escrito no hub. Vale só em mês sem comentário no Trello.")
    return POR_TIPO.get(t, ("—", "—", ""))


def _idade(iso: str | None):
    if not iso:
        return None, "—"
    dias = (datetime.now() - datetime.fromisoformat(iso)).days
    return dias, f"{dias} dia{'s' if dias != 1 else ''}"


def _velha(rotulo: str, f: dict) -> bool:
    dias = _idade(f.get("modificado"))[0]
    return dias is not None and dias > DIAS_VELHA and not _historica(rotulo)


def _linha_fonte(rotulo: str, f: dict) -> str:
    txt = _idade(f.get("modificado"))[1]
    marca = f'<span class="chip warn">{html.escape(txt)}</span>' if _velha(rotulo, f) else html.escape(txt)
    quando = (f.get("modificado") or "—").replace("T", " ")
    alimenta = ALIMENTA.get(rotulo) or ("Comparativo: safra encerrada" if _historica(rotulo) else "—")
    return f"""        <tr>
          <td>{html.escape(alimenta)}</td>
          <td class="file">{html.escape(f.get("arquivo") or "—")}</td>
          <td class="tight">{html.escape(DONO.get(rotulo, "Haras"))}</td>
          <td class="file">{html.escape(f.get("caminho") or "—")}</td>
          <td class="num tight">{html.escape(quando)}</td>
          <td class="num tight">{marca}</td>
        </tr>"""


def _linha_slide(sl: dict, fontes: dict) -> str:
    """Uma linha por slide com dado: de onde vem, por qual regra, e se saiu."""
    n = sl.get("n")
    rotulo, aba, regra = _meta(sl)
    if sl.get("t") == "pendente":
        situacao = '<span class="chip bad">pendente</span>'
        regra = f'<b>{html.escape(sl.get("motivo") or "")}</b>'
    elif sl.get("oculto"):
        situacao = '<span class="chip warn">oculto</span>'
        regra = html.escape(regra)
    else:
        situacao = '<span class="chip ok">com dado</span>'
        regra = html.escape(regra)
    f = fontes.get(rotulo) or {}
    caminho = f.get("caminho") or rotulo
    idade = _idade(f.get("modificado"))[1]
    idade = f'<span class="chip warn">{html.escape(idade)}</span>' if _velha(rotulo, f) else html.escape(idade)
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


def _cartoes_limite(fontes: dict, pendentes: list) -> str:
    cartoes = []
    velhas = [(r, f) for r, f in fontes.items() if _velha(r, f)]
    if velhas:
        itens = " · ".join(f"{html.escape(r)} {_idade(f['modificado'])[1]}" for r, f in velhas)
        cartoes.append(f"""    <div class="card">
      <h3>Fonte parada há mais de {DIAS_VELHA} dias</h3>
      <p>O deck publica o mês, mas estas fontes não mudaram desde então. O número
      sai, e sai velho; quem resolve é a rotina de origem.</p>
      <div class="figures">{itens}</div>
    </div>""")
    if pendentes:
        nomes = ", ".join((p.get("titulo") or "")[:40] for p in pendentes)
        cartoes.append(f"""    <div class="card">
      <h3>Slide sem conteúdo</h3>
      <p>Exposições, manejo e fotos são escritos no hub a cada mês. Sem isso o slide
      fica marcado como pendente em vez de herdar o texto de outro mês.</p>
      <div class="figures">{len(pendentes)} pendente(s): {html.escape(nomes)}</div>
    </div>""")
    dre, anual = fontes.get("DRE histórico"), fontes.get("DRE anual (Haras)")
    if dre and anual and dre.get("modificado") and anual.get("modificado"):
        atras = anual["modificado"] > dre["modificado"]
        situ = ("a base está atrás do arquivo da controladoria: rode o extrator"
                if atras else "a base está em dia com o arquivo da controladoria")
        cartoes.append(f"""    <div class="card">
      <h3>A base do DRE é derivada</h3>
      <p>O deck lê o <code>DRE_Historico.xlsx</code>, que o extrator gera a partir dos
      arquivos anuais. A controladoria corrige meses já fechados; o incremental do mês
      relê o período e regera as bases longas, e o build avisa quando o anual fica mais
      novo que a base.</p>
      <div class="figures">base {html.escape(dre["modificado"].replace("T", " "))} · anual {html.escape(anual["modificado"].replace("T", " "))} · {situ}</div>
    </div>""")
    cartoes.append("""    <div class="card">
      <h3>A categoria do estoque é a da planilha</h3>
      <p>O deck conta pela coluna CATEGORIA do controle de plantel. O relatório do
      haras reclassifica alguns animais à mão, sem regra na planilha; a divisão por
      categoria pode diferir enquanto o total, o patrimônio e o valor médio batem.</p>
      <div class="figures">jul/26 · total, patrimônio e valor médio batem</div>
    </div>""")
    return "\n".join(cartoes)


# Prosa que não envelhece: o que já estava errado e foi corrigido, e onde o
# relatório montado à mão diverge. Fica escrita porque é história; nenhum
# build sabe disso. Sem valor em reais nem nome de animal: este arquivo é público.
HISTORICO = """<section>
  <h2>O que já estava errado</h2>
  <p class="lede">Cada item foi corrigido e fica registrado porque explica uma regra
  que hoje pode parecer arbitrária. Os de agosto são da primeira rodada; os de
  setembro vieram da conferência do deck de agosto com o haras.</p>

  <div class="cards">
    <div class="card">
      <h3>Análise de custos contra a DRE-Compet</h3>
      <p>No relatório de agosto do haras as colunas estavam deslocadas. Conferido
      linha a linha, por grupo, subgrupo e natureza, o deck bate com a DRE-Compet.</p>
      <div class="figures">set/26 · 59 de 59 linhas batem</div>
    </div>
    <div class="card">
      <h3>Arredondamento diferente da planilha</h3>
      <p>O navegador arredondava o meio para cima e a planilha para longe do zero:
      valor negativo terminado em ,5 saía com 1 real de diferença, e variação que
      arredonda a zero perdia o sinal que a planilha mostra.</p>
      <div class="figures">set/26 · 76 de 76 valores das faces batem</div>
    </div>
    <div class="card">
      <h3>Deduções da Casa zeradas no acumulado</h3>
      <p>A face chama a linha de Deduções e a base de Deduções e Impostos. Sem o
      sinônimo, a linha saía zerada enquanto a Receita Líquida já vinha descontada.</p>
      <div class="figures">set/26</div>
    </div>
    <div class="card">
      <h3>Inadimplência de outra carteira</h3>
      <p>O painel filtrava a carteira da Carla. Desde agosto mostra a carteira CAR
      inteira, como o dash do hub P&amp;C, e o build confere os cartões contra o
      histórico de KPIs do dash.</p>
      <div class="figures">set/26 · julho mantém a regra com que foi apresentado</div>
    </div>
    <div class="card">
      <h3>Movimentação fora do hub</h3>
      <p>O slide lia o mapa da controladoria, e o fechamento do plantel já era feito
      na aba Plantel do hub. Agora roda o mesmo motor da aba sobre os dados do hub.</p>
      <div class="figures">set/26</div>
    </div>
    <div class="card">
      <h3>Comentários genéricos</h3>
      <p>O slide trazia um resumo digitado no hub. A controladoria já escreve todo mês,
      no Trello, a explicação de cada variação; o deck passou a ler o card do mês.</p>
      <div class="figures">set/26 · quadro Fluxo de Caixa</div>
    </div>
    <div class="card">
      <h3>ETL que não via correção retroativa</h3>
      <p>Com o mês já na base, o extrator do DRE respondia que não havia nada a
      atualizar, e o do fluxo de caixa trocava só o mês pedido. A reestruturação de
      investimentos que a controladoria fez em 25/09 em todos os meses de 2026 só
      entraria com rebuild completo. Os dois passaram a reler o período.</p>
      <div class="figures">set/26 · incremental igual ao rebuild completo, 0 diferenças</div>
    </div>
    <div class="card">
      <h3>O deck lia a cópia velha do DRE</h3>
      <p>Havia duas saídas do extrator e o comitê lia a do Drive, parada em julho sem
      realizado. Hoje há uma base só, e o build avisa quando o arquivo anual fica mais
      novo que ela.</p>
      <div class="figures">ago/26 · deck 06/2026 → 07/2026</div>
    </div>
    <div class="card">
      <h3>Contagem de hoje em deck de outro mês</h3>
      <p>O slide do plantel lia a aba CONTAGEM, que não tem dimensão de mês: o deck de
      junho mostrava a contagem de agosto. Passou a ler o fechamento do próprio mês.</p>
      <div class="figures">ago/26</div>
    </div>
    <div class="card">
      <h3>Fotos herdadas de junho</h3>
      <p>As fotos vinham de imagens extraídas do deck de junho, e todo mês herdava as
      mesmas. Hoje vêm do conteúdo de cada mês no hub, agrupadas por tema.</p>
      <div class="figures">ago/26</div>
    </div>
  </div>
</section>

<section>
  <h2>Onde o relatório do haras diverge</h2>
  <p class="lede">Conferido com a versão corrigida de agosto (v3). Nestes pontos o
  deck segue a fonte, e a diferença é do relatório montado à mão.</p>
  <div class="cards">
    <div class="card">
      <h3>Análise de custos e despesas</h3>
      <p>Colunas deslocadas: o valor de uma natureza aparece na linha de outra.</p>
      <div class="figures">fonte: DRE-Compet</div>
    </div>
    <div class="card">
      <h3>Acumulado do ano</h3>
      <p>Números de julho no acumulado de agosto. O deck usa as colunas YTD da face
      do mês.</p>
      <div class="figures">fonte: Real x Orçado (Comp), colunas YTD</div>
    </div>
    <div class="card">
      <h3>Contratos de embrião</h3>
      <p>Linhas repetidas no lugar de outras, que sumiram. Na planilha cada contrato
      aparece uma vez.</p>
      <div class="figures">fonte: EMBRIOES A ENTREGAR - A RECEBER</div>
    </div>
    <div class="card">
      <h3>Vendas</h3>
      <p>Julho fora do acumulado. O mapa do fechamento tem o mês completo.</p>
      <div class="figures">fonte: mapa de vendas do fechamento</div>
    </div>
    <div class="card">
      <h3>Movimentação e inadimplência</h3>
      <p>Saldo da movimentação deslocado de mês, e o painel de inadimplência rotulado
      31/07 com os números de 31/08.</p>
      <div class="figures">fonte: aba Plantel do hub · dash de inadimplência</div>
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
    ocultos = [s for s in deck if s.get("oculto")]
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
    <span>{len(deck)} slides · {len(ocultos)} ocultos · {len(pendentes)} pendentes</span>
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
  <p class="lede">Fontes que atrasam, conteúdo que depende de alguém escrever e regras
  da planilha que o relatório do haras não segue.</p>
  <div class="cards">
{_cartoes_limite(fontes, pendentes)}
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
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){color-scheme:dark;--ground:#121614;--surface:#1A1F1C;
--surface-alt:#222824;--line:#333B36;--line-soft:#272E2A;--ink:#E6EAE5;--ink-soft:#AFB8B1;
--ink-mute:#838D86;--accent:#CE8462;--accent-soft:#35231B;--ok:#6FBE8F;--ok-bg:#1B2E23;
--warn:#D6A947;--warn-bg:#322913;--bad:#E08078;--bad-bg:#33201E;--zebra:#1D2320}}
:root[data-theme="dark"]{color-scheme:dark;--ground:#121614;--surface:#1A1F1C;--surface-alt:#222824;--line:#333B36;
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
