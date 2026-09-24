"""Monta a especificação do Comitê Mensal HPG a partir das bases reais.

O deck é a saída, não a fonte: cada slide vira um objeto {t: <tipo>, ...} num
spec JSON, e daí saem as DUAS renderizações — o HTML (hub/comite.html) e o PPTX
(botão "Exportar PPTX"). Fonte única, dois formatos.

Mapa completo de slide × fonte: `_docs/COMITE_MAPEAMENTO.md`.

**Um deck por mês.** O DRE não é lido pela aba "Real x Orçado", que mostra só o
mês em que o operador deixou o arquivo — é lido da aba `DRE-Compet`, que tem
TODOS os meses lado a lado (a linha 6 marca, em cada bloco, o número do mês na
coluna do Realizado). Assim o deck tem seletor de mês e nunca fica preso num mês
velho. As bases não-DRE (plantel, estação, vendas) entram no mês que elas têm;
quando o mês pedido não existe na base, o slide vira `pendente` dizendo isso.

Uso:
    python tools/build_comite.py          # todos os meses com dado
    python tools/build_comite.py 06/2026  # só esse mês
"""
import json
import base64
import hashlib
import io
import re
import sys
import unicodedata
from datetime import date, datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
import requests
from dotenv import dotenv_values

# tools/ vive na raiz do repo desde a reorganizacao: a raiz E o site.
REPO = Path(__file__).resolve().parent.parent
HUB = REPO
OUT = HUB / "assets" / "comite"
sys.path.insert(0, str(REPO / "scripts"))

# Caminhos do Drive e helpers vêm do pipeline que já roda — não duplicar.
from PGSemanalReport import (                                    # noqa: E402
    DRIVE_ROOT, EMB_COMERCIAIS, ESTACAO_MONTA_BASE, MAPA_VENDAS_DIR,
    _controle_plantel, _estacao_dirs, _latest_by_yymmdd, _latest_estacao_master,
    _load, _norm, _s, _to_num, caminho_curto, headcount_de,
)
# os resolvedores de fonte compartilhados anotam ali o arquivo que escolheram
from PGSemanalReport import _FONTES_USADAS as _FONTES_COMPARTILHADAS   # noqa: E402

ROTINAS = Path(r"C:/Users/Arthur/repos/LuxorMonthlyP-CRoutines")
# Saída do extractor: ele grava AO LADO DE SI MESMO (ver DRE_HIST, abaixo).
DRE_DIR = ROTINAS / "DRE Data"
PLANTEL_DIR = ROTINAS / "PlantelHPG"

# Os xlsx anuais do DRE, na pasta que o próprio LxDREdataExtractor usa como fonte
# (BASE_REPORTS). Havia cópia deles em `Ambiente de testes`, e era de lá que o slide de
# Investimentos lia — a de 2026 estava parada em 18/03/2026, cinco meses atrás, sem
# nada sinalizando: `pend()` só dispara quando o arquivo SOME, e a cópia velha existe.
# `Ambiente de testes` está deprecated (26/08/2026): a Controladoria migrou para os
# repositórios do GitHub e o que ficou lá ninguém mais atualiza.
DRE_ANUAL_DIR = Path(
    r"G:/Drives compartilhados/Luxor Controladoria/Relatórios Gerenciais"
    r"/RELATORIOS - OPERAÇÃO HARAS E FAZENDA PG"
)


def _dre_anual(ano: int, entidade: str) -> Path:
    """xlsx anual do DRE, na VERSÃO mais nova. `entidade` é o sufixo do nome:
    'HPG - HARAS', 'FPG - CASA'.

    A controladoria versiona no nome — 'DRE 2026 HPG - HARAS.xlsx', depois ' v2',
    ' v3'. O caminho era fixo no nome sem sufixo, que é justamente o mais ANTIGO:
    em 21/09/2026 a pasta tinha v3 de 16/09 17:45 e o comitê lia a cópia de 11/09
    17:59. Escolhe pelo número de versão e, no empate, pelo mtime."""
    pasta = DRE_ANUAL_DIR / str(ano) / "Fluxo de Caixa e DRE"
    base = f"DRE {ano} {entidade}"
    cands = [f for f in pasta.glob(f"{base}*.xlsx") if not f.name.startswith("~$")]
    if not cands:
        return pasta / f"{base}.xlsx"      # inexistente: quem chama avisa

    def _versao(f: Path) -> tuple[int, float]:
        m = re.search(r"\bv(\d+)\b", f.stem, re.I)
        return (int(m.group(1)) if m else 0, f.stat().st_mtime)

    escolhido = max(cands, key=_versao)
    if len(cands) > 1:
        print(f"  [DRE] {len(cands)} versões de {base!r}; usando {escolhido.name} "
              f"(outras: " + ", ".join(sorted(f.name for f in cands
                                              if f != escolhido)) + ")")
    return escolhido


DRE_HARAS = _dre_anual(2026, "HPG - HARAS")
# NÃO é fonte do comitê — a seção CASA/FPG sai do DRE_Historico, igual ao resto do
# financeiro (ver _docs/COMITE_MAPEAMENTO.md, pendência 3). Fica declarada só para
# quem for procurar o arquivo do ano não concluir que ele foi esquecido.
DRE_CASA = _dre_anual(2026, "FPG - CASA")  # noqa: F401
BASE_BI = REPO / "bases" / "base_bi.parquet"

MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
         "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
ABR = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]

avisos = []
def aviso(m):
    if m not in avisos:
        avisos.append(m)
        print(f"  [aviso] {m}")


def _json_default(o):
    if isinstance(o, np.generic):
        return o.item()
    if isinstance(o, (datetime, date)):
        return o.isoformat()[:10]
    raise TypeError(f"tipo não serializável: {type(o).__name__}")


def num(v):
    if v is None or isinstance(v, str):
        return None
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return None if pd.isna(f) else f


def pend(n, titulo, sub, fonte, motivo, edita=None):
    """Slide sem conteúdo ainda.

    `edita` diz QUAL editor abre esse slide quando ele é de conteúdo humano
    (comentários, exposições, manejo, fotos). Sem isso o deck marcava o slide
    como `t="pendente"` e o botão Editar ficava apagado — o slide só virava
    editável DEPOIS de já ter conteúdo, que é o contrário do que se precisa:
    quem abre o deck vazio é justamente quem vai escrever."""
    s = {"t": "pendente", "n": n, "titulo": titulo, "sub": sub, "fonte": fonte, "motivo": motivo}
    if edita:
        s["edita"] = edita
    return s


def brl_curto(v):
    if v is None:
        return "—"
    s, a = ("-" if v < 0 else ""), abs(v)
    if a >= 1e6:
        return f"{s}R$ {a/1e6:.1f}M".replace(".", ",")
    if a >= 1e3:
        return f"{s}R$ {a/1e3:.0f}k"
    return f"{s}R$ {a:.0f}"


def brl_cheio(v):
    """R$ 102.500 — valor inteiro com milhar, como o relatório escreve."""
    if v is None:
        return "—"
    return ("-" if v < 0 else "") + "R$ " + f"{abs(v):,.0f}".replace(",", ".")


def brl_k(v):
    """R$30k / R$105k — a coluna VALOR dos contratos de embrião. Arredonda como o
    relatório (meio vai pro par: 22.500 sai R$22k, 51.750 sai R$52k)."""
    if not v:
        return "—"
    return ("-" if v < 0 else "") + f"R${round(abs(v) / 1000):.0f}k"


# ------------------------------------------------------------ nomes como o haras
# As planilhas guardam nome de animal e de pessoa em caixa alta e sem acento
# ("JUSTICA DA PAO GRANDE", "SERGIO MUNIZ"); o relatório da Ana escreve com acento
# e abrevia o sufixo do haras ("Damasco da PG"). O dicionário cobre as palavras
# que aparecem nos slides — palavra fora dele fica como veio, sem adivinhação.
ACENTOS = {
    "JUSTICA": "JUSTIÇA", "MUSICA": "MÚSICA", "BEGONIA": "BEGÔNIA", "TRES": "TRÊS",
    "CORACOES": "CORAÇÕES", "IMPERIO": "IMPÉRIO", "INVENCIVEL": "INVENCÍVEL",
    "XODO": "XODÓ", "FEITICO": "FEITIÇO", "LEGITIMO": "LEGÍTIMO", "PAVAO": "PAVÃO",
    "PETUNIA": "PETÚNIA", "ACAUA": "ACAUÃ", "CAPADOCIA": "CAPADÓCIA", "RETA": "RETÃ",
    "ATLANTICO": "ATLÂNTICO", "ATALNTICO": "ATLÂNTICO", "OPERA": "ÓPERA",
    "NEGOCIO": "NEGÓCIOS", "NEGOCIOS": "NEGÓCIOS", "LEILAO": "LEILÃO",
    "EMBRIAO": "EMBRIÃO", "EMBRIOES": "EMBRIÕES", "RECEPTORA": "RECEPTORA",
    "PECUARIA": "PECUÁRIA", "AGROPECUARIA": "AGROPECUÁRIA", "JOSE": "JOSÉ",
    "JOAO": "JOÃO", "SERGIO": "SÉRGIO", "TARCISIO": "TARCÍSIO", "VINICIUS": "VINÍCIUS",
    "JULIO": "JÚLIO", "NIOBIO": "NIÓBIO", "OASIS": "OÁSIS", "PLATAO": "PLATÃO",
    "PICANCO": "PICANÇO", "ANTONIO": "ANTÔNIO", "MARCIO": "MÁRCIO", "FABIO": "FÁBIO",
    "ROGERIO": "ROGÉRIO", "LUCIO": "LÚCIO", "LIDIA": "LÍDIA", "LUDIA": "LÚDIA",
    "CAMPEA": "CAMPEÃ", "CAMPEAO": "CAMPEÃO", "COBERTURA": "COBERTURA",
    "SEMEN": "SÊMEN", "ALCATEIA": "ALCATEIA", "ALCATÉIA": "ALCATEIA",
    "CANCELAMENTO": "CANCELAMENTO", "REPOSICAO": "REPOSIÇÃO", "CONFIRMACAO": "CONFIRMAÇÃO",
    "GARANHAO": "GARANHÃO", "MEAIPE": "MEAÍPE", "NACIONAL": "NACIONAL", "EXPOSICAO": "EXPOSIÇÃO",
}
PARTICULAS = {"DA", "DE", "DO", "DAS", "DOS", "E", "COM", "PARA", "EM", "A", "O", "AS", "OS",
              "NA", "NO", "NAS", "NOS", "À", "AO", "POR"}


def _acentua(txt: str) -> str:
    return " ".join(ACENTOS.get(p, p) for p in str(txt or "").split())


def nome_animal(txt: str, curto: bool = True) -> str:
    """Nome de animal em CAIXA ALTA, acentuado; `curto` troca 'DA PAO GRANDE' por
    'DA PG' (é como a Ana escreve nos slides de estação)."""
    s = _acentua(" ".join(str(txt or "").upper().split()))
    if curto:
        s = re.sub(r"\bDA PAO GRANDE\b", "DA PG", s)
    return s


def titulo_pt(txt: str) -> str:
    """'DAMASCO DA PAO GRANDE' -> 'Damasco da PG'; 'A DEFINIR' -> 'A Definir'."""
    s = _acentua(" ".join(str(txt or "").upper().split()))
    s = re.sub(r"\bPAO GRANDE\b", "PG", s)
    out = []
    for i, p in enumerate(s.split()):
        if p in ("PG", "MH", "EAO", "ABCCMM", "RJ", "CTE", "GTA", "DNA", "TI", "II", "III",
                 "IV", "V", "VI", "XV", "XVI", "XXIX", "XX", "XXX", "N19", "I"):
            out.append(p)
        elif i and p in PARTICULAS:
            out.append(p.lower())
        else:
            out.append("-".join(x[:1] + x[1:].lower() for x in p.split("-")))
    return " ".join(out)


def pessoa_curta(txt: str) -> str:
    """Contraparte do contrato como o relatório mostra: sem o percentual, no máximo
    três nomes e sem terminar em 'de/da'. Contrato com dois compradores ('25% FULANO
    / 25% BELTRANO') vira 'Fulano Sobrenome / Beltrano'."""
    s = str(txt or "").strip()
    if not s:
        return "—"
    s = re.split(r"\s+-\s+", s)[0]                   # 'MARCOS REZENDE - HARAS ...'
    partes = [re.sub(r"^\s*\d+([.,]\d+)?\s*%\s*", "", p).strip() for p in s.split("/")]
    partes = [p for p in partes if p]

    def encurta(p, n):
        ws = titulo_pt(p).split()[:n]
        while len(ws) > 1 and ws[-1].upper() in PARTICULAS:
            ws.pop()
        return " ".join(ws)
    if len(partes) > 1:
        return f"{encurta(partes[0], 2)} / {encurta(partes[1], 1)}"
    return encurta(partes[0], 3)


# ==================================================================== DRE
# Fonte: DRE_Historico.xlsx — a base consolidada que o LxDREdataExtractor gera e
# que o LuxorP&CHub já lê. É ela, e NÃO o "DRE 2026 HPG - HARAS.xlsx":
#   - o arquivo do ano tem só o mês em que o operador deixou a aba de resumo, e
#     em 03/08/2026 estava em fevereiro, com a coluna Orçado zerada;
#   - o histórico traz os 12 meses, Haras (HPG) e Casa (FPG), Competência e
#     Caixa, com Orçado e Realizado de verdade (jun/26 bate com o deck oficial).
# O DRE_Historico é DERIVADO: quem o gera é o LxDREdataExtractor, e ele grava o
# resultado AO LADO DE SI MESMO (OUTPUT_PATH = pasta do próprio script). Havia uma
# segunda cópia do extractor no Drive e, portanto, uma segunda saída; o build escolhia a
# mais recente das duas. Não escolhe mais — a do Drive é a saída de um extractor de
# 15/07, versão anterior à do repo, e a pasta está deprecated.
#
# Fonte única cobra um preço: ela pode estar velha sem parecer velha. Daí a guarda em
# `_dre_hist` — se o xlsx anual mudou DEPOIS do último rebuild, o fechamento novo ainda
# não entrou na base. Foi exatamente o caso em 26/08/2026: julho ajustado em 18/08 e a
# base do repo parada em 11/08, sem EMBRIÕES PRODUZIDOS (82.500) nem as baixas de
# estoque. O deck sairia com julho pela metade e sem reclamar de nada.
DRE_HIST = DRE_DIR / "DRE_Historico.xlsx"
_hist_cache = {}

# rótulo -> arquivo efetivamente aberto neste run. Vai para o spec e de lá para a
# auditoria: o que o build LEU, não o que alguém escreveu que ele lê.
_FONTES: dict = {}


def _registra(rotulo, caminho):
    """Anota a fonte e devolve o caminho, para caber em linha de chamada."""
    if caminho is not None:
        _FONTES[rotulo] = Path(caminho)
    return caminho


_dre_hist_cache = []


def _quando(p: Path) -> datetime:
    return datetime.fromtimestamp(p.stat().st_mtime)


def _dre_hist():
    """A base do repo de rotinas, avisando se ela ficou para trás dos xlsx anuais.

    O aviso sai UMA vez por run: a função é chamada tanto na checagem de existência
    quanto na leitura."""
    if _dre_hist_cache:
        return _dre_hist_cache[0]
    if not DRE_HIST.exists():
        return None
    dia = lambda p: _quando(p).strftime("%d/%m/%Y")
    # a base é derivada dos anuais; anual mais novo que ela = fechamento fora da base
    atrasos = [f for f in (DRE_HARAS, DRE_CASA)
               if f.exists() and _quando(f) > _quando(DRE_HIST)]
    if atrasos:
        aviso(f"DRE_Historico é de {dia(DRE_HIST)}, mais antigo que "
              + ", ".join(f"{f.name} ({dia(f)})" for f in atrasos)
              + f" — o fechamento novo NÃO está na base. Rode LxDREdataExtractor.py "
                f"em {DRE_DIR} e responda TUDO (o modo incremental não recalcula mês "
                f"que já existe; _rebuild.py não serve, derruba Base DRE Geral e Base YTD)")
    print(f"  [dre] base de {dia(DRE_HIST)} — {DRE_HIST}")
    _dre_hist_cache.append(_registra("DRE histórico", DRE_HIST))
    return DRE_HIST


def le_historico():
    """Base DRE Geral (mês) + Base YTD (acumulado), já filtradas para 2026."""
    if _hist_cache:
        return _hist_cache
    DRE_HIST = _dre_hist()
    if DRE_HIST is None:
        return {}
    geral = pd.read_excel(DRE_HIST, sheet_name="Base DRE Geral")
    ytd = pd.read_excel(DRE_HIST, sheet_name="Base YTD")
    for d in (geral, ytd):
        d["dt"] = pd.to_datetime(d["Data de Fechamento"])
        d["mes"] = d["dt"].dt.month
        d["ano"] = d["dt"].dt.year
    _hist_cache["geral"], _hist_cache["ytd"] = geral, ytd
    return _hist_cache


def meses_fechados(cc="HPG", modelo="Competência", ano=2026):
    """Mês só entra no deck quando ACABOU e tem realizado lançado.

    Duas condições, e a primeira não era necessária até 26/08/2026. O
    `_rebuild.py` corta o mês corrente (`r["data"] <= cutoff`); o `rebuild_all()`
    do próprio extractor, não — e é ele que precisa ser usado, porque o
    `_rebuild.py` só escreve a Base DRE. Resultado: rodar o rebuild no dia 26
    trouxe agosto com 26 dias de lançamento, e o deck padrão virou um mês pela
    metade. Mês corrente parece fechado porque tem realizado; só a data diz que
    não está."""
    h = le_historico()
    if not h:
        return []
    hoje = datetime.now()
    limite = 12 if ano < hoje.year else hoje.month - 1   # ano futuro não chega aqui
    g = h["geral"]
    g = g[(g["Centro de Custo"] == cc) & (g["Modelo"] == modelo) & (g["ano"] == ano)]
    lancados = g.groupby("mes")["Realizado"].apply(lambda x: x.abs().sum())
    return sorted(int(m) for m, s in lancados.items() if s and int(m) <= limite)


def pct(orc, real):
    """∆ % = (realizado − orçado) / |orçado|. Sem orçado não existe percentual."""
    if not orc:
        return None
    return (real - orc) / abs(orc)


def _linhas_dre(df, col_orc, col_real, so_subtotal=False, so_com_valor=False):
    """Três níveis, tirados de Grupo/Subgrupo/É Subtotal:
        0 = linha do GRUPO (Receita Bruta, Custos e Despesas, Resultado…)
        1 = SUBGRUPO (Volumoso e Concentrado, Despesas com Pessoal…)
        2 = natureza folha
    Antes tudo que era `É Subtotal` virava nível 0, e o slide saía com quase toda
    linha em dourado e negrito — sem hierarquia nenhuma pra ler.

    Nome repetido em grupos diferentes (Sanidade aparece em Custos e em Despesas)
    ganha o grupo entre parênteses; sem isso a mesma palavra aparecia duas vezes
    com números diferentes.
    """
    linhas = []
    for _, r in df.sort_values("Ordem").iterrows():
        sub = bool(r["É Subtotal"])
        if so_subtotal and not sub:
            continue
        orc, real = num(r[col_orc]) or 0.0, num(r[col_real]) or 0.0
        if so_com_valor and not orc and not real and not sub:
            continue
        nome = str(r["Natureza de Lançamento"]).strip()
        grupo = "" if pd.isna(r["Grupo"]) else str(r["Grupo"]).strip()
        subg = "" if pd.isna(r["Subgrupo"]) else str(r["Subgrupo"]).strip()
        nivel = 0 if (not subg or subg == nome and not grupo) else (1 if sub else 2)
        if subg == nome and grupo and grupo != nome:
            nivel = 1
        if not subg or grupo == nome:
            nivel = 0
        linhas.append({"nome": nome.title(), "nivel": nivel, "grupo": grupo.title(),
                       "v": [orc, real, (real - orc) / 1000.0, pct(orc, real)]})
    vistos = {}
    for l in linhas:
        vistos[l["nome"]] = vistos.get(l["nome"], 0) + 1
    for l in linhas:
        if vistos[l["nome"]] > 1 and l["grupo"] and l["grupo"] != l["nome"]:
            l["nome"] = f"{l['nome']} ({l['grupo']})"
        l["total"] = l["nivel"] == 0        # compat: o render antigo lia `total`
        del l["grupo"]
    return linhas


# ---------------------------------------------------------------- gabarito oficial
# As abas "Real x Orçado" das planilhas anuais são a FACE do relatório: a mesma
# lista de linhas, na mesma ordem, que a Ana levava ao comitê. O slide era montado
# com `so_subtotal=True` sobre a base long-format, e isso derrubava 12 linhas do
# resumo do Haras — entre elas "Baixa de Estoque por Venda" (-81.500) e "por
# Mortes e Doações" (-241.250), que o haras confere todo mês e que são naturezas,
# não subtotais.
#
# Agora a aba oficial dá a ORDEM e o RÓTULO, e a base long-format dá o VALOR (é
# ela que tem o histórico; a aba só traz o mês corrente e o YTD). Linha do
# gabarito que não existir na base sai do slide em vez de sair zerada, senão o
# deck inventaria uma linha que o relatório não tem.
_GABARITO_CACHE: dict = {}

# A aba escreve alguns rótulos de um jeito e a base de outro. Chave = rótulo
# oficial normalizado; valor = rótulo na base.
SINONIMOS_DRE = {
    "RECEITABRUTA": "RECEITA OPERACIONAL BRUTA",
    "RECEITAOPERACIONALLIQUIDA": "RECEITA OPERACIONAL LIQUIDA",
    "CUSTOSDEVENDA": "CUSTOS DE VENDAS",
    "CUSTOSEDESPESAS": "CUSTOS E DESPESAS OPERACIONAIS",
    "CUSTOS": "CUSTOS INDIRETOS DE PRODUÇÃO",
    "CLETADESEMEN": "COLETA DE SEMEN",
    "RESULTADOAPOSINVESTIMENTOS": "RESULTADO APÓS OS INVESTIMENTOS",
    "ARRENDAMENTODEPASTODLUDIA": "DESPESAS - ARRENDAMENTO D. LÚDIA - HARAS",
    "ARRENDAMENTOVASSOURAS": "DESPESAS - ARRENDAMENTO Vassouras - HARAS",
    # Casa/FPG: a face e a base batizam as receitas de formas diferentes
    "CASA": "RECEITAS - CASA",
    "RECEITASCOMLOCACAO": "LOCAÇÃO DA CASA",
    "RECEITASFINANCEIRAS": "RECEITAS ADM/FINANCEIRAS",
    "RECEITALIQUIDALOCACAO": "RECEITA LIQUIDA - LOCAÇÃO",
    "DESPESAS": "DESPESAS - GERAIS",
    # No modelo CAIXA a face chama de "Resultado" o que a base chama de "Fluxo
    # de Caixa" — mesma linha, nome diferente. Como alternativa (o Competência
    # tem "Resultado Operacional" de verdade), entra na lista de fallback.
}

# Alternativas testadas em ordem quando o rótulo da face não existe na base.
ALTERNATIVAS_DRE = {
    "RESULTADOOPERACIONAL": ["FLUXO DE CAIXA ANTES DOS INVESTIMENTOS"],
    "RESULTADOAPOSINVESTIMENTOS": ["FLUXO DE CAIXA APÓS OS INVESTIMENTOS",
                                   "FLUXO DE CAIXA LÍQUIDO APÓS INVESTIMENTOS"],
}

# Subtotais que a FACE imprime e a base não tem como linha própria: somam-se os
# componentes, que é o que a planilha faz na célula.
SOMAS_DRE = {
    "DEDUCOESECANCELAMENTOS": ["Cancelamentos", "Custos de Venda"],
    "DESPESASARRENDAMENTOS": ["Arrendamento de Pasto - D. Lúdia", "Arrendamento Vassouras"],
}


def _chave_dre(x) -> str:
    x = unicodedata.normalize("NFKD", str(x)).encode("ascii", "ignore").decode().upper()
    return re.sub(r"[^A-Z0-9]", "", x)


def gabarito(arquivo: Path, aba: str) -> list[str]:
    """Rótulos do resumo oficial, na ordem em que ele os imprime."""
    chave = (str(arquivo), aba)
    if chave in _GABARITO_CACHE:
        return _GABARITO_CACHE[chave]
    fora: list[str] = []
    try:
        d = pd.read_excel(arquivo, sheet_name=aba, header=None)
        for i in range(len(d)):
            nome = str(d.iat[i, 0]).strip()
            # as duas primeiras linhas são título e cabeçalho das colunas
            if i < 2 or nome in ("nan", "") or nome.lower().startswith("dre "):
                continue
            if nome not in fora:
                fora.append(nome)
    except Exception as exc:
        aviso(f"não deu pra ler o gabarito {arquivo.name}/{aba}: {exc!r} — "
              f"o resumo sai na ordem da base")
    _GABARITO_CACHE[chave] = fora
    return fora


def _na_ordem_oficial(linhas: list[dict], rotulos: list[str]) -> list[dict]:
    """Reordena e renomeia as linhas da base conforme o gabarito."""
    if not rotulos:
        return linhas
    por_chave: dict[str, dict] = {}
    # `_linhas_dre` desambigua nome repetido pondo o grupo entre parênteses
    # ("Sanidade (Despesas)" e "Sanidade (Custos E Despesas Operacionais)"). O
    # relatório oficial imprime só uma "Sanidade", a do bloco de Custos — então
    # o nome-base também vira chave, e quando há mais de um candidato vence o de
    # maior valor absoluto, que é o do bloco principal (Sanidade -23.143 contra
    # -5.249; Manutenção -84.875 contra -10).
    # Desempate entre "Sanidade (Despesas)" e "Sanidade (Custos ...)": vence o
    # SUBTOTAL, que é a linha que a face imprime. Só no empate entre dois
    # subtotais (ou dois detalhes) o maior valor decide — escolher pelo valor
    # sozinho elegia a natureza errada assim que o YTD trouxe mais linhas.
    def _peso(l):
        return (1 if l.get("nivel", 2) <= 1 else 0,
                max(abs(l["v"][0] or 0), abs(l["v"][1] or 0)))
    for l in linhas:
        por_chave.setdefault(_chave_dre(l["nome"]), l)
        base = re.sub(r"\s*\(.*\)\s*$", "", l["nome"])
        if base != l["nome"]:
            ch_base = _chave_dre(base)
            se_ja = por_chave.get(ch_base)
            if se_ja is None or _peso(l) > _peso(se_ja):
                por_chave[ch_base] = l
    fora = []
    for rot in rotulos:
        ch = _chave_dre(rot)
        l = por_chave.get(ch) or por_chave.get(_chave_dre(SINONIMOS_DRE.get(ch, "")))
        for alt in ALTERNATIVAS_DRE.get(ch, []):
            if l is not None:
                break
            l = por_chave.get(_chave_dre(alt))
        if l is None and ch in SOMAS_DRE:
            partes = [por_chave.get(_chave_dre(x)) or
                      por_chave.get(_chave_dre(SINONIMOS_DRE.get(_chave_dre(x), "")))
                      for x in SOMAS_DRE[ch]]
            partes = [x for x in partes if x]
            if partes:
                o = sum(x["v"][0] or 0 for x in partes)
                r = sum(x["v"][1] or 0 for x in partes)
                fora.append({"nome": rot, "nivel": 0, "total": True,
                             "v": [o, r, (r - o) / 1000.0, pct(o, r)]})
                continue
        if l is None:
            # O relatório imprime a linha mesmo zerada (ÓVULOS, REAVALIAÇÃO DO
            # PLANTEL e FORMAÇÃO DE PASTAGEM não tiveram movimento em agosto). A
            # base só guarda linha com valor, então a linha sai daqui com zero —
            # some do slide era pior: quem confere procura a linha e não acha.
            fora.append({"nome": rot, "nivel": 2, "total": False,
                         "v": [0.0, 0.0, 0.0, None]})
            continue
        # o rótulo que vale é o do relatório oficial, não o da base
        fora.append({**l, "nome": rot})
    return fora


def dre_mes(cc, modelo, ano, m, **kw):
    h = le_historico()
    if not h:
        return []
    g = h["geral"]
    df = g[(g["Centro de Custo"] == cc) & (g["Modelo"] == modelo)
           & (g["ano"] == ano) & (g["mes"] == m)]
    return _linhas_dre(df, "Orçado", "Realizado", **kw)


def dre_ytd(cc, modelo, ano, m, **kw):
    h = le_historico()
    if not h:
        return []
    y = h["ytd"]
    faixa = f"{m:02d}-Jan a {ABR[m-1]}"
    df = y[(y["Centro de Custo"] == cc) & (y["Modelo"] == modelo)
           & (y["ano"] == ano) & (y["Acumulado"] == faixa)]
    if df.empty:                       # a faixa é rotulada pelo mês final
        df = y[(y["Centro de Custo"] == cc) & (y["Modelo"] == modelo)
               & (y["ano"] == ano) & (y["mes"] == m) & (y["Acumulado"].str.startswith(f"{m:02d}-"))]
    return _linhas_dre(df, "Orçado YTD", "Realizado YTD", **kw)


# ------------------------------------------------ face do relatório (as linhas dela)
# O resumo do relatório da Ana não é a aba oficial inteira (48 linhas): é um
# recorte de 19, com três pesos de linha — grupo (preto), detalhe (cinza recuado)
# e resultado (faixa azul em negrito), fechando com a linha invertida em navy.
# Cada linha aponta para o rótulo da aba "Real x Orçado", que é de onde o VALOR
# sai (via _na_ordem_oficial, conferido 158/158 com a face em agosto/2026).
#   (rótulo no slide, rótulo na face oficial, estilo)
GAB_RESUMO = [
    ("Receita Bruta", "Receita Bruta", "grupo"),
    ("Venda de Produtos", "Venda de Produtos", "det"),
    ("Receitas Financeiras", "Receitas Financeiras", "det"),
    ("Deduções e Cancelamentos", "Deduções e Cancelamentos", "grupo"),
    ("Cancelamentos", "Cancelamentos", "det"),
    ("Custos de Venda", "Custos de Venda", "det"),
    ("Receita Líquida", "Receita operacional - Líquida", "banda"),
    ("Custos e Despesas", "Custos E Despesas", "banda"),
    ("Custos", "Custos", "det"),
    ("Despesas", "Despesas", "det"),
    ("Arrendamentos", "Despesas - Arrendamentos", "det"),
    ("Resultado Operacional", "Resultado Operacional", "banda"),
    ("Variação Patrimonial", "Variação Patrimonial", "grupo"),
    ("Baixas de Estoque", "Baixas De Estoque", "grupo"),
    ("Por Venda", "Baixa De Estoque Por Venda", "det"),
    ("Mortes e Doações", "Baixa De Estoque Por Mortes E Doações", "det"),
    ("Resultado Patrimonial", "Resultado Patrimonial", "banda"),
    ("Investimentos", "Investimentos", "grupo"),
    ("Resultado após Invest.", "Resultado após Investimentos", "fim"),
]
GAB_CAIXA = [
    ("Receita Bruta", "Receita Bruta", "grupo"),
    ("Receita Líquida", "Receita operacional - Líquida", "banda"),
    ("Custos e Despesas", "Custos E Despesas", "banda"),
    ("Resultado Operacional", "Resultado Operacional", "banda"),
    ("Investimentos", "Investimentos", "grupo"),
    ("Resultado após Invest.", "Resultado após Investimentos", "fim"),
]
GAB_CASA = [
    ("CASA", None, "rotulo"),
    ("Receita Bruta", "Receita Bruta", "grupo"),
    ("Deduções", "Deduções", "det"),
    ("Receitas Financeiras", "Receitas Financeiras", "det"),
    ("Receita Líquida — Locação", "Receita Liquida - Locação", "banda"),
    ("Despesas Gerais", "Despesas Gerais", "grupo"),
    ("Desp. com Pessoal", "Despesa Com Pessoal", "det"),
    ("Manutenção", "Manutenção ( Servicos E Materiais)", "det"),
    ("Contas de Consumo", "Contas De Consumo", "det"),
    ("Materiais de Consumo", "Materiais De Consumo", "det"),
    ("Desp. Administrativas", "Despesas Administrativas", "det"),
    ("Tributos", "Tributos", "det"),
    ("Resultado Operacional", "Resultado Operacional", "banda"),
    ("Investimentos", "Investimentos", "grupo"),
    ("Resultado após Investimentos", "Resultado após Investimentos", "fim"),
]
# Linhas de Despesas Gerais que a Ana não trazia (em julho estavam zeradas).
# Com valor no mês elas ENTRAM, como detalhe — senão os detalhes não fecham com
# o total do grupo (agosto: Marketing -1.100 e Hospedagem Família +2.400).
CASA_EXTRAS = [("Marketing", "Marketing"), ("Hospedagem Família", "Hospedagem Familia")]


def linhas_face(face: list[dict], gab: list, extras=None, depois_de=None) -> list[dict]:
    """Recorta a face oficial (saída de _na_ordem_oficial) no gabarito do slide."""
    por = {_chave_dre(l["nome"]): l for l in face}
    out = []
    for rot, rot_face, estilo in gab:
        if rot_face is None:
            out.append({"nome": rot, "estilo": estilo, "v": [None, None, None, None]})
            continue
        l = por.get(_chave_dre(rot_face))
        v = list(l["v"]) if l else [0.0, 0.0, 0.0, None]
        out.append({"nome": rot, "estilo": estilo, "v": v})
        if extras and rot == depois_de:
            for rot_x, face_x in extras:
                lx = por.get(_chave_dre(face_x))
                if lx and (abs(lx["v"][0] or 0) >= 0.5 or abs(lx["v"][1] or 0) >= 0.5):
                    out.append({"nome": rot_x, "estilo": "det", "v": list(lx["v"])})
    return out


# ------------------------------------------------------ análise de custos/despesas
# A lista é a do relatório dela: por subgrupo, as naturezas que ela acompanha,
# na ordem em que ela as apresenta — e não "tudo que teve valor". Duas páginas
# por tema, a linha TOTAL no topo de cada uma. Cada linha é casada na base pelo
# (Subgrupo, Natureza) — o nome repete entre blocos (Sanidade e Reprodução
# existem em Custos e em Despesas/Vassouras).
#   (rótulo, estilo, grupo da base, subgrupo, natureza)
_CUS = "CUSTOS E DESPESAS OPERACIONAIS"
_DES = "DESPESAS"
ANALISE_CUSTOS = [
    [("CUSTOS TOTAIS", "total", _CUS, "CUSTOS INDIRETOS DE PRODUÇÃO", "CUSTOS INDIRETOS DE PRODUÇÃO"),
     ("Volumoso e Concentrado", "sub", _CUS, "VOLUMOSO E CONCENTRADO", "VOLUMOSO E CONCENTRADO"),
     ("Manutenção de Pastagem", "folha", _CUS, "VOLUMOSO E CONCENTRADO", "MANUTENÇÃO DE PASTAGEM"),
     ("Ração", "folha", _CUS, "VOLUMOSO E CONCENTRADO", "RAÇÃO"),
     ("Feno e Alfafa", "folha", _CUS, "VOLUMOSO E CONCENTRADO", "FENO E ALFAFA"),
     ("Sanidade", "sub", _CUS, "SANIDADE", "SANIDADE"),
     ("Vacinas", "folha", _CUS, "SANIDADE", "VACINAS"),
     ("Internações e Tratamentos", "folha", _CUS, "SANIDADE", "INTERNAÇÕES E TRATAMENTOS"),
     ("Medic. Clínica", "folha", _CUS, "SANIDADE", "MEDICAMENTOS DE CLINICA"),
     ("Exames Clínicos", "folha", _CUS, "SANIDADE", "EXAMES CLINICOS"),
     ("Reprodução", "sub", _CUS, "REPRODUÇÃO", "REPRODUÇÃO"),
     ("Vet. Reprodução (Transf. Embrião)", "folha", _CUS, "REPRODUÇÃO",
      "VETERINARIOS DE REPRODUÇÃO - TRANSF. EMBRIÃO"),
     ("Fretes e Transportes (Reprodução)", "folha", _CUS, "REPRODUÇÃO", "FRETES E TRANSPORTES PARA REPRODUÇÃO"),
     ("Coleta Coberturas/Sêmen", "folha", _CUS, "REPRODUÇÃO", "COLETA DE COBERTURAS/SEMEN"),
     ("Medic. Reprodução", "folha", _CUS, "REPRODUÇÃO", "MEDICAMENTOS DE REPRODUÇÃO"),
     ("GTA e Exames (Reprodução)", "folha", _CUS, "REPRODUÇÃO", "GTA E EXAMES DE REPRODUÇÃO")],
    [("CUSTOS TOTAIS", "total", _CUS, "CUSTOS INDIRETOS DE PRODUÇÃO", "CUSTOS INDIRETOS DE PRODUÇÃO"),
     ("Pista", "sub", _CUS, "PISTA", "PISTA"),
     ("Com Exposição", "folha", _CUS, "PISTA", "COM EXPOSIÇÃO"),
     ("Consultoria Pista", "folha", _CUS, "PISTA", "CONSULTORIA PISTA"),
     ("Medicamento de Pista", "folha", _CUS, "PISTA", "MEDICAMENTOS E MATERIAL DE PISTA"),
     ("Dentista", "folha", _CUS, "PISTA", "DENTISTA"),
     ("CTE (Centro de Treinamento)", "folha", _CUS, "PISTA", "CTE (CENTRO DE TREINAMENTO)"),
     ("Suplemento Pista", "folha", _CUS, "PISTA", "SUPLEMENTO PISTA"),
     ("Frete de Pista", "folha", _CUS, "PISTA", "FRETE DE PISTA"),
     ("GTA e Exames de Pista", "folha", _CUS, "PISTA", "GTA E EXAMES DE PISTA"),
     ("Registros e Transf.", "sub", _CUS, "REGISTROS E TRANSFERENCIAS", "REGISTROS E TRANSFERENCIAS"),
     ("Técnico", "folha", _CUS, "REGISTROS E TRANSFERENCIAS", "TENICO"),
     ("ABCCMM (Associação)", "folha", _CUS, "REGISTROS E TRANSFERENCIAS", "ABCCMM (ASSOCIAÇÃO)")],
]
ANALISE_DESPESAS = [
    [("DESPESAS TOTAIS", "total", _DES, None, "DESPESAS"),
     ("Marketing", "sub", _DES, "MARKETING", "MARKETING"),
     ("Manutenção", "sub", _DES, "MANUTENÇÃO", "MANUTENÇÃO"),
     ("Veículos - Caminhão", "folha", _DES, "MANUTENÇÃO", "VEICULOS - CAMINHÃO"),
     ("Manutenção Instalações", "folha", _DES, "MANUTENÇÃO", "MANUTENÇÃO DE INSTALAÇÕES"),
     ("Veículos - Trator e Implementos", "folha", _DES, "MANUTENÇÃO", "VEÍCULOS - TRATOR E IMPLEMENTOS"),
     ("Manut. Máquinas e Equipamentos", "folha", _DES, "MANUTENÇÃO", "MANUTENÇÃO DE MÁQUINAS E EQUIPAMENTOS"),
     ("Consumo de Água e Luz", "sub", _DES, "CONSUMO DE ÁGUA E LUZ", "CONSUMO DE ÁGUA E LUZ"),
     ("Despesas com Pessoal", "sub", _DES, "DESPESAS COM PESSOAL", "DESPESAS COM PESSOAL"),
     ("Outras Desp. Pessoal", "folha", _DES, "DESPESAS COM PESSOAL", "OUTRAS DESPESAS COM PESSOAL"),
     ("Férias", "folha", _DES, "DESPESAS COM PESSOAL", "FÉRIAS"),
     ("Cesta Básica", "folha", _DES, "DESPESAS COM PESSOAL", "CESTA BÁSICA"),
     ("Rescisões", "folha", _DES, "DESPESAS COM PESSOAL", "RESCISÕES"),
     ("INSS", "folha", _DES, "DESPESAS COM PESSOAL", "INSS"),
     ("Uniformes/EPIs", "folha", _DES, "DESPESAS COM PESSOAL", "UNIFORMES E EPI´S"),
     ("Salário", "folha", _DES, "DESPESAS COM PESSOAL", "SALÁRIO"),
     ("FGTS", "folha", _DES, "DESPESAS COM PESSOAL", "FGTS")],
    [("DESPESAS TOTAIS", "total", _DES, None, "DESPESAS"),
     ("Desp. Administrativas", "sub", _DES, "DESPESAS ADMINISTRATIVAS", "DESPESAS ADMINISTRATIVAS"),
     ("Cobrança Adm. e Jurídica", "folha", _DES, "DESPESAS ADMINISTRATIVAS", "COBRANÇA ADM E JURÍDICA"),
     ("Custo RJ", "folha", _DES, "DESPESAS ADMINISTRATIVAS", "CUSTO RJ"),
     ("Despesas com Visitas", "folha", _DES, "DESPESAS ADMINISTRATIVAS", "DESPESAS COM VISITAS"),
     ("Arrendamento D. Lúdia", "sub", _DES, "DESPESAS - ARRENDAMENTO D. LÚDIA - HARAS",
      "DESPESAS - ARRENDAMENTO D. LÚDIA - HARAS"),
     ("Arrendamento de Pasto", "folha", _DES, "DESPESAS - ARRENDAMENTO D. LÚDIA - HARAS", "ARRENDAMENTO DE PASTO"),
     ("Arrendamento Vassouras", "sub", _DES, "DESPESAS - ARRENDAMENTO Vassouras - HARAS",
      "DESPESAS - ARRENDAMENTO Vassouras - HARAS"),
     ("Volumes e Concentrados", "folha", _DES, "VOLUMES E CONCENTRADOS", "VOLUMES E CONCENTRADOS"),
     ("Pessoal", "folha", _DES, "PESSOAL", "PESSOAL"),
     ("Reprodução", "folha", _DES, "REPRODUÇÃO", "REPRODUÇÃO"),
     ("Sanidade", "folha", _DES, "SANIDADE", "SANIDADE"),
     ("Resultado Operacional", "banda", "RESULTADO OPERACIONAL", None, "RESULTADO OPERACIONAL")],
]


def linhas_analise(ano: int, m: int, pagina: list) -> list[dict]:
    """Uma página da análise: cada linha do gabarito casada na base do mês."""
    h = le_historico()
    if not h:
        return []
    g = h["geral"]
    df = g[(g["Centro de Custo"] == "HPG") & (g["Modelo"] == "Competência")
           & (g["ano"] == ano) & (g["mes"] == m)]
    ch = lambda x: _chave_dre(x) if isinstance(x, str) else ""
    grupo_k = df["Grupo"].map(ch)
    sub_k = df["Subgrupo"].map(ch)
    nat_k = df["Natureza de Lançamento"].map(ch)
    out = []
    for rot, estilo, grupo, sub, nat in pagina:
        sel = (nat_k == _chave_dre(nat)) & (grupo_k == _chave_dre(grupo))
        if sub:
            sel &= sub_k == _chave_dre(sub)
        cand = df[sel]
        if len(cand) > 1:
            # A base traz o bloco de Manutenção duas vezes (o principal e um bloco
            # residual de -760). O relatório mostra o principal: o de maior valor.
            cand = cand.assign(_p=cand["Orçado"].abs().fillna(0) + cand["Realizado"].abs().fillna(0)) \
                       .sort_values("_p", ascending=False).head(1)
        if cand.empty:
            orc = real = 0.0
        else:
            r = cand.iloc[0]
            orc, real = num(r["Orçado"]) or 0.0, num(r["Realizado"]) or 0.0
        out.append({"nome": rot, "estilo": estilo,
                    "v": [orc, real, (real - orc) / 1000.0, pct(orc, real)]})
    return out


# =========================================================== Investimentos (S09)
def slide_investimentos(m, ano):
    if not DRE_HARAS.exists():
        return pend(9, f"INVESTIMENTOS — COMENTÁRIOS {ano}", "", DRE_HARAS.name, "arquivo não encontrado")
    import openpyxl
    # único slide que não sai do DRE_Historico; sem registrar, o caminho dele não
    # aparecia na auditoria — e foi por isso que a cópia de março passou meses despercebida
    _registra("DRE anual (Haras)", DRE_HARAS)
    wb = openpyxl.load_workbook(DRE_HARAS, data_only=True, read_only=True)
    ws = wb["Investimentos"]
    BLOCOS = ("INFRAESTRUTURA", "COMPRA DE ANIMAIS E PRODUTOS", "MÁQUINAS E EQUIPAMENTOS",
              "INSTALAÇÕES", "FORMAÇÃO DE PASTAGEM")
    meses, atual, bl = [], None, None
    for r in ws.iter_rows(values_only=True):
        a = str(r[0]).strip().upper() if r[0] is not None else ""
        b = str(r[1]).strip() if len(r) > 1 and r[1] is not None else ""
        v = num(r[2]) if len(r) > 2 else None
        if a.startswith("INVESTIMENTOS -"):
            nome = a.split("-", 1)[1].strip().split("/")[0].title()
            atual = {"mes": nome, "total": 0.0, "itens": []}
            meses.append(atual); bl = None
            continue
        if atual is None:
            continue
        if a in BLOCOS:
            bl = a
            if bl == "COMPRA DE ANIMAIS E PRODUTOS" and v:
                atual["total"] = v
            continue
        if bl == "COMPRA DE ANIMAIS E PRODUTOS" and v is not None:
            atual["itens"].append({"desc": desc_investimento(b, a), "valor": v})
    wb.close()
    # o slide é acumulado do ano: mostra de janeiro até o mês do deck
    idx = {nm.lower(): i + 1 for i, nm in enumerate(MESES)}
    meses = [x for x in meses if idx.get(x["mes"].lower(), 99) <= m]
    for x in meses:
        k = idx.get(x["mes"].lower(), 0)
        # rótulo do relatório: 'Jan/26'; o mês do deck vai na faixa dourada
        x["rotulo"] = f"{ABR[k-1]}/{str(ano)[2:]}" if k else x["mes"]
        x["atual"] = k == m
        if not x["itens"]:
            x["itens"] = [{"desc": "Sem compra de animais e produtos registrada no mês", "valor": 0.0}]
    return {"t": "lista_mes", "n": 9, "titulo": f"INVESTIMENTOS — COMENTÁRIOS {ano}",
            "sub": f"Compra de Animais e Produtos  ·  Janeiro a {MESES[m-1]}", "meses": meses}


def desc_investimento(desc: str, quem: str) -> str:
    """Linha da compra como o relatório escreve: 'Ref. Canc. 25% Nióbio da PG — IV
    Semana de Negócios PG (Vitor Bezerra de Menezes Picanço)'. Tira a data que a
    controladoria cola no fim e troca os separadores; o texto continua o dela."""
    d = " ".join(str(desc or "").split())
    # a controladoria cola no fim o mês de referência ('- AGOSTO/2026') e, às
    # vezes, a cláusula inteira do contrato; no slide a compra é uma linha só
    d = re.sub(r"(?i)\s*[-–]?\s*(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)[A-ZÇ]*/20\d\d\b.*$", "", d)
    d = re.sub(r"\s*-?\s*\d{2}/\d{2}/\d{4}\s*$", "", d)            # data no fim
    d = re.sub(r"(?i)\bREF\.?\s*CANCELAMENTO\b", "REF. CANC.", d)
    d = re.sub(r"(?i)^(COMPRA DE )?0(\d)\b", lambda mm: (mm.group(1) or "") + mm.group(2), d)
    partes = [x.strip() for x in re.split(r"\s+-\s+", d) if x.strip()]
    curto = partes[0] if partes else d
    for x in partes[1:]:
        if len(curto) + len(x) > 75:
            break
        curto += " - " + x
    t = titulo_pt(curto).replace(" X ", " × ").replace(" - ", " — ")
    t = " ".join(w.lower() if i and w.upper() in MINUSCULAS_DESC else w for i, w in enumerate(t.split()))
    t = re.sub(r"^Ref\. canc\.", "Ref. Canc.", t, flags=re.I)
    # nome inteiro, como o relatório; sequência longa de dígito (CPF/CNPJ colado
    # no nome do favorecido) nunca vai pro slide
    q = titulo_pt(re.sub(r"\d{6,}", "", str(quem or "")).split(" - ")[0]) if quem else ""
    return f"{t} ({q})" if q else t


# palavras que ficam em minúscula no meio da descrição (o resto segue o
# título, porque quase tudo ali é nome de animal ou de pessoa)
MINUSCULAS_DESC = PARTICULAS | {"PELO", "PELA", "PELOS", "PELAS", "QUE", "MAS", "FOI", "AQUI", "SEM",
                                "UM", "UMA", "SE", "SUA", "SEU", "FAZER", "DAR", "TINHAM", "ESTAVAM"}


# ============================================================ Plantel (S11/S12/S37)
# S11 — estoque em equinos. Regra do guia: status PLANTEL e sufixo EXATO
# 'DA PAO GRANDE' ou 'OUTRO'; variação com percentual ficava fora porque o
# animal dividido já aparece pela cota e contá-lo de novo o duplicaria.
SUFIXOS_S11 = ("DA PAO GRANDE", "OUTRO")
# O `E 100%` (100% do Eduardo) NÃO entra aqui. A mudança de 04/08/2026 que o
# incluiu vale só para o headcount da ATUALIZAÇÃO SEMANAL — confirmado pelo
# Arthur em 23/09/2026 —, e trazê-la pro comitê descolou o slide do histórico
# que o haras confere: julho saía 193 animais / R$ 18,44M contra os 189 /
# R$ 16,0M do relatório oficial da Ana. Sem o E 100%, bate na vírgula.
SUFIXOS_EXTRA_S11 = ()
_base_bi_cache = None


def base_bi():
    global _base_bi_cache
    if _base_bi_cache is None and BASE_BI.exists():
        d = pd.read_parquet(BASE_BI)
        d["mes"] = pd.to_datetime(d["mes_referencia"]).dt.strftime("%Y-%m")
        _base_bi_cache = d
    return _base_bi_cache


# A base guarda a categoria no singular e sem acento (EMBRIAO, GARANHAO); o
# relatório da Ana imprime no plural acentuado. E ela não lista a cauda inteira:
# o que fica abaixo das principais vira uma linha "Outros", senão o slide ganha
# oito linhas de uma unidade cada.
CATEGORIA_PLURAL = {
    "EMBRIAO": "Embriões", "POTRA": "Potras", "POTRO": "Potros",
    "GARANHAO": "Garanhões", "DOADORA": "Doadoras", "CASTRADO": "Castrados",
    "EGUA DE PISTA": "Éguas de Pista", "CAVALO DE PISTA": "Cavalos de Pista",
    "CAVALO": "Cavalos", "RECEPTORA": "Receptoras", "MATRIZ": "Matrizes",
    "APOSENTADO": "Aposentados", "POTRA DE PISTA": "Potras de Pista",
    "POTRO DE PISTA": "Potros de Pista",
}
CATEGORIAS_NA_FACE = 7


def _linhas_categoria(cat, total):
    """Linhas da tabela de categorias, no formato do relatório."""
    itens = list(cat.items())
    principais, cauda = itens[:CATEGORIAS_NA_FACE], itens[CATEGORIAS_NA_FACE:]
    linhas = [[CATEGORIA_PLURAL.get(_norm_nome_cat(k), str(k).title()), int(v),
               f"{v / total * 100:.0f}%"] for k, v in principais]
    resto = sum(int(v) for _, v in cauda)
    if resto:
        linhas.append(["Outros", resto, f"{resto / total * 100:.0f}%"])
    return linhas


def _norm_nome_cat(k):
    import unicodedata as _u
    return _u.normalize("NFKD", str(k)).encode("ascii", "ignore").decode().upper().strip()


def slide_estoque(m, ano):
    d = base_bi()
    if d is None:
        return pend(11, "ESTOQUE EM EQUINOS — FAZENDA PAO GRANDE", "", "bases/base_bi.parquet",
                    "rode python scripts/PGBaseBI.py")
    alvo = f"{ano}-{m:02d}"
    if alvo not in set(d["mes"]):
        ult = sorted(d["mes"].unique())[-1]
        aviso(f"base_bi vai até {ult} — S11 fica pendente nos meses seguintes")
        return pend(11, "ESTOQUE EM EQUINOS — FAZENDA PAO GRANDE", "", "bases/base_bi.parquet",
                    f"a base vai até {ult}; sem o mês {alvo}. Rode scripts/PGDataExtractor.py + PGBaseBI.py")
    x = d[(d["mes"] == alvo) & (d["status_plantel"] == "PLANTEL")
          & (d["sufixo_grupo"].isin(SUFIXOS_S11) | d["sufixo"].isin(SUFIXOS_EXTRA_S11))]
    patrim = float(x["patrimonio_proporcional"].sum())
    # Média sobre TODOS os animais do slide, não só os avaliados — é como o
    # relatório oficial calcula: julho dá R$ 366k (69,163M / 189), e dividindo
    # pelos 187 avaliados daria R$ 370k. Animal sem valor entra no denominador
    # porque ele existe no plantel; o que falta é a avaliação dele.
    medio = float(x["valor_100"].sum()) / len(x) if len(x) else 0.0
    # O patrimônio do cartão é o saldo do Resumo Contábil liberado — o mesmo
    # número do slide de movimentação (jul/26: R$ 15.970.552,61, "R$ 16,0M").
    # A soma do parquet (15,94M) sai do cálculo por cota e não do divulgado.
    rc = resumo_contabil(ano, m).get(m, {})
    if rc.get("saldo_fim"):
        patrim = rc["saldo_fim"]
    cat = x["categoria"].value_counts()
    # Cartões como os dela: só valor e rótulo, e o de animais com a regra do
    # sufixo embaixo. Quantos estão sem avaliação não aparece no slide — o valor
    # médio já é sobre todos (ver acima).
    return {"t": "estoque", "n": 11, "titulo": "ESTOQUE EM EQUINOS — FAZENDA PAO GRANDE",
            "sub": (f"Composição patrimonial do plantel  ·  {MESES[m-1].upper()} {ano}  ·  "
                    f"{len(x)} animais  ·  Status PLANTEL  ·  Sufixo: Da PG / Outros"),
            "kpis": [{"v": f"{len(x)}", "l": "Animais Ativos", "s": "DA PAO GRANDE + OUTROS",
                      "cor": "navy", "pt": 28},
                     {"v": brl_curto(patrim), "l": "Patrimônio HPG", "s": "", "cor": "ouro", "pt": 20},
                     {"v": brl_curto(medio), "l": "Valor Médio", "s": "", "cor": "azul", "pt": 24}],
            "rows": _linhas_categoria(cat, len(x))}


MOV_LINHAS = [("saldo_ini", "Saldo Inicial"), ("compra", "(+) Compras"), ("producao", "(+) Prod. Emb."),
              ("venda", "(-) Baixa Vendas"), ("morte", "(-) Baixa Mortes"), ("doacao", "(-) Doações"),
              ("reaval", "(±) Reavaliação"), ("saiu_controle", "(-) Saiu do Controle"),
              ("saldo_fim", "Saldo Final")]


# Resumo da movimentação = a aba `Resumo Contabil` do mapa de movimentações da
# Controladoria, que é o número LIBERADO (e o que o slide da Ana reproduz na
# vírgula: julho/26 fecha em R$ 15.970.552,61). A cascata calculada
# (mov_cascata.parquet) somava também o Eduardo e saía do divulgado.
MAPA_MOV_DIR = Path(r"G:\Drives compartilhados\Luxor Controladoria\Relatórios Gerenciais"
                    r"\RELATORIOS - OPERAÇÃO HARAS E FAZENDA PG\Posição Equinos"
                    r"\PLANTEL - Movimentações")
MOV_ROTULOS = {"SALDO INICIAL": "saldo_ini", "(+) COMPRAS": "compras",
               "(+) PRODUCAO EMBRIOES": "producao", "(-) BAIXA VENDAS": "vendas",
               "(-) BAIXA MORTES E DOACOES": "mortes", "(+/-) REAVALIACOES": "reaval",
               "SALDO FINAL": "saldo_fim"}
MOV_LINHAS_ANA = [("saldo_ini", "Saldo Inicial"), ("compras", "(+) Compras"),
                  ("producao", "(+) Prod. Emb."), ("vendas", "(-) Baixa Vendas"),
                  ("mortes", "(-) Baixa Mortes"), ("reaval", "(+/-) Reaval."),
                  ("saldo_fim", "Saldo Final")]


def _mapa_mov(ano: int, m: int):
    """Mapa daquele fechamento ('... (Jul 2026).xlsx'); sem ele, o mais novo do ano.
    Mês fechado lê o arquivo do próprio mês: o mais novo pode ter revisado o
    passado, e o deck de julho não muda porque agosto foi publicado."""
    pasta = MAPA_MOV_DIR / str(ano)
    if not pasta.exists():
        return None
    todos = [f for f in pasta.glob("*.xlsx") if "Movimenta" in f.name and not f.name.startswith("~$")]
    do_mes = [f for f in todos if f"({ABR[m-1]} {ano})" in f.name]
    cands = do_mes or todos
    return max(cands, key=lambda f: f.stat().st_mtime) if cands else None


def _rotulo(v) -> str:
    """Rótulo de linha comparável: sem acento, caixa alta, espaço simples. O
    `_norm` do pipeline só põe em caixa alta — 'TÍTULO' não casava com 'TITULO'
    e o slide caía na cascata calculada, somando o Eduardo."""
    t = unicodedata.normalize("NFKD", str(v or "")).encode("ascii", "ignore").decode().upper()
    return " ".join(t.split())


def resumo_contabil(ano: int, m: int) -> dict:
    """{'2026-01': {saldo_ini, compras, ...}} da aba Resumo Contabil."""
    f = _mapa_mov(ano, m)
    if f is None:
        return {}
    wb = _load(_registra("resumo contábil (mapa)", f))
    aba = next((s for s in wb.sheetnames if "RESUMO" in _norm(s)), None)
    out, col_mes = {}, {}
    if aba:
        for r in wb[aba].iter_rows(values_only=True):
            rot = _rotulo(r[1]) if len(r) > 1 else ""
            if rot == "TITULO":
                for j, c in enumerate(r):
                    k = next((i + 1 for i, a in enumerate(ABR) if _rotulo(c)[:3] == _rotulo(a)), None)
                    if k and j > 1:
                        col_mes[j] = k
                continue
            campo = MOV_ROTULOS.get(rot)
            if not campo:
                continue
            for j, k in col_mes.items():
                v = r[j] if j < len(r) else None
                if isinstance(v, (int, float)):
                    out.setdefault(k, {})[campo] = float(v)
    wb.close()
    return out


def slide_movimentacao(m, ano):
    rc = resumo_contabil(ano, m)
    meses = [k for k in range(1, m + 1) if k in rc]
    if meses:
        u = rc[meses[-1]]
        mes_nome = MESES[meses[-1] - 1]
        rows = [[rot] + [rc[k].get(campo, 0.0) for k in meses] for campo, rot in MOV_LINHAS_ANA]
        ab = ABR[meses[-1] - 1]
        return {"t": "movimentacao", "n": 12, "titulo": f"RESUMO DA MOVIMENTAÇÃO DO PLANTEL — {ano}",
                "sub": "Saldo mensal · Compras, produções, vendas e baixas",
                "kpis": [{"v": ("+" if u.get("producao", 0) > 0 else "") + brl_curto(u.get("producao", 0)),
                          "l": f"Produção Emb. {ab}", "s": f"{mes_nome} {ano}", "cor": "navy"},
                         {"v": brl_curto(u.get("vendas", 0)), "l": f"Baixa Vendas {ab}",
                          "s": f"{mes_nome} {ano}", "cor": "vinho"},
                         {"v": brl_curto(u.get("mortes", 0)), "l": f"Mortes/Doações {ab}",
                          "s": f"{mes_nome} {ano}", "cor": "vermelho"},
                         {"v": brl_curto(u.get("saldo_fim", 0)), "l": f"Saldo Final {ab}",
                          "s": "Haras PG", "cor": "azul"}],
                "cols": ["TÍTULO"] + [ABR[k - 1].upper() for k in meses], "rows": rows}
    f = PLANTEL_DIR / "mov_cascata.parquet"
    if not f.exists():
        return pend(12, f"RESUMO DA MOVIMENTAÇÃO DO PLANTEL — {ano}", "", f.name,
                    "rode o LxMovimentacao.py no repo LuxorMonthlyP-CRoutines")
    c = pd.read_parquet(f)
    a = c[(c["mes"] >= f"{ano}-01") & (c["mes"] <= f"{ano}-{m:02d}")].sort_values("mes")
    if a.empty:
        ult = c["mes"].max()
        return pend(12, f"RESUMO DA MOVIMENTAÇÃO DO PLANTEL — {ano}", "", f.name,
                    f"a cascata vai até {ult}; sem meses de {ano} até {ABR[m-1]}")
    cols = ["TÍTULO"] + [ABR[int(x.split("-")[1]) - 1].upper() for x in a["mes"]]
    rows = [[rot] + [a[k].tolist()[i] for i in range(len(a))] for k, rot in MOV_LINHAS]
    u = a.iloc[-1]
    return {"t": "matriz", "n": 12, "titulo": f"RESUMO DA MOVIMENTAÇÃO DO PLANTEL — {ano}",
            "sub": "Saldo mensal · compras, produções, vendas e baixas",
            "kpis": [{"v": brl_curto(u["producao"]), "l": "Produção Emb.", "s": cols[-1]},
                     {"v": brl_curto(u["venda"]), "l": "Baixa Vendas", "s": cols[-1]},
                     {"v": brl_curto(u["morte"] + u["doacao"]), "l": "Mortes/Doações", "s": cols[-1]},
                     {"v": brl_curto(u["saldo_fim"]), "l": "Saldo Final", "s": "Haras PG"}],
            "cols": cols, "rows": rows}


RE_PREFIXO_DATA = re.compile(r"^(\d{6})")


def _fechamento_do_mes(ano: int, m: int):
    """CONTROLE_DE_PLANTEL do FECHAMENTO daquele mês, em qualquer pasta de estação.

    O nome traz o mês de REFERÊNCIA (`..._AGO_26.xlsx`) e o prefixo é a data em que
    o arquivo foi gerado, já no mês seguinte. Cópia de trabalho ('EDITAR ...') fica
    de fora: ela é editada durante o mês seguinte e mistura dois meses — mesma
    regra do módulo de plantel (tools/seed_plantel_hub.py)."""
    alvo = _norm(f"_{ABR[m - 1]}_")
    padrao = re.compile(re.escape(alvo) + r"(20)?" + str(ano)[2:] + r"(?!\d)")
    cands = []
    for d in _estacao_dirs():
        for f in d.glob("*CONTROLE_DE_PLANTEL_PAO_GRANDE_*.xlsx"):
            n = _norm(f.name)
            if f.name.startswith("~$") or "EDITAR" in n or not padrao.search(n):
                continue
            cands.append(f)
    if not cands:
        return None
    # mais recente pela data no prefixo do nome, e depois por mtime — a mesma
    # regra de escolha de versão do módulo de plantel
    return max(cands, key=lambda f: ((RE_PREFIXO_DATA.match(f.name) or [""])[0]
                                     if RE_PREFIXO_DATA.match(f.name) else "",
                                     f.stat().st_mtime))


def _receptoras_do_fechamento(ref):
    """ARRENDAMENTOS E RECEPTORAS como estava quando aquele fechamento foi gerado.

    O arquivo de receptoras não traz mês de referência no nome, só o prefixo com a
    data em que foi salvo. A âncora, então, é o prefixo do próprio fechamento: vale
    a versão mais recente que NÃO seja posterior a ele. Sem isso, o deck de um mês
    passado contaria as receptoras de hoje contra o roster daquele mês.

    None quando não há candidato anterior — aí quem chama usa o resolvedor padrão."""
    pref = RE_PREFIXO_DATA.match(ref.name) if ref else None
    if not pref:
        return None
    limite = pref.group(1)
    cands = []
    for d in _estacao_dirs():
        for f in d.glob("*PLANTEL ARRENDAMENTOS E RECEPTORAS.xlsx"):
            mm = RE_PREFIXO_DATA.match(f.name)
            if f.name.startswith("~$") or not mm or mm.group(1) > limite:
                continue
            cands.append((mm.group(1), f))
    if not cands:
        return None
    return max(cands)[1]


def slide_contagem(m, ano):
    """S37 — contagem por local no FECHAMENTO do mês do deck.

    Passou por três fontes. A aba CONTAGEM do CONTROLE PLANTEL não serve: é um
    retrato AO VIVO, sem dimensão de mês — o deck de JUNHO/2026 exibia 203 animais
    (a contagem de 14/08) enquanto junho fechou com 206.

    Depois veio o snapshot do fechamento SEMANAL, pegando o último do mês. Também
    não é o fechamento: o último semanal de agosto/2026 é 28/08 e agosto fechou em
    31/08 — o que entrou e saiu no 29, 30 e 31 ficava de fora, e quem cobre o fim
    de agosto é o snapshot de 04/09, que um filtro por prefixo `2026-08` nunca
    alcança. Em junho/2026 o último semanal era 26/06. Só bate com o fechamento
    quando a sexta cai no último dia do mês, o que é coincidência de calendário.

    A fonte é o CONTROLE_DE_PLANTEL do mês (o que o mapeamento do comitê já
    dizia: S11, S12 e S37 saem dele), contado pelas MESMAS regras do fechamento
    — `headcount_de` é a função do semanal, com os arquivos daquele mês no lugar
    dos de hoje. Sem o arquivo do mês, a pendência é explícita, em vez de mostrar
    o número de outro mês."""
    src = _fechamento_do_mes(ano, m)
    if src is None:
        return pend(37, "PLANTEL — PAO GRANDE, ARRENDAMENTO E SÓCIOS",
                    f"{MESES[m-1].upper()} {ano}", "CONTROLE_DE_PLANTEL_PAO_GRANDE_*.xlsx",
                    f"nenhum fechamento do haras para {ABR[m-1]}/{str(ano)[2:]} nas "
                    f"pastas de estação (cópia 'EDITAR' não conta: mistura dois meses)")
    _registra("contagem do fechamento", src)
    rec = _receptoras_do_fechamento(src)
    if rec is not None:
        _registra("receptoras do fechamento", rec)
    # último dia do mês do deck: override registrado DEPOIS disso não vale aqui
    # (o de 04/09/2026 não pode apagar animal do fechamento de junho). O semanal
    # não passa data e segue com todos, como sempre.
    fim_do_mes = date(ano + m // 12, m % 12 + 1, 1) - timedelta(days=1)
    try:
        hc, _ = headcount_de(src, rec, fim_do_mes)
    except Exception as e:
        return pend(37, "PLANTEL — PAO GRANDE, ARRENDAMENTO E SÓCIOS",
                    f"{MESES[m-1].upper()} {ano}", src.name, f"não consegui contar: {e}")

    det = {k: v for k, v in (hc.get("detalhe") or {}).items() if k != "TOTAL GERAL"}
    ordem = [k for k in ("FAZENDA", "ARRENDAMENTO", "CTE", "SOCIO") if k in det]
    ordem += [k for k in det if k not in ordem]
    rows = [[k.title(), int(det[k]["animais"]), int(det[k]["receptoras"]),
             int(det[k]["total"])] for k in ordem]
    total = int(hc.get("total") or sum(r[-1] for r in rows))
    kp = [{"v": f"{r[-1]}", "l": r[0].upper(),
           "s": f"{r[-1]/total*100:.0f}% do total" if total else "—"} for r in rows[:3]]
    kp.append({"v": f"{total}", "l": "TOTAL GERAL", "s": "sob responsabilidade da PG"})
    return {"t": "kpis_tabela", "n": 37, "titulo": "PLANTEL — PAO GRANDE, ARRENDAMENTO E SÓCIOS",
            "sub": f"{MESES[m-1].upper()} {ano} · fechamento mensal do haras · {src.name}",
            "kpis": kp,
            "tabela": {"cols": ["LOCAL", "ANIMAIS", "RECEPTORAS", "TOTAL"], "rows": rows}}


# ============================================================ Estação (S16–S20)
def safra_do_deck(ano: int, m: int) -> str:
    """Safra da estação do MÊS DO DECK. A estação corre de agosto a julho (ver
    MESES_ESTACAO), então o deck de AGOSTO/2026 já é 2026/2027 e o de julho/2026
    ainda é 2025/2026.

    Vinha de SAFRA_ATUAL, que é calculado por `date.today()` (a virada de 26/27
    está cadastrada em 04/09/2026). Com isso o rótulo dependia de QUANDO o build
    rodou, não do mês do deck: o deck de agosto gerado em 20/08 saía 25/26, e
    reconstruir hoje o deck de JULHO sai 26/27."""
    return f"{ano}/{ano + 1}" if m >= 8 else f"{ano - 1}/{ano}"


def _master_da_safra(safra: str):
    """Master 'ESTACAO DE MONTA.xlsx' da pasta daquela safra.

    `_latest_estacao_master()` devolve a pasta mais nova por mtime — é o certo pro
    semanal, que quer HOJE, e o errado pro deck de um mês passado, que pediria a
    safra antiga dentro do arquivo da safra nova. Sem a pasta da safra, cai no
    comportamento antigo em vez de quebrar o deck inteiro."""
    a, b = safra.split("/")
    for nome in (f"Estação {a}-{b}", f"Estacao {a}-{b}"):
        d = ESTACAO_MONTA_BASE / nome
        if d.is_dir():
            cands = [f for f in d.glob("*ESTACAO DE MONTA.xlsx") if not f.name.startswith("~$")]
            if cands:
                return max(cands, key=lambda f: f.stat().st_mtime)
    return _latest_estacao_master()


def _estacao_wb():
    return _load(_latest_estacao_master())


def slides_estacao(safra: str):
    """S16 funil, S17 garanhões, S18 comparativo, S19/S20 doadoras A e B.

    A estação é da SAFRA — o mesmo conteúdo vale para qualquer mês DA MESMA safra.
    Definições (do guia): absorção = perda antes dos 60d; aborto = embrião
    confirmado que não nasceu; óbito = nasceu e morreu.
    """
    try:
        src = _master_da_safra(safra)
        wb = _load(src)
    except Exception as e:
        p = pend(16, "ESTAÇÃO DE MONTA — EMBRIÕES E PRENHEZES", "", "ESTACAO DE MONTA.xlsx",
                 f"não consegui abrir: {e}")
        return [p, dict(p, n=17, titulo="ESTAÇÃO DE MONTA — GARANHÕES"),
                dict(p, n=18, titulo="ESTAÇÃO DE MONTA — COMPARATIVO COM ANOS ANTERIORES"),
                dict(p, n=19, titulo="ESTAÇÃO DE MONTA — DOADORAS TIME A"),
                dict(p, n=20, titulo="ESTAÇÃO DE MONTA — DOADORAS TIME B")]
    out = [funil(wb, safra), garanhoes(wb, safra), comparativo(wb, safra)] + doadoras(wb, safra)
    wb.close()
    return out


def funil(wb, safra):
    """S16 — aba ESTAÇÃO. Colunas (1-based): 11 LAVADO, 13 15D, 14 30D, 15 45D,
    16 60D, 17 ABORTO, 36 ESTAÇÃO. Confirmado = lavado+ e 15d+ e (30/45/60 '+' ou
    vazio), menos aborto=SIM."""
    ws = wb["ESTAÇÃO"]
    tent = lav = p15 = p30 = p45 = p60 = ab = 0
    for i, r in enumerate(ws.iter_rows(values_only=True), 1):
        if i < 3 or r[0] is None or _s(r[35]) != safra:
            continue
        tent += 1
        if _norm(r[10]) != "+":
            continue
        lav += 1
        # o funil é ENCADEADO: só chega em 30d quem passou em 15d, e assim por
        # diante. Contar cada coluna solta dava 85 confirmados contra 56 do
        # relatório oficial, porque a linha que nunca chegou aos 60 dias tem a
        # célula vazia e entrava como se tivesse passado.
        if _norm(r[12]) != "+":
            continue
        p15 += 1
        if _norm(r[13]) not in ("+", ""):
            continue
        p30 += 1
        if _norm(r[14]) not in ("+", ""):
            continue
        p45 += 1
        if _norm(r[15]) not in ("+", ""):
            continue
        p60 += 1
        if _norm(r[16]) == "SIM":     # aborto só conta depois de confirmado >60d
            ab += 1
    conf = p60 - ab
    ref = lambda v: f"{v/lav*100:.0f}% dos lavados" if lav else "—"
    curta = f"{safra[2:4]}/{safra[-2:]}"
    taxa = f"{lav/tent*100:.0f}%" if tent else "—"
    return {"t": "funil", "n": 16, "titulo": f"ESTAÇÃO DE MONTA {safra} — EMBRIÕES E PRENHEZES",
            "sub": (f"{conf} embriões confirmados  ·  Taxa recuperação: {taxa}  ·  {tent} tentativas"
                    if tent else "sem tentativas na safra"),
            "kpis": [{"v": str(conf), "l": "EMBRIÕES CONF.", "s": f"Estação {curta}", "cor": "navy"},
                     {"v": str(lav), "l": "LAVADOS +", "s": f"{taxa} de positivos", "cor": "azul"},
                     {"v": taxa, "l": "TAXA RECUP.", "s": curta, "cor": "ouro"},
                     {"v": str(ab), "l": "ABORTOS", "s": "Confirmados >60d", "cor": "vinho"}],
            "cab": f"FUNIL DE PRENHEZ — ESTAÇÃO {curta}",
            "rows": [["Tentativas", tent, "100%"],
                     ["Lavados (+)", lav, taxa],
                     ["Prenhez 15d", p15, ref(p15)], ["Prenhez 30d", p30, ref(p30)],
                     ["Prenhez 45d", p45, ref(p45)], ["Prenhez 60d", p60, ref(p60)],
                     ["(−) Abortos", ab, ">60 dias confirmados"],
                     ["Confirmados", conf, ref(conf)]]}


def garanhoes(wb, safra):
    """S17 — aba GARANHOES: 3 garanhão, 4 tipo de sêmen, 5 total lavados,
    6 lavados positivos, 7 %, 8 embriões confirmados, 9 prenhez, 10 aborto,
    11 total confirmados."""
    ws = wb["GARANHOES"]
    # no fim da aba há um bloco de legenda com os tipos de sêmen — ele entrava na
    # lista como se fosse garanhão ("Fresco", "Refrigerado", "Congelado"). O
    # bloco guarda também a REFERÊNCIA de cada tipo (coluna D: 60% refrigerado,
    # 69% fresco), que é o "ref.:" dos cartões do relatório.
    LEGENDA = {"FRESCO", "REFRIGERADO", "CONGELADO", "TIPO DE SEMEN", "TIPO DE SÊMEN"}
    aba = {}
    for i, r in enumerate(ws.iter_rows(values_only=True), 1):
        if i < 4 or len(r) < 11 or r[2] is None:
            continue
        nome = _s(r[2])
        n = _norm(r[2])
        if not nome or n.startswith("TOTAL") or n in LEGENDA:
            continue
        aba[_chave_garanhao(n)] = {"nome": nome_animal(nome, curto=False), "t": (_s(r[3]) or "")[:1].upper(),
                                   "lav": int(_to_num(r[4]) or 0), "pos": int(_to_num(r[5]) or 0),
                                   "conf": int(_to_num(r[10]) or 0)}
    # A aba é mantida à mão e esquece garanhão: em 25/26 faltavam Latino, Esteio
    # e Invencível, que o relatório trouxe da aba ESTAÇÃO. Mesma regra aqui —
    # quem teve tentativa na safra e não está na aba entra com a conta da ESTAÇÃO.
    for k, d in _garanhoes_da_estacao(wb, safra).items():
        if k not in aba or not aba[k]["lav"]:
            if not d["lav"]:
                continue
            base = aba.get(k, {"t": ""})
            # o nome vem da ESTAÇÃO junto com os números (a aba grafa 'MARADA')
            aba[k] = {**base, "nome": nome_animal(d["nome"], curto=False),
                      "lav": d["lav"], "pos": d["pos"], "conf": d["conf"]}
    rows = [x for x in aba.values() if x["lav"]]
    for x in rows:
        if not x["t"]:
            x["t"] = TIPO_SEMEN.get(_chave_garanhao(_norm(x["nome"])), "")
    rows.sort(key=lambda x: (-x["conf"], -x["lav"]))
    tent = sum(x["lav"] for x in rows)
    pos = sum(x["pos"] for x in rows)
    conf = sum(x["conf"] for x in rows)
    por_tipo = {}
    for x in rows:
        if x["t"]:
            a, b = por_tipo.get(x["t"], (0, 0))
            por_tipo[x["t"]] = (a + x["pos"], b + x["lav"])
    nome_tipo = {"R": "REFRIGERADO", "C": "CONGELADO", "F": "FRESCO"}
    cores = {"R": "navy", "C": "azul", "F": "ardosia"}
    kpis = [{"v": f"{p/t*100:.0f}%", "l": nome_tipo[k], "cor": cores[k], "s": REF_SEMEN[k]}
            for k in ("R", "C", "F") if k in por_tipo and por_tipo[k][1]
            for p, t in [por_tipo[k]]]
    fun = funil(wb, safra)
    tent_s, lav_s = fun["rows"][0][1], fun["rows"][1][1]
    return {"t": "garanhoes", "n": 17, "titulo": f"ESTAÇÃO DE MONTA {safra} — GARANHÕES",
            "sub": (f"{tent_s} tentativas  ·  {lav_s} lavados positivos "
                    f"({lav_s/tent_s*100:.0f}%)  ·  {fun['rows'][-1][1]} embriões confirmados"
                    if tent_s else f"{tent} lavados  ·  {conf} embriões confirmados"),
            "kpis": kpis, "rows": rows}


# Referência de aproveitamento por tipo de sêmen, como o relatório imprime em
# cada cartão. Não sai de planilha: a coluna que a trazia na aba GARANHOES sumiu
# na cópia de 01/08/2026 (ficou só a taxa realizada).
REF_SEMEN = {"R": "ref.: 60%", "C": "ref.: 50-60%", "F": "ref.: 70%"}
# Tipo de sêmen de garanhão que não está na aba GARANHOES da safra (a aba é
# mantida à mão). Os tipos são os que o relatório de jul/2026 atribui a eles;
# a aba 26/27 confirma Latino (fresco) e Xodó (congelado).
TIPO_SEMEN = {"LATINO PAO GRANDE": "F", "DAMASCO PAO GRANDE": "F", "XODO PORTEIRA AZUL": "C", "ESTEIO TRES CORACOES": "C",
              "INVENCIVEL LUA PRATA": "C", "ENCANTADO AGROTEXAS": "C", "ATREVIDO MORADA NOVA": "R",
              "FUTURO MYLA": "C"}


def _chave_garanhao(n: str) -> str:
    """Chave de nome de garanhão entre as abas: a mesma égua aparece como
    'FAVACHO ALCATÉIA' numa e 'FAVACHO ALCATEIA' noutra, 'ESTEIO TRES CORAÇÕES'
    e 'ESTEIO DE TRES CORACOES', 'MARADA'/'MORADA'. Tira acento e partícula."""
    s = unicodedata.normalize("NFKD", str(n)).encode("ascii", "ignore").decode().upper()
    s = s.replace("MARADA", "MORADA")
    return " ".join(p for p in re.findall(r"[A-Z0-9]+", s) if p not in PARTICULAS)


def _garanhoes_da_estacao(wb, safra: str) -> dict:
    """Tentativas, lavados + e confirmados por garanhão, pela aba ESTAÇÃO — com o
    mesmo funil encadeado do slide de prenhezes."""
    out = {}
    ws = wb["ESTAÇÃO"]
    for i, r in enumerate(ws.iter_rows(values_only=True), 1):
        if i < 3 or r[0] is None or len(r) < 36 or _s(r[35]) != safra or not r[3]:
            continue
        k = _chave_garanhao(_norm(r[3]))
        d = out.setdefault(k, {"nome": _s(r[3]), "lav": 0, "pos": 0, "conf": 0})
        d["lav"] += 1
        if _norm(r[10]) != "+":
            continue
        d["pos"] += 1
        if _norm(r[12]) != "+" or any(_norm(r[j]) not in ("+", "") for j in (13, 14, 15)):
            continue
        if _norm(r[16]) == "SIM":
            continue
        d["conf"] += 1
    return out


# Mês da estação de monta: começa em agosto e fecha em julho.
MESES_ESTACAO = [8, 9, 10, 11, 12, 1, 2, 3, 4, 5, 6, 7]
# Meta atingida das safras já fechadas, como o relatório oficial de jul/2026
# publicou. A meta de uma safra vive no PLANEJAMENTO do master DELA, e os
# masters antigos (formato CONTROLE_DOADORAS) não têm essa aba — então o número
# não é recalculável e fica registrado aqui, com a fonte.
META_PCT_OFICIAL = {"2022/2023": 0.71, "2023/2024": 0.75, "2024/2025": 0.89}


def _stats_safra_estacao(ws, safra_alvo=None) -> dict:
    """Por safra, da aba ESTAÇÃO (formato novo): confirmados por MÊS DA IA —
    que é como o relatório data o embrião (24/25 bate mês a mês) —, total e
    doadoras distintas com embrião confirmado."""
    por = {}
    for i, r in enumerate(ws.iter_rows(values_only=True), 1):
        if i < 3 or r[0] is None or len(r) < 36:
            continue
        sf = _s(r[35])
        if not sf or "/" not in sf or (safra_alvo and sf != safra_alvo):
            continue
        d = por.setdefault(sf, {"meses": {}, "conf": 0, "doad": set()})
        if _norm(r[10]) != "+" or _norm(r[12]) != "+":
            continue
        if any(_norm(r[j]) not in ("+", "") for j in (13, 14, 15)) or _norm(r[16]) == "SIM":
            continue
        d["conf"] += 1
        d["doad"].add(_norm(r[2]))
        ia = r[7]
        if hasattr(ia, "month"):
            d["meses"][ia.month] = d["meses"].get(ia.month, 0) + 1
    return por


def _stats_safra_antiga(safra: str):
    """Safra anterior à reorganização da planilha: o master da pasta DELA, aba
    CONTROLE_DOADORAS (10 lavado, 15 15D, 16-18 30/45/60D, 23 data do aborto,
    45 safra)."""
    a, b = safra.split("/")
    pasta = ESTACAO_MONTA_BASE / f"Estação {a}-{b}"
    if not pasta.is_dir():
        return None
    cands = [f for f in pasta.glob("*ESTACAO DE MONTA*.xlsx")
             if not f.name.startswith("~$") and "PBI" not in f.name.upper()]
    if not cands:
        return None
    f = max(cands, key=lambda x: (RE_PREFIXO_DATA.match(x.name).group(1)
                                  if RE_PREFIXO_DATA.match(x.name) else "", x.stat().st_mtime))
    try:
        wb = _load(_registra(f"estação {safra}", f))
    except Exception:
        return None
    if "CONTROLE_DOADORAS" not in wb.sheetnames:
        wb.close()
        return None
    d = {"meses": {}, "conf": 0, "doad": set()}
    for i, r in enumerate(wb["CONTROLE_DOADORAS"].iter_rows(values_only=True), 1):
        if i < 6 or r[0] is None or len(r) < 46 or _s(r[45]) != safra:
            continue
        if _norm(r[10]) != "+" or _norm(r[15]) != "+":
            continue
        if any(_norm(r[j]) not in ("+", "") for j in (16, 17, 18)) or r[23] is not None:
            continue
        d["conf"] += 1
        d["doad"].add(_norm(r[2]))
        if hasattr(r[7], "month"):
            d["meses"][r[7].month] = d["meses"].get(r[7].month, 0) + 1
    wb.close()
    return d


def _meta_safra(safra: str):
    """Meta total de embriões do PLANEJAMENTO do master da safra (a célula B1 diz
    de que safra é a aba). None quando não há."""
    try:
        wb = _load(_master_da_safra(safra))
    except Exception:
        return None
    try:
        ws = wb["PLANEJAMENTO"]
        if _s(next(ws.iter_rows(min_row=1, max_row=1, values_only=True))[1]) != safra:
            return None
        cab = [re.sub(r"\s+", "", _norm(x or "")) for x in
               next(ws.iter_rows(min_row=3, max_row=3, values_only=True))]
        c_nome = next((j for j, h in enumerate(cab) if h == "NOME"), None)
        c_meta = next((j for j, h in enumerate(cab) if h == "METATOTAL"), None)
        if c_nome is None or c_meta is None:
            return None
        tot = 0
        for r in ws.iter_rows(min_row=4, values_only=True):
            if c_nome < len(r) and r[c_nome] and _s(r[c_nome]) and c_meta < len(r):
                tot += _to_num(r[c_meta]) or 0
        return tot or None
    finally:
        wb.close()


def comparativo(wb, safra: str):
    """S18 — as quatro últimas safras, confirmados por mês da IA, com a meta
    atingida embaixo de cada uma. A aba COMPARATIVO da planilha está congelada em
    20/21–23/24 e não serve; a aba ESTAÇÃO tem a safra de cada embrião."""
    ini = int(safra[:4])
    safras = [f"{a}/{a + 1}" for a in range(ini - 3, ini + 1)]
    novas = _stats_safra_estacao(wb["ESTAÇÃO"])
    blocos = []
    for sf in safras:
        d = novas.get(sf) or _stats_safra_antiga(sf)
        if not d:
            continue
        meta = _meta_safra(sf)
        pct_meta = (d["conf"] / meta) if meta else META_PCT_OFICIAL.get(sf)
        curta = f"{sf[2:4]}/{sf[-2:]}"
        blocos.append({"rotulo": f"{curta}  ({d['conf']} emb / {len(d['doad'])} doad)",
                       "curta": curta, "atual": sf == safra, "total": d["conf"],
                       "meses": [d["meses"].get(mm) for mm in MESES_ESTACAO],
                       "meta": f"Meta: {pct_meta*100:.0f}%" if pct_meta is not None else "Meta: —"})
    if not blocos:
        return pend(18, "ESTAÇÃO DE MONTA — COMPARATIVO COM ANOS ANTERIORES", "",
                    "ESTACAO DE MONTA.xlsx, aba ESTAÇÃO", "nenhuma safra encontrada na coluna ESTAÇÃO")
    curtas = [b["curta"] for b in blocos]
    lista = (", ".join(curtas[:-1]) + f" e {curtas[-1]}") if len(curtas) > 1 else curtas[0]
    return {"t": "comparativo", "n": 18, "titulo": "ESTAÇÃO DE MONTA — COMPARATIVO COM ANOS ANTERIORES",
            "sub": f"Embriões confirmados por mês  ·  Estações {lista}",
            "meses": [ABR[mm - 1] for mm in MESES_ESTACAO], "safras": blocos}


def doadoras(wb, safra):
    """S19/S20 — meta × realizado por doadora, do PLANEJAMENTO (NOME, TIME, META
    TOTAL, TOTAL EMBRIÕES). Colunas pelo CABEÇALHO (linha 3): na safra 2026/2027 o
    haras reorganizou a aba — saíram CATEGORIA, LOCAL e TIME, e META/TOTAL
    EMBRIÕES andaram duas colunas. Com índice fixo o slide de agosto saía vazio.

    Time: o do PLANEJAMENTO; doadora sem time lá (Música, em 25/26) usa o da aba
    REC. EMBR., que é como o relatório a classificou. A META do time soma só quem
    tem o time no PLANEJAMENTO — é o número que o relatório imprime (Time B: 42).
    Sem coluna TIME na safra, sai um slide só, sem inventar divisão."""
    ws = wb["PLANEJAMENTO"]
    # o cabeçalho tem quebra de linha dentro da célula ('META' + quebra + 'TOTAL'):
    # compara sem nenhum espaço em branco, senão nenhuma coluna casa e o slide
    # sai vazio — foi o DOADORAS em branco do deck de agosto
    cab = [re.sub(r"\s+", "", _norm(x or "")) for x in
           next(ws.iter_rows(min_row=3, max_row=3, values_only=True))]

    def _col(*nomes):
        for j, h in enumerate(cab):
            if any(h == re.sub(r"\s+", "", _norm(n)) for n in nomes):
                return j
        return None
    c_nome, c_time = _col("NOME"), _col("TIME")
    c_meta = _col("META TOTAL", "META  TOTAL")
    c_real = _col("TOTAL EMBRIOES", "TOTAL EMBRIÕES")
    if c_nome is None or c_meta is None or c_real is None:
        aviso(f"PLANEJAMENTO sem NOME/META/TOTAL EMBRIÕES no cabeçalho ({safra}) — S19/S20 pendentes")
        return [pend(19, f"ESTAÇÃO DE MONTA {safra} — DOADORAS", "", "ESTACAO DE MONTA.xlsx",
                     "a aba PLANEJAMENTO mudou de layout e não achei as colunas")]
    time_rec = {}
    for i, r in enumerate(wb["REC. EMBR."].iter_rows(values_only=True), 1):
        if i >= 3 and len(r) > 3 and r[2] and c_time is not None:
            t = _norm(r[3])
            if t in ("A", "B"):
                time_rec[_norm(r[2])] = t
    vistos, linhas = set(), []
    for i, r in enumerate(ws.iter_rows(values_only=True), 1):
        if i < 4 or c_nome >= len(r) or r[c_nome] is None:
            continue
        nome = _s(r[c_nome])
        if not nome or _norm(nome).startswith("TOTAL"):
            continue
        meta = int(_to_num(r[c_meta]) or 0) if c_meta < len(r) else 0
        real = int(_to_num(r[c_real]) or 0) if c_real < len(r) else 0
        # a planilha repete doadora (Charmosa aparece como PG e como sócio) e traz
        # linha sem meta nem embrião — nenhuma das duas é doadora a acompanhar
        if _norm(nome) in vistos or (not meta and not real):
            continue
        vistos.add(_norm(nome))
        t_plan = _norm(r[c_time]) if c_time is not None and c_time < len(r) and r[c_time] else ""
        t = t_plan or time_rec.get(_norm(nome), "")
        linhas.append({"nome": nome_animal(nome), "chave": _norm(nome), "meta": meta, "real": real,
                       "time": t if c_time is not None else "",
                       "conta_meta": bool(t_plan) or c_time is None, "ordem": i})
    # taxa de recuperação e prenhez do time, pela aba ESTAÇÃO (mesmo funil do S16)
    taxas = {}
    grupo_de = {l["chave"]: l["time"] for l in linhas}
    for i, r in enumerate(wb["ESTAÇÃO"].iter_rows(values_only=True), 1):
        if i < 3 or r[0] is None or len(r) < 36 or _s(r[35]) != safra:
            continue
        t = grupo_de.get(_norm(r[2]))
        if t is None:
            continue
        x = taxas.setdefault(t, [0, 0, 0])
        x[0] += 1
        if _norm(r[10]) == "+":
            x[1] += 1
            if _norm(r[12]) == "+":
                x[2] += 1
    times = ["A", "B"] if c_time is not None else [""]
    out = []
    for t in times:
        rows = [l for l in linhas if l["time"] == t]
        if not rows:
            continue
        # com embrião primeiro, do maior realizado pro menor; os zerados depois,
        # na ordem da planilha
        rows.sort(key=lambda l: (0, -l["real"], l["ordem"]) if l["real"] else (1, 0, l["ordem"]))
        meta = sum(l["meta"] for l in rows if l["conta_meta"])
        real = sum(l["real"] for l in rows)
        tent, lav, p15 = taxas.get(t, [0, 0, 0])
        partes = [f"Meta: {meta} embriões", f"Realizado: {real} embriões"]
        if tent:
            partes.append(f"Rec. Embrionária: {lav/tent*100:.0f}%")
        if lav:
            partes.append(f"Prenhez: {p15/lav*100:.0f}%")
        out.append({"t": "doadoras", "n": 20 if t == "B" else 19,
                    "titulo": (f"ESTAÇÃO DE MONTA {safra} — DOADORAS TIME {t}" if t
                               else f"ESTAÇÃO DE MONTA {safra} — DOADORAS"),
                    "sub": "  ·  ".join(partes),
                    "rows": [{"nome": l["nome"], "meta": l["meta"], "real": l["real"]} for l in rows]})
    return out or [pend(19, f"ESTAÇÃO DE MONTA {safra} — DOADORAS", "", "ESTACAO DE MONTA.xlsx",
                        "o PLANEJAMENTO da safra não tem doadora com meta")]


# ========================================================= Coberturas (S21)
# O guia chama de "COBERTURAS_CAVALOS_FORA.xlsx"; no Drive o arquivo é
# REPRODUÇÃO/COBERTURAS - CAVALOS DE FORA NÃO USADAS.xlsx.
COBERTURAS = None            # resolvido em tempo de execução (ver _coberturas_path)
# Excluídos por decisão do haras (constam no guia).
COBERTURAS_FORA = ("TRILHO DA ZIZICA", "QUANTUM DE ALCATEIA")


def _coberturas_path():
    from _pg_common import DRIVE_ROOT
    p = DRIVE_ROOT / "REPRODUÇÃO" / "COBERTURAS - CAVALOS DE FORA NÃO USADAS.xlsx"
    return p if p.exists() else None


def slide_coberturas(safra):
    """S21 — aba Planilha2: 2 garanhão, 3 qtd comprada, 4 utilizadas, 5 saldo.
    A Planilha1 é o log de compra (uma linha por negócio); a Planilha2 é o
    consolidado por garanhão, que é o que o slide mostra."""
    f = _coberturas_path()
    if f is None:
        return pend(21, f"ESTAÇÃO DE MONTA {safra} — COBERTURAS DISPONÍVEIS",
                    "Saldo por garanhão de fora",
                    "REPRODUÇÃO/COBERTURAS - CAVALOS DE FORA NÃO USADAS.xlsx",
                    "arquivo não encontrado no Drive")
    wb = _load(_registra("coberturas de fora", f))
    ws = wb["Planilha2"]
    rows = []
    for i, r in enumerate(ws.iter_rows(values_only=True), 1):
        if i < 3 or len(r) < 5 or r[1] is None:
            continue
        n = _norm(r[1])
        # abaixo de 'ARQUIVO MORTO' a aba guarda os garanhões encerrados, com o
        # próprio cabeçalho repetido — entravam no slide como se fossem linhas
        if n.startswith("ARQUIVO MORTO"):
            break
        if not n or n.startswith("TOTAL") or n in COBERTURAS_FORA or n in ("GARANHAO", "GARANHÃO"):
            continue
        saldo = int(_to_num(r[4]) or 0)
        if saldo > 0:
            rows.append({"nome": titulo_pt(r[1]), "saldo": saldo})
    wb.close()
    # do maior saldo pro menor; no empate, a ordem da planilha (sort estável)
    rows.sort(key=lambda x: -x["saldo"])
    return {"t": "coberturas", "n": 21,
            "titulo": f"ESTAÇÃO DE MONTA {safra} — COBERTURAS DISPONÍVEIS",
            "sub": "Coberturas de garanhões de fora com saldo disponível  ·  Fonte: planilha de controle de coberturas",
            "rows": rows}


# ====================================================== Inadimplência (S31)
# Fonte: as fotos mensais que o ControleInadimplencia.py arquiva em
# output_pbi/historico (uma por data de referência, com a carteira inteira).
# A saída "viva" (indicadores_kpi.xlsx) é sobrescrita a cada rodada e não serve
# pra mês fechado: regerar o deck de agosto em 18/09 trocava a posição de 31/08
# pela de setembro. O slide só mostra agregados — nenhum nome de devedor sai daqui.
INAD_DIR = Path(r"C:/Users/Arthur/repos/controle-de-inadimplencia/output_pbi")


INAD_HIST = INAD_DIR / "historico"
INAD_CARTEIRA = "Carla"      # o painel que o relatório mostra é o da carteira da Carla
STATUS_VENCIDO = ("Aberto vencido", "Aberto parcialmente")


def _inad_foto(data_ref: date):
    """Foto da carteira na data (historico/fato_titulos_AAAA-MM-DD.parquet)."""
    f = INAD_HIST / f"fato_titulos_{data_ref.isoformat()}.parquet"
    return f if f.exists() else None


def _inad_kpis(f: Path) -> dict:
    """Os seis cartões e a quebra por ano, com a MESMA regra do dashboard de
    conferência (ControleInadimplencia.py, computeKPIs/aggregateByYear): título
    vencido = status vencido/parcial E mais de 7 dias; Ação Judicial, Não
    Entregues e Inadimplentes são a quebra dos vencidos; o resto é A Vencer."""
    d = pd.read_parquet(f)
    d = d[(d["tipo_conta"] == "CAR") & (d["carteira"] == INAD_CARTEIRA)]
    venc = d["status_vencimento"].astype(str).str.strip().isin(STATUS_VENCIDO) & (d["dias_atraso"] > 7)
    cat = d["categoria"].astype(str)
    k = {"total": float(d["vl_em_aberto"].sum()),
         "venc": float(d.loc[venc, "vl_em_aberto"].sum()),
         "aj": float(d.loc[venc & (cat == "ACAO JUDICIAL"), "vl_em_aberto"].sum()),
         "ne": float(d.loc[venc & (cat == "NAO ENTREGUES"), "vl_em_aberto"].sum()),
         "clientes": int(d["nome_pessoa"].nunique()),
         "clientes_venc": int(d.loc[venc, "nome_pessoa"].nunique())}
    k["inad"] = k["venc"] - k["aj"] - k["ne"]
    k["avencer"] = k["total"] - k["venc"]
    anos = []
    for ano_e, g in d.groupby("ano_emissao"):
        if pd.isna(ano_e) or not g["vl_em_aberto"].sum():
            continue
        v = venc.loc[g.index]
        c = cat.loc[g.index]
        tot = float(g["vl_em_aberto"].sum())
        vv = float(g.loc[v, "vl_em_aberto"].sum())
        aj = float(g.loc[v & (c == "ACAO JUDICIAL"), "vl_em_aberto"].sum())
        ne = float(g.loc[v & (c == "NAO ENTREGUES"), "vl_em_aberto"].sum())
        anos.append({"ano": int(ano_e), "venc": vv / tot, "inad": (vv - aj - ne) / tot,
                     "aj": aj / tot, "ne": ne / tot})
    k["anos"] = anos
    return k


def slide_inadimplencia(m, ano):
    """S31 — o painel de cobrança na posição do FIM do mês do deck, comparado com
    o fim do mês anterior. O relatório colava o print do dashboard; aqui o painel
    é desenhado com os mesmos números, tirados da foto que o próprio
    ControleInadimplencia.py arquiva a cada fechamento (historico/). Sem foto do
    mês, cai na posição arquivada pelo build (_cache) e, por último, na viva."""
    fim_mes = date(ano + m // 12, m % 12 + 1, 1) - timedelta(days=1)
    f = _inad_foto(fim_mes)
    if f is not None:
        _registra("inadimplência (foto do mês)", f)
        k = _inad_kpis(f)
        ini = date(fim_mes.year, fim_mes.month, 1) - timedelta(days=1)
        fa = _inad_foto(ini)
        ka = _inad_kpis(fa) if fa is not None else None
        return {"t": "inadimplencia", "n": 31, "titulo": "VENDAS — INADIMPLÊNCIAS E RECEBÍVEIS",
                "sub": f"Posição {MESES[m-1].upper()}/{ano}  ·  Fonte: Dashboard de Gestão de Cobrança",
                "k": {x: k[x] for x in ("total", "venc", "aj", "ne", "inad", "avencer",
                                         "clientes", "clientes_venc")},
                "ant": ({x: ka[x] for x in ("total", "venc", "aj", "ne", "inad", "avencer")}
                        if ka else None),
                "ref_ant": ini.strftime("%d/%m/%Y") if ka else None,
                "anos": k["anos"]}
    # Sem a foto do fim do mês o slide fica pendente. Cair na planilha viva (ou
    # numa posição arquivada de outra data) punha no deck de maio a carteira de
    # 31/08 — número de outro mês com cara de certo.
    return pend(31, "VENDAS — INADIMPLÊNCIAS E RECEBÍVEIS", f"Posição {MESES[m-1].upper()}/{ano}",
                "controle-de-inadimplencia → output_pbi/historico",
                f"sem foto da carteira em {fim_mes.strftime('%d/%m/%Y')}; rode o ControleInadimplencia.py "
                f"com a base do fechamento")


# ============================================================== Vendas (S29–S35)
VENDEDOR_COMITE = "CARLA"


def _mapa_vendas_do_mes(ano: int, m: int):
    """Mapa de vendas do FECHAMENTO do mês: a primeira versão gerada depois do fim
    do mês (o prefixo do nome é a data). O mais novo pode ter revisado o passado
    — foi o caso de março/26, que caiu de 749.050 para 670.050 na Semana entre as
    versões de julho e de agosto —, e o deck de um mês fechado não muda por isso.
    Sem versão posterior ao mês (mês corrente), vale a mais nova."""
    raiz = Path(MAPA_VENDAS_DIR).parent
    fim = date(ano + m // 12, m % 12 + 1, 1)
    cands = []
    for f in raiz.rglob("*_PG_Mapa Vendas*.xlsx"):
        mm = RE_PREFIXO_DATA.match(f.name)
        if f.name.startswith("~$") or not mm:
            continue
        try:
            d = datetime.strptime(mm.group(1), "%y%m%d").date()
        except ValueError:
            continue
        cands.append((d, f.stat().st_mtime, f))
    depois = sorted(c for c in cands if c[0] >= fim)
    if depois:
        return depois[0][2]
    return max(cands)[2] if cands else None


def _evento_vendas(tipo, nome) -> str:
    """Rótulo da origem como o relatório: 'Venda Direta' ou o NOME do evento
    ('XVI Semana de Negócios PG'). O tipo ('LEILAO PROPRIO') não serve de rótulo,
    e a planilha grafa a mesma Semana com um e com dois espaços."""
    if "DIRETA" in _norm(tipo or "") or not _s(nome):
        return "Venda Direta"
    n = " ".join(_norm(nome).split())
    return titulo_pt(n)


def slides_vendas(m, ano, meta_anual=4_500_000):
    """S29 resultado acumulado e S30 detalhamento. Filtro obrigatório: VENDEDOR =
    CARLA, sem CANCELADO (regra do guia). Colunas do MAPA VENDAS (1-based): 8
    valor da venda, 11 tipo de evento, 12 nome do evento, 15 vendedor, 19 status
    contrato, 22 ano, 23 mês."""
    try:
        src = _mapa_vendas_do_mes(ano, m) or _latest_by_yymmdd(MAPA_VENDAS_DIR, "*_PG_Mapa Vendas.xlsx",
                                                                "mapa de vendas")
        wb = _load(_registra("mapa de vendas", src))
    except Exception as e:
        p = pend(29, "VENDAS — RESULTADO ACUMULADO", "", "PG_Mapa Vendas.xlsx, aba MAPA VENDAS",
                 f"não consegui abrir: {e}")
        return [p, dict(p, n=30, titulo="VENDAS — DETALHAMENTO POR MÊS E EVENTO")]
    ws = wb["MAPA VENDAS"]
    por_mes = {}
    for i, r in enumerate(ws.iter_rows(values_only=True), 1):
        if i < 3 or len(r) < 23 or r[21] is None:
            continue
        # a coluna traz o nome completo do vendedor ("CARLA ...") — comparar por
        # igualdade zerava o slide inteiro
        if VENDEDOR_COMITE not in _norm(r[14]):
            continue
        if "CANCELAD" in _norm(r[18]):
            continue
        a, mm = _to_num(r[21]), _to_num(r[22])
        if a != ano or not mm or mm > m:
            continue
        v = _to_num(r[7]) or 0
        ev = _evento_vendas(r[10], r[11])
        por_mes.setdefault(int(mm), {}).setdefault(ev, 0)
        por_mes[int(mm)][ev] += v
    wb.close()
    ytd = sum(sum(d.values()) for d in por_mes.values())
    mes_v = sum(por_mes.get(m, {}).values())
    pct = ytd / meta_anual if meta_anual else 0
    mes_nome = MESES[m - 1]
    meses = [{"mes": MESES[mm - 1].upper(), "abr": ABR[mm - 1], "total": sum(por_mes[mm].values()),
              "eventos": [[ev, v] for ev, v in sorted(por_mes[mm].items(), key=lambda kv: -kv[1])]}
             for mm in sorted(por_mes)]
    return [
        {"t": "vendas_acum", "n": 29, "titulo": f"VENDAS {ano} — RESULTADO ACUMULADO — {VENDEDOR_COMITE}",
         "sub": (f"Meta anual: {brl_cheio(meta_anual)}  ·  Acumulado Jan–{ABR[m-1]}: {brl_curto(ytd)}"
                 f"  ·  Vendedor: {VENDEDOR_COMITE.title()}"),
         "kpis": [{"v": brl_cheio(mes_v), "l": f"Vendas {mes_nome}", "s": "Realizado no mês", "cor": "ouro"},
                  {"v": brl_curto(ytd), "l": "Acumulado YTD", "s": f"Jan–{ABR[m-1]} {ano}", "cor": "navy"},
                  {"v": brl_curto(meta_anual), "l": "Meta Anual", "s": f"Objetivo {ano}", "cor": "ardosia"},
                  {"v": brl_curto(max(meta_anual - ytd, 0)), "l": "Saldo para Meta", "s": "Ainda a realizar",
                   "cor": "azul"}],
         "pct": pct, "barra": f"{pct*100:.0f}% da meta atingida",
         "colunas": [{"rot": x["abr"], "v": x["total"]} for x in meses]},
        {"t": "vendas_mes", "n": 30,
         "titulo": f"VENDAS — JANEIRO A {mes_nome.upper()}/{str(ano)[2:]} — {VENDEDOR_COMITE}",
         "sub": f"Detalhamento por mês e evento  ·  Filtro: Vendedor = {VENDEDOR_COMITE.title()}",
         "meses": meses},
    ]


# ENTREGAR (1-based): 2 doadora, 3 garanhão, 4 data venda, 6 comprador, 7 cota,
# 11 valor, 12 status pgto, 13 status embrião.  RECEBER: 6 vendedor.
S32 = ("QUITADO", "PAGANDO")
S33 = ("PAUSAD", "APOS CONF", "APÓS CONF")
S34 = ("DIREITO", "TROCA")


def _data_contrato(v):
    """Data como vem na planilha (datetime, 'dd/mm/aaaa' ou só o ano) -> (date|None, texto)."""
    if hasattr(v, "strftime"):
        return (v.date() if hasattr(v, "date") else v), v.strftime("%d/%m/%y")
    t = _s(v)
    mm = re.match(r"^(\d{2})/(\d{2})/(\d{4})$", t or "")
    if mm:
        d = date(int(mm.group(3)), int(mm.group(2)), int(mm.group(1)))
        return d, d.strftime("%d/%m/%y")
    return None, (t or "—")


def _cota_txt(c):
    if not c:
        return "—"
    return f"{c:g}"


def slides_embrioes(m, ano):
    """S32–S35 — contratos de embrião, na ordem da planilha (que já vem por
    doadora) e com a altura de linha padrão que o haras pediu em junho. Contrato
    vendido DEPOIS do mês do deck não entra: o deck de agosto não mostra venda
    de setembro."""
    try:
        wb = _load(_registra("embriões a entregar", EMB_COMERCIAIS))
    except Exception as e:
        base = pend(32, "VENDAS — EMBRIÕES VENDIDOS A FAZER", "", EMB_COMERCIAIS.name,
                    f"não consegui abrir: {e}")
        return [base, dict(base, n=33), dict(base, n=34),
                dict(base, n=35, titulo="ESTAÇÃO DE MONTA — EMBRIÕES COMPRADOS A RECEBER")]
    fim_mes = date(ano + m // 12, m % 12 + 1, 1) - timedelta(days=1)

    def linhas(aba):
        out = []
        for i, r in enumerate(wb[aba].iter_rows(values_only=True), 1):
            if i < 4 or len(r) < 13 or r[1] is None:
                continue
            d, dtxt = _data_contrato(r[3])
            if d and d > fim_mes:
                continue
            out.append({"doadora": nome_animal(r[1], curto=False),
                        "garanhao": titulo_pt(r[2]) if _s(r[2]) else "—",
                        "data": dtxt, "contraparte": pessoa_curta(r[5]),
                        "cota": _cota_txt(_to_num(r[6])), "valor": brl_k(_to_num(r[10])),
                        "pgto": _s(r[11]) or "", "status": _s(r[12]) or ""})
        return out

    ent = linhas("ENTREGAR")
    rec = linhas("RECEBER")
    wb.close()
    af = lambda x: _norm(x["status"]).startswith("A FAZER")
    tem = lambda x, ks: any(k in _norm(x["pgto"]) for k in ks)

    def slide(n, titulo, sub, lst, rotulo_contra="COMPRADOR", col_doadora="DOADORA", pgto=None):
        return {"t": "contratos", "n": n, "titulo": titulo,
                "sub": f"{len(lst)} contratos  ·  {sub}",
                "cols": [col_doadora, "GARANHÃO", "DATA", rotulo_contra, "CT", "VALOR", "PGTO"],
                "rows": [[x["doadora"], x["garanhao"], x["data"], x["contraparte"], x["cota"],
                          x["valor"], (pgto(x) if pgto else x["pgto"])] for x in lst]}

    s33 = [x for x in ent if af(x) and tem(x, S33 + ("A PAGAR",))]
    tem_a_pagar = any("A PAGAR" in _norm(x["pgto"]) for x in s33)
    s34 = [x for x in ent if _norm(x["status"]).startswith("REPOSI") or (af(x) and tem(x, S34))]
    return [
        slide(32, "VENDAS — EMBRIÕES VENDIDOS A FAZER (QUITADO / PAGANDO)",
              "Status: A fazer  ·  Pgto: Quitado ou Pagando", [x for x in ent if af(x) and tem(x, S32)]),
        slide(33, "VENDAS — EMBRIÕES VENDIDOS A FAZER (PGTO PAUSADO / APÓS CONF.)",
              "Status: A fazer  ·  Pgto: Pausado" + (", Após confirmação ou A pagar" if tem_a_pagar
                                                    else " ou Após confirmação"), s33),
        # na reposição o que importa é o status do embrião ('Reposição'); no
        # contrato de direito, o 'Direito' — é o que a coluna mostra no relatório
        slide(34, "VENDAS — EMBRIÕES DE DIREITO / REPOSIÇÃO",
              "Status: Reposição ou A fazer  ·  Pgto: Direito / Troca", s34,
              pgto=lambda x: ("Reposição" if _norm(x["status"]).startswith("REPOSI")
                              else ("Direito" if "DIREITO" in _norm(x["pgto"]) else x["pgto"]))),
        slide(35, "ESTAÇÃO DE MONTA — EMBRIÕES COMPRADOS A RECEBER",
              'Status "A Fazer" — ainda não produzidos  ·  Fonte: aba RECEBER', [x for x in rec if af(x)],
              "VENDEDOR", "DOADORA (ORIGEM)"),
    ]


# ============================= Conteúdo escrito à mão (S08, S23–S27, S38, S39)
# Comentário do DRE, exposição, manejo e foto não saem de planilha: são escritos
# todo mês. Até 31/08/2026 ficavam só em `_docs/comite_conteudo.json`, editado
# na mão por quem tinha o repo aberto. Agora a fonte PRINCIPAL é a tabela
# comite_conteudo no Supabase do hub — editável direto pelo hub (Ana/Aline/
# Arthur). O JSON local virou seed/backup: só entra se o Supabase não
# responder OU não tiver aquele mês ainda (ex.: mês seguinte, antes de
# alguém editar pelo hub — sem isso o slide vira placeholder à toa se o
# conteúdo já existe no JSON de uma migração antiga).
CONTEUDO = REPO / "_docs" / "comite_conteudo.json"
CAMPOS_CONTEUDO = ("comentarios", "exposicoes", "manejo", "fotos", "pendencias")
FALTA_CONTEUDO = "escreva o conteúdo desse mês pelo hub (aba Comitê) ou em _docs/comite_conteudo.json"


def _supabase_env():
    cfg = dotenv_values(REPO / ".env")
    url = cfg.get("SUPABASE_URL", "").rstrip("/")
    key = cfg.get("SUPABASE_SERVICE_ROLE_KEY")
    return (url, key) if url and key else (None, None)


def le_conteudo():
    """Supabase primeiro (fonte principal); JSON local só tapa buraco —
    mês que o Supabase não tem ainda, ou o serviço fora do ar."""
    local = {}
    if CONTEUDO.exists():
        d = json.loads(CONTEUDO.read_text(encoding="utf-8"))
        local = {k: v for k, v in d.items() if re.fullmatch(r"\d{4}-\d{2}", k)}

    url, key = _supabase_env()
    if not url:
        print("  [conteudo] sem SUPABASE_URL/SERVICE_ROLE no .env — usando só o JSON local")
        return local
    try:
        r = requests.get(
            # select=* e não a lista de colunas: `pendencias` entrou depois
            # (migration 20260924) e pedir coluna que o banco ainda não tem é 400
            f"{url}/rest/v1/comite_conteudo?select=*",
            headers={"apikey": key, "Authorization": f"Bearer {key}"}, timeout=15,
        )
        r.raise_for_status()
        remoto = {row["mes"]: {k: row.get(k) for k in CAMPOS_CONTEUDO} for row in r.json()}
    except Exception as exc:
        print(f"  [conteudo] Supabase indisponível ({exc!r}) — usando só o JSON local")
        return local

    # remoto manda; mês que só existe no JSON (não migrado/editado ainda) sobrevive
    faltando_no_remoto = sorted(set(local) - set(remoto))
    if faltando_no_remoto:
        print(f"  [conteudo] só no JSON local (não migrado pro Supabase ainda): "
              f"{', '.join(faltando_no_remoto)}")
    return {**local, **remoto}


def conteudo_do_mes(todos, chave):
    """Só conteúdo do próprio mês — puxar de um mês futuro colocaria no deck de
    janeiro a exposição que ainda não tinha acontecido."""
    return todos.get(chave, {})


def slide_comentarios(c, m, ano):
    """S8 — comentários do MÊS do deck.

    Era `VARIAÇÕES YTD JAN–<mês>`, e entrava logo depois do DRE acumulado: quem
    escrevia o conteúdo comentava o acumulado do ano, porque era o que o slide
    anterior mostrava. O deck é mensal — o comentário acompanha o mês, e por isso
    o slide agora fica junto do bloco mensal (resumo, custos e despesas do mês)."""
    itens = c.get("comentarios") or []
    titulo = f"COMENTÁRIOS — {MESES[m-1].upper()} {ano}"
    if not itens:
        return pend(8, titulo, "Principais destaques do mês por categoria",
                    "_docs/comite_conteudo.json → comentarios", FALTA_CONTEUDO,
                    edita="comentarios")
    return {"t": "comentarios", "n": 8, "titulo": titulo,
            "sub": "DRE 2026 | HPG  ·  Principais destaques do mês por categoria",
            "itens": itens}


def slide_pendencias(c, m, ano):
    """S03 — o que ficou combinado na apresentação ANTERIOR, em forma de lista de
    verificação. É o terceiro slide do relatório dela, logo depois da agenda."""
    ant = MESES[(m - 2) % 12].upper()
    titulo = f"PENDÊNCIAS DA APRESENTAÇÃO DE {ant}"
    itens = [x for x in (c.get("pendencias") or []) if str(x).strip()]
    if not itens:
        return pend(3, titulo, "", "comite_conteudo → pendencias", FALTA_CONTEUDO, edita="pendencias")
    return {"t": "pendencias", "n": 3, "titulo": titulo, "itens": itens}


def slides_exposicoes(c, ano):
    exp = c.get("exposicoes") or {}
    prog, res = exp.get("programacao") or [], exp.get("resultados") or []
    out = []
    if prog:
        out.append({"t": "tabela", "n": 23, "titulo": f"EXPOSIÇÕES {ano} — PROGRAMAÇÃO",
                    "sub": "Calendário de participações previstas",
                    "cols": ["EVENTO", "DATA", "LOCAL", "STATUS"], "rows": prog})
    else:
        out.append(pend(23, f"EXPOSIÇÕES {ano} — PROGRAMAÇÃO", "Calendário de participações",
                        "_docs/comite_conteudo.json → exposicoes.programacao", FALTA_CONTEUDO,
                        edita="exposicoes"))
    if res:
        for k, r in enumerate(res):
            out.append({"t": "resultados", "n": 24 + k, "titulo": r["titulo"],
                        "sub": r.get("sub", ""), "animais": r["animais"]})
    else:
        out.append(pend(24, "RESULTADOS DAS EXPOSIÇÕES", "Animais, títulos e colocações",
                        "_docs/comite_conteudo.json → exposicoes.resultados", FALTA_CONTEUDO,
                        edita="exposicoes"))
    return out


def slide_manejo(c, m, ano):
    itens = c.get("manejo") or []
    if not itens:
        return pend(38, "MANEJO — PONTOS DE MELHORIA E DECISÕES", "Histórico de intervenções",
                    "_docs/comite_conteudo.json → manejo", FALTA_CONTEUDO,
                    edita="manejo")
    return {"t": "manejo", "n": 38, "titulo": "MANEJO — PONTOS DE MELHORIA E DECISÕES",
            "sub": f"Histórico de intervenções Jan–{ABR[m-1]} {ano}", "itens": itens,
            "atual": ABR[m - 1]}


# Fotos do mês, como o haras as manda: pasta única por ano, arquivos do WhatsApp.
FOTOS_DIR_BASE = DRIVE_ROOT / "ATA & APRESENTACOES MENSAIS"   # /<ano>/FOTOS
FOTO_EXTS = (".jpeg", ".jpg", ".png")

# 6 fotos por slide: mais que isso e cada foto vira selo; menos, sobra tela.
FOTOS_POR_SLIDE = 6
# Grade por quantidade, porque o último slide raramente fecha com 6 — e julho de 2026
# tem UMA foto no mês. Com as 3 colunas fixas de antes, essa foto virava uma tira de
# 1/3 de largura no canto. A grade vai no spec, e não em cada renderizador, para o
# HTML e o PPTX não divergirem.
GRADE_FOTOS = {1: (1, 1), 2: (2, 1), 3: (3, 1), 4: (2, 2), 5: (3, 2), 6: (3, 2)}
# 12 por mês = 2 slides. Junho de 2026 tem 54 arquivos, 45 depois do dedup por hash.
FOTOS_POR_MES = 12

# "WhatsApp Image 2026-06-19 at 14.31.07 (2).jpeg" — data e hora estão no nome porque
# o WhatsApp REMOVE o EXIF (confirmado: 0 dos 80 arquivos tem tag de data).
#
# O nome marca o ENCAMINHAMENTO, não a foto. Duas medições mostram isso:
#   * 9 pares são byte-idênticos com nomes de dias diferentes (08/06 07.31.42 ==
#     19/06 14.31.03) — alguém reenviou o mesmo lote;
#   * as 45 fotos únicas de junho ocupam 8 minutos distintos de nome, porque um lote
#     inteiro sai com segundos consecutivos.
# Consequência prática: a hora não separa cenas. Colapsar "rajada" por janela de tempo
# derrubou junho de 45 para 6 e jogou fora foto legítima. Comparei as 45 por dHash e a
# distância mínima entre quaisquer duas passa de 22 bits: são 45 imagens distintas, não
# há rajada nenhuma. Só o hash de bytes remove repetição de verdade.
# A DATA continua servindo: é o mês em que o registro chegou ao grupo, que é o que o
# comitê discute.
_RE_FOTO_DT = re.compile(r"(20\d\d)-(\d\d)-(\d\d)(?:\D+(\d\d)\.(\d\d)\.(\d\d))?")


def _foto_quando(p: Path) -> datetime:
    """Data do nome; sem data no nome, o mtime do arquivo."""
    m = _RE_FOTO_DT.search(p.name)
    if not m:
        return datetime.fromtimestamp(p.stat().st_mtime)
    a, s_, d, hh, mm, ss = m.groups()
    return datetime(int(a), int(s_), int(d), int(hh or 0), int(mm or 0), int(ss or 0))


def _fotos_do_mes(ano: int, mes: int) -> list[Path]:
    """Fotos do mês, sem repetição e espalhadas pelos dias.

    Dois filtros:

      1. hash — arquivo byte-idêntico reenviado noutro dia é uma foto só;
      2. espalhamento — com mais foto que vaga, roda um dia por vez em vez de pegar
         as primeiras. Sem isso junho seria 12 fotos do dia 19, que tem 28 das 45; os
         dias 02, 08, 11 e 24 sumiriam do slide.
    """
    pasta = FOTOS_DIR_BASE / str(ano) / "FOTOS"
    if not pasta.exists():
        return []
    todas = [p for p in pasta.rglob("*") if p.suffix.lower() in FOTO_EXTS]
    do_mes = [p for p in todas
              if (q := _foto_quando(p)).year == ano and q.month == mes]
    if not do_mes:
        return []

    # 1. mesma foto reenviada: fica a de nome mais antigo (primeira aparição)
    por_hash: dict[str, Path] = {}
    for p in sorted(do_mes, key=_foto_quando):
        por_hash.setdefault(hashlib.md5(p.read_bytes()).hexdigest(), p)
    cenas = sorted(por_hash.values(), key=_foto_quando)
    if len(cenas) <= FOTOS_POR_MES:
        return cenas

    # 2. round-robin pelos dias até fechar a cota
    por_dia: dict[int, list[Path]] = {}
    for p in cenas:
        por_dia.setdefault(_foto_quando(p).day, []).append(p)
    escolhidas = []
    while len(escolhidas) < FOTOS_POR_MES:
        rodada = [fila.pop(0) for _, fila in sorted(por_dia.items()) if fila]
        if not rodada:
            break
        escolhidas += rodada[:FOTOS_POR_MES - len(escolhidas)]
    return sorted(escolhidas, key=_foto_quando)
# As fotos NAO ficam no repo (publico) nem no site (publico): entram embutidas no
# spec.json, que sai pelo bucket privado. Cru sao 12 MB em 28 arquivos; reduzidas
# pro tamanho em que o deck as mostra (6 por slide, ~1/3 de tela) dao ~2 MB.
FOTO_LADO_MAX = 760
FOTO_QUALIDADE = 68


def _foto_embutida(caminho: Path) -> str | None:
    """Foto como data URI reduzida. None se nao der pra ler — slide vira pendencia
    em vez de imagem quebrada, que e o que acontecia com caminho relativo depois
    que as fotos sairam do repo."""
    try:
        from PIL import Image
    except ImportError:
        print("  [fotos] Pillow ausente (pip install Pillow) — fotos ficam de fora")
        return None
    try:
        im = Image.open(caminho)
        im.thumbnail((FOTO_LADO_MAX, FOTO_LADO_MAX))
        if im.mode != "RGB":
            im = im.convert("RGB")
        buf = io.BytesIO()
        im.save(buf, "JPEG", quality=FOTO_QUALIDADE, optimize=True)
        return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()
    except Exception as exc:
        print(f"  [fotos] {caminho.name}: {exc!r}")
        return None


def _foto_do_bucket(path: str) -> str | None:
    """Foto já reduzida (upload/edição pelo hub sempre grava otimizada — ver
    tools/migrar_comite_conteudo.py e o upload do próprio hub) — baixa e
    embute direto, sem passar pelo PIL de novo."""
    url, key = _supabase_env()
    if not url:
        return None
    try:
        r = requests.get(f"{url}/storage/v1/object/comite-fotos/{path}",
                          headers={"apikey": key, "Authorization": f"Bearer {key}"}, timeout=15)
        r.raise_for_status()
        return "data:image/jpeg;base64," + base64.b64encode(r.content).decode()
    except Exception as exc:
        print(f"  [fotos] bucket {path}: {exc!r}")
        return None


def _item_foto(uri: str, video: str | None = None) -> dict:
    """Item da grade com a proporção da imagem. O relatório mostra a foto INTEIRA
    (sem recorte, cada uma na sua proporção, lado a lado ocupando a altura), e o
    layout só consegue fazer isso sabendo largura e altura de cada uma."""
    w = h = None
    try:
        from PIL import Image
        im = Image.open(io.BytesIO(base64.b64decode(uri.split(",", 1)[1])))
        w, h = im.size
    except Exception:
        pass
    d = {"img": uri, "w": w, "h": h}
    if video:
        d["video"] = video
    return d


# path no formato "AAAA-MM/arquivo.jpg" = veio do Supabase (comite_conteudo,
# tools/migrar_comite_conteudo.py ou upload pelo hub). Formato antigo local
# ("fotos/arquivo.jpg", só em conteúdo ainda não migrado) lê do disco.
_RE_PATH_BUCKET = re.compile(r"^\d{4}-\d{2}/")

# Vídeo entra na mesma grade das fotos, mas o que se embute no spec é o POSTER
# (`<path>.poster.jpg`, gravado junto no upload) — nunca o arquivo: um mp4 em
# base64 no spec.js deixaria o deck impossível de carregar. O item vira
# {"img": <poster>, "video": <path>}; o deck monta o link a partir do path.
_RE_VIDEO = re.compile(r"\.(mp4|webm|mov|m4v|ogv)$", re.I)


def _fotos_grupo_por_tema(grupos, m, ano):
    """Um ou mais slides por TEMA, igual ao deck oficial — 'Obras e melhorias
    realizadas · Banqueta' com 1-6 fotos, nunca misturando tema no mesmo slide.
    Tema com mais de FOTOS_POR_SLIDE vira '(1/2)' só dentro dele mesmo."""
    out = []
    for g in grupos:
        tema, arquivos = g.get("tema") or "", g.get("arquivos") or []
        embutidas = []
        for f in arquivos:
            if _RE_VIDEO.search(f):
                # sem poster o item ficaria sem nada pra mostrar no slide baked;
                # o conteúdo ao vivo ainda o resgata com a marca de play
                poster = _foto_do_bucket(f + ".poster.jpg") if _RE_PATH_BUCKET.match(f) else None
                if poster:
                    embutidas.append(_item_foto(poster, f))
                continue
            uri = _foto_do_bucket(f) if _RE_PATH_BUCKET.match(f) else _foto_embutida(OUT / f)
            if uri:
                embutidas.append(_item_foto(uri))
        if not embutidas:
            continue
        n = (len(embutidas) + FOTOS_POR_SLIDE - 1) // FOTOS_POR_SLIDE
        for k in range(n):
            bloco = embutidas[k * FOTOS_POR_SLIDE:(k + 1) * FOTOS_POR_SLIDE]
            cols, rows = GRADE_FOTOS[len(bloco)]
            sub = f"Obras e melhorias realizadas  ·  {tema}" if tema else f"Registros de {MESES[m-1]} {ano}"
            if n > 1:
                sub += f" ({k+1}/{n})"
            out.append({"t": "fotos", "n": 39, "titulo": "MANEJO — FOTOS E REGISTROS",
                        "sub": sub, "grade": [cols, rows], "fotos": bloco})
    return out


def slides_fotos(c, m, ano):
    """Fotos do mês pedido. Sem herdar mês anterior — é isso que estava errado.

    `c` é o conteúdo do mês, e o resolvedor dele cai no mês anterior mais recente
    quando o pedido não existe. Para texto isso é razoável (o histórico de manejo é
    cumulativo); para foto, não: o deck de janeiro saía com as fotos de junho.

    Ordem de fonte, ao contrário do dashboard semanal:
      1. comite_conteudo.json — extraído do PRÓPRIO deck oficial daquele mês
         (tools/extrair_conteudo.py), com o agrupamento por TEMA que o comitê usa.
      2. pasta do Drive (ATA & APRESENTACOES MENSAIS/<ano>/FOTOS) — só quando o
         mês ainda não foi extraído. Achado em 28/08/2026: a pasta do Drive pra
         julho/2026 tem 1 foto só (é alimentada semana a semana, sem curadoria);
         o deck oficial de julho tem 47 fotos em 13 temas. Pra ESTE deck, o
         extraído do oficial é o que importa — a pasta do Drive é o fallback
         genérico, não o contrário."""
    legado = c.get("fotos") or []
    if legado and isinstance(legado[0], dict):
        # formato novo: [{tema, arquivos}], vindo do Supabase (hub) ou extraído
        # do PPTX oficial do mês
        out = _fotos_grupo_por_tema(legado, m, ano)
        if out:
            _registra("fotos do mês", CONTEUDO)
            n_fotos = sum(len(g["fotos"]) for g in out)
            print(f"  [fotos] {MESES[m-1]}: {n_fotos} fotos em "
                  f"{len({g['sub'].split(' (')[0] for g in out})} temas "
                  f"(comite_conteudo.json)")
            return out
    if legado and isinstance(legado[0], str) and _RE_PATH_BUCKET.match(legado[0]):
        # lista solta MAS já migrada pro Supabase (junho/2026, formato antigo
        # sem tema — ver tools/migrar_comite_conteudo.py): trata como 1 grupo
        # sem tema, mesmo caminho do formato novo (baixa do bucket).
        out = _fotos_grupo_por_tema([{"tema": "", "arquivos": legado}], m, ano)
        if out:
            _registra("fotos do mês", CONTEUDO)
            n_fotos = sum(len(g["fotos"]) for g in out)
            print(f"  [fotos] {MESES[m-1]}: {n_fotos} fotos (comite_conteudo.json, sem tema)")
            return out

    caminhos = _fotos_do_mes(ano, m)
    fonte = FOTOS_DIR_BASE / str(ano) / "FOTOS"
    if caminhos:
        _registra("fotos do mês", fonte)
    else:
        # legado antigo local (pré-migração): lista solta de arquivos, sem
        # tema, ainda apontando pro disco (não pro bucket).
        caminhos = [OUT / f for f in legado if isinstance(f, str)]
        if caminhos:
            print(f"  [fotos] {MESES[m-1]}: pasta do Drive sem foto, usando as "
                  f"{len(caminhos)} do comite_conteudo.json (formato antigo, sem tema)")
    if not caminhos:
        return [pend(39, "MANEJO — FOTOS E REGISTROS", f"Registros de {MESES[m-1]}",
                     caminho_curto(fonte),
                     f"nenhuma foto de {MESES[m-1].lower()}/{str(ano)[2:]} na pasta")]
    # troca o caminho pela imagem embutida; some quem nao carregou
    embutidas, kb = [], 0
    for f in caminhos:
        uri = _foto_embutida(f)
        if uri:
            embutidas.append(_item_foto(uri))
            kb += len(uri) // 1024
    if not embutidas:
        return [pend(39, "MANEJO — FOTOS E REGISTROS", f"Registros de {MESES[m-1]}",
                     caminho_curto(fonte), "nenhuma das fotos do mês pôde ser lida")]
    dias = sorted({_foto_quando(f).day for f in caminhos})
    print(f"  [fotos] {MESES[m-1]}: {len(embutidas)} de {len(caminhos)} "
          f"({kb // 1024 or 1} MB) — dias {', '.join(f'{d:02d}' for d in dias)}")
    fs = embutidas
    n = (len(fs) + FOTOS_POR_SLIDE - 1) // FOTOS_POR_SLIDE
    out = []
    for k in range(n):
        bloco = fs[k * FOTOS_POR_SLIDE:(k + 1) * FOTOS_POR_SLIDE]
        cols, rows = GRADE_FOTOS[len(bloco)]
        out.append({"t": "fotos", "n": 39, "titulo": "MANEJO — FOTOS E REGISTROS",
                    "sub": f"Registros de {MESES[m-1]} {ano}" + (f" · {k+1}/{n}" if n > 1 else ""),
                    "grade": [cols, rows], "fotos": bloco})
    return out


# ==================================================================== deck
def divisor(n, titulo, sub):
    return {"t": "divisor", "n": n, "titulo": titulo, "sub": sub}


# Contratos de embrião: com a altura de linha que o haras pediu (0,27in) cabem
# 16 por slide, que é o que o relatório de julho tem no maior deles. O desenho
# encolhe a linha até 23; passou disso, continua noutro slide.
MAX_CONTRATOS = 23


def divide_contratos(slide):
    rows = slide.get("rows") or []
    if len(rows) <= MAX_CONTRATOS:
        return [slide]
    n = (len(rows) + MAX_CONTRATOS - 1) // MAX_CONTRATOS
    passo = (len(rows) + n - 1) // n
    return [dict(slide, rows=rows[k * passo:(k + 1) * passo],
                 titulo=f"{slide['titulo']} ({k + 1}/{n})") for k in range(n)]


def so_mensal(slides):
    """Marca os slides que recortam o MÊS — os que não vão no trimestral.

    Marcar em vez de montar um segundo deck: o trimestral é o mesmo conteúdo com
    um recorte a menos, e duplicar o build significaria manter duas versões da
    mesma regra."""
    for x in (slides if isinstance(slides, list) else [slides]):
        x["so_mensal"] = True
    return slides if isinstance(slides, list) else [slides]


def oculto(slide):
    """Slide que existe no arquivo mas não entra na apresentação — o 'ocultar
    slide' do PowerPoint. O relatório dela esconde o acumulado do Haras, o da
    Casa e os comentários; o deck mantém os três (quem quiser, navega até eles)
    e o PPTX sai com eles marcados como ocultos."""
    slide["oculto"] = True
    return slide


def monta_deck(m, ano, ctx):
    cont = conteudo_do_mes(ctx["conteudo"], f"{ano}-{m:02d}")
    safra = safra_do_deck(ano, m)
    if safra not in ctx["estacao_por_safra"]:
        ctx["estacao_por_safra"][safra] = (slides_estacao(safra), slide_coberturas(safra))
    est_slides, cob_slide = ctx["estacao_por_safra"][safra]
    MES, yy = MESES[m - 1].upper(), str(ano)[2:]
    s = [
        {"t": "capa", "titulo": "RELATÓRIO DE DESEMPENHO ESTRATÉGICO",
         "mes": f"{MES} / {ano}", "org": "HARAS PAO GRANDE"},
        {"t": "agenda", "titulo": "AGENDA",
         "sub": f"RELATÓRIO DESEMPENHO ESTRATÉGICO — {MES}/{yy}",
         "itens": [{"n": "01", "titulo": "FINANCEIRO", "sub": "DRE Haras · Caixa · Plantel"},
                   {"n": "02", "titulo": "ESTAÇÃO DE MONTA", "sub": "Embriões · Doadoras · Garanhões"},
                   {"n": "03", "titulo": "EXPOSIÇÕES", "sub": "Programação e resultados"},
                   {"n": "04", "titulo": "VENDAS", "sub": "Pipeline e contratos"},
                   {"n": "05", "titulo": "DECISÕES E MANEJO", "sub": "Plantel · Obras · Casa"}]},
        slide_pendencias(cont, m, ano),
        divisor(1, "FINANCEIRO", f"DRE Haras  ·  Caixa  ·  Plantel  |  {MES} {ano}"),
    ]
    # O relatório abrevia o mês no resumo do Haras ("JUL/26") e escreve por
    # extenso no Caixa e na Casa ("JULHO/26"). É o título que o haras conhece.
    mesano, mesano_ext = f"{ABR[m-1].upper()}/{yy}", f"{MES}/{yy}"

    def dre(n, titulo, sub, linhas, layout):
        if not linhas:
            return pend(n, titulo, sub, "DRE_Historico.xlsx", "sem linha para esse recorte no histórico")
        return {"t": "dre", "n": n, "titulo": titulo, "sub": sub, "layout": layout, "linhas": linhas}

    face_comp = _na_ordem_oficial(dre_mes("HPG", "Competência", ano, m),
                                  gabarito(DRE_HARAS, "Real x Orçado (Comp)"))
    s += so_mensal(dre(4, f"RESUMO FINANCEIRO — HARAS COMPETÊNCIA — ORÇADO X REALIZADO {mesano}",
                       "DRE 2026 | HPG  ·  Competência Mensal  ·  Fonte: aba Real x Orçado (Comp)",
                       linhas_face(face_comp, GAB_RESUMO), "resumo"))
    for k, pag in enumerate(ANALISE_CUSTOS, 1):
        s += so_mensal(dre(5, f"ANÁLISE DE CUSTOS — {MES} {ano}",
                           f"DRE Haras  ·  Custos Indiretos de Produção  ·  Fonte: DRE-Compet  ·  "
                           f"Parte {k} de {len(ANALISE_CUSTOS)}",
                           linhas_analise(ano, m, pag), "analise"))
    for k, pag in enumerate(ANALISE_DESPESAS, 1):
        s += so_mensal(dre(6, f"ANÁLISE DE DESPESAS — {MES} {ano}",
                           f"DRE Haras  ·  Despesas Operacionais  ·  Fonte: DRE-Compet  ·  "
                           f"Parte {k} de {len(ANALISE_DESPESAS)}",
                           linhas_analise(ano, m, pag), "analise"))
    face_ytd = _na_ordem_oficial(dre_ytd("HPG", "Competência", ano, m),
                                 gabarito(DRE_HARAS, "Real x Orçado (Comp)"))
    s.append(oculto(dre(7, f"HARAS COMPETÊNCIA — ACUMULADO JAN–{ABR[m-1].upper()} {ano} (YTD)",
                        f"DRE 2026 | HPG  ·  Competência  ·  Janeiro a {MESES[m-1]}  ·  Fonte: Base YTD",
                        linhas_face(face_ytd, GAB_RESUMO), "resumo")))
    com = slide_comentarios(cont, m, ano)
    s.append(oculto(com) if com["t"] == "pendente" else com)
    s.append(slide_investimentos(m, ano))
    face_cx = _na_ordem_oficial(dre_mes("HPG", "Caixa", ano, m), gabarito(DRE_HARAS, "Real x Orçado (Caixa)"))
    s += so_mensal(dre(10, f"HARAS CAIXA — ORÇADO X REALIZADO {mesano_ext}",
                       "FC 2026 | HPG  ·  Caixa Mensal  ·  Dados confirmados",
                       linhas_face(face_cx, GAB_CAIXA), "caixa"))
    s.append(slide_estoque(m, ano))
    s.append(slide_movimentacao(m, ano))
    face_casa = _na_ordem_oficial(dre_mes("FPG", "Caixa", ano, m), gabarito(DRE_CASA, "Real x Orçado"))
    s += so_mensal(dre(13, f"RESUMO FINANCEIRO — CASA/FPG — ORÇADO X REALIZADO {mesano_ext}",
                       f"FC {ano} | FPG  ·  Caixa  ·  {MESES[m-1]} {ano}  ·  Fonte: aba Real x Orçado",
                       linhas_face(face_casa, GAB_CASA, CASA_EXTRAS, "Tributos"), "casa"))
    face_casa_ytd = _na_ordem_oficial(dre_ytd("FPG", "Caixa", ano, m), gabarito(DRE_CASA, "Real x Orçado"))
    s.append(oculto(dre(14, f"CASA/FPG — ORÇADO X REALIZADO ACUMULADO JAN–{ABR[m-1].upper()} {ano}",
                        f"FC {ano} | FPG  ·  Caixa  ·  Janeiro a {MESES[m-1]}  ·  Fonte: Base YTD",
                        linhas_face(face_casa_ytd, GAB_CASA, CASA_EXTRAS, "Tributos"), "casa")))

    s.append(divisor(2, "ESTAÇÃO DE MONTA", f"Embriões  ·  Doadoras  ·  Garanhões  |  {safra}"))
    s += est_slides
    s.append(cob_slide)

    s.append(divisor(3, "EXPOSIÇÕES", f"Programação  ·  Resultados  |  {MES} {ano}"))
    s += slides_exposicoes(cont, ano)

    s.append(divisor(4, "VENDAS", f"Pipeline Comercial  ·  Contratos {ano}  |  {MES} {ano}"))
    s += slides_vendas(m, ano)
    s.append(slide_inadimplencia(m, ano))
    for x in slides_embrioes(m, ano):
        s += divide_contratos(x)

    s.append(divisor(5, "DECISÕES E MANEJO", f"Plantel  ·  Obras  ·  Casa  |  {MES} {ano}"))
    # a contagem por local não está no relatório dela; fica no arquivo, oculta
    s.append(oculto(slide_contagem(m, ano)))
    s.append(slide_manejo(cont, m, ano))
    s += slides_fotos(cont, m, ano)
    s.append({"t": "encerramento", "titulo": "HARAS PAO GRANDE",
              "sub": f"Relatório de Desempenho Estratégico  ·  {MES} {ano}"})
    return s


# comite.html é estático e carrega deck.js/deck.css/spec.js. Sem versão na URL, o
# navegador serve o cache e a mudança não aparece — o arquivo no disco está certo e a
# tela continua velha. O carimbo abaixo invalida o cache a cada build.
RE_ASSET = re.compile(r'((?:href|src)="assets/comite/(?:deck\.js|deck\.css|spec\.js|layout\.js))(?:\?v=\d+)?"')


def _versiona_assets():
    alvo = REPO / "comite.html"
    if not alvo.exists():
        return
    html = alvo.read_text(encoding="utf-8")
    v = datetime.now().strftime("%Y%m%d%H%M")
    novo = RE_ASSET.sub(lambda m: f'{m.group(1)}?v={v}"', html)
    if novo != html:
        alvo.write_text(novo, encoding="utf-8")
        print(f"  [cache] assets do deck versionados: v={v}")


def build(so_mes=None):
    ano = so_mes.year if so_mes else 2026
    ctx = {}
    if _dre_hist() is None:
        aviso(f"DRE_Historico.xlsx não encontrado em {DRE_DIR} — seção financeira fica "
              "pendente. É saída do LxDREdataExtractor.py, que grava na pasta dele mesmo")
        meses = []
    else:
        meses = meses_fechados(ano=ano)
        if meses:
            print(f"  [dre] meses fechados em {ano}: "
                  f"{', '.join(ABR[x-1] for x in meses)}")
    if not meses:
        meses = [so_mes.month] if so_mes else [date.today().month]
        aviso("nenhum mês com realizado no DRE — deck sai só com as bases não-financeiras")

    # a estação é da safra do MÊS DO DECK, e um build pode montar meses de safras
    # diferentes (julho é 25/26, agosto é 26/27) — resolve por safra, uma vez cada
    ctx["estacao_por_safra"] = {}
    ctx["conteudo"] = le_conteudo()

    alvo = [so_mes.month] if so_mes else meses
    decks = {}
    for m in alvo:
        decks[f"{ano}-{m:02d}"] = monta_deck(m, ano, ctx)

    chaves = sorted(decks)
    # As fontes que os resolvedores compartilhados registram (roster, receptoras,
    # controle mensal) ficam em PGSemanalReport._FONTES_USADAS; as deste módulo, em
    # _FONTES. A auditoria quer as duas na mesma lista.
    fontes = {}
    for rotulo, caminho in list(_FONTES_COMPARTILHADAS.items()) + list(_FONTES.items()):
        p = Path(caminho)
        try:
            quando = datetime.fromtimestamp(p.stat().st_mtime).isoformat(timespec="minutes")
        except OSError:
            quando = None
        fontes[rotulo] = {"arquivo": p.name, "caminho": caminho_curto(p),
                          "pasta": p.parent.name, "modificado": quando}

    payload = {"meses": chaves, "padrao": chaves[-1], "avisos": avisos,
               "labels": {k: f"{MESES[int(k[5:]) - 1]} {k[:4]}" for k in chaves},
               "fontes": fontes, "decks": decks}
    OUT.mkdir(parents=True, exist_ok=True)
    js = json.dumps(payload, ensure_ascii=False, separators=(",", ":"), default=_json_default)
    (OUT / "spec.json").write_text(js, encoding="utf-8")
    (OUT / "spec.js").write_text(f"window.COMITE_SPEC = {js};\n", encoding="utf-8")
    _versiona_assets()
    n = len(decks[chaves[-1]])
    p = sum(1 for x in decks[chaves[-1]] if x["t"] == "pendente")
    print(f"[comite] {len(chaves)} meses ({chaves[0]} … {chaves[-1]}) · {n} slides "
          f"({n - p} com conteúdo, {p} pendentes) · {len(js)//1024} KB -> assets/comite/spec.js")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        try:
            mm, aaaa = sys.argv[1].split("/")
            build(date(int(aaaa), int(mm), 1))
        except ValueError:
            sys.exit(f"mês inválido: {sys.argv[1]!r} — use MM/AAAA (ex.: 06/2026)")
    else:
        build()
