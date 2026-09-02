"""Gera o artefato da aba Plantel / Movimentação do hub.

O que este script faz é o que o Resumo Contábil era feito à mão: pega dois
snapshots mensais do plantel do haras, decompõe a variação de patrimônio animal
por animal, e cruza cada efeito com o log de MOVIMENTAÇÕES da própria planilha do
haras. O que casa vira movimento classificado; o que não casa vira fila para
alguém decidir na tela — é a parte que não se automatiza e nunca deve ser
adivinhada.

    saída: assets/plantel/spec.js  (window.PLANTEL_MOV = {...})

Fontes (nenhuma é escrita):
  base_plantel.parquet  — série mensal 2021→ construída pelo LxEtlPlantel, no
                          repo LuxorMonthlyP-CRoutines/PlantelHPG. É o histórico
                          fechado; o hub só lê.
  CONTROLE_DE_PLANTEL_..._<MES>.xlsx — planilha do haras no Drive. Serve para dois
                          fins: o snapshot do mês ABERTO (que ainda não entrou na
                          base) e a aba MOVIMENTAÇÕES, que é o log de ocorrências.

Decisões de modelagem, todas medidas contra os dados reais (02/09/2026):

1. `RECEPTORAS <n>` é UMA linha agregada com a contagem dentro do nome (cotas 1,
   valor ~R$249k). Quando o rebanho vai de 126 para 124 cabeças, a chave muda e o
   diff lê saída de R$261k + entrada de R$257k. No histórico isso gerou R$16,0M de
   `saiu_controle` e R$16,3M de `compra` falsos — 37% de todo o pendente. Aqui a
   chave é fixa (PSEUDO:RECEPTORAS): a variação de valor vira reavaliação, que é o
   que ela é.

2. Renome e nascimento trocam a identidade do animal. O log registra isso em texto
   ("MUDOU DE/O NOME - ESTAVA X", "NASCEU - ESTAVA COMO <cruzamento>"), e sem
   costurar as duas pontas um renome aparece como sumiço + produção do mesmo valor.
   A ponte é conservadora de propósito (ver `_pontes`): casar por prefixo curto
   fundia animais distintos e chegou a inventar R$4M de compra num teste.

3. Escopo contábil = sufixo exatamente `DA PAO GRANDE` ou `OUTRO` (o plantel da
   Carla). Os `DA PAO GRANDE - E xx%` são parceria com o Eduardo e ficam fora:
   com esse corte o saldo reproduz o Resumo Contábil manual, e a diferença que
   sobra é exatamente a comissão de 8,5% de 6 animais (R$28.432,50 em jul/26), que
   o mapa do haras embute no valor do animal e a nossa base guarda em coluna
   própria.

Uso:
    python tools/build_plantel_mov.py                 # mês aberto = o mais recente
    python tools/build_plantel_mov.py --mes 2026-08
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import unicodedata
from datetime import date, datetime
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent.parent
SAIDA = BASE_DIR / "assets" / "plantel" / "spec.js"

# A base histórica e o extrator vivem no repo das rotinas mensais. Import por
# caminho (e não cópia) de propósito: duplicar o ETL faria as duas cópias
# divergirem na primeira mudança de layout da planilha do haras.
PROTO_DIR = Path(r"C:\Users\Arthur\repos\LuxorMonthlyP-CRoutines\PlantelHPG")
BASE_PARQUET = PROTO_DIR / "base_plantel.parquet"
CHECKS_PARQUET = PROTO_DIR / "base_plantel_checks.parquet"

sys.path.insert(0, str(PROTO_DIR))
sys.path.insert(0, str(BASE_DIR / "scripts"))

ESCOPO = ("DA PAO GRANDE", "OUTRO")
PSEUDO = "PSEUDO:RECEPTORAS"
CAUSAS = ["compra", "inclusao_embriao", "producao", "venda", "morte", "doacao",
          "reaval", "transferencia", "saiu_controle"]
# causa -> linha do Resumo Contábil manual, para a tela sair na ordem que o comitê
# já conhece
LINHAS_RESUMO = [
    ("compras", ["compra"]),
    ("producao_embrioes", ["inclusao_embriao", "producao"]),
    ("baixa_vendas", ["venda"]),
    ("baixa_mortes_doacoes", ["morte", "doacao"]),
    ("reavaliacoes", ["reaval"]),
    ("sem_classificacao", ["saiu_controle", "transferencia"]),
]


def _norm(s) -> str:
    s = "" if s is None else str(s)
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip().upper().strip('"')


def _f(x) -> float:
    try:
        v = float(x)
        return 0.0 if pd.isna(v) else v
    except (TypeError, ValueError):
        return 0.0


def _chave(nome, letra="") -> str:
    n = _norm(nome)
    if n.startswith("RECEPTORAS"):
        return PSEUDO
    return f"{n}|{_norm(letra)}"


# ===================== log do haras =====================
RX_PONTE = re.compile(r"ESTAVA (?:COMO )?\"?(.+?)\"?(?: - | PASSOU | E O | E A |, |$)")
# Ordem importa: 'NASCEU ... ESTAVA <cruzamento>' é troca de IDENTIDADE (o valor do
# embrião já estava no plantel), não entrada financeira nova.
REGRAS_LOG = [
    ("1-nome", "renome", r"MUDOU DE NOME|MUDOU O NOME|ADICAO DE SUFIXO|MUDOU DE SUFIXO"),
    ("1-nome", "nasceu_renomeia", r"NASCEU.*ESTAVA"),
    ("2-local", "mudanca_local", r"MUDOU O LOCAL|MUDOU DE LOCAL|FOI PARA O CENTRO DE TREINAMENTO"),
    ("3-financeira", "inclusao_embriao", r"EMBRIAO CONFIRMADO|EMBRIAO COMPRADO|INSERIDO NO PLANTEL"),
    ("3-financeira", "nascimento", r"NASCEU"),
    ("3-financeira", "morte", r"MORREU|OBITO|ABORTOU"),
    ("3-financeira", "doacao", r"\bDOAD|DOACAO"),
    ("3-financeira", "venda", r"VENDID|VENDA|COMPRADOR"),
    ("3-financeira", "compra", r"\bCOMPRA|COMPRADO"),
    ("3-financeira", "cota_valor", r"REAVALIACAO|ALTEROU O VALOR|MUDOU A %|MUDOU A PORCENTAGEM|"
                                   r"ZEROU A COTA|ZERADA A %"),
    ("3-financeira", "saida", r"SAIU DO HARAS|DEVOLVID|ENTREGUE"),
]


def classifica_log(ocorrencia: str) -> tuple[str, str]:
    """(tipo, subtipo) na taxonomia do controle: 1-Nome, 2-Local, 3-Financeira."""
    o = _norm(ocorrencia)
    for tipo, sub, rx in REGRAS_LOG:
        if re.search(rx, o):
            return tipo, sub
    return "nao-classificado", ""


def le_log(controle: Path) -> list[dict]:
    import PGSemanalReport as R

    wb = R._load(controle)
    aba = [s for s in wb.sheetnames if "MOVIMENT" in s.upper()][0]
    out = []
    for i, r in enumerate(wb[aba].iter_rows(values_only=True), start=1):
        if i < 3 or r[2] is None:
            continue
        d = r[3].date() if isinstance(r[3], datetime) else r[3]
        if not isinstance(d, date):
            continue
        oc = (R._s(r[4]) or "").strip()
        tipo, sub = classifica_log(oc)
        out.append({"linha": R._s(r[1]), "produto": R._s(r[2]) or "", "data": d,
                    "ocorrencia": oc, "tipo": tipo, "subtipo": sub,
                    "mes": f"{d.year:04d}-{d.month:02d}"})
    wb.close()
    out.sort(key=lambda x: x["data"])
    return out


# ===================== snapshots =====================
def prepara(df: pd.DataFrame) -> pd.DataFrame:
    d = df.copy()
    d["k"] = [_chave(n, l) for n, l in zip(d["nome"].fillna(""), d["letra"].fillna(""))]
    d["suf"] = d["sufixo"].fillna("").map(_norm)
    d["no_escopo"] = d["suf"].isin(ESCOPO)
    return d


def snapshot_mes_aberto(controle: Path, ano: int, mes: int, log: list[dict],
                        anterior: pd.DataFrame) -> pd.DataFrame:
    """Snapshot do mês que ainda não fechou, a partir da planilha de trabalho.

    A planilha de trabalho do haras é editada durante o mês SEGUINTE, então ela
    já traz movimentação posterior ao fim do mês (em 02/09/2026 eram 15 linhas,
    incluindo 13 doações para o Mato Grosso feitas em 01/09). Essas linhas são
    revertidas por carry-forward do mês anterior — senão o fechamento do mês sai
    contaminado com o que aconteceu depois dele.
    """
    import LxEtlPlantel as E

    raw, _totais = E._read_sheet(controle)
    atual = prepara(E._enrich(raw, ano, mes, controle.name))
    fim = date(ano, mes, 1) + pd.offsets.MonthEnd(0)
    fim = fim.date() if hasattr(fim, "date") else fim

    depois = {_norm(x["produto"]) for x in log
              if x["data"] > fim
              and re.search(r"DOADOS PARA O MATO GROSSO|SAIU DO HARAS|VENDIDO E ENTREGUE|\bDOAD",
                            _norm(x["ocorrencia"]))}
    if not depois:
        return atual, 0
    ant = anterior.set_index("k")
    revertidos = 0
    for i in atual.index:
        k, nome = atual.at[i, "k"], _norm(atual.at[i, "nome"])
        if nome not in depois or k not in ant.index:
            continue
        linha = ant.loc[k]
        if isinstance(linha, pd.DataFrame):
            linha = linha.iloc[0]
        for c in ("status_plantel", "cotas", "valor_100", "patrimonio_proporcional",
                  "local", "comissao"):
            atual.at[i, c] = linha[c]
        revertidos += 1
    return atual, revertidos


# ===================== pontes de identidade =====================
def _indice(d: pd.DataFrame) -> dict:
    ix = {}
    for n, k in zip(d["nome"].fillna(""), d["k"]):
        if n:
            ix.setdefault(_norm(n), k)
    return ix


def _acha(nome_norm: str, ix: dict):
    """Exato, ou prefixo ÚNICO de 25+ caracteres.

    O nome no texto do log vem truncado ("ESTAVA JAVA DA PAO GRANDE X QUEBRUTO DE
    ALCATEIA 17/09/20") e a base tem o nome inteiro, então casamento exato perde
    pontes reais. Mas prefixo curto é pior: com 18 caracteres 'ADRENALINA DA PAO
    GRANDE' casa com 'ADRENALINA DA PAO GRANDE X ...' e o teste inventou R$4M de
    compra e R$2,4M de venda. 25 + unicidade foi o ponto onde nenhum falso
    positivo sobrou.
    """
    if nome_norm in ix:
        return ix[nome_norm]
    if len(nome_norm) < 25:
        return None
    cands = {k for real, k in ix.items()
             if real.startswith(nome_norm[:25]) or nome_norm.startswith(real[:25])}
    return next(iter(cands)) if len(cands) == 1 else None


def _pontes(log: list[dict], ant: pd.DataFrame, atual: pd.DataFrame) -> tuple[dict, int]:
    """{chave_no_mes_anterior: chave_no_mes_atual} para renome/nascimento."""
    brutas = {}
    for x in log:
        o = _norm(x["ocorrencia"])
        if "ESTAVA" not in o:
            continue
        m = RX_PONTE.search(o + " $")
        if not m:
            continue
        antigo, novo = _norm(m.group(1)), _norm(x["produto"])
        if antigo and antigo != novo and len(antigo) >= 6:
            brutas[antigo] = novo

    ix_ant, ix_atual = _indice(ant), _indice(atual)
    ks_ant, ks_atual = set(ant["k"]), set(atual["k"])
    remap, recusadas = {}, 0
    for antigo, novo in brutas.items():
        k_old, k_new = _acha(antigo, ix_ant), _acha(_norm(novo), ix_atual)
        if not (k_old and k_new) or k_old == k_new:
            continue
        # Renome de verdade: a identidade antiga desapareceu e a nova não existia.
        # Se as duas convivem nos dois snapshots, são animais distintos que o
        # prefixo aproximou — recusa e deixa a diferença aparecer na fila.
        if k_old in ks_atual or k_new in ks_ant:
            recusadas += 1
            continue
        remap[k_old] = k_new
    return remap, recusadas


# ===================== efeitos =====================
def agrega(d: pd.DataFrame) -> pd.DataFrame:
    return d.groupby("k").agg(nome=("nome", "first"), suf=("suf", "first"),
                              no_escopo=("no_escopo", "max"), cotas=("cotas", "sum"),
                              valor=("valor_100", "first"),
                              patr=("patrimonio_proporcional", "sum"),
                              status=("status_plantel", "first"),
                              categoria=("categoria", "first"),
                              comissao=("comissao", "sum"))


def efeitos_do_mes(ant: pd.DataFrame, atual: pd.DataFrame, mes: str,
                   log_mes: list[dict]) -> list[dict]:
    """Decompõe a variação de patrimônio de cada animal em quantidade × preço.

    Δcotas × preço médio  = mudou a fatia (compra/venda/morte/doação)
    Δvalor × cotas médias = mudou o preço  (reavaliação)
    Os dois efeitos podem coexistir no mesmo animal no mesmo mês.
    """
    remap, recusadas = _pontes(log_mes, ant, atual)
    ant = ant.copy()
    ant["k"] = ant["k"].map(lambda k: remap.get(k, k))

    ga, gp = agrega(atual), agrega(ant)
    ka, kp = set(ga.index), set(gp.index)

    por_nome = {}
    for x in log_mes:
        por_nome.setdefault(_norm(x["produto"]), []).append(x)

    def log_de(nome):
        n = _norm(nome)
        itens = list(por_nome.get(n, []))
        if not itens and len(n) >= 25:
            for chave_log, l in por_nome.items():
                if chave_log.startswith(n[:25]) or n.startswith(chave_log[:25]):
                    itens += l
        return itens

    def resumo_log(itens):
        return [{"data": x["data"].isoformat(), "tipo": x["tipo"], "subtipo": x["subtipo"],
                 "ocorrencia": x["ocorrencia"]} for x in itens]

    out = []
    for k in ka & kp:
        v0, v1 = _f(gp.at[k, "valor"]), _f(ga.at[k, "valor"])
        q0, q1 = _f(gp.at[k, "cotas"]), _f(ga.at[k, "cotas"])
        nome = ga.at[k, "nome"]
        itens = log_de(nome)
        subs = {x["subtipo"] for x in itens}
        comum = dict(mes=mes, k=k, nome=nome, sufixo=ga.at[k, "suf"],
                     no_escopo=bool(ga.at[k, "no_escopo"]), categoria=ga.at[k, "categoria"],
                     valor_ant=round(v0, 2), valor_atual=round(v1, 2),
                     cotas_ant=q0, cotas_atual=q1,
                     status_ant=gp.at[k, "status"], status_atual=ga.at[k, "status"],
                     log=resumo_log(itens))
        # Saída/entrada TOTAL não se decompõe. Quando a cota vai a zero o haras
        # zera o valor junto, e a decomposição preço×quantidade partia a baixa em
        # metade `venda` e metade `reaval` — QUANTICO saindo por R$36k aparecia
        # como R$18k de venda + R$18k de reavaliação. O mapa do haras joga 100%
        # na baixa (ALTEZA N19: valor 250k, baixa venda −250k, final 0), e é isso
        # que o Resumo Contábil espera.
        if (q0 and not q1) or (q1 and not q0):
            total = round(q1 * v1 - q0 * v0, 2)
            if total:
                if q1:
                    tipo = "inclusao_embriao" if "inclusao_embriao" in subs else "compra"
                else:
                    tipo = ("morte" if "morte" in subs else "doacao" if "doacao" in subs else "venda")
                out.append({**comum, "tipo": tipo, "efeito": total})
            continue

        ef_qtd = (q1 - q0) * ((v0 + v1) / 2)
        ef_prc = (v1 - v0) * ((q0 + q1) / 2)
        if round(ef_qtd, 2):
            if q1 < q0:
                tipo = ("morte" if "morte" in subs else "doacao" if "doacao" in subs else "venda")
            else:
                tipo = "inclusao_embriao" if "inclusao_embriao" in subs else "compra"
            out.append({**comum, "tipo": tipo, "efeito": round(ef_qtd, 2)})
        if round(ef_prc, 2):
            out.append({**comum, "tipo": "reaval", "efeito": round(ef_prc, 2)})

    for k in ka - kp:
        v, q = _f(ga.at[k, "valor"]), _f(ga.at[k, "cotas"])
        nome = ga.at[k, "nome"]
        itens = log_de(nome)
        subs = {x["subtipo"] for x in itens}
        cat = _norm(ga.at[k, "categoria"])
        tipo = ("inclusao_embriao" if ("inclusao_embriao" in subs or "EMBRI" in cat)
                else "producao" if ("nascimento" in subs or "POTR" in cat) else "compra")
        out.append(dict(mes=mes, k=k, nome=nome, sufixo=ga.at[k, "suf"],
                        no_escopo=bool(ga.at[k, "no_escopo"]), categoria=ga.at[k, "categoria"],
                        tipo=tipo, efeito=round(v * q, 2), valor_ant=0.0,
                        valor_atual=round(v, 2), cotas_ant=0.0, cotas_atual=q,
                        status_ant="(novo)", status_atual=ga.at[k, "status"],
                        log=resumo_log(itens)))

    for k in kp - ka:
        v, q = _f(gp.at[k, "valor"]), _f(gp.at[k, "cotas"])
        nome = gp.at[k, "nome"]
        itens = log_de(nome)
        subs = {x["subtipo"] for x in itens}
        tipo = ("venda" if "venda" in subs else "doacao" if "doacao" in subs
                else "morte" if "morte" in subs else "saiu_controle")
        out.append(dict(mes=mes, k=k, nome=nome, sufixo=gp.at[k, "suf"],
                        no_escopo=bool(gp.at[k, "no_escopo"]), categoria=gp.at[k, "categoria"],
                        tipo=tipo, efeito=round(-v * q, 2), valor_ant=round(v, 2),
                        valor_atual=0.0, cotas_ant=q, cotas_atual=0.0,
                        status_ant=gp.at[k, "status"], status_atual="(sumiu)",
                        log=resumo_log(itens)))
    return out, recusadas, remap


# ===================== montagem =====================
def monta(mes_aberto: str | None) -> dict:
    import LxEtlPlantel as E  # noqa: F401  (garante que o import por caminho funciona)

    if not BASE_PARQUET.exists():
        raise SystemExit(f"base não encontrada: {BASE_PARQUET}\n"
                         "Rode o LxEtlPlantel no repo LuxorMonthlyP-CRoutines primeiro.")
    b = pd.read_parquet(BASE_PARQUET)
    b["ym"] = b.mes_referencia.dt.to_period("M").astype(str)
    fechados = sorted(b["ym"].unique())
    fechado_ate = fechados[-1]

    controle = _acha_controle()
    log = le_log(controle)

    # mês aberto = o seguinte ao último fechado, salvo pedido explícito
    if not mes_aberto:
        ano, mm = map(int, fechado_ate.split("-"))
        ano, mm = (ano + 1, 1) if mm == 12 else (ano, mm + 1)
        mes_aberto = f"{ano:04d}-{mm:02d}"

    snaps = {m: prepara(b[b.ym == m]) for m in fechados}
    ano, mm = map(int, mes_aberto.split("-"))
    aberto, revertidos = snapshot_mes_aberto(controle, ano, mm, log, snaps[fechado_ate])
    snaps[mes_aberto] = aberto
    ordem = fechados + [mes_aberto]

    # `cadeia` acumula as pontes de todos os meses: um animal renomeado duas vezes
    # no ano tem de ser seguido até a identidade final, senão a abertura anual
    # (valor inicial -> valor final por animal) abre em dois — o velho zerando e o
    # novo aparecendo do nada, como acontecia em 49 dos 270 animais.
    todos_efeitos, recusadas_tot, cadeia = [], 0, {}
    for i in range(1, len(ordem)):
        m = ordem[i]
        log_mes = [x for x in log if x["mes"] == m]
        efs, rec, remap = efeitos_do_mes(snaps[ordem[i - 1]], snaps[m], m, log_mes)
        todos_efeitos += efs
        recusadas_tot += rec
        cadeia.update(remap)

    ef = pd.DataFrame(todos_efeitos)
    return _artefato(ef, snaps, ordem, log, mes_aberto, fechado_ate, controle,
                     revertidos, recusadas_tot, cadeia)


def _acha_controle() -> Path:
    """Planilha de trabalho/fechamento mais recente do haras, no Drive."""
    from _pg_common import DRIVE_ROOT

    base = DRIVE_ROOT / "PLANTEL"
    cands = [f for d in sorted(base.glob("Estação *"), reverse=True)
             for f in d.glob("*CONTROLE_DE_PLANTEL_PAO_GRANDE_*.xlsx")
             if not f.name.startswith("~$")]
    if not cands:
        raise SystemExit(f"Nenhum CONTROLE_DE_PLANTEL em {base}")
    return max(cands, key=lambda f: f.stat().st_mtime)


def _artefato(ef, snaps, ordem, log, mes_aberto, fechado_ate, controle,
              revertidos, recusadas, cadeia=None) -> dict:
    cadeia = cadeia or {}

    def identidade_final(k):
        visto = set()
        while k in cadeia and k not in visto:
            visto.add(k)
            k = cadeia[k]
        return k

    resumo, saldos = {}, {}
    for m in ordem:
        d = snaps[m]
        e = d[d.no_escopo]
        saldos[m] = {"patr": round(float(e["patrimonio_proporcional"].sum()), 2),
                     "comissao": round(float(e["comissao"].sum()), 2),
                     "cabecas": int(len(e))}
    for i in range(1, len(ordem)):
        m, ant = ordem[i], ordem[i - 1]
        e = ef[(ef.mes == m) & (ef.no_escopo)] if len(ef) else ef
        causas = {c: round(float(e[e.tipo == c].efeito.sum()), 2) for c in CAUSAS} if len(e) else \
                 {c: 0.0 for c in CAUSAS}
        linhas = {rot: round(sum(causas[c] for c in cs), 2) for rot, cs in LINHAS_RESUMO}
        ini, fim = saldos[ant]["patr"], saldos[m]["patr"]
        resumo[m] = {"saldo_ini": ini, "saldo_fim": fim, "causas": causas, "linhas": linhas,
                     "erro_identidade": round(ini + sum(causas.values()) - fim, 2),
                     "comissao_ini": saldos[ant]["comissao"], "comissao_fim": saldos[m]["comissao"],
                     "cabecas": saldos[m]["cabecas"],
                     "fechado": m <= fechado_ate}

    # abertura por animal do ano corrente: valor inicial + causas = valor final,
    # no mesmo formato do mapa que o haras mantém à mão
    ano = mes_aberto[:4]
    meses_ano = [m for m in ordem if m.startswith(ano)]
    m0 = ordem[ordem.index(meses_ano[0]) - 1] if ordem.index(meses_ano[0]) else meses_ano[0]
    ini_snap = snaps[m0].copy()
    ini_snap["k"] = ini_snap["k"].map(identidade_final)     # segue renome/nascimento
    ini_df, fim_df = agrega(ini_snap), agrega(snaps[ordem[-1]])
    aberturas = []
    e_ano = ef[(ef.mes >= meses_ano[0]) & (ef.no_escopo)] if len(ef) else ef
    for k in sorted(set(ini_df.index) | set(fim_df.index)):
        no_escopo = (bool(fim_df.at[k, "no_escopo"]) if k in fim_df.index
                     else bool(ini_df.at[k, "no_escopo"]))
        if not no_escopo:
            continue
        sel = e_ano[e_ano.k == k] if len(e_ano) else []
        linha = {"k": k,
                 "nome": (fim_df.at[k, "nome"] if k in fim_df.index else ini_df.at[k, "nome"]),
                 "sufixo": (fim_df.at[k, "suf"] if k in fim_df.index else ini_df.at[k, "suf"]),
                 "categoria": (fim_df.at[k, "categoria"] if k in fim_df.index else ini_df.at[k, "categoria"]),
                 "status": (fim_df.at[k, "status"] if k in fim_df.index else "(saiu)"),
                 "cota": _f(fim_df.at[k, "cotas"]) if k in fim_df.index else 0.0,
                 "valor_ini": round(_f(ini_df.at[k, "patr"]) if k in ini_df.index else 0.0, 2),
                 "valor_fim": round(_f(fim_df.at[k, "patr"]) if k in fim_df.index else 0.0, 2)}
        for rot, causas_rot in LINHAS_RESUMO:
            linha[rot] = round(float(sum(sel[sel.tipo == c].efeito.sum() for c in causas_rot)), 2) if len(sel) else 0.0
        linha["erro"] = round(linha["valor_ini"] + sum(linha[r] for r, _ in LINHAS_RESUMO)
                              - linha["valor_fim"], 2)
        if any(linha[r] for r, _ in LINHAS_RESUMO) or linha["valor_ini"] or linha["valor_fim"]:
            aberturas.append(linha)

    # fila: efeito sem explicação, e log financeiro sem efeito no patrimônio
    e_ab = ef[(ef.mes == mes_aberto) & (ef.no_escopo)] if len(ef) else pd.DataFrame()
    pendencias = []
    if len(e_ab):
        p = e_ab[(e_ab.tipo == "saiu_controle") & (e_ab.efeito.abs() >= 1)]
        pendencias = json.loads(p.to_json(orient="records"))

    nomes_com_efeito = {_norm(n) for n in (e_ab["nome"] if len(e_ab) else [])}
    divergencias = []
    for x in [y for y in log if y["mes"] == mes_aberto and y["tipo"] == "3-financeira"]:
        if _norm(x["produto"]) in nomes_com_efeito:
            continue
        d_fim = agrega(snaps[mes_aberto])
        k = _chave(x["produto"])
        estado = None
        for kk in d_fim.index:
            if _norm(d_fim.at[kk, "nome"]) == _norm(x["produto"]):
                estado = {"cota": _f(d_fim.at[kk, "cotas"]),
                          "patrimonio": round(_f(d_fim.at[kk, "patr"]), 2),
                          "status": d_fim.at[kk, "status"], "sufixo": d_fim.at[kk, "suf"]}
                break
        divergencias.append({"data": x["data"].isoformat(), "produto": x["produto"],
                             "subtipo": x["subtipo"], "ocorrencia": x["ocorrencia"],
                             "estado_no_plantel": estado})

    checks = []
    if CHECKS_PARQUET.exists():
        c = pd.read_parquet(CHECKS_PARQUET)
        c["mes"] = c.mes_referencia.dt.to_period("M").astype(str)
        checks = json.loads(c.drop(columns=["mes_referencia"]).to_json(orient="records"))

    log_mes = [{"data": x["data"].isoformat(), "produto": x["produto"],
                "ocorrencia": x["ocorrencia"], "tipo": x["tipo"], "subtipo": x["subtipo"]}
               for x in log if x["mes"] == mes_aberto]

    return {
        "gerado_em": datetime.now().isoformat(timespec="seconds"),
        "mes_aberto": mes_aberto,
        "fechado_ate": fechado_ate,
        "meses": [m for m in ordem if m >= f"{ano}-01"],
        "escopo": list(ESCOPO),
        "fonte": {"controle": controle.name, "base": BASE_PARQUET.name,
                  "revertidos_pos_mes": revertidos, "pontes_recusadas": recusadas},
        "resumo": resumo,
        "log_mes": log_mes,
        "efeitos": json.loads(ef[ef.mes >= f"{ano}-01"].to_json(orient="records")) if len(ef) else [],
        "aberturas": aberturas,
        "pendencias": pendencias,
        "divergencias": divergencias,
        "checks": checks,
    }


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--mes", help="mês aberto no formato AAAA-MM")
    args = ap.parse_args()

    spec = monta(args.mes)
    SAIDA.parent.mkdir(parents=True, exist_ok=True)
    corpo = json.dumps(spec, ensure_ascii=False)
    # .js pra abrir a página direto na máquina (demo offline); .json é o que o
    # publish_hub sobe pro bucket privado e o porteiro injeta em produção
    SAIDA.write_text("window.PLANTEL_MOV = " + corpo + ";\n", encoding="utf-8")
    SAIDA.with_suffix(".json").write_text(corpo, encoding="utf-8")

    r = spec["resumo"].get(spec["mes_aberto"], {})
    print(f"mês aberto      : {spec['mes_aberto']} (fechado até {spec['fechado_ate']})")
    print(f"fonte           : {spec['fonte']['controle']}")
    print(f"  revertidos    : {spec['fonte']['revertidos_pos_mes']} linha(s) de movimentação posterior ao mês")
    print(f"  pontes recusadas: {spec['fonte']['pontes_recusadas']}")
    print(f"saldo           : {r.get('saldo_ini', 0):,.2f} -> {r.get('saldo_fim', 0):,.2f}")
    for rot, _ in LINHAS_RESUMO:
        v = r.get("linhas", {}).get(rot, 0)
        if v:
            print(f"  {rot:22s} {v:>16,.2f}")
    print(f"erro de identidade: {r.get('erro_identidade', 0):,.2f}")
    print(f"pendências      : {len(spec['pendencias'])}")
    print(f"divergências log x patrimônio: {len(spec['divergencias'])}")
    print(f"aberturas       : {len(spec['aberturas'])} animais")
    print(f"checks          : {sum(1 for c in spec['checks'] if c.get('ok'))}/{len(spec['checks'])} meses ok")
    print(f"→ {SAIDA.relative_to(BASE_DIR)} ({SAIDA.stat().st_size/1024:.0f} KB)")


if __name__ == "__main__":
    main()
