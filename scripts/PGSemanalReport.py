"""
PGSemanalReport — gera a ATUALIZAÇÃO SEMANAL do Haras Pão Grande a partir das
fontes REORGANIZADAS do Drive (layout 2026-07).

Substitui o pipeline antigo (PGSemanalExtractor/_pg_semanal), que apontava pra
paths e modelos de planilha que não existem mais.

Fontes (ver memória project-hpg-semanal):
  - Headcount   : ATUALIZACAO SEMANAL/CONTROLE PLANTEL.xlsx  aba CONTAGEM (pré-agregada)
  - Receptoras  : idem, aba 'RECEPTORAS ' (censo; filtro do report a definir)
  - Produção    : REPRODUÇÃO/ESTAÇÃO DE MONTA/Estação 2025-2026/{YYMMDD} ESTACAO DE MONTA.xlsx
                  aba ESTAÇÃO (embrião confirmado = coluna 60D == '+')
  - Movimentação: PLANILHAS SEMANAIS/SAIDA E ENTRADA DE ANIMAIS - MODELO ENVIAR NO GRUPO.xlsx
  - Comerciais  : REPRODUÇÃO/EMBRIOES A ENTREGAR - A RECEBER.xlsx  abas PAINEL/ENTREGAR/RECEBER

Uso:
    python PGSemanalReport.py                 # semana termina hoje
    python PGSemanalReport.py 17/07/2026       # semana termina nessa data
    python PGSemanalReport.py 03/07/2026 17/07/2026   # janela explícita

Saída:
    - imprime as 5 seções no console, comparadas aos alvos do docx 17-07 (quando aplicável)
    - grava semanal_data.json (consumido pelo dashboard HTML)
"""

from __future__ import annotations

import json
import re
import sys
import unicodedata
from dataclasses import dataclass, field, asdict
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl

from _pg_common import DRIVE_ROOT, ensure_cache

BASE_DIR = Path(__file__).resolve().parent.parent  # raiz do projeto (scripts/ fica 1 nível abaixo)

# ------------------------------------------------------------------
# Localização das fontes (layout novo do Drive)
# ------------------------------------------------------------------
CONTROLE_PLANTEL_SEMANAL = DRIVE_ROOT / "ATUALIZACAO SEMANAL" / "CONTROLE PLANTEL.xlsx"
SAIDA_ENTRADA = (
    DRIVE_ROOT / "PLANILHAS SEMANAIS"
    / "SAIDA E ENTRADA DE ANIMAIS - MODELO ENVIAR NO GRUPO.xlsx"
)
# "EMBRIÕES E MATRIZES - MODELO ENVIAR NO GRUPO 3 <DD-MM>.xlsx" — a planilha que o
# haras manda no grupo e que É a fonte do acumulado na estação. O sufixo de data
# muda, então resolvemos por glob (mais recente por mtime). Mora em ATUALIZACAO
# SEMANAL desde a reorganização do Drive (antes era PLANILHAS SEMANAIS).
EMB_MATRIZES_DIR = DRIVE_ROOT / "ATUALIZACAO SEMANAL"
EMB_MATRIZES_GLOB = "EMBRI*E MATRIZES*.xlsx"
EMB_COMERCIAIS = DRIVE_ROOT / "REPRODUÇÃO" / "EMBRIOES A ENTREGAR - A RECEBER.xlsx"
ESTACAO_MASTER_DIR = (
    DRIVE_ROOT / "REPRODUÇÃO" / "ESTAÇÃO DE MONTA" / "Estação 2025-2026"
)
# Receptoras (aba ATUALIZAÇÃO SEMANAL pré-agregada) e Mapa de Vendas — por data no prefixo
RECEPTORAS_DIR = DRIVE_ROOT / "PLANTEL" / "Estação 2025-2026"
MAPA_VENDAS_DIR = DRIVE_ROOT / "VENDAS" / "MAPAS DE VENDAS" / "Estação 2025-2026"
# CONTROLE_DE_PLANTEL mensal (aba MOVIMENTAÇÕES datada) — mesma pasta das receptoras
CONTROLE_MENSAL_DIR = RECEPTORAS_DIR
# Animais para sair (vendidos/sociedade pendentes) — aba ANIMAIS VENDIDOS.
# O arquivo mudou de pasta (saiu de PLANILHAS PARA O EDUARDO, hoje mora em
# ATUALIZACAO SEMANAL), então procuramos nas duas e ficamos com o mais recente.
ANIMAIS_SAIR_DIRS = (
    DRIVE_ROOT / "ATUALIZACAO SEMANAL",
    DRIVE_ROOT / "PLANILHAS PARA O EDUARDO",
)

HIST_HEADCOUNT = BASE_DIR / "_cache" / "headcount_history.json"
HIST_SNAPSHOTS = BASE_DIR / "_cache" / "semanal_snapshots.json"

SAFRA_ATUAL = "2025/2026"
BASES_DIR = BASE_DIR / "bases"
JSON_OUT = BASES_DIR / "semanal_data.json"

# Alvos do docx 17-07-26 (semana 03/07-17/07) — só pra validação em tela
DOCX_1707 = {
    "acumulado_estacao": 61,
    "confirmados_semana": 2,
    "nascimentos": 1,
    "abortos_obitos": 1,
    "receptoras_total": 63,
    "receptoras_prenhas": 36,
    "receptoras_vazias": 27,
    "headcount_total": 206,
    "headcount_fpg": 104,
    "headcount_arr": 43,
    "headcount_cte": 1,
    "headcount_soc": 58,
    "saidas_semana": 8,
    "vendidos_pendentes": 2,
    "sociedade_pendentes": 2,
}


# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------
def _norm(v) -> str:
    return str(v).strip().upper() if v is not None else ""


def _s(v):
    if v is None:
        return None
    s = str(v).strip()
    return s or None


def _dt(v):
    if isinstance(v, datetime):
        return v.date()
    return None


def _to_num(v):
    try:
        return float(v) if v is not None else None
    except (TypeError, ValueError):
        return None


def _load(path: Path):
    cached = ensure_cache(path)
    return openpyxl.load_workbook(cached, data_only=True, read_only=True)


def _latest_by_mtime(folder: Path, pattern: str) -> Path:
    """Arquivo mais recente por MTIME (data de modificação real). Inclui as cópias
    'EDITAR ...' — o operador trabalha nelas (versão viva), então são as MAIS frescas.
    Os arquivos {YYMMDD} congelados são snapshots antigos."""
    cands = [f for f in folder.glob(pattern) if not f.name.startswith("~$")]
    if not cands:
        raise FileNotFoundError(f"Nenhum arquivo {pattern} em {folder}")
    return max(cands, key=lambda f: f.stat().st_mtime)


def _latest_estacao_master() -> Path:
    return _latest_by_mtime(ESTACAO_MASTER_DIR, "*ESTACAO DE MONTA.xlsx")


def _latest_by_yymmdd(folder: Path, pattern: str) -> Path:
    return _latest_by_mtime(folder, pattern)


# ------------------------------------------------------------------
# Estrutura do relatório
# ------------------------------------------------------------------
@dataclass
class Report:
    semana_inicio: str
    semana_fim: str
    fontes: dict = field(default_factory=dict)
    producao: dict = field(default_factory=dict)
    receptoras: dict = field(default_factory=dict)
    headcount: dict = field(default_factory=dict)
    terceiros: dict = field(default_factory=dict)
    saidas: dict = field(default_factory=dict)
    detalhe: dict = field(default_factory=dict)  # tabelas de detalhe p/ o dashboard
    eventos: dict = field(default_factory=dict)  # listas COMPLETAS datadas (filtro client-side)
    calendario: list = field(default_factory=list)  # semanas travadas p/ o seletor
    semana_atual: str = ""  # id (segunda) da semana selecionada por padrão
    snapshots: dict = field(default_factory=dict)  # snapshot por semana (histórico local)
    roster: list = field(default_factory=list)  # nomes do plantel (p/ diff saídas/entradas)
    receptoras_locais: dict = field(default_factory=dict)  # {animal: LOCAL} p/ diff de transferência
    populacao: list = field(default_factory=list)  # plantel + receptoras contadas (p/ diff saídas/entradas)
    saidas_planilha: dict | None = None  # aba SAIDAS-ENTRADAS, quando o haras preencher
    confirmed: list = field(default_factory=list)  # embriões confirmados (+/-=OK) p/ diff semanal
    docx_ref: dict = field(default_factory=dict)  # números dos relatórios oficiais (SÓ p/ validar)


# ------------------------------------------------------------------
# Seção 1 — PRODUÇÃO (master da estação de monta)
# ------------------------------------------------------------------
# Abas da planilha do grupo que compõem o acumulado.
#   (aba, fatia fixa ou None = deduzir do STATUS)
# As COLUNAS são resolvidas pelo CABEÇALHO (linha 3), nunca por índice fixo. Em
# 07/08/2026 a coluna STATUS foi apagada da aba de PAO GRANDE, tudo à direita
# andou uma casa e o índice de ESTAÇÃO passou a cair em COTA PG: nenhuma linha
# batia a safra, a fatia "pg" zerou e o acumulado caiu de 61 pra 27 sem erro
# nenhum. Cabeçalho por nome faz a planilha poder mexer coluna sem quebrar.
ABAS_ACUMULADO = (
    ("EMBRIÕES PAO GRANDE", "pg"),
    ("EMBRIOES SOCIOS - VENDIDOS", None),
)
HDR_ROW_ACUMULADO = 3  # linha do cabeçalho nas duas abas


def _sem_acento(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", str(s))
                   if unicodedata.category(c) != "Mn")


def _col_idx(hdr, *nomes) -> int:
    """Índice da coluna pelo nome do cabeçalho (sem acento, case-insensitive).
    Explode se não achar — coluna que sumiu tem de virar erro, não zero calado."""
    alvo = {_sem_acento(n).strip().upper() for n in nomes}
    for i, h in enumerate(hdr):
        if h is not None and _sem_acento(h).strip().upper() in alvo:
            return i
    achados = [str(h) for h in hdr if h is not None]
    raise KeyError(f"coluna {nomes[0]!r} não está no cabeçalho: {achados}")


def _latest_emb_matrizes() -> Path:
    cands = [f for f in EMB_MATRIZES_DIR.glob(EMB_MATRIZES_GLOB)
             if not f.name.startswith("~$")]
    if not cands:
        raise FileNotFoundError(f"Nenhuma planilha {EMB_MATRIZES_GLOB} em {EMB_MATRIZES_DIR}")
    return max(cands, key=lambda f: f.stat().st_mtime)


def _acumulado_grupo() -> dict:
    """Embriões vivos da safra na planilha 'EMBRIÕES E MATRIZES' — a que o haras
    manda no grupo e usa como fonte oficial do acumulado.

    Ela é um retrato do que está EM PÉ: a linha é apagada quando o embrião sai da
    conta, seja por parição (vira potro) ou por aborto. Por isso o acumulado da
    estação = linhas vivas + parições da safra. Aborto NÃO volta — é baixa
    definitiva, confirmado com o haras em 2026-07-31 (os 3 abortos da safra também
    sumiram da planilha e não entram no número oficial).

    A aba ICSI fica de fora: hoje só tem safra 2023/2024."""
    src = _latest_emb_matrizes()
    wb = _load(src)
    out = {"pg": 0, "socio": 0, "vendido": 0, "fonte": src.name}
    for aba, fatia_fixa in ABAS_ACUMULADO:
        if aba not in wb.sheetnames:
            print(f"  [acumulado] aba {aba!r} não existe em {src.name} — fatia zerada")
            continue
        iest = ist = None
        for i, r in enumerate(wb[aba].iter_rows(values_only=True), start=1):
            if i == HDR_ROW_ACUMULADO:
                iest = _col_idx(r, "ESTACAO")
                ist = None if fatia_fixa else _col_idx(r, "STATUS")
                continue
            if i <= HDR_ROW_ACUMULADO or r[1] is None or not str(r[1]).strip():
                continue
            if _s(r[iest]) != SAFRA_ATUAL:
                continue
            fatia = fatia_fixa or ("vendido" if _norm(r[ist]) == "VENDIDO" else "socio")
            out[fatia] += 1
    wb.close()
    out["total"] = out["pg"] + out["socio"] + out["vendido"]
    return out


def _mortes_do_plantel(ini: date, fim: date) -> list:
    """Óbitos registrados na aba 'CONFIRMAÇÕES, ABORTOS, MORTES' do controle mensal.

    Reunião 2026-07-31: a estação de monta só registra aborto/absorção, morte de
    receptora prenha e potro que nasce e morre em seguida. Óbito de animal já no
    plantel só aparece aqui. Colunas: B produto, C data, D observação (texto livre,
    por isso a classificação é por palavra-chave).
    """
    try:
        src = _latest_by_yymmdd(CONTROLE_MENSAL_DIR, "*CONTROLE_DE_PLANTEL_PAO_GRANDE_*.xlsx")
    except FileNotFoundError:
        return []
    wb = _load(src)
    aba = "CONFIRMAÇÕES, ABORTOS, MORTES"
    if aba not in wb.sheetnames:
        wb.close()
        return []
    achados = []
    for i, r in enumerate(wb[aba].iter_rows(values_only=True), start=1):
        if i < 3 or r[1] is None:
            continue
        d = _dt(r[2])
        if not d or not (ini <= d <= fim):
            continue
        obs = _norm(r[3])
        if "ABORT" in obs:            # aborto já vem da estação; aqui só morte
            continue
        if any(p in obs for p in ("MORREU", "MORTE", "OBITO", "MORTO")):
            achados.append({"animal": _s(r[1]), "data": d.isoformat(),
                            "ocorrencia": _s(r[3]), "origem": "plantel"})
    wb.close()
    return achados


def _acumulado_planejamento(wb) -> int:
    """Total da estação de monta = soma de 'TOTAL EMBRIÕES' (REAL, idx8) da aba
    PLANEJAMENTO, linhas de doadora (col0 numérica). NÃO é mais o acumulado
    publicado — é referência de conferência: a estação só registra o que passou
    pela FPG, então fica abaixo do oficial."""
    if "PLANEJAMENTO" not in wb.sheetnames:
        return 0
    ws = wb["PLANEJAMENTO"]
    total = 0.0
    for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if i < 4 or row[0] is None:
            continue
        if not str(row[0]).strip().isdigit():
            continue
        v = _to_num(row[8])
        if v:
            total += v
    return int(round(total))


def build_producao(rep: Report, ini: date, fim: date):
    master = _latest_estacao_master()
    rep.fontes["estacao_master"] = master.name
    wb = _load(master)
    ws = wb["ESTAÇÃO"]
    embrioes = []
    for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if i < 3 or r[0] is None:
            continue
        if _s(r[35]) != SAFRA_ATUAL:
            continue
        ia = _dt(r[7])
        # confirmação oficial = coluna +/- (idx17) == 'OK' (validado: os embriões
        # confirmados no docx têm +/-='OK', 60D às vezes nem marcado). Conta = PLANEJAMENTO.
        confirmado = _norm(r[17]) == "OK"
        data_conf = (ia + timedelta(days=60)) if ia else None
        cotas = _to_num(r[5])
        # split PG / sócio / vendido (defs do comitê: total produzido aberto nessas 3 fatias)
        comprador = _s(r[34])
        if cotas is not None and cotas == 0:
            fatia = "vendido"
        elif (cotas is not None and 0 < cotas < 1) or _s(r[6]):
            fatia = "socio"
        else:
            fatia = "pg"
        embrioes.append({
            "doadora": _s(r[2]), "garanhao": _s(r[3]), "local": _s(r[4]),
            "cotas_pg": cotas, "socio": _s(r[6]), "fatia": fatia,
            "data_ia": ia.isoformat() if ia else None,
            "receptora": _s(r[11]),
            "confirmado": confirmado,
            "data_confirmacao": data_conf.isoformat() if data_conf else None,
            "sexo_potro": _s(r[24]), "nome_potro": _s(r[25]),
            "data_paricao": _dt(r[23]).isoformat() if _dt(r[23]) else None,
            "data_aborto": _dt(r[21]).isoformat() if _dt(r[21]) else None,
            "data_obito": _dt(r[28]).isoformat() if _dt(r[28]) else None,
            "status": _s(r[32]), "categoria": _s(r[33]), "comprador": comprador,
        })

    def _in_week(iso):
        return bool(iso and ini <= date.fromisoformat(iso) <= fim)

    def _in_month(iso):
        if not iso:
            return False
        d = date.fromisoformat(iso)
        return d.year == fim.year and d.month == fim.month

    def _split(items):
        return {f: sum(1 for e in items if e["fatia"] == f) for f in ("pg", "socio", "vendido")}

    # ACUMULADO NA ESTAÇÃO (2026-07-31, reunião com o haras):
    #   linhas vivas da planilha "EMBRIÕES E MATRIZES" (PG + sócios) + parições da safra.
    # A planilha do grupo é o retrato do que está em pé — some a linha quando pare ou
    # aborta —, então as parições voltam pra conta e os abortos não. A estação de monta
    # sozinha não serve: só enxerga o que passou pela FPG (56 vs 61 hoje). Fica como
    # conferência em `acumulado_estacao_monta`.
    grupo = _acumulado_grupo()
    rep.fontes["embrioes_matrizes"] = grupo["fonte"]
    acumulado_planejamento = _acumulado_planejamento(wb)

    confirmados = [e for e in embrioes if e["confirmado"]]
    for e in confirmados:
        e["key"] = f"{e['doadora']}|{e['garanhao']}|{e['receptora']}|{e['data_ia']}"
    rep.confirmed = confirmados
    na_semana = []          # confirmados_semana e acumulado_mes: preenchidos por _compute_confirmados_diff
    nascimentos = [e for e in embrioes if _in_week(e["data_paricao"])]
    # aborto = embrião confirmado que não nasceu (data aborto na semana);
    # óbito = nasceu e morreu (data óbito na semana). (absorção = perda pré-60d, não confirmada)
    abortos = [e for e in embrioes if e["confirmado"] and _in_week(e["data_aborto"])]
    obitos = [e for e in embrioes if _in_week(e["data_obito"])]
    # óbito de animal já no plantel não passa pela estação — vem do controle mensal
    mortes_plantel = _mortes_do_plantel(ini, fim)
    if mortes_plantel:
        print(f"  [produção] +{len(mortes_plantel)} óbito(s) da aba "
              f"'CONFIRMAÇÕES, ABORTOS, MORTES' do controle mensal")

    # parições da safra inteira (não só da semana) — voltam pro acumulado
    paridos_safra = [e for e in embrioes if e["data_paricao"]]
    acumulado = grupo["total"] + len(paridos_safra)
    split = {"pg": grupo["pg"], "socio": grupo["socio"], "vendido": grupo["vendido"]}
    for e in paridos_safra:
        split[e["fatia"]] = split.get(e["fatia"], 0) + 1
    if acumulado != acumulado_planejamento:
        print(f"  [produção] acumulado {acumulado} "
              f"(planilha do grupo {grupo['total']} + {len(paridos_safra)} parições) "
              f"vs {acumulado_planejamento} na estação de monta")

    rep.producao = {
        "acumulado_estacao": acumulado,
        "acumulado_estacao_split": split,
        "acumulado_estacao_monta": acumulado_planejamento,   # conferência
        "acumulado_grupo_vivos": grupo["total"],
        "acumulado_paricoes_safra": len(paridos_safra),
        "acumulado_estacao_split_confirmados": _split(confirmados),  # split antigo (estação)
        "confirmados_semana": None,   # _compute_confirmados_diff
        "acumulado_mes": None,        # _compute_confirmados_diff
        "nascimentos": len(nascimentos),
        "abortos_obitos": len(abortos) + len(obitos) + len(mortes_plantel),
    }
    def _produto(e):   # nome do animal nascido (ou descrição sexo — doadora × garanhão)
        base = f"{e.get('doadora') or ''} × {e.get('garanhao') or ''}".strip(" ×")
        sx = {"M": "Macho", "F": "Fêmea"}.get((e.get("sexo_potro") or "").upper(), e.get("sexo_potro") or "")
        return e.get("nome_potro") or (f"{sx} — {base}" if base else sx) or "--"
    rep.detalhe["confirmados_semana"] = na_semana
    rep.detalhe["nascimentos_semana"] = [dict(e, produto=_produto(e)) for e in nascimentos]
    rep.detalhe["abortos_obitos_semana"] = abortos + obitos + mortes_plantel
    rep.detalhe["embrioes_confirmados"] = confirmados
    # listas COMPLETAS datadas (o dashboard filtra por semana no cliente)
    rep.eventos["confirmados"] = [dict(e, data=e["data_confirmacao"]) for e in confirmados if e["data_confirmacao"]]
    rep.eventos["nascimentos"] = [dict(e, data=e["data_paricao"]) for e in embrioes if e["data_paricao"]]
    rep.eventos["abortos_obitos"] = [
        dict(e, data=(e["data_aborto"] or e["data_obito"]))
        for e in embrioes if (e["confirmado"] and e["data_aborto"]) or e["data_obito"]
    ]
    wb.close()


# ------------------------------------------------------------------
# Seção 2 — RECEPTORAS (censo + candidatos de filtro)
# ------------------------------------------------------------------
RECEPTORAS_LOCAIS_ATIVOS = ("PAO GRANDE", "ARRENDAMENTO CESAR FURTADO")


def build_receptoras(rep: Report):
    """Rebanho ATIVO = aba ANIMAIS do PLANTEL ARRENDAMENTOS E RECEPTORAS, filtrando
    LOCAL em (PAO GRANDE, ARRENDAMENTO CESAR FURTADO) e STATUS prenha/vazia.
    NÃO usa mais a aba 'ATUALIZAÇÃO SEMANAL' (aba a ser aposentada). Validado vs
    docx 24/07: prenhas 34 / vazias 28 / total 62 — bate EXATO.
    Colunas ANIMAIS (linha 3 header, dados r4+): 1 ANIMAL, 2 STATUS, 3 LOCAL."""
    src = _latest_by_yymmdd(RECEPTORAS_DIR, "*PLANTEL ARRENDAMENTOS E RECEPTORAS.xlsx")
    rep.fontes["receptoras"] = src.name
    wb = _load(src)
    ws = wb["ANIMAIS"]
    pren = vaz = 0
    for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if i < 4 or r[1] is None:
            continue
        if _norm(r[3]) not in RECEPTORAS_LOCAIS_ATIVOS:
            continue
        st = _norm(r[2])
        if st.startswith("PRENHA"):
            pren += 1
        elif st.startswith("VAZIA"):
            vaz += 1
    wb.close()
    doadoras_plantel = _count_doadoras()             # CATEGORIA='DOADORA' no plantel
    doadoras_fpg = _count_doadoras("FAZENDA PAO GRANDE")   # só referência
    doadoras = DOADORAS_INDICE or doadoras_plantel
    rep.receptoras = {
        "total": pren + vaz,
        "prenhas": pren,
        "vazias": vaz,
        "doadoras": doadoras,
        "doadoras_fonte": "fixo" if DOADORAS_INDICE else "plantel",
        "doadoras_plantel": doadoras_plantel,
        "doadoras_plantel_fpg": doadoras_fpg,
        "indice_eficiencia": round(vaz / doadoras, 1) if doadoras else None,
    }
    if DOADORAS_INDICE and DOADORAS_INDICE != doadoras_plantel:
        print(f"  [receptoras] índice usa {DOADORAS_INDICE} doadoras (fixo); "
              f"no plantel há {doadoras_fpg} na FPG e {doadoras_plantel} no total")


# Denominador do índice de eficiência.
# Era FIXO em 10 pela reunião de 2026-07-31, porque nenhuma leitura do plantel dava
# esse número. Em 07/08/2026 o relatório oficial trocou o divisor: 29 vazias com
# índice 2,4 = 29/12, e 12 é exatamente CATEGORIA='DOADORA' no plantel (nas 4 semanas
# anteriores o divisor implícito no docx era 10). Ou seja, o haras passou a usar o
# contado — então paramos de fixar e contamos.
# None = usar as doadoras contadas no plantel. Só voltar a pôr número aqui se o
# haras decidir travar o divisor de novo.
DOADORAS_INDICE = None


def _count_doadoras(local: str | None = None) -> int:
    """Doadoras no plantel = CATEGORIA 'DOADORA' na aba PLANTEL do CONTROLE PLANTEL
    semanal. Com `local`, conta só as daquele LOCAL. Referência/conferência —
    o índice usa DOADORAS_INDICE."""
    wb = _load(CONTROLE_PLANTEL_SEMANAL)
    ws = wb["PLANTEL"]
    n = 0
    for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if i < 2 or r[0] is None:
            continue
        if _norm(r[2]) != "DOADORA":
            continue
        if local is not None and _norm(r[4]) != local:
            continue
        n += 1
    wb.close()
    return n


# ------------------------------------------------------------------
# Seção 3 — HEADCOUNT (calculado do roster, espelhando a aba CONTAGEM)
# ------------------------------------------------------------------
# A CONTAGEM é COUNTIF puro sobre o LOCAL do roster:
#     C3 =COUNTIF(PLANTEL!E:E;"FAZENDA PAO GRANDE")      D3 = 25   (digitado)
#     C4 =COUNTIF(PLANTEL!E:E;"ARRENDAMENTO CESAR FURTADO")  D4 = 37   (digitado)
#     C5 =COUNTIF(PLANTEL!E:E;"OUTROS")                  D5 = 0     (rótulo "CTE")
#     C6 =COUNTIF(PLANTEL!E:E;"SOCIO")                   D6 = 0
#     E7 =SUM(E3:E6)
# Reproduzimos a MESMA regra (conta linha por LOCAL, sem filtrar STATUS — por isso
# os vendidos pendentes de saída entram) e o MESMO conjunto de buckets. Só muda de
# onde vem o número das receptoras: contado da fonte de receptoras (a mesma da
# seção 2) em vez de digitado na mão. Confere — PAO GRANDE 25 e ARRENDAMENTO CESAR
# FURTADO 37, idênticos ao que estava fixo em D3/D4.
#
# MATO GROSSO fica FORA do headcount por decisão de negócio (confirmado em
# 2026-07-31): os animais de lá não entram na contagem do plantel, e é por isso
# que a CONTAGEM nunca teve linha para eles.
#
# LOCAL vazio não entra (o COUNTIF também ignora), o que descarta de graça a
# pseudo-linha "RECEPTORAS 67" que mora no meio do roster.

# LOCAL no roster -> chave no relatório (rótulo que a CONTAGEM usa)
HEADCOUNT_BUCKETS = {
    "FAZENDA PAO GRANDE": ("FAZENDA", "fazenda_pg"),
    "ARRENDAMENTO CESAR FURTADO": ("ARRENDAMENTO", "arrendamento"),
    "OUTROS": ("CTE", "cte"),
    "SOCIO": ("SOCIO", "socio"),
}
# LOCAL que existe no roster mas não conta no headcount.
HEADCOUNT_LOCAIS_FORA = ("MATO GROSSO",)
# LOCAL da fonte de receptoras -> LOCAL do roster
RECEPTORAS_PARA_BUCKET = {
    "PAO GRANDE": "FAZENDA PAO GRANDE",
    "ARRENDAMENTO CESAR FURTADO": "ARRENDAMENTO CESAR FURTADO",
}


def _slug_local(local: str) -> str:
    return local.lower().replace(" ", "_")


def _contagem_declarada() -> dict:
    """Valores da aba CONTAGEM — só para conferência, não é fonte."""
    wb = _load(CONTROLE_PLANTEL_SEMANAL)
    ws = wb["CONTAGEM"]
    out = {}
    for r in ws.iter_rows(values_only=True):
        label = _norm(r[1])
        if label in ("FAZENDA", "ARRENDAMENTO", "CTE", "SOCIO", "TOTAL GERAL"):
            out[label] = {"animais": r[2], "receptoras": r[3], "total": r[4]}
    wb.close()
    return out


def _receptoras_por_local() -> dict:
    """{LOCAL do roster: nº de receptoras}. Mesma regra da seção 2 (prenha/vazia)."""
    src = _latest_by_yymmdd(RECEPTORAS_DIR, "*PLANTEL ARRENDAMENTOS E RECEPTORAS.xlsx")
    wb = _load(src)
    ws = wb["ANIMAIS"]
    out = {}
    for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if i < 4 or r[1] is None:
            continue
        loc = _norm(r[3])
        if loc not in RECEPTORAS_LOCAIS_ATIVOS:
            continue
        st = _norm(r[2])
        if st.startswith("PRENHA") or st.startswith("VAZIA"):
            out[RECEPTORAS_PARA_BUCKET[loc]] = out.get(RECEPTORAS_PARA_BUCKET[loc], 0) + 1
    wb.close()
    return out


def build_headcount(rep: Report):
    # 1) animais: COUNTIF por LOCAL no roster
    wb = _load(CONTROLE_PLANTEL_SEMANAL)
    ws = wb["PLANTEL"]
    animais: dict[str, int] = {}
    for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if i < 2 or r[0] is None or not str(r[0]).strip():
            continue
        local = _norm(r[4])
        if not local:
            continue
        animais[local] = animais.get(local, 0) + 1
    wb.close()

    # 2) receptoras: contadas da fonte de receptoras
    receptoras = _receptoras_por_local()

    # 3) monta os buckets. LOCAL fora da contagem é ignorado; LOCAL desconhecido
    #    também não entra (a CONTAGEM não o teria), mas vira aviso — assim um
    #    local novo no roster aparece pra alguém decidir, em vez de sumir calado.
    detalhe, chaves, fora = {}, {}, {}
    for local in sorted(set(animais) | set(receptoras)):
        a = animais.get(local, 0)
        rc = receptoras.get(local, 0)
        if local in HEADCOUNT_LOCAIS_FORA or local not in HEADCOUNT_BUCKETS:
            fora[local] = a + rc
            continue
        rotulo, chave = HEADCOUNT_BUCKETS[local]
        detalhe[rotulo] = {"animais": a, "receptoras": rc, "total": a + rc}
        chaves[chave] = a + rc

    total = sum(v["total"] for v in detalhe.values())
    detalhe["TOTAL GERAL"] = {
        "animais": sum(v["animais"] for k, v in detalhe.items() if k != "TOTAL GERAL"),
        "receptoras": sum(v["receptoras"] for k, v in detalhe.items() if k != "TOTAL GERAL"),
        "total": total,
    }

    rep.headcount = {"total": total, **chaves, "detalhe": detalhe, "fora_da_contagem": fora}

    # 4) conferência contra a aba manual — divergência vira aviso, não erro
    decl = _contagem_declarada()
    dif = []
    for rotulo, v in detalhe.items():
        d = decl.get(rotulo)
        if d is None:
            dif.append(f"{rotulo}: {v['total']} calculado, sem linha na CONTAGEM")
        elif d.get("total") != v["total"]:
            dif.append(f"{rotulo}: {v['total']} calculado vs {d.get('total')} na CONTAGEM")
    declarado_total = (decl.get("TOTAL GERAL") or {}).get("total")
    if declarado_total is not None and declarado_total != total:
        dif.append(f"TOTAL GERAL: {total} calculado vs {declarado_total} na CONTAGEM")
    rep.headcount["conferencia_contagem"] = dif
    if dif:
        print("  [headcount] divergência vs aba CONTAGEM (usando o calculado):")
        for d in dif:
            print(f"    - {d}")
    desconhecidos = [l for l in fora if l not in HEADCOUNT_LOCAIS_FORA]
    if desconhecidos:
        print("  [headcount] LOCAL novo no roster, FORA da contagem — conferir:")
        for l in desconhecidos:
            print(f"    - {l}: {fora[l]}")


# ------------------------------------------------------------------
# Seção 5 — SAÍDAS / ENTRADAS / TRANSFERÊNCIAS  (CONTROLE_DE_PLANTEL aba MOVIMENTAÇÕES)
# ------------------------------------------------------------------
def _categorize_mov(obs: str):
    if "MUDOU O LOCAL PARA FAZENDA PAO GRANDE" in obs or "MUDOU O LOCAL PARA ARRENDAMENTO" in obs:
        return "TRANSFERENCIA"
    if "SAIU DO HARAS" in obs:
        return "SAIDA"
    if "CHEGOU NO HARAS" in obs:
        return "ENTRADA"
    return None


# Aba SAIDAS-ENTRADAS do controle mensal (reunião 2026-07-31): fonte oficial de
# saída/entrada. O haras começou a preencher em 12/08/2026.
# Colunas: B animal, C local saída, D local entrada, E data, F classificação.
# Δ combinado: entrada = nascimento + compra; saída = venda + morte.
#
# A classificação vem digitada com o sentido colado no motivo — a primeira linha real
# foi 'SAIDA-SOCIO'. Só as palavras de motivo abaixo não bastavam: 'SAIDA-SOCIO' não
# casava com nada, caía no `else None` e era descartada em silêncio. Como a aba já
# tinha linha, ela vencia como fonte oficial e a saída da semana virava 0 — foi assim
# que LINDEZA DA PAO GRANDE (12/08/2026) sumiu do fechamento de 14/08.
CLASSIF_ENTRADA = ("NASCIMENTO", "COMPRA")
CLASSIF_SAIDA = ("VENDA", "MORTE", "SOCIO")
# Saída da FAZENDA que continua no PLANTEL: a aba CONTAGEM conta o sócio, então o
# animal que vai pro sócio é saída na seção 5 e NÃO mexe no Δ do headcount. É o que o
# relatório oficial faz em 14/08/2026: 'Saídas na semana: 01' com 'Δ +00 / -00'.
CLASSIF_FORA_DO_DELTA = ("SOCIO",)


def _classificar_se(classif: str):
    """(sentido, afeta_headcount) da classificação da aba SAIDAS-ENTRADAS.
    (None, None) = vocabulário desconhecido; quem chama tem de avisar, não engolir."""
    def _sai():
        return "SAIDA", not any(c in classif for c in CLASSIF_FORA_DO_DELTA)
    if classif.startswith("ENTRADA"):
        return "ENTRADA", True
    if classif.startswith("SAIDA"):
        return _sai()
    if any(c in classif for c in CLASSIF_ENTRADA):
        return "ENTRADA", True
    if any(c in classif for c in CLASSIF_SAIDA):
        return _sai()
    return None, None


def _saidas_entradas_planilha(wb, ini: date, fim: date):
    """Lê a aba SAIDAS-ENTRADAS. Devolve None se ela não existir ou estiver vazia,
    pra que o cálculo caia na diferença de roster (comportamento atual)."""
    if "SAIDAS-ENTRADAS" not in wb.sheetnames:
        return None
    evs = {"SAIDA": [], "ENTRADA": []}
    achou = False
    desconhecidas = []
    for i, r in enumerate(wb["SAIDAS-ENTRADAS"].iter_rows(values_only=True), start=1):
        if i < 3 or r[1] is None or not str(r[1]).strip():
            continue
        achou = True
        d = _dt(r[4])
        if not d or not (ini <= d <= fim):
            continue
        classif = _norm(r[5])
        alvo, afeta = _classificar_se(classif)
        if alvo is None:
            # classificação nova na planilha: avisar, nunca virar zero calado
            desconhecidas.append(f"linha {i} ({_s(r[1])}): {_s(r[5])!r}")
            continue
        evs[alvo].append({"animal": _s(r[1]), "data": d.isoformat(),
                          "classificacao": _s(r[5]), "afeta_headcount": afeta,
                          "local_saida": _s(r[2]), "local_entrada": _s(r[3])})
    if desconhecidas:
        print(f"  [SAIDAS-ENTRADAS] classificação não reconhecida em "
              f"{len(desconhecidas)} linha(s) da janela — NÃO entraram na conta: "
              + "; ".join(desconhecidas))
    return evs if achou else None


# TRANSFERÊNCIA INTERNA = animal que trocou de LOCAL entre os dois locais próprios
# (FPG <-> arrendamento) na aba ANIMAIS do PLANTEL ARRENDAMENTOS E RECEPTORAS.
# A aba MOVIMENTAÇÕES do controle mensal NÃO serve: a última transferência lançada
# lá é de setembro/2025. Em 07/08/2026 o relatório oficial disse 14 transferências e
# o diff de LOCAL dá exatamente 14 (5 arrendamento->FPG, 9 FPG->arrendamento).
# Igual a saídas/entradas, é diff: precisa do mapa da semana anterior. Sem ele
# (primeira captura), fica em branco — nunca zero.
def _receptoras_arquivos() -> list:
    """Arquivos de receptoras, do mais recente pro mais antigo (por mtime)."""
    cands = [f for f in RECEPTORAS_DIR.glob("*PLANTEL ARRENDAMENTOS E RECEPTORAS.xlsx")
             if not f.name.startswith("~$")]
    return sorted(cands, key=lambda f: f.stat().st_mtime, reverse=True)


def _receptoras_info(src: Path | None = None) -> dict:
    """{ANIMAL: {local, status, embriao, obs}} da aba ANIMAIS — TODAS as linhas,
    inclusive fora dos nossos locais, pra saber pra onde o animal foi."""
    if src is None:
        src = _latest_by_yymmdd(RECEPTORAS_DIR, "*PLANTEL ARRENDAMENTOS E RECEPTORAS.xlsx")
    wb = _load(src)
    ws = wb["ANIMAIS"]
    out = {}
    for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if i < 4 or r[1] is None:
            continue
        out[_norm(r[1])] = {"local": _norm(r[3]), "status": _s(r[2]),
                            "embriao": _s(r[4]), "obs": _s(r[5])}
    wb.close()
    return out


def _receptoras_locais(src: Path | None = None) -> dict:
    """{ANIMAL: LOCAL} da aba ANIMAIS, só quem está num local nosso."""
    return {k: v["local"] for k, v in _receptoras_info(src).items()
            if v["local"] in RECEPTORAS_LOCAIS_ATIVOS}


def _transferencias_internas(rep: Report) -> list | None:
    """Diff do LOCAL vs o mapa da semana anterior. None = sem base de comparação."""
    hist = _load_hist()
    prev = None
    for wid in sorted(hist):
        if wid < rep.semana_atual and hist[wid].get("receptoras_locais"):
            prev = hist[wid]["receptoras_locais"]
    cur = rep.receptoras_locais or {}
    if not cur:
        return None
    if not prev:
        # Bootstrap: nenhuma semana anterior guardou o mapa (o campo é novo). Cai no
        # arquivo de receptoras anterior, que é o retrato de onde os animais estavam.
        anteriores = _receptoras_arquivos()[1:]
        if not anteriores:
            print("  [transferências] sem mapa de LOCAL da semana anterior e sem "
                  "arquivo anterior de receptoras — fica EM BRANCO, não zero")
            return None
        prev = _receptoras_locais(anteriores[0])
        print(f"  [transferências] primeira semana com mapa de LOCAL: comparando com "
              f"{anteriores[0].name} (da próxima em diante, compara com o snapshot)")
    return [{"animal": k, "local_saida": prev[k], "local_entrada": cur[k]}
            for k in cur if k in prev and prev[k] != cur[k]]


def build_movimentacao(rep: Report, ini: date, fim: date):
    src = _latest_by_yymmdd(CONTROLE_MENSAL_DIR, "*CONTROLE_DE_PLANTEL_PAO_GRANDE_*.xlsx")
    rep.fontes["controle_plantel_mensal"] = src.name
    wb = _load(src)
    rep.saidas_planilha = _saidas_entradas_planilha(wb, ini, fim)
    ws = wb["MOVIMENTAÇÕES"]
    evs = {"SAIDA": [], "ENTRADA": [], "TRANSFERENCIA": []}
    ultimo = None  # último lançamento DATADO da aba, de qualquer tipo
    for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if i < 3 or r[3] is None:
            continue
        d = _dt(r[3])
        if not d:
            continue
        if ultimo is None or d > ultimo:
            ultimo = d
        tipo = _categorize_mov(str(r[4] or "").upper())
        if tipo is None:
            continue
        evs[tipo].append({"animal": _s(r[2]), "data": d.isoformat(), "ocorrencia": _s(r[4])})
    inw = lambda x: ini <= date.fromisoformat(x["data"]) <= fim
    # A aba só mede a semana se ela chegou na semana. Parada antes da janela, contar
    # zero afirmaria "não houve movimentação" quando o que existe é falta de
    # lançamento — foi o que fez transferências virar 0 com 14 no relatório oficial.
    rep.saidas = {
        "saidas_semana": sum(1 for x in evs["SAIDA"] if inw(x)),
        "entradas_semana": sum(1 for x in evs["ENTRADA"] if inw(x)),
    }
    # transferências não vêm daqui (ver comentário acima de _receptoras_locais)
    rep.receptoras_locais = _receptoras_locais()
    transf = _transferencias_internas(rep)
    rep.saidas["transferencias_semana"] = len(transf) if transf is not None else None
    rep.detalhe["transferencias_internas"] = transf
    rep.saidas["movimentacao_ultimo_lancamento"] = ultimo.isoformat() if ultimo else None
    if ultimo is None or ultimo < ini:
        rep.saidas["movimentacao_defasada"] = True
        print(f"  [movimentação] {src.name}: último lançamento datado é "
              f"{ultimo.strftime('%d/%m/%Y') if ultimo else 'nenhum'}, antes da janela "
              f"({ini.strftime('%d/%m')}) — a aba MOVIMENTAÇÕES não mede esta semana")
    # listas COMPLETAS datadas p/ filtro client-side
    rep.eventos["saidas"] = evs["SAIDA"]
    rep.eventos["entradas"] = evs["ENTRADA"]
    rep.eventos["transferencias"] = evs["TRANSFERENCIA"]
    wb.close()


# ------------------------------------------------------------------
# Seção 4b — PENDENTES DE SAÍDA / TERCEIROS  (CONTROLE PLANTEL aba PLANTEL)
# ------------------------------------------------------------------
# Reunião com o haras (2026-07-31): vendidos pendentes e doadoras de terceiros
# passam a sair do roster, pela coluna STATUS PLANTEL. Os valores abaixo ainda NÃO
# existem na planilha — o haras vai preencher. Enquanto não houver nenhuma linha
# marcada, cada indicador cai na fonte antiga (vendidos) ou fica em branco
# (doadoras de terceiros), e o run avisa qual fonte usou.
STATUS_VENDIDO_PENDENTE = "VENDIDO PENDENTE"
STATUS_TERCEIRO = "TERCEIRO"


def _plantel_por_status() -> dict:
    """Lê o roster uma vez e devolve o que depende de STATUS PLANTEL / CATEGORIA.
    Colunas: A nome, C categoria, D status plantel, E local, F status."""
    wb = _load(CONTROLE_PLANTEL_SEMANAL)
    ws = wb["PLANTEL"]
    roster, vendidos_pend, doadoras_terc = [], [], []
    for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if i < 2 or r[0] is None:
            continue
        nome = _s(r[0])
        if not nome:
            continue
        roster.append(nome)
        status_plantel = _norm(r[3])
        categoria, local = _norm(r[2]), _norm(r[4])
        if STATUS_VENDIDO_PENDENTE in status_plantel:
            vendidos_pend.append({"nome": nome, "local": _s(r[4]), "cota": None,
                                  "comprador": None, "tipo": "VENDA", "obs": _s(r[5]),
                                  "reposicao": False})
        if (categoria == "DOADORA" and STATUS_TERCEIRO in status_plantel
                and local == "FAZENDA PAO GRANDE"):
            doadoras_terc.append({"nome": nome, "local": _s(r[4]), "categoria": categoria})
    wb.close()
    return {"roster": sorted(set(roster)),
            "vendidos_pendentes": vendidos_pend,
            "doadoras_terceiros": doadoras_terc}


# Embrião comercial pendente de saída: aba ENTREGAR do "EMBRIOES A ENTREGAR - A
# RECEBER". 'Status embrião' PRONTO-* = feito e ainda não entregue (os outros estados
# — A FAZER, EM ANDAMENTO, ENTREGUE, NASCIDO, CANCELADO, REPOSIÇÃO — não são pendência
# de saída). 'Cota PG' < 1 = sociedade; = 1 = venda 100%.
# É essa a fonte dos "5 embriões" que o relatório de 07/08/2026 somou no card de
# sociedade; a aba EMBRIOES VENDIDOS do "Animais para sair" está vazia e não é usada.
EMB_PENDENTE_PREFIXO = "PRONTO"


def _embrioes_pendentes() -> list:
    """Embriões prontos e não entregues, com tipo SOCIEDADE (cota parcial) ou VENDA."""
    wb = _load(EMB_COMERCIAIS)
    ws = wb["ENTREGAR"]
    out, cols = [], None
    for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if i == 3:
            cols = {n: _col_idx(r, n) for n in
                    ("ID Embrião", "Doadora", "Garanhão", "Comprador", "Cota PG",
                     "Status embrião", "Observação")}
            continue
        if cols is None or r[cols["ID Embrião"]] is None:
            continue
        if not _norm(r[cols["Status embrião"]]).startswith(EMB_PENDENTE_PREFIXO):
            continue
        cota = r[cols["Cota PG"]]
        try:
            parcial = cota is not None and float(cota) < 1
        except (TypeError, ValueError):
            parcial = False
        out.append({
            "nome": f'{_s(r[cols["Doadora"]])} x {_s(r[cols["Garanhão"]])}',
            "id": _s(r[cols["ID Embrião"]]), "local": None, "cota": cota,
            "comprador": _s(r[cols["Comprador"]]),
            "tipo": "SOCIEDADE" if parcial else "VENDA",
            "obs": _s(r[cols["Status embrião"]]), "reposicao": False,
            "especie": "EMBRIAO",
        })
    wb.close()
    return out


def build_pendentes(rep: Report):
    plantel = _plantel_por_status()
    rep.roster = plantel["roster"]

    # VENDIDOS / SOCIEDADE pendentes = aba ANIMAIS VENDIDOS do "Animais para sair"
    # (validado vs docx 17/07: VENDA≠REPOSIÇÃO=2 vendidos; SOCIEDADE=2). col5=tipo, col6=obs.
    pend, pend_emb = [], []
    try:
        src = _latest_animais_sair()
        wb = _load(src)
        rep.fontes["animais_para_sair"] = src.name
        ws = wb["ANIMAIS VENDIDOS"]
        for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
            if i < 4 or r[1] is None:
                continue
            tipo = _norm(r[5])
            obs = _norm(r[6])
            if tipo not in ("VENDA", "SOCIEDADE"):
                continue
            pend.append({"nome": _s(r[1]), "local": _s(r[2]), "cota": r[3],
                         "comprador": _s(r[4]), "tipo": tipo, "obs": _s(r[6]),
                         "reposicao": obs == "REPOSICAO"})
        wb.close()
    except FileNotFoundError as exc:
        # NÃO engolir: sem esse arquivo, vendidos/sociedade pendentes e a lista da
        # seção 5 viram zero — e zero aqui é indistinguível de "não tem nenhum".
        print(f"  [pendentes] FONTE AUSENTE: {exc} -> vendidos/sociedade pendentes "
              f"e pendentes de saída ficam ZERADOS nesta semana")
    pend_emb = _embrioes_pendentes()

    # VENDIDOS PENDENTES: fonte nova é o roster (STATUS PLANTEL). Enquanto o haras
    # não marcar ninguém lá, segue valendo o "Animais para sair".
    vendidos_plantel = plantel["vendidos_pendentes"]
    if vendidos_plantel:
        vendidos = vendidos_plantel
        fonte_vendidos = "plantel"
    else:
        vendidos = [p for p in pend if p["tipo"] == "VENDA" and not p["reposicao"]]
        fonte_vendidos = "animais_para_sair"
        print(f"  [terceiros] nenhum '{STATUS_VENDIDO_PENDENTE}' no STATUS PLANTEL; "
              f"vendidos pendentes ainda saindo do Animais para sair")

    # SOCIEDADE pendente = animais + embriões (regra do relatório desde 07/08/2026)
    soc_animais = [p for p in pend if p["tipo"] == "SOCIEDADE"]
    soc_embrioes = [p for p in pend_emb if p["tipo"] == "SOCIEDADE"]
    sociedade = soc_animais + soc_embrioes
    rep.fontes["embrioes_pendentes"] = EMB_COMERCIAIS.name

    # DOADORAS DE TERCEIROS: roster com CATEGORIA=DOADORA, STATUS PLANTEL de terceiro
    # e LOCAL=FAZENDA PAO GRANDE. Sem marcação na planilha, fica em branco (não zero:
    # zero afirmaria que não há nenhuma, e o que temos é ausência de informação).
    doadoras_terc = plantel["doadoras_terceiros"]
    if not doadoras_terc:
        print(f"  [terceiros] nenhuma doadora com STATUS PLANTEL de {STATUS_TERCEIRO} "
              f"na FPG; card fica em branco")

    rep.terceiros.update({
        "vendidos_pendentes": len(vendidos),
        "vendidos_pendentes_fonte": fonte_vendidos,
        "sociedade_pendentes": len(sociedade),
        "sociedade_pendentes_animais": len(soc_animais),
        "sociedade_pendentes_embrioes": len(soc_embrioes),
        "terceiros_propriedade": len(vendidos),   # terceiros na propriedade = vendidos (docx)
        "doadoras_terceiros": len(doadoras_terc) if doadoras_terc else None,
        "outros_terceiros": None,
    })
    rep.detalhe["doadoras_terceiros"] = doadoras_terc
    rep.detalhe["terceiros_vendidos"] = vendidos          # só os 2 vendidos (bate KPI seção 4)
    rep.detalhe["terceiros_sociedade"] = sociedade        # sociedade pendente de saída, listada igual
    rep.detalhe["pendentes_saida"] = pend + pend_emb       # lista completa (seção 5)


def _latest_animais_sair() -> Path:
    """Arquivo 'Animais para sair*.xlsx' mais recente (por mtime — nome tem o ano).
    Procura em todas as pastas de ANIMAIS_SAIR_DIRS: o arquivo já andou de pasta."""
    cands = [f for d in ANIMAIS_SAIR_DIRS for f in d.glob("Animais para sair*.xlsx")
             if not f.name.startswith("~$")]
    if not cands:
        raise FileNotFoundError(
            "Nenhum 'Animais para sair*.xlsx' em: "
            + " | ".join(str(d) for d in ANIMAIS_SAIR_DIRS))
    return max(cands, key=lambda f: f.stat().st_mtime)


# ------------------------------------------------------------------
# Δ headcount vs run anterior (histórico local leve, automático)
# ------------------------------------------------------------------
def build_headcount_delta(rep: Report, fim: date):
    HIST_HEADCOUNT.parent.mkdir(parents=True, exist_ok=True)
    hist = {}
    if HIST_HEADCOUNT.exists():
        try:
            hist = json.loads(HIST_HEADCOUNT.read_text(encoding="utf-8"))
        except Exception:
            hist = {}
    total = rep.headcount.get("total")
    # run anterior = maior data < fim
    prev = None
    for k in sorted(hist):
        if k < fim.isoformat():
            prev = hist[k]
    if prev is not None and total is not None:
        rep.headcount["delta"] = total - prev.get("total", total)
    else:
        rep.headcount["delta"] = None
    atual = {"total": total, "fpg": rep.headcount.get("fazenda_pg"),
             "arr": rep.headcount.get("arrendamento"),
             "cte": rep.headcount.get("cte"), "soc": rep.headcount.get("socio")}
    # CONTAGEM idêntica à da semana passada, local por local, quase sempre significa
    # que a aba não foi atualizada — não que nada mudou. Em 31/07/2026 isso aconteceu:
    # o snapshot repetiu 205 de 24/07, Δ saiu 0, e o relatório oficial dizia 204 / -01.
    if prev is not None and atual == prev:
        print(f"  [headcount] CONTAGEM idêntica à de {max(k for k in hist if k < fim.isoformat())} "
              f"em todos os locais ({total} total) — conferir se a aba foi atualizada; "
              f"Δ desta semana sai 0 por isso")
    # grava snapshot desta run (idempotente por data)
    hist[fim.isoformat()] = atual
    HIST_HEADCOUNT.write_text(json.dumps(hist, ensure_ascii=False, indent=2), encoding="utf-8")


# ------------------------------------------------------------------
# Seção 4 — TERCEIROS / COMERCIAIS (embriões a entregar / a receber)
# ------------------------------------------------------------------
def build_comerciais(rep: Report):
    wb = _load(EMB_COMERCIAIS)
    painel = {}
    ws = wb["PAINEL"]
    grid = list(ws.iter_rows(values_only=True))
    for r in grid:
        cells = [_s(c) for c in r]
        for j, c in enumerate(cells):
            if c and c.upper() in ("A FAZER", "EM ANDAMENTO", "TOTAL VENDIDOS", "TOTAL COMPRADOS"):
                # o número costuma estar mais à direita na mesma linha
                nums = [x for x in r if isinstance(x, (int, float))]
                if nums:
                    painel[c] = nums[-1]
    rep.terceiros = {"painel": painel}
    # tabelas detalhe
    def _tab(sheet, hdr_row=3):
        ws = wb[sheet]
        data = list(ws.iter_rows(values_only=True))
        hdr = [_s(c) for c in data[hdr_row - 1]]
        rows = []
        for r in data[hdr_row:]:
            if all(c is None for c in r):
                continue
            if _s(r[0]) is None and _s(r[1]) is None:
                continue
            rows.append({(hdr[i] or f"c{i}"): _s(v) for i, v in enumerate(r) if hdr[i]})
        return rows
    rep.detalhe["a_entregar"] = _tab("ENTREGAR")
    rep.detalhe["a_receber"] = _tab("RECEBER")
    wb.close()


# ------------------------------------------------------------------
# Orquestração
# ------------------------------------------------------------------
def _calendario_dos_snapshots(hist: dict) -> list:
    """Seletor = semanas com snapshot (docx-semente ou capturadas pelo script).
    Chave = data de referência (fim da semana). Janela = docx anterior+1 .. este."""
    keys = sorted(k for k in hist if _is_iso(k))
    semanas = []
    for i, wid in enumerate(keys):
        ref = date.fromisoformat(wid)
        ini = (date.fromisoformat(keys[i - 1]) + timedelta(days=1)) if i > 0 else (ref - timedelta(days=7))
        semanas.append({
            "id": wid,
            "ini": ini.isoformat(),
            "fim": wid,
            "iso": ref.isocalendar()[1],
            "source": hist[wid].get("source"),
        })
    return semanas


def _is_iso(s: str) -> bool:
    try:
        date.fromisoformat(s)
        return True
    except (ValueError, TypeError):
        return False


def build_report(ini: date, fim: date) -> Report:
    rep = Report(semana_inicio=ini.isoformat(), semana_fim=fim.isoformat())
    build_producao(rep, ini, fim)
    build_receptoras(rep)
    build_headcount(rep)
    build_headcount_delta(rep, fim)
    build_movimentacao(rep, ini, fim)
    build_comerciais(rep)
    build_pendentes(rep)
    rep.semana_atual = fim.isoformat()           # semana de referência = data do fechamento
    rep.docx_ref = _load_docx_ref()               # relatórios oficiais (validação + seed do 1º caso)
    _compute_movimento(rep)                       # saídas/entradas = diff da população contada
    _compute_confirmados_diff(rep)                # confirmados na semana = diff de confirmados (forward)
    _persist_snapshot(rep)                        # congela snapshot CALCULADO desta semana
    rep.calendario = _calendario_dos_snapshots(rep.snapshots)
    return rep


def _load_hist() -> dict:
    if HIST_SNAPSHOTS.exists():
        try:
            return json.loads(HIST_SNAPSHOTS.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def _populacao_contada(roster, receptoras_locais) -> list:
    """Conjunto que o headcount conta: animais do plantel + receptoras nos NOSSOS
    locais. Receptora que vai pro sócio sai da contagem (CONTAGEM só tem receptora
    na fazenda e no arrendamento), então tem de contar como saída."""
    return sorted(set(roster or []) | set(receptoras_locais or {}))


def _descreve_mov(nome: str, info_ant: dict, info_atual: dict) -> dict:
    """Uma saída/entrada com contexto: '309' virou 'receptora 309, prenha de JAVA x
    QUEBRUTO, Pao Grande -> sócio'."""
    ant, atual = info_ant.get(nome), info_atual.get(nome)
    ref = atual or ant or {}
    return {
        "animal": nome,
        "tipo": "RECEPTORA" if (ant or atual) else "ANIMAL",
        "local_saida": (ant or {}).get("local"),
        "local_entrada": (atual or {}).get("local"),
        "status": ref.get("status"),
        "embriao": ref.get("embriao"),
        "obs": ref.get("obs"),
    }


def _compute_movimento(rep: Report):
    """Saídas/entradas na semana.

    Ordem de preferência:
      1. aba SAIDAS-ENTRADAS do controle mensal, classificada — entrada = nascimento
         ou compra, saída = venda ou morte. É a fonte oficial quando o haras preencher.
      2. diff da POPULAÇÃO CONTADA vs o snapshot anterior.
      3. bootstrap do relatório oficial em Word, na primeira captura.

    O diff era só do roster do plantel e por isso vivia dando 0: o plantel (aba
    PLANTEL) não mexe quando a movimentação é de receptora. Em 07/08/2026 a única
    saída da semana foi a receptora 309 indo pro sócio — diff de roster: 0; diff da
    população contada: 1 saída, 0 entradas, que é o '+00 / -01' do relatório e fecha
    com o Δ do headcount (204 -> 203).
    """
    if rep.saidas_planilha:
        ent, sai = rep.saidas_planilha["ENTRADA"], rep.saidas_planilha["SAIDA"]
        rep.saidas["saidas_semana"] = len(sai)
        rep.saidas["entradas_semana"] = len(ent)
        rep.saidas["fonte"] = "SAIDAS-ENTRADAS"
        rep.detalhe["saidas_diff"] = sai
        rep.detalhe["entradas_diff"] = ent
        _conferir_delta(rep)
        return

    rep.populacao = _populacao_contada(rep.roster, rep.receptoras_locais)
    hist = _load_hist()
    prev_pop = None
    for wid in sorted(hist):
        if wid < rep.semana_atual and hist[wid].get("populacao"):
            prev_pop = hist[wid]["populacao"]
    if prev_pop is None:
        # Bootstrap: nenhuma semana anterior guardou a população (campo novo). Monta a
        # anterior com o roster daquele snapshot + o arquivo de receptoras anterior.
        prev_snap = None
        for wid in sorted(hist):
            if wid < rep.semana_atual and hist[wid].get("roster"):
                prev_snap = hist[wid]
        anteriores = _receptoras_arquivos()[1:]
        if prev_snap and anteriores:
            prev_pop = _populacao_contada(prev_snap["roster"],
                                          _receptoras_locais(anteriores[0]))
            print(f"  [saídas/entradas] primeira semana com população guardada: "
                  f"receptoras da semana anterior vindas de {anteriores[0].name}")

    cur = set(rep.populacao)
    if prev_pop and cur:
        rep.saidas["fonte"] = "diff_populacao"
        prev = set(prev_pop)
        saidas = sorted(prev - cur)
        entradas = sorted(cur - prev)
        rep.saidas["saidas_semana"] = len(saidas)
        rep.saidas["entradas_semana"] = len(entradas)
        # nome sozinho não diz nada — receptora é número ("309"). Anexa o que ela é,
        # de onde saiu, pra onde foi e a observação da planilha.
        info_atual = _receptoras_info()
        info_ant = _receptoras_info(_receptoras_arquivos()[1]) if len(_receptoras_arquivos()) > 1 else {}
        rep.detalhe["saidas_diff"] = [_descreve_mov(n, info_ant, info_atual) for n in saidas]
        rep.detalhe["entradas_diff"] = [_descreve_mov(n, info_ant, info_atual) for n in entradas]
    else:
        # BOOTSTRAP: 1ª captura, sem semana anterior p/ diff → semeia do relatório oficial
        dx = (rep.docx_ref or {}).get(rep.semana_atual, {}).get("saidas", {})
        rep.saidas["saidas_semana"] = dx.get("saidas_semana")
        rep.saidas["entradas_semana"] = dx.get("entradas")
        rep.saidas["fonte"] = "docx"
        rep.saidas["_seed"] = "docx" if dx else None
    _conferir_delta(rep)


def _conferir_delta(rep: Report):
    """Δ do headcount = entradas - saídas. As duas contas vêm de fontes diferentes
    (CONTAGEM vs diff da população), então uma confere a outra. Divergir significa
    movimentação que não passou pelas planilhas — tem de aparecer, não sumir."""
    ent, sai = rep.saidas.get("entradas_semana"), rep.saidas.get("saidas_semana")
    rep.headcount["delta_entradas"] = ent
    rep.headcount["delta_saidas"] = sai
    delta = rep.headcount.get("delta")
    if None in (ent, sai) or delta is None:
        return
    if ent - sai != delta:
        print(f"  [Δ] headcount variou {delta:+d} mas o diff da população dá "
              f"+{ent}/-{sai} (líquido {ent - sai:+d}) — conferir: uma das duas "
              f"fontes não registrou alguma movimentação")


def _compute_confirmados_diff(rep: Report):
    """Confirmados na semana = embriões que viraram +/-=OK vs o snapshot anterior
    (novos no conjunto de confirmados). Forward: precisa de 2 semanas capturadas."""
    hist = _load_hist()
    prev_keys = None
    for wid in sorted(hist):
        if wid < rep.semana_atual and hist[wid].get("confirmed_keys") is not None:
            prev_keys = hist[wid]["confirmed_keys"]
    cur = {e["key"]: e for e in rep.confirmed}
    if prev_keys is not None:
        novos = [e for k, e in cur.items() if k not in set(prev_keys)]
        rep.producao["confirmados_semana"] = len(novos)
        rep.detalhe["confirmados_semana"] = novos
    else:
        # BOOTSTRAP: 1ª captura → semeia do relatório oficial
        dx = (rep.docx_ref or {}).get(rep.semana_atual, {}).get("producao", {})
        rep.producao["confirmados_semana"] = dx.get("confirmados_semana")
        rep.detalhe["confirmados_semana"] = []

    # ACUMULADO NO MÊS = novos confirmados desde o último snapshot ANTES do mês corrente
    # (docx "--" = 0; acumulado parado no mês → 0). Não usa IA+60 (proxy ruim).
    month_start = rep.semana_atual[:8] + "01"          # YYYY-MM-01
    prev_month_keys = None
    for wid in sorted(hist):
        if wid < month_start and hist[wid].get("confirmed_keys") is not None:
            prev_month_keys = hist[wid]["confirmed_keys"]
    if prev_month_keys is not None:
        pm = set(prev_month_keys)
        rep.producao["acumulado_mes"] = sum(1 for k in cur if k not in pm)
    else:
        dxp = (rep.docx_ref or {}).get(rep.semana_atual, {}).get("producao", {})
        rep.producao["acumulado_mes"] = dxp.get("acumulado_mes") or 0   # "--" = 0


def _snap_from_rep(rep: Report) -> dict:
    """Snapshot completo desta run (mesmo schema do _map_docx_to_snap)."""
    return {
        "source": "extractor",
        "acumulado_estacao": rep.producao.get("acumulado_estacao"),
        "confirmados_semana": rep.producao.get("confirmados_semana"),
        "acumulado_mes": rep.producao.get("acumulado_mes"),
        "nascimentos": rep.producao.get("nascimentos"),
        "abortos_obitos": rep.producao.get("abortos_obitos"),
        "acumulado_estacao_split": rep.producao.get("acumulado_estacao_split"),
        "receptoras": rep.receptoras,
        "headcount": {k: v for k, v in rep.headcount.items() if k != "detalhe"},
        "terceiros": {k: v for k, v in rep.terceiros.items() if k != "painel"},
        "movimento": {"saidas": rep.saidas.get("saidas_semana"),
                      "entradas": rep.saidas.get("entradas_semana"),
                      "transferencias": rep.saidas.get("transferencias_semana")},
        "detalhe": {
            "confirmados": rep.detalhe.get("confirmados_semana"),
            "nascimentos": rep.detalhe.get("nascimentos_semana"),
            "abortos_obitos": rep.detalhe.get("abortos_obitos_semana"),
            "saidas": rep.detalhe.get("saidas_diff"),
            "entradas": rep.detalhe.get("entradas_diff"),
            "pendentes_saida": rep.detalhe.get("pendentes_saida"),
            "terceiros_vendidos": rep.detalhe.get("terceiros_vendidos"),
            "terceiros_sociedade": rep.detalhe.get("terceiros_sociedade"),
            "transferencias": rep.detalhe.get("transferencias_internas"),
        },
        "roster": rep.roster,
        "receptoras_locais": rep.receptoras_locais,
        "populacao": rep.populacao,
        "confirmed_keys": [e["key"] for e in rep.confirmed],
    }


def _persist_snapshot(rep: Report):
    """Congela o snapshot CALCULADO desta semana. O que vai pro dash é sempre calculado
    aqui — o docx nunca vira dado, só valida (ver rep.docx_ref)."""
    HIST_SNAPSHOTS.parent.mkdir(parents=True, exist_ok=True)
    hist = _load_hist()
    hist[rep.semana_atual] = _snap_from_rep(rep)
    HIST_SNAPSHOTS.write_text(json.dumps(hist, ensure_ascii=False, indent=2), encoding="utf-8")
    rep.snapshots = hist


def _load_docx_ref() -> dict:
    """Números dos relatórios oficiais (semanal_docx.json), keyed by data de referência.
    Usado SÓ como overlay de validação no dash — nunca como dado."""
    fp = BASES_DIR / "semanal_docx.json"
    if not fp.exists():
        return {}
    out = {}
    for w in json.loads(fp.read_text(encoding="utf-8")):
        out[w["ref"]] = w
    return out


def _cmp(label, got, target):
    flag = "" if target is None else ("  OK" if got == target else f"  (docx: {target})")
    print(f"    {label:38} {got}{flag}")


def print_report(rep: Report):
    print("=" * 66)
    print(f"ATUALIZAÇÃO SEMANAL — {rep.semana_inicio} a {rep.semana_fim}")
    print("=" * 66)
    print("Fontes:")
    for k, v in rep.fontes.items():
        print(f"    {k}: {v}")
    print("\n1) PRODUÇÃO")
    _cmp("Acumulado na estação", rep.producao["acumulado_estacao"], DOCX_1707["acumulado_estacao"])
    print(f"        split PG/sócio/vendido: {rep.producao['acumulado_estacao_split']}")
    _cmp("Embriões confirmados na semana", rep.producao["confirmados_semana"], DOCX_1707["confirmados_semana"])
    _cmp("Acumulado no mês", rep.producao["acumulado_mes"], None)
    _cmp("Nascimentos na semana", rep.producao["nascimentos"], DOCX_1707["nascimentos"])
    _cmp("Abortos / óbitos na semana", rep.producao["abortos_obitos"], DOCX_1707["abortos_obitos"])
    print("\n2) RECEPTORAS (rebanho ativo = FPG + ARRENDAMENTO)")
    _cmp("Total receptoras", rep.receptoras["total"], DOCX_1707["receptoras_total"])
    _cmp("Prenhas", rep.receptoras["prenhas"], DOCX_1707["receptoras_prenhas"])
    _cmp("Vazias", rep.receptoras["vazias"], DOCX_1707["receptoras_vazias"])
    _cmp("Índice eficiência (vazias/doadoras)", rep.receptoras["indice_eficiencia"], None)
    print("\n3) HEADCOUNT")
    _cmp("Total geral", rep.headcount["total"], DOCX_1707["headcount_total"])
    _cmp("Fazenda Pão Grande", rep.headcount["fazenda_pg"], DOCX_1707["headcount_fpg"])
    _cmp("Arrendamento", rep.headcount["arrendamento"], DOCX_1707["headcount_arr"])
    _cmp("Centro de Treinamento", rep.headcount["cte"], DOCX_1707["headcount_cte"])
    _cmp("Sócios", rep.headcount["socio"], DOCX_1707["headcount_soc"])
    _cmp("Δ vs semana anterior", rep.headcount.get("delta"), None)
    print("\n4) TERCEIROS / COMERCIAIS")
    _cmp("Vendidos pendentes de saída", rep.terceiros.get("vendidos_pendentes"), DOCX_1707["vendidos_pendentes"])
    _cmp("Em sociedade pendentes", rep.terceiros.get("sociedade_pendentes"), DOCX_1707["sociedade_pendentes"])
    _cmp("Terceiros na propriedade", rep.terceiros.get("terceiros_propriedade"), None)
    print(f"    painel comerciais: {rep.terceiros.get('painel')}")
    print("\n5) SAÍDAS")
    _cmp("Saídas na semana", rep.saidas["saidas_semana"], DOCX_1707["saidas_semana"])
    _cmp("Entradas na semana", rep.saidas["entradas_semana"], None)
    _cmp("Transferências internas", rep.saidas["transferencias_semana"], None)


def _parse_d(s: str) -> date:
    s = s.strip()
    if re.match(r"^\d{2}/\d{2}/\d{4}$", s):
        d, m, y = s.split("/"); return date(int(y), int(m), int(d))
    return date.fromisoformat(s)


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    args = sys.argv[1:]
    if len(args) >= 2:
        ini, fim = _parse_d(args[0]), _parse_d(args[1])
    elif len(args) == 1:
        fim = _parse_d(args[0]); ini = fim - timedelta(days=14)
    else:
        fim = date.today(); ini = fim - timedelta(days=7)
    rep = build_report(ini, fim)
    print_report(rep)
    JSON_OUT.parent.mkdir(parents=True, exist_ok=True)
    JSON_OUT.write_text(json.dumps(asdict(rep), ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n-> {JSON_OUT.name} gravado ({JSON_OUT.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
