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
EMB_MATRIZES_GRUPO = (
    DRIVE_ROOT / "PLANILHAS SEMANAIS"
    / "EMBRIÕES E MATRIZES - MODELO ENVIAR NO GRUPO 3.xlsx"
)
EMB_COMERCIAIS = DRIVE_ROOT / "REPRODUÇÃO" / "EMBRIOES A ENTREGAR - A RECEBER.xlsx"
ESTACAO_MASTER_DIR = (
    DRIVE_ROOT / "REPRODUÇÃO" / "ESTAÇÃO DE MONTA" / "Estação 2025-2026"
)
# Receptoras (aba ATUALIZAÇÃO SEMANAL pré-agregada) e Mapa de Vendas — por data no prefixo
RECEPTORAS_DIR = DRIVE_ROOT / "PLANTEL" / "Estação 2025-2026"
MAPA_VENDAS_DIR = DRIVE_ROOT / "VENDAS" / "MAPAS DE VENDAS" / "Estação 2025-2026"
# CONTROLE_DE_PLANTEL mensal (aba MOVIMENTAÇÕES datada) — mesma pasta das receptoras
CONTROLE_MENSAL_DIR = RECEPTORAS_DIR
# Animais para sair (vendidos/sociedade pendentes) — aba ANIMAIS VENDIDOS
ANIMAIS_SAIR_DIR = DRIVE_ROOT / "PLANILHAS PARA O EDUARDO"

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
    confirmed: list = field(default_factory=list)  # embriões confirmados (+/-=OK) p/ diff semanal
    docx_ref: dict = field(default_factory=dict)  # números dos relatórios oficiais (SÓ p/ validar)


# ------------------------------------------------------------------
# Seção 1 — PRODUÇÃO (master da estação de monta)
# ------------------------------------------------------------------
def _acumulado_planejamento(wb) -> int:
    """Acumulado na estação = soma da coluna 'TOTAL EMBRIÕES' (REAL, idx8) da aba PLANEJAMENTO,
    linhas de doadora (col0 numérica). Regra oficial confirmada pelo usuário."""
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

    # ACUMULADO NA ESTAÇÃO = SUM(PLANEJAMENTO 'TOTAL EMBRIÕES' REAL) — regra confirmada pelo
    # usuário (2026-07-21). É o número oficial do relatório; bate quando o arquivo está fresco.
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

    rep.producao = {
        "acumulado_estacao": acumulado_planejamento,
        "acumulado_estacao_split": _split(confirmados),
        "confirmados_semana": None,   # _compute_confirmados_diff
        "acumulado_mes": None,        # _compute_confirmados_diff
        "nascimentos": len(nascimentos),
        "abortos_obitos": len(abortos) + len(obitos),
    }
    def _produto(e):   # nome do animal nascido (ou descrição sexo — doadora × garanhão)
        base = f"{e.get('doadora') or ''} × {e.get('garanhao') or ''}".strip(" ×")
        sx = {"M": "Macho", "F": "Fêmea"}.get((e.get("sexo_potro") or "").upper(), e.get("sexo_potro") or "")
        return e.get("nome_potro") or (f"{sx} — {base}" if base else sx) or "--"
    rep.detalhe["confirmados_semana"] = na_semana
    rep.detalhe["nascimentos_semana"] = [dict(e, produto=_produto(e)) for e in nascimentos]
    rep.detalhe["abortos_obitos_semana"] = abortos + obitos
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
    doadoras = _count_doadoras()   # denominador do índice
    rep.receptoras = {
        "total": pren + vaz,
        "prenhas": pren,
        "vazias": vaz,
        "doadoras": doadoras,
        "indice_eficiencia": round(vaz / doadoras, 1) if doadoras else None,
    }


def _count_doadoras() -> int:
    """Doadoras do plantel = CATEGORIA 'DOADORA' na aba PLANTEL do CONTROLE PLANTEL semanal.
    (denominador do índice de eficiência vazias/doadoras)."""
    wb = _load(CONTROLE_PLANTEL_SEMANAL)
    ws = wb["PLANTEL"]
    n = 0
    for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if i < 2 or r[0] is None:
            continue
        if _norm(r[2]) == "DOADORA":
            n += 1
    wb.close()
    return n


# ------------------------------------------------------------------
# Seção 3 — HEADCOUNT (aba CONTAGEM, pré-agregada — bate exato)
# ------------------------------------------------------------------
def build_headcount(rep: Report):
    wb = _load(CONTROLE_PLANTEL_SEMANAL)
    ws = wb["CONTAGEM"]
    m = {}
    for r in ws.iter_rows(values_only=True):
        label = _norm(r[1])
        if label in ("FAZENDA", "ARRENDAMENTO", "CTE", "SOCIO", "TOTAL GERAL"):
            animais = r[2]; recept = r[3]; total = r[4]
            m[label] = {"animais": animais, "receptoras": recept, "total": total}
    rep.headcount = {
        "total": m.get("TOTAL GERAL", {}).get("total"),
        "fazenda_pg": m.get("FAZENDA", {}).get("total"),
        "arrendamento": m.get("ARRENDAMENTO", {}).get("total"),
        "cte": m.get("CTE", {}).get("total"),
        "socio": m.get("SOCIO", {}).get("total"),
        "detalhe": m,
    }
    wb.close()


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


def build_movimentacao(rep: Report, ini: date, fim: date):
    src = _latest_by_yymmdd(CONTROLE_MENSAL_DIR, "*CONTROLE_DE_PLANTEL_PAO_GRANDE_*.xlsx")
    rep.fontes["controle_plantel_mensal"] = src.name
    wb = _load(src)
    ws = wb["MOVIMENTAÇÕES"]
    evs = {"SAIDA": [], "ENTRADA": [], "TRANSFERENCIA": []}
    for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if i < 3 or r[3] is None:
            continue
        d = _dt(r[3])
        if not d:
            continue
        tipo = _categorize_mov(str(r[4] or "").upper())
        if tipo is None:
            continue
        evs[tipo].append({"animal": _s(r[2]), "data": d.isoformat(), "ocorrencia": _s(r[4])})
    inw = lambda x: ini <= date.fromisoformat(x["data"]) <= fim
    rep.saidas = {
        "saidas_semana": sum(1 for x in evs["SAIDA"] if inw(x)),
        "entradas_semana": sum(1 for x in evs["ENTRADA"] if inw(x)),
        "transferencias_semana": sum(1 for x in evs["TRANSFERENCIA"] if inw(x)),
    }
    # listas COMPLETAS datadas p/ filtro client-side
    rep.eventos["saidas"] = evs["SAIDA"]
    rep.eventos["entradas"] = evs["ENTRADA"]
    rep.eventos["transferencias"] = evs["TRANSFERENCIA"]
    wb.close()


# ------------------------------------------------------------------
# Seção 4b — PENDENTES DE SAÍDA / TERCEIROS  (CONTROLE PLANTEL aba PLANTEL)
# ------------------------------------------------------------------
def build_pendentes(rep: Report):
    # roster do plantel (p/ diff de saídas/entradas) — do CONTROLE PLANTEL
    wb = _load(CONTROLE_PLANTEL_SEMANAL)
    ws = wb["PLANTEL"]
    roster = []
    for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if i < 2 or r[0] is None:
            continue
        nome = _s(r[0])
        if nome:
            roster.append(nome)
    wb.close()
    rep.roster = sorted(set(roster))

    # VENDIDOS / SOCIEDADE pendentes = aba ANIMAIS VENDIDOS do "Animais para sair"
    # (validado vs docx 17/07: VENDA≠REPOSIÇÃO=2 vendidos; SOCIEDADE=2). col5=tipo, col6=obs.
    pend = []
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
    except FileNotFoundError:
        pass

    vendidos = [p for p in pend if p["tipo"] == "VENDA" and not p["reposicao"]]
    sociedade = [p for p in pend if p["tipo"] == "SOCIEDADE"]
    rep.terceiros.update({
        "vendidos_pendentes": len(vendidos),
        "sociedade_pendentes": len(sociedade),
        "terceiros_propriedade": len(vendidos),   # terceiros na propriedade = vendidos (docx)
        "doadoras_terceiros": None,
        "outros_terceiros": None,
    })
    rep.detalhe["terceiros_vendidos"] = vendidos          # só os 2 vendidos (bate KPI seção 4)
    rep.detalhe["pendentes_saida"] = pend                 # lista completa (seção 5)


def _latest_animais_sair() -> Path:
    """Arquivo 'Animais para sair*.xlsx' mais recente (por mtime — nome tem o ano)."""
    cands = [f for f in ANIMAIS_SAIR_DIR.glob("Animais para sair*.xlsx")
             if not f.name.startswith("~$")]
    if not cands:
        raise FileNotFoundError(f"Nenhum 'Animais para sair' em {ANIMAIS_SAIR_DIR}")
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
    # grava snapshot desta run (idempotente por data)
    hist[fim.isoformat()] = {"total": total, "fpg": rep.headcount.get("fazenda_pg"),
                             "arr": rep.headcount.get("arrendamento"),
                             "cte": rep.headcount.get("cte"), "soc": rep.headcount.get("socio")}
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
    # saídas/entradas ficam com o log MOVIMENTAÇÕES (build_movimentacao). NÃO uso diff de roster:
    # animal vendido continua no plantel até entrega, e o diff confundia nascimento com entrada.
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


def _compute_movimento(rep: Report):
    """Saídas/entradas na semana = diff do roster do plantel vs o snapshot anterior
    (lógica do operador). Precisa de 2 semanas capturadas pelo script; antes disso, None."""
    hist = _load_hist()
    prev_roster = None
    for wid in sorted(hist):
        if wid < rep.semana_atual and hist[wid].get("roster"):
            prev_roster = hist[wid]["roster"]
    cur = set(rep.roster or [])
    if prev_roster is not None and cur:
        prev = set(prev_roster)
        saidas = sorted(prev - cur)
        entradas = sorted(cur - prev)
        rep.saidas["saidas_semana"] = len(saidas)
        rep.saidas["entradas_semana"] = len(entradas)
        rep.detalhe["saidas_diff"] = [{"animal": n} for n in saidas]
        rep.detalhe["entradas_diff"] = [{"animal": n} for n in entradas]
    else:
        # BOOTSTRAP: 1ª captura, sem semana anterior p/ diff → semeia do relatório oficial
        dx = (rep.docx_ref or {}).get(rep.semana_atual, {}).get("saidas", {})
        rep.saidas["saidas_semana"] = dx.get("saidas_semana")
        rep.saidas["entradas_semana"] = dx.get("entradas")
        rep.saidas["_seed"] = "docx" if dx else None


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
        },
        "roster": rep.roster,
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
