"""Popula a aba Plantel / Movimentação do hub com o histórico que já existe.

A aba lê o mês do Supabase (`plantel_snapshot`). Sem semear, ela abre vazia e
espera alguém importar arquivo — o que não faz sentido para meses que já estão
fechados e cujos arquivos estão todos no Drive. Este script faz a carga inicial:

  para cada mês, o FECHAMENTO do haras (aba PLANTEL + aba MOVIMENTAÇÕES) e, do
  mapa do Luxor daquele mês, a ATRIBUIÇÃO de quem é cada animal (Carla/Eduardo),
  que não sai do arquivo do haras e não se deriva do sufixo.

Formato gravado em `linhas` (jsonb): {cab, ix, rows}. É autocontido de propósito
— o cabeçalho e o mapa de colunas viajam junto com as linhas, porque a mesma
planilha muda de layout entre versões (a aba MOVIMENTAÇÕES começa na coluna B, o
`VALOR` já se chamou 'VALOR 100% ...' e hoje é 'VALOR (R$)'). Guardar só as
linhas obrigaria a tela a adivinhar em que coluna está o quê.

Requer SUPABASE_SERVICE_ROLE_KEY no .env (mesma credencial do publish_hub.py):
a RLS de escrita exige sessão de admin, e a anon key do site não passa.

Uso:
    python tools/seed_plantel_hub.py                 # dez/25 a jul/26
    python tools/seed_plantel_hub.py --de 2026-03 --ate 2026-07
    python tools/seed_plantel_hub.py --listar        # só mostra o que achou
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import unicodedata
from datetime import date, datetime
from pathlib import Path

import requests
from dotenv import dotenv_values

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR / "scripts"))

TABELA = "plantel_snapshot"
MESES_ABR = {"JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4, "MAI": 5, "JUN": 6,
             "JUL": 7, "AGO": 8, "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12}
MAPA_DIR = Path(r"G:\Drives compartilhados\Luxor Controladoria\Relatórios Gerenciais"
                r"\RELATORIOS - OPERAÇÃO HARAS E FAZENDA PG\Posição Equinos"
                r"\PLANTEL - Movimentações")

# mesmos alvos do plantel.js: a coluna é achada pelo RÓTULO, nunca por índice
ALVOS = {
    "qtde": ["QTDE"], "letra": ["LETRA"], "sufixo": ["SUFIXO"], "nome": ["NOME"],
    "sexo": ["SEXO"], "categoria": ["CATEGORIA"], "status": ["STATUS PLANTEL"],
    "local": ["LOCAL"], "mae": ["MAE"], "pai": ["PAI"], "nascimento": ["NASCIMENTO"],
    "cota": ["COTAS"], "socio": ["NOME SOCIO"], "condicao": ["CONDICAO ATUAL"],
    "valor": ["VALOR (R$)", "VALOR 100%", "VALOR"], "comissao": ["COMISSAO", "COMISS"],
    "hpgCota": ["PLANTEL HPG"], "edCota": ["PLANTEL EDUARDO"],
}


def norm(s) -> str:
    s = "" if s is None else str(s)
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip().upper().strip('"')


def idx_de(cab, alvos):
    for alvo in alvos:
        for i, c in enumerate(cab):
            if norm(c).startswith(alvo):
                return i
    return -1


def mapa_colunas(cab) -> dict:
    ix = {campo: idx_de(cab, alvos) for campo, alvos in ALVOS.items()}
    ix["hpgVal"] = ix["hpgCota"] + 1 if ix["hpgCota"] >= 0 else -1
    ix["edVal"] = ix["edCota"] + 1 if ix["edCota"] >= 0 else -1
    return ix


def _json_seguro(v):
    """Excel devolve datetime/Decimal; o JSON do Postgres não aceita. Data vira
    ISO, que é o que o `new Date(...)` da tela entende."""
    if isinstance(v, (datetime, date)):
        return v.isoformat()
    if isinstance(v, float) and (v != v):     # NaN
        return None
    return v


def le_planilha(caminho: Path) -> dict:
    """{cab, ix, rows, log} da planilha do haras (ou do mapa do Luxor)."""
    import PGSemanalReport as R

    wb = R._load(caminho)
    nome_pl = next((s for s in wb.sheetnames if norm(s).startswith("PLANTEL")), None)
    if not nome_pl:
        raise SystemExit(f"{caminho.name}: não achei a aba PLANTEL")
    tudo = [list(r) for r in wb[nome_pl].iter_rows(values_only=True)]
    i_cab = next((i for i, r in enumerate(tudo)
                  if any(norm(c) == "NOME" for c in r)
                  and any(norm(c).startswith("CATEGORIA") for c in r)), -1)
    if i_cab < 0:
        raise SystemExit(f"{caminho.name}: não achei o cabeçalho da aba PLANTEL")
    cab = [_json_seguro(c) for c in tudo[i_cab]]
    ix = mapa_colunas(tudo[i_cab])
    rows = [[_json_seguro(c) for c in r] for r in tudo[i_cab + 1:] if r[ix["nome"]]]

    # log de ocorrências: PRODUTO / DATA / OCORRENCIA (só o arquivo do haras tem)
    log = []
    nome_mv = next((s for s in wb.sheetnames if norm(s).startswith("MOVIMENTA")), None)
    if nome_mv:
        linhas_mv = [list(r) for r in wb[nome_mv].iter_rows(values_only=True)]
        i_c = next((i for i, r in enumerate(linhas_mv)
                    if any(norm(c).startswith("OCORRENCIA") for c in r)), -1)
        if i_c >= 0:
            c = linhas_mv[i_c]
            ip, idt, io = idx_de(c, ["PRODUTO"]), idx_de(c, ["DATA"]), idx_de(c, ["OCORRENCIA"])
            if min(ip, idt, io) >= 0:
                for r in linhas_mv[i_c + 1:]:
                    if len(r) <= max(ip, idt, io) or not r[ip] or not isinstance(r[idt], (datetime, date)):
                        continue
                    log.append({"produto": str(r[ip]).strip(),
                                "data": _json_seguro(r[idt]),
                                "ocorrencia": str(r[io] or "").strip()})
    wb.close()
    return {"cab": cab, "ix": ix, "rows": rows, "log": log}


def chave(row, ix) -> str:
    return norm(row[ix["nome"]]) + "|" + norm(row[ix["letra"]] if ix["letra"] >= 0 else "")


def atribuicao_do_mapa(mes: str) -> dict:
    """{chave: hpg|eduardo|nenhum} a partir do mapa do Luxor daquele mês.

    A atribuição é POR MÊS: o dono muda no ano (animal da Carla em janeiro vira
    'nenhum' depois de repassado). Usar uma atribuição só para todos os meses põe
    o saldo de janeiro R$ 2,49 milhões abaixo do liberado."""
    ano, mm = mes.split("-")
    pasta = MAPA_DIR / ano
    if not pasta.exists():
        return {}
    abr = [k for k, v in MESES_ABR.items() if v == int(mm)][0].capitalize()
    cand = [f for f in pasta.iterdir()
            if "Movimenta" in f.name and f"({abr} {ano})" in f.name and not f.name.startswith("~$")]
    if not cand:
        return {}
    d = le_planilha(max(cand, key=lambda f: f.stat().st_mtime))
    ix = d["ix"]
    if ix["hpgCota"] < 0:
        return {}
    out = {}
    for r in d["rows"]:
        pega = lambda i: (r[i] if i >= 0 and i < len(r) else None) or 0
        tem_hpg = pega(ix["hpgCota"]) or pega(ix["hpgVal"])
        tem_ed = pega(ix["edCota"]) or pega(ix["edVal"])
        out[chave(r, ix)] = "hpg" if tem_hpg else "eduardo" if tem_ed else "nenhum"
    return out


def resumo_liberado() -> dict:
    """{mes: {saldo_ini, saldo_fim, compras, producao, vendas, mortes, reaval}} da
    aba `Resumo Contabil` do mapa mais recente.

    É o número que foi de fato liberado, e o hub confere contra ele em vez de
    tentar reproduzi-lo às cegas. Importa porque as duas fontes já divergiram:
    a coluna PLANTEL HPG do mapa de JANEIRO soma R$ 18.053.880, enquanto o
    Resumo Contábil dá R$ 18.445.077 para o mesmo mês — janeiro foi revisado
    depois que aquele mapa foi feito. Nos outros seis meses de 2026 as duas
    fontes batem na casa dos centavos."""
    import PGSemanalReport as R

    pasta = MAPA_DIR / "2026"
    mapas = [f for f in pasta.iterdir() if "Movimenta" in f.name and not f.name.startswith("~$")]
    if not mapas:
        return {}
    wb = R._load(max(mapas, key=lambda f: f.stat().st_mtime))
    aba = next((s for s in wb.sheetnames if "RESUMO" in norm(s)), None)
    if not aba:
        wb.close()
        return {}
    rotulos = {
        "SALDO INICIAL": "saldo_ini", "SALDO FINAL": "saldo_fim",
        "(+) COMPRAS": "compras", "(+) PRODUCAO EMBRIOES": "producao",
        "(-) BAIXA VENDAS": "vendas", "(-) BAIXA MORTES E DOACOES": "mortes",
        "(+/-) REAVALIACOES": "reaval",
    }
    col_mes, out = {}, {}
    for r in wb[aba].iter_rows(values_only=True):
        rot = norm(r[1]) if len(r) > 1 else ""
        if rot == "TITULO":
            for j, c in enumerate(r):
                m = MESES_ABR.get(norm(c)[:3])
                if m and j > 1:
                    col_mes[j] = f"2026-{m:02d}"
            continue
        campo = rotulos.get(rot)
        if not campo:
            continue
        for j, mes in col_mes.items():
            v = r[j] if j < len(r) else None
            if isinstance(v, (int, float)):
                out.setdefault(mes, {})[campo] = round(float(v), 2)
    wb.close()
    return out


def acha_fechamento(mes: str) -> Path | None:
    """Fechamento do haras daquele mês, na pasta PLANTEL do Drive.

    O nome traz o MÊS DE REFERÊNCIA (`..._JUL_26.xlsx`) e o prefixo é a data em
    que foi gerado, já no mês seguinte. Cópia de trabalho ('EDITAR ...') fica de
    fora: ela é editada durante o mês seguinte e mistura dois meses."""
    from _pg_common import DRIVE_ROOT

    ano, mm = mes.split("-")
    abr = [k for k, v in MESES_ABR.items() if v == int(mm)][0]
    alvo = f"_{abr}_{ano[2:]}"
    cands = []
    for d in sorted((DRIVE_ROOT / "PLANTEL").glob("Estação *"), reverse=True):
        for f in d.glob("*CONTROLE_DE_PLANTEL_PAO_GRANDE_*.xlsx"):
            n = norm(f.name)
            if alvo in n and not n.startswith("EDITAR") and not f.name.startswith("~$"):
                cands.append(f)
    if not cands:
        return None
    # mais recente por data no prefixo do nome, e depois por mtime — a mesma
    # regra de escolha de versão do LxEtlPlantel
    return max(cands, key=lambda f: (re.match(r"(\d{6})", f.name) and f.name[:6] or "", f.stat().st_mtime))


def env():
    cfg = dotenv_values(BASE_DIR / ".env")
    url = (cfg.get("SUPABASE_URL") or "").rstrip("/")
    key = cfg.get("SUPABASE_SERVICE_ROLE_KEY")
    if not url or not key:
        sys.exit("Faltam SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY no .env da raiz.")
    return url, key


def sobe(url, key, linha: dict) -> None:
    r = requests.post(
        f"{url}/rest/v1/{TABELA}?on_conflict=mes",
        headers={"apikey": key, "Authorization": f"Bearer {key}",
                 "Content-Type": "application/json",
                 "Prefer": "resolution=merge-duplicates,return=minimal"},
        data=json.dumps(linha, ensure_ascii=False).encode("utf-8"), timeout=300)
    if r.status_code >= 300:
        raise SystemExit(f"{linha['mes']}: HTTP {r.status_code} — {r.text[:400]}")


def meses_entre(de: str, ate: str) -> list[str]:
    out, a, m = [], *map(int, de.split("-"))
    fim_a, fim_m = map(int, ate.split("-"))
    while (a, m) <= (fim_a, fim_m):
        out.append(f"{a:04d}-{m:02d}")
        a, m = (a + 1, 1) if m == 12 else (a, m + 1)
    return out


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--de", default="2025-12")
    ap.add_argument("--ate", default="2026-07")
    ap.add_argument("--listar", action="store_true", help="mostra os arquivos e sai")
    args = ap.parse_args()

    alvos = []
    for mes in meses_entre(args.de, args.ate):
        f = acha_fechamento(mes)
        alvos.append((mes, f))
        print(f"{mes}: {f.name if f else '— sem fechamento no Drive'}")
    if args.listar:
        return
    faltando = [m for m, f in alvos if not f]
    if faltando:
        print(f"\nsem fechamento: {', '.join(faltando)} — esses meses não entram")

    url, key = env()
    liberado = resumo_liberado()
    print(f"resumo contábil liberado: {len(liberado)} meses")
    for mes, f in alvos:
        if not f:
            continue
        d = le_planilha(f)
        at = atribuicao_do_mapa(mes)
        payload = {
            "mes": mes,
            "arquivo": f.name,
            # o `liberado` viaja junto com o mês pra a aba Checks confrontar sem
            # precisar de outra tabela
            "linhas": {"cab": d["cab"], "ix": d["ix"], "rows": d["rows"],
                       "liberado": liberado.get(mes)},
            "log": d["log"],
            "atribuicao": at,
        }
        tam = len(json.dumps(payload, ensure_ascii=False)) / 1024
        sobe(url, key, payload)
        print(f"  {mes}: {len(d['rows'])} animais · {len(d['log'])} ocorrências · "
              f"{len(at)} atribuições · {tam:.0f} KB")

    print("\npronto — a aba Plantel / Movimentação abre com esses meses.")


if __name__ == "__main__":
    main()
