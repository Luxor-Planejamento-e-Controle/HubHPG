"""Gera os dados das abas Comitê Mensal e Plantel/Movimentação do Hub HPG.

Lê as saídas que as rotinas JÁ produzem — não recalcula nada:
  - Comitê  : HubHPG/bases/base_bi.parquet          (PGBaseBI.py, neste repo)
  - Plantel : LuxorMonthlyP-CRoutines/PlantelHPG/   (LxEtlPlantel.py + LxMovimentacao.py)

Saída em dois formatos, igual ao LuxorP&CHub:
  assets/data/<nome>.js    -> window.<VAR> = {...}   (demo offline; file:// não faz fetch)
  assets/data/<nome>.json  -> mesmo payload          (fase gold: vai pro bucket privado)

Uso: python hub/tools/build_data.py
Requer: pandas, pyarrow
"""
import json
import sys
from pathlib import Path

import pandas as pd

HUB = Path(__file__).resolve().parent.parent
REPO = HUB.parent
OUT = HUB / "assets" / "data"

BASE_BI = REPO / "bases" / "base_bi.parquet"
# Fonte da verdade do plantel/movimentação = repo local do LuxorMonthlyP-CRoutines.
PLANTEL_DIR = Path(r"C:/Users/Arthur/repos/LuxorMonthlyP-CRoutines/PlantelHPG")


def write(nome, var, payload):
    OUT.mkdir(parents=True, exist_ok=True)
    js = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    (OUT / f"{nome}.json").write_text(js, encoding="utf-8")
    (OUT / f"{nome}.js").write_text(f"window.{var} = {js};\n", encoding="utf-8")
    print(f"[{nome}] {len(js) // 1024} KB -> assets/data/{nome}.js (+ .json)")


def tabela(df, cols):
    """Formato compacto {cols, rows}: o nome da coluna não se repete por linha,
    o que corta o JSON quase pela metade. O app.js reidrata com objify()."""
    df = df[cols].where(pd.notna(df[cols]), None)
    return {"cols": list(cols), "rows": df.to_numpy().tolist()}


def limpa(v):
    """NaN/NaT viram None (viram null no JSON); o resto passa como está."""
    return None if pd.isna(v) else v


# ---------------------------------------------------------------- Comitê
def build_comite():
    if not BASE_BI.exists():
        print(f"[comite] pulado — {BASE_BI} não existe (rode python scripts/PGBaseBI.py)")
        return
    d = pd.read_parquet(BASE_BI)
    d["mes"] = pd.to_datetime(d["mes_referencia"]).dt.strftime("%Y-%m")
    d = d.rename(columns={"status_plantel": "status", "in_slide_1": "s1", "in_slide_2": "s2"})
    d["sexo"] = d["sexo"].astype("string").str.strip().replace("", pd.NA)
    d["idade"] = pd.to_numeric(d["idade_anos"], errors="coerce").round(1)
    for c in ("cotas", "valor_100", "patrimonio_proporcional"):
        d[c] = pd.to_numeric(d[c], errors="coerce")
    d = d.rename(columns={"patrimonio_proporcional": "patrimonio"})
    d["s1"] = d["s1"].fillna(False).astype(bool)
    d["s2"] = d["s2"].fillna(False).astype(bool)

    cols = ["mes", "nome", "categoria", "sexo", "local", "status", "nota",
            "idade", "cotas", "valor_100", "patrimonio", "s1", "s2"]
    meses = sorted(d["mes"].unique())
    write("comite", "COMITE_DATA", {
        "fonte": BASE_BI.name,
        "meses": meses,
        "tabela": tabela(d.sort_values(["mes", "nome"]), cols),
    })


# ------------------------------------------------------- Plantel / Movimentação
def build_plantel():
    casc_f = PLANTEL_DIR / "mov_cascata.parquet"
    det_f = PLANTEL_DIR / "mov_detalhe.parquet"
    if not casc_f.exists() or not det_f.exists():
        print(f"[plantel] pulado — parquets não encontrados em {PLANTEL_DIR}\n"
              f"          rode o LxMovimentacao.py no repo LuxorMonthlyP-CRoutines")
        return

    casc = pd.read_parquet(casc_f).sort_values("mes")
    casc_cols = ["mes", "saldo_ini", "saldo_fim", "d_total", "compra", "producao",
                 "venda", "morte", "doacao", "reaval", "transferencia",
                 "saiu_controle", "check_dif"]
    faltando = [c for c in casc_cols if c not in casc.columns]
    if faltando:
        sys.exit(f"[plantel] mov_cascata mudou de forma — sem as colunas {faltando}")

    det = pd.read_parquet(det_f).sort_values(["mes", "nome"])
    det_cols = ["mes", "nome", "categoria", "tipo", "status_ant", "status_atual",
                "valor_ant", "valor_atual", "efeito"]

    write("plantel", "PLANTEL_DATA", {
        "fontes": {"cascata": casc_f.name, "detalhe": det_f.name},
        "cascata": tabela(casc, casc_cols),
        "detalhe": tabela(det, det_cols),
    })


if __name__ == "__main__":
    build_comite()
    build_plantel()
