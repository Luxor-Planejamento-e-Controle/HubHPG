"""
PGDataExtractor — extrai e enriquece o plantel de UM mês, salvando como parquet
no cache local. Não gera arquivos finais — esse é o passo 1 do pipeline.

Fluxo:
  1. Prompt MM/AAAA
  2. Localiza planilha correspondente no Drive (preferindo REAVALIAÇÃO)
  3. Cache local da fonte (read_only — nunca modifica)
  4. Lê sheet PLANTEL, enriquece (normalização + flags + patrimônio)
  5. Salva `_cache/parquet/{YYYY-MM}.parquet`

O passo 2 (PGBaseBI.py) consolida todos os parquets em `base_bi.xlsx`.

Uso interativo:
    python PGDataExtractor.py

Uso programático:
    from PGDataExtractor import extract_mes
    from _pg_common import parse_mes_aaaa
    extract_mes(parse_mes_aaaa("04/2026"))
"""

from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path

import openpyxl
import pandas as pd

from _pg_common import (
    PARQUET_DIR,
    MesRef,
    ensure_cache,
    find_compatible_source,
    prompt_mes_aaaa,
)

# --- regras de negócio (slides 1 e 2) ---
LOCAIS_SLIDE2 = ["FAZENDA PAO GRANDE", "SOCIO", "ARRENDAMENTO CESAR FURTADO"]
SUFIXOS_SLIDE1 = ["DA PAO GRANDE", "OUTRO"]
NOTAS_VALIDAS = {"A+", "A", "A-", "B+", "B", "B-", "C", "D"}

# --- extração ---
PLANTEL_HEADER_ROW = 4
PLANTEL_DATA_START = 5
PLANTEL_NCOLS = 30
_DATE_RE = re.compile(r"(\d{1,2})/(\d{1,2})/(\d{2,4})")


def _clean_header(value) -> str:
    if value is None:
        return ""
    return " ".join(str(value).replace("\n", " ").replace("\r", " ").split())


def _dedupe(headers: list[str]) -> list[str]:
    seen: dict[str, int] = {}
    out: list[str] = []
    for h in headers:
        base = h or "col"
        if base not in seen:
            seen[base] = 1
            out.append(base)
        else:
            seen[base] += 1
            out.append(f"{base}_{seen[base]}")
    return out


def _read_plantel_raw(source_path: Path) -> pd.DataFrame:
    """Lê apenas a sheet PLANTEL, sem tratamento."""
    wb = openpyxl.load_workbook(source_path, data_only=True, read_only=True)
    try:
        ws = wb["PLANTEL"]
        headers: list[str] | None = None
        rows: list[list] = []
        for r, row in enumerate(ws.iter_rows(values_only=True), start=1):
            if r == PLANTEL_HEADER_ROW:
                slice_ = list(row[:PLANTEL_NCOLS])
                slice_ += [None] * (PLANTEL_NCOLS - len(slice_))
                headers = _dedupe([_clean_header(v) for v in slice_])
                continue
            if r < PLANTEL_DATA_START:
                continue
            slice_ = list(row[:PLANTEL_NCOLS])
            slice_ += [None] * (PLANTEL_NCOLS - len(slice_))
            key = slice_[0]
            if key is None or (isinstance(key, str) and not key.strip()):
                if all(v in (None, "") for v in slice_):
                    break
                continue
            rows.append(slice_)
        if headers is None:
            raise ValueError("Header da PLANTEL não encontrado na linha 4")
        return pd.DataFrame(rows, columns=headers)
    finally:
        wb.close()


def _date_in_name(name: str) -> date | None:
    matches = _DATE_RE.findall(name)
    if not matches:
        return None
    d, m, y = matches[-1]
    yyyy = int(y)
    if yyyy < 100:
        yyyy += 2000
    try:
        return date(yyyy, int(m), int(d))
    except ValueError:
        return None


def _pick_valor_nota(raw: pd.DataFrame) -> tuple[str, str | None]:
    """VALOR: coluna começando com 'VALOR 100%' (mais recente).
    NOTA : coluna contendo 'GIAN' mas não começando com 'VALOR' (mais recente).
    """
    valor_cols = [c for c in raw.columns if str(c).upper().startswith("VALOR 100%")]
    if not valor_cols:
        raise ValueError("Nenhuma coluna 'VALOR 100% ...' no plantel")
    valor_col = max(valor_cols, key=lambda c: (_date_in_name(c) or date.min))

    nota_cands = [
        c for c in raw.columns
        if "GIAN" in str(c).upper() and not str(c).upper().startswith("VALOR")
    ]
    nota_col = None
    if nota_cands:
        nota_col = max(nota_cands, key=lambda c: (_date_in_name(c) or date.min))
    return valor_col, nota_col


def _norm_text(value) -> str | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = " ".join(str(value).strip().split())
    return s if s else None


def _sufixo_grupo(s: str | None) -> str | None:
    if s is None:
        return None
    if s in ("DA PAO GRANDE", "OUTRO"):
        return s
    if s.startswith("DA PAO GRANDE - E "):
        return "PARCERIA"
    return "OUTRO_NAO_PADRAO"


def _norm_nota(value) -> str | None:
    n = _norm_text(value)
    if n is None:
        return None
    n = n.upper()
    return n if n in NOTAS_VALIDAS else None


def _enrich(raw: pd.DataFrame, mes: MesRef) -> pd.DataFrame:
    """Transforma o plantel bruto em um DataFrame-fato pronto para o BI."""
    col_valor, col_nota = _pick_valor_nota(raw)

    sufixo = raw["SUFIXO"].apply(_norm_text)
    categoria = raw["CATEGORIA"].apply(_norm_text)
    status = raw["STATUS PLANTEL"].apply(_norm_text)
    local = raw["LOCAL"].apply(_norm_text)
    sufixo_grp = sufixo.apply(_sufixo_grupo)
    cotas = pd.to_numeric(raw["COTAS (%)"], errors="coerce").fillna(0.0)
    valor_100 = pd.to_numeric(raw[col_valor], errors="coerce")
    nota = raw[col_nota].apply(_norm_nota) if col_nota else pd.Series([None] * len(raw))

    mes_date = pd.Timestamp(year=mes.yyyy, month=mes.mm, day=1)

    fato = pd.DataFrame({
        "mes_referencia": mes_date,
        "qtde": raw["QTDE."],
        "letra": raw["LETRA"],
        "nome": raw["NOME"],
        "sufixo": sufixo,
        "sufixo_grupo": sufixo_grp,
        "categoria": categoria,
        "sexo": raw["SEXO"],
        "status_plantel": status,
        "local": local,
        "cotas": cotas,
        "valor_100": valor_100,
        "patrimonio_proporcional": valor_100.fillna(0.0) * cotas,
        "nota": nota,
        "mae": raw["MAE"],
        "pai": raw["PAI"],
        "nascimento": pd.to_datetime(raw["NASCIMENTO"], errors="coerce"),
        "idade_anos": pd.to_numeric(raw["IDADE (ANO)"], errors="coerce"),
        "safra": raw["SAFRA"],
        "pelagem": raw["PELAGEM"],
        "nome_socio": raw["NOME SOCIO"],
        "in_slide_1": (status == "PLANTEL") & sufixo.isin(SUFIXOS_SLIDE1),
        "in_slide_2": (status == "PLANTEL") & local.isin(LOCAIS_SLIDE2),
    })
    return fato


def extract_mes(mes: MesRef) -> Path:
    """Extrai e enriquece o plantel do mês indicado. Salva parquet no cache."""
    source, rejeitados = find_compatible_source(mes)
    print(f"  [{mes}] fonte: {source.name}")
    for r in rejeitados:
        print(f"    [pulado, layout diferente] {r}")
    cached_src = ensure_cache(source)
    raw = _read_plantel_raw(cached_src)
    fato = _enrich(raw, mes)

    PARQUET_DIR.mkdir(parents=True, exist_ok=True)
    out = PARQUET_DIR / f"{mes.tag}.parquet"
    fato.to_parquet(out, engine="pyarrow", index=False)
    print(f"  [{mes}] {len(fato)} animais -> {out.relative_to(PARQUET_DIR.parent.parent)}")
    return out


def main() -> None:
    mes = prompt_mes_aaaa()
    extract_mes(mes)
    print(
        "\nFeito. Rode `python PGBaseBI.py` para regenerar base_bi.xlsx "
        "incluindo esse mês."
    )


if __name__ == "__main__":
    main()
