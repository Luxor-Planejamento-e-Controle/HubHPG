"""
PGBaseBI — consolida todos os parquets mensais em `base_bi.xlsx` (e .parquet).

A saída é UMA tabela fato (`fato_plantel`) com uma linha por animal por mês.
O Power BI conecta a esta planilha, usa `mes_referencia` num slicer, e calcula
todas as métricas (patrimônio, valor médio, contagens por sufixo/categoria/
nota/local) via DAX usando as flags `in_slide_1` e `in_slide_2`.

Colunas da fato_plantel:
    mes_referencia          : 1º dia do mês (datetime, usar no slicer)
    qtde, letra, nome       : identificação do animal
    sufixo                  : sufixo normalizado (sem espaços extras)
    sufixo_grupo            : DA PAO GRANDE / OUTRO / PARCERIA / OUTRO_NAO_PADRAO
    categoria               : DOADORA / MATRIZ / GARANHAO / POTRA / POTRO / EMBRIAO ...
    sexo                    : F / M
    status_plantel          : PLANTEL / VENDIDO / DOADO / OBITO / DE TERCEIRO / VENDIDO E ENTREGUE
    local                   : FAZENDA PAO GRANDE / SOCIO / ARRENDAMENTO CESAR FURTADO / OUTROS / MATO GROSSO
    cotas                   : % do Pão Grande (0.0 a 1.0)
    valor_100               : avaliação Gianfranco mais recente, R$ (pode ser NaN)
    patrimonio_proporcional : valor_100 × cotas (0 se valor_100 NaN)
    nota                    : A+ / A / A- / B+ / B / B- / C / D (pode ser NaN)
    mae, pai                : genealogia
    nascimento, idade_anos, safra, pelagem, nome_socio
    in_slide_1              : True se entra no universo do slide 1 (PLANTEL e sufixo em {DA PAO GRANDE, OUTRO})
    in_slide_2              : True se entra no universo do slide 2 (PLANTEL e local em 3 valores válidos)

DAX sugerido no Power BI:
    Patrimônio = CALCULATE(SUMX(fato_plantel, [valor_100] * [cotas]), fato_plantel[in_slide_1])
    Valor médio = CALCULATE(AVERAGE(fato_plantel[valor_100]), fato_plantel[in_slide_1])
    Animais slide 2 = CALCULATE(COUNTROWS(fato_plantel), fato_plantel[in_slide_2])

Uso:
    python PGBaseBI.py
"""

from __future__ import annotations

import pandas as pd

from _pg_common import BASE_BI_PARQUET, BASE_BI_XLSX, PARQUET_DIR


def build_base() -> pd.DataFrame:
    if not PARQUET_DIR.exists():
        raise FileNotFoundError(
            f"Cache de parquets não existe: {PARQUET_DIR}. "
            "Rode `python PGDataExtractor.py` ou `python PGRunHistorico.py` antes."
        )
    parquets = sorted(PARQUET_DIR.glob("*.parquet"))
    if not parquets:
        raise FileNotFoundError(f"Nenhum parquet em {PARQUET_DIR}.")

    print(f"Consolidando {len(parquets)} meses:")
    frames = []
    for p in parquets:
        df = pd.read_parquet(p, engine="pyarrow")
        frames.append(df)
        print(f"  {p.stem}: {len(df)} animais")

    fato = pd.concat(frames, ignore_index=True)
    fato = fato.sort_values(["mes_referencia", "qtde"]).reset_index(drop=True)
    return fato


def write_outputs(fato: pd.DataFrame) -> None:
    print(f"\nEscrevendo {BASE_BI_PARQUET.name} ({len(fato)} linhas)...")
    fato.to_parquet(BASE_BI_PARQUET, engine="pyarrow", index=False)

    print(f"Escrevendo {BASE_BI_XLSX.name}...")
    with pd.ExcelWriter(BASE_BI_XLSX, engine="openpyxl") as w:
        fato.to_excel(w, sheet_name="fato_plantel", index=False)


def main() -> None:
    fato = build_base()
    write_outputs(fato)
    meses = fato["mes_referencia"].drop_duplicates().sort_values()
    print(
        f"\nBase pronta: {len(fato)} linhas, "
        f"{len(meses)} meses ({meses.min().strftime('%Y-%m')} -> {meses.max().strftime('%Y-%m')})"
    )


if __name__ == "__main__":
    main()
