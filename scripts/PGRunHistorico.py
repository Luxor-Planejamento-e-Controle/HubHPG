"""
PGRunHistorico — bootstrap do histórico: roda PGDataExtractor para todos os meses
disponíveis no Drive, depois consolida tudo em `base_bi.xlsx` via PGBaseBI.

Uso:
    python PGRunHistorico.py            # extrai meses ainda sem parquet + consolida
    python PGRunHistorico.py --force    # re-extrai todos os meses do zero
"""

from __future__ import annotations

import argparse
import time

from _pg_common import PARQUET_DIR, list_available_meses
from PGBaseBI import build_base, write_outputs
from PGDataExtractor import extract_mes


def run(force: bool = False) -> None:
    meses = list_available_meses()
    print(f"{len(meses)} meses com planilha padronizada no Drive.\n")

    sucesso: list[str] = []
    falha: list[tuple[str, str]] = []

    for mes in meses:
        cached_parquet = PARQUET_DIR / f"{mes.tag}.parquet"
        if cached_parquet.exists() and not force:
            print(f"  [{mes}] parquet já existe (skip)")
            sucesso.append(str(mes))
            continue
        try:
            t0 = time.time()
            extract_mes(mes)
            print(f"    {time.time()-t0:.1f}s")
            sucesso.append(str(mes))
        except Exception as exc:
            motivo = str(exc).split("\n")[0]
            falha.append((str(mes), motivo))
            print(f"  [{mes}] FALHA: {motivo}")

    print()
    print("=" * 60)
    print(f"Extração: {len(sucesso)}/{len(meses)} OK")
    if falha:
        print(f"Falhas ({len(falha)}):")
        for m, mot in falha:
            print(f"  - {m}: {mot}")

    print("\nConsolidando base_bi...")
    fato = build_base()
    write_outputs(fato)
    meses_ok = fato["mes_referencia"].drop_duplicates().sort_values()
    print(
        f"\nBase pronta: {len(fato)} linhas, "
        f"{len(meses_ok)} meses ({meses_ok.min().strftime('%Y-%m')} -> "
        f"{meses_ok.max().strftime('%Y-%m')})"
    )


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--force", action="store_true",
                        help="ignora parquets em cache e re-extrai do zero")
    args = parser.parse_args()
    run(force=args.force)


if __name__ == "__main__":
    main()
