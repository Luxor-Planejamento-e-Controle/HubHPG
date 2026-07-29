"""
PGSemanalValidar — valida o extractor AUTOMÁTICO contra os relatórios oficiais (docx).

Para cada semana do docx (semanal_docx.json), roda o extractor na janela
(docx anterior+1 .. este docx) e compara os números. Foca nas métricas de EVENTO
(confirmados/nascimentos/abortos/saídas/entradas/transferências), que os logs
datados (ESTAÇÃO + MOVIMENTAÇÕES) conseguem reproduzir p/ semanas passadas.

Snapshot (acumulado/receptoras/headcount) só bate na semana mais recente — arquivos
de estado são sobrescritos, sem histórico. É esperado divergir no passado.

Uso: python PGSemanalValidar.py
"""

from __future__ import annotations

import json
import sys
from datetime import date, timedelta
from pathlib import Path

from PGSemanalReport import build_report

BASE_DIR = Path(__file__).resolve().parent.parent  # raiz do projeto (scripts/ fica 1 nível abaixo)
DOCX_JSON = BASE_DIR / "bases" / "semanal_docx.json"


def _d(iso):
    return date.fromisoformat(iso)


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    weeks = json.loads(DOCX_JSON.read_text(encoding="utf-8"))
    print(f"Validando {len(weeks)} semanas do docx contra o extractor automático\n")
    print(f"{'semana':11} {'conf':>10} {'nasc':>10} {'saídas':>12} {'entradas':>10} {'transf':>10}")
    print("-" * 68)

    def cell(got, tgt):
        if tgt is None:
            tgt = 0
        ok = "OK" if got == tgt else "≠"
        return f"{got}/{tgt}{ok:>3}"

    tot = {"conf": [0, 0], "nasc": [0, 0], "sai": [0, 0]}
    for i, w in enumerate(weeks):
        ref = _d(w["ref"])
        ini = (_d(weeks[i - 1]["ref"]) + timedelta(days=1)) if i > 0 else (ref - timedelta(days=7))
        rep = build_report(ini, ref)
        p, s = rep.producao, rep.saidas
        dp, ds = w["producao"], w["saidas"]
        print(f"{w['ref']:11} "
              f"{cell(p['confirmados_semana'], dp['confirmados_semana']):>10} "
              f"{cell(p['nascimentos'], dp['nascimentos']):>10} "
              f"{cell(s['saidas_semana'], ds['saidas_semana']):>12} "
              f"{cell(s['entradas_semana'], ds['entradas']):>10} "
              f"{cell(s['transferencias_semana'], ds['transferencias']):>10}")
        for key, got, tgt in [("conf", p['confirmados_semana'], dp['confirmados_semana']),
                              ("nasc", p['nascimentos'], dp['nascimentos']),
                              ("sai", s['saidas_semana'], ds['saidas_semana'])]:
            if tgt is not None:
                tot[key][1] += 1
                if got == (tgt or 0):
                    tot[key][0] += 1
    print("-" * 68)
    print(f"match: confirmados {tot['conf'][0]}/{tot['conf'][1]} · "
          f"nascimentos {tot['nasc'][0]}/{tot['nasc'][1]} · saídas {tot['sai'][0]}/{tot['sai'][1]}")


if __name__ == "__main__":
    main()
