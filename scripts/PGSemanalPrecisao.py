"""PGSemanalPrecisao — placar de acerto do cálculo contra os relatórios oficiais.

Roda as MESMAS 19 comparações do fechamento, mas para todas as semanas já
congeladas, e mostra a evolução. Serve para responder "estamos ganhando ou perdendo
precisão?" sem depender da memória de quem rodou.

Marca a confiabilidade do relatório de cada semana, porque comparar contra um docx
que não fecha consigo mesmo mede a digitação, não o cálculo:
  reaproveitado = a linha 'Semana:' aponta outra semana que a do nome do arquivo
  não fecha     = a soma dos locais difere do total geral declarado

Uso: python scripts/PGSemanalPrecisao.py
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import PGSemanalReport as R


def _linhas(snap: dict, dx: dict) -> list:
    """(rótulo, calculado, docx) — espelha o _validacao do orquestrador."""
    hc = snap.get("headcount") or {}
    rc = snap.get("receptoras") or {}
    mv = snap.get("movimento") or {}
    tc = snap.get("terceiros") or {}
    dp, dr, dh, ds = dx["producao"], dx["receptoras"], dx["headcount"], dx["saidas"]
    return [
        ("Acumulado estação", snap.get("acumulado_estacao"), dp["acumulado_estacao"]),
        # transição de estação: linha só existe quando o relatório publica as duas
        ("Acumulado estação (safra nova)", snap.get("acumulado_estacao_proxima"),
         dp.get("acumulado_estacao_proxima")),
        ("Confirmados semana", snap.get("confirmados_semana"), dp["confirmados_semana"]),
        ("Nascimentos", snap.get("nascimentos"), dp["nascimentos"]),
        ("Acumulado no mês", snap.get("acumulado_mes"), dp.get("acumulado_mes")),
        ("Abortos / óbitos", snap.get("abortos_obitos"), dp.get("abortos_obitos")),
        ("Receptoras total", rc.get("total"), dr["total"]),
        ("Receptoras prenhas", rc.get("prenhas"), dr["prenhas"]),
        ("Receptoras vazias", rc.get("vazias"), dr["vazias"]),
        ("Índice eficiência", rc.get("indice_eficiencia"), dr.get("indice")),
        ("Headcount total", hc.get("total"), dh["total"]),
        ("Fazenda Pao Grande", hc.get("fazenda_pg"), dh["fazenda_pg"]),
        ("Arrendamento", hc.get("arrendamento"), dh["arrendamento"]),
        ("Centro de Treinamento", hc.get("cte"), dh.get("cte")),
        ("Sócios", hc.get("socio"), dh["socio"]),
        # ver PGSemanal._validacao: o Δ do relatório é a abertura, aberta em duas linhas
        ("Saídas efetivas (Δ)", hc.get("delta_saidas"), dh.get("delta_saidas")),
        ("Entradas efetivas (Δ)", hc.get("delta_entradas"), dh.get("delta_entradas")),
        ("Saídas semana", mv.get("saidas"),
         ds["saidas_semana"] if ds["saidas_semana"] is not None else dh.get("delta_saidas")),
        # ver PGSemanal._validacao: nascimento não é entrada, então não há atalho
        # pelo Δ do relatório
        ("Entradas semana", mv.get("entradas"), ds.get("entradas")),
        ("Transferências internas", mv.get("transferencias"), ds.get("transferencias")),
        ("Vendidos pendentes", tc.get("vendidos_pendentes"), ds["vendidos_pendentes"]),
        # total (animais + embriões), como o relatório publica. Cheguei a comparar só
        # animais por ter lido "01" como sendo o animal — está escrito "01 (embrião)".
        ("Sociedade pendentes", tc.get("sociedade_pendentes"),
         ds.get("sociedade_pendentes")),
        ("Total terceiros", tc.get("terceiros_propriedade"), dx["terceiros"].get("total")),
        ("Outros terceiros", tc.get("outros_terceiros"), dx["terceiros"].get("outros")),
    ]


def _soc_docx(dx):
    """Animais em sociedade segundo o relatorio: a primeira parte da abertura
    quando ele abre, senao o numero publicado."""
    partes = ((dx.get("aberturas") or {}).get("sociedade_pendentes") or {}).get("partes") or []
    if partes:
        return partes[0]
    return (dx.get("saidas") or {}).get("sociedade_pendentes")

def _eq(a, b) -> bool:
    na = 0 if a in (None, "") else a
    nb = 0 if b in (None, "") else b
    return na == nb


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    hist = R._load_hist()
    docx = R._load_docx_ref()
    semanas = [w for w in sorted(hist) if R._is_iso(w) and w in docx]
    if not semanas:
        sys.exit("Nenhuma semana com snapshot E relatório oficial")

    # Relatório reaproveitado descreve OUTRA semana (a linha 'Semana:' aponta para
    # outra data). Comparar o snapshot desta semana com ele não mede nada — é o
    # número de uma semana contra o texto de outra. Fica fora do placar.
    def _comparavel(w):
        return docx[w].get("ref_confere") is not False

    por_linha = {}
    print(f"{'semana':12} {'placar':>7}  {'relatório':<22} divergências")
    print("-" * 96)
    for w in semanas:
        dx, snap = docx[w], hist[w]
        if not _comparavel(w):
            print(f"{w:12} {'--':>7}  {'reaproveitado':<22} "
                  f"o docx declara '{dx.get('semana_txt')}' — descreve outra semana")
            continue
        linhas = _linhas(snap, dx)
        ok = [l for l in linhas if _eq(l[1], l[2])]
        ruins = [l for l in linhas if not _eq(l[1], l[2])]
        for lab, got, tgt in linhas:
            d = por_linha.setdefault(lab, {"ok": 0, "n": 0})
            d["n"] += 1
            d["ok"] += 1 if _eq(got, tgt) else 0
        selo = []
        if dx.get("ref_confere") is False:
            selo.append("reaproveitado")
        if dx["headcount"].get("coerente") is False:
            selo.append("não fecha")
        print(f"{w:12} {len(ok):>2}/{len(linhas):<4}  {', '.join(selo) or 'coerente':<22} "
              + "; ".join(f"{lab}={got}/{tgt}" for lab, got, tgt in ruins[:4])
              + (f" (+{len(ruins)-4})" if len(ruins) > 4 else ""))

    print("\n" + "=" * 96)
    print("ACERTO POR INDICADOR (todas as semanas comparáveis)")
    print("-" * 96)
    for lab, d in sorted(por_linha.items(), key=lambda x: (x[1]["ok"] / x[1]["n"], x[0])):
        pct = d["ok"] / d["n"]
        barra = "#" * round(pct * 20)
        print(f"  {lab:24} {d['ok']}/{d['n']}  {pct:5.0%}  {barra}")

    # o placar médio só é honesto sobre semanas cujo relatório fecha
    comparaveis = [w for w in semanas if _comparavel(w)]
    limpas = [w for w in comparaveis
              if docx[w]["headcount"].get("coerente") is not False]
    def _media(ws):
        if not ws:
            return None
        tot = sum(len([l for l in _linhas(hist[w], docx[w]) if _eq(l[1], l[2])]) for w in ws)
        return tot / (len(ws) * len(_linhas(hist[ws[0]], docx[ws[0]])))
    print("\n" + "=" * 96)
    m_comp, m_limpas = _media(comparaveis), _media(limpas)
    print(f"  semanas congeladas ................................. {len(semanas)}")
    print(f"  fora do placar (docx reaproveitado) ................ "
          f"{len(semanas) - len(comparaveis)}")
    print(f"  média sobre as {len(comparaveis)} comparáveis ..................... {m_comp:.0%}")
    if m_limpas is not None:
        print(f"  média sobre as {len(limpas)} cujo relatório fecha ........... {m_limpas:.0%}")


if __name__ == "__main__":
    main()
