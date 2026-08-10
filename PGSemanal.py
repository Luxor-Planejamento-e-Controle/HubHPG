"""
PGSemanal — ORQUESTRADOR do fechamento semanal do Haras Pao Grande.

Roda o pipeline inteiro em sequência, num comando só:
  1. parseia os relatórios oficiais (docx)  -> alvo de validação
  2. calcula a semana a partir das planilhas -> semanal_data.json + snapshot congelado
  3. gera o dashboard HTML
  4. abre o dashboard no navegador
  5. imprime o placar de validação (calculado vs docx daquela semana)

Uso:
    python PGSemanal.py                 # semana de referência = hoje (prompt confirma)
    python PGSemanal.py 17/07/2026      # semana de referência = essa data
    python PGSemanal.py --no-open       # não abre o navegador
    python PGSemanal.py --no-docx       # pula o reparse dos docx (mais rápido)

Cada etapa é isolada; se uma falhar, o orquestrador avisa e para.
"""

from __future__ import annotations

import os
import sys
import webbrowser
from datetime import date, timedelta
from pathlib import Path

# os módulos do pipeline vivem em scripts/ (este orquestrador fica na raiz)
sys.path.insert(0, str(Path(__file__).resolve().parent / "scripts"))

import PGSemanalDocx
import PGSemanalReport as R
import PGSemanalDashboard as D


def _log(step, msg):
    print(f"[{step}] {msg}")


def _parse_ref(args) -> date:
    for a in args:
        if a.startswith("--"):
            continue
        return R._parse_d(a)
    return date.today()


def _janela(ref: date) -> tuple[date, date]:
    """Janela da semana: do último snapshot capturado (+1 dia) até ref; senão ref-7."""
    hist = R._load_hist()
    prev = None
    for wid in sorted(hist):
        if R._is_iso(wid) and wid < ref.isoformat():
            prev = date.fromisoformat(wid)
    ini = (prev + timedelta(days=1)) if prev else (ref - timedelta(days=7))
    return ini, ref


def _validacao(rep):
    """Placar calculado vs docx da semana de referência (se houver relatório)."""
    dx = rep.docx_ref.get(rep.semana_atual)
    if not dx:
        print("    (sem relatório oficial nessa data — nada p/ validar)")
        return
    p, r, h, s = rep.producao, rep.receptoras, rep.headcount, rep.saidas
    dp, dr, dh, ds = dx["producao"], dx["receptoras"], dx["headcount"], dx["saidas"]
    linhas = [
        ("Acumulado estação", p["acumulado_estacao"], dp["acumulado_estacao"]),
        ("Confirmados semana", p["confirmados_semana"], dp["confirmados_semana"]),
        ("Nascimentos", p["nascimentos"], dp["nascimentos"]),
        ("Receptoras total", r["total"], dr["total"]),
        ("Receptoras prenhas", r["prenhas"], dr["prenhas"]),
        ("Receptoras vazias", r["vazias"], dr["vazias"]),
        ("Índice eficiência", r.get("indice_eficiencia"), dr.get("indice")),
        ("Headcount total", h["total"], dh["total"]),
        ("Fazenda Pao Grande", h["fazenda_pg"], dh["fazenda_pg"]),
        ("Arrendamento", h["arrendamento"], dh["arrendamento"]),
        ("Centro de Treinamento", h.get("cte"), dh.get("cte")),
        ("Sócios", h["socio"], dh["socio"]),
        ("Δ headcount", h.get("delta"), dh.get("delta_net")),
        ("Saídas semana", s["saidas_semana"], ds["saidas_semana"]),
        ("Entradas semana", s.get("entradas_semana"), ds.get("entradas")),
        ("Transferências internas", s.get("transferencias_semana"), ds.get("transferencias")),
        ("Vendidos pendentes", rep.terceiros.get("vendidos_pendentes"), ds["vendidos_pendentes"]),
        ("Sociedade pendentes", rep.terceiros.get("sociedade_pendentes"),
         ds.get("sociedade_pendentes")),
        ("Total terceiros", rep.terceiros.get("terceiros_propriedade"),
         dx["terceiros"].get("total")),
    ]
    def _eq(a, b):
        # docx "-" vira None = zero/nada; trata None e 0 como iguais
        na = 0 if a in (None, "") else a
        nb = 0 if b in (None, "") else b
        return na == nb
    ok = 0
    for lab, got, tgt in linhas:
        flag = "OK " if _eq(got, tgt) else "≠  "
        if _eq(got, tgt):
            ok += 1
        print(f"    {flag} {lab:23} calc={got}  docx={tgt}")
    print(f"    -> {ok}/{len(linhas)} batem")
    _conferir_docx(rep, dx)


def _conferir_docx(rep, dx):
    """O relatório oficial é digitado à mão. Antes de culpar o cálculo, checa se o
    próprio docx fecha — e se o docx da semana anterior fecha, porque o Δ desta
    semana depende dele."""
    def _avisos(wid, dxw, papel):
        h = dxw["headcount"]
        if dxw.get("ref_confere") is False:
            print(f"    !  {papel} ({wid}) declara '{dxw.get('semana_txt')}' — o arquivo "
                  f"foi reaproveitado como rascunho de outra semana, então os números "
                  f"dele estão misturados")
        if h.get("coerente") is False:
            print(f"    !  {papel} ({wid}) não fecha: locais somam {h['soma_locais']} "
                  f"e o total declarado é {h['total']}")

    _avisos(rep.semana_atual, dx, "relatório desta semana")
    ant = None
    for wid in sorted(rep.docx_ref):
        if wid < rep.semana_atual:
            ant = wid
    if not ant:
        return
    _avisos(ant, rep.docx_ref[ant], "relatório da semana anterior")
    # O Δ é medido contra o nosso snapshot da semana anterior. Se ele discordar do
    # total oficial daquela semana, o Δ desta semana nasce errado — e foi assim que
    # 31/07 passou batido: a aba CONTAGEM não tinha sido atualizada, o snapshot
    # repetiu o total de 24/07 e o Δ daquela semana saiu 0 em vez de -1.
    nosso = (rep.snapshots.get(ant) or {}).get("headcount", {}).get("total")
    oficial = rep.docx_ref[ant]["headcount"].get("total")
    if nosso and oficial and nosso != oficial:
        print(f"    !  base do Δ divergente: nosso snapshot de {ant} tem headcount "
              f"{nosso} e o relatório daquela semana diz {oficial} — corrigir o "
              f"snapshot de {ant} antes de confiar no Δ desta semana")


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    args = sys.argv[1:]
    do_open = "--no-open" not in args
    do_docx = "--no-docx" not in args
    ref = _parse_ref(args)
    ini, fim = _janela(ref)

    print("=" * 66)
    print(f"FECHAMENTO SEMANAL HPG — referência {fim.strftime('%d/%m/%Y')} "
          f"(janela {ini.strftime('%d/%m')}–{fim.strftime('%d/%m')})")
    print("=" * 66)

    # 1) relatórios oficiais (validação)
    if do_docx:
        _log("1/4 DOCX", "parseando relatórios oficiais...")
        try:
            import json
            weeks = PGSemanalDocx.build_all()
            PGSemanalDocx.JSON_OUT.write_text(
                json.dumps(weeks, ensure_ascii=False, indent=2), encoding="utf-8")
            _log("1/4 DOCX", f"{len(weeks)} relatórios -> {PGSemanalDocx.JSON_OUT.name}")
        except Exception as exc:
            _log("1/4 DOCX", f"FALHOU: {exc!r} (segue sem validação)")
    else:
        _log("1/4 DOCX", "pulado (--no-docx)")

    # 2) calcula a semana
    _log("2/4 CALC", "lendo planilhas e calculando a semana...")
    rep = R.build_report(ini, fim)
    import json
    from dataclasses import asdict
    R.JSON_OUT.parent.mkdir(parents=True, exist_ok=True)
    R.JSON_OUT.write_text(json.dumps(asdict(rep), ensure_ascii=False, indent=2), encoding="utf-8")
    _log("2/4 CALC", f"{R.JSON_OUT.name} gravado; snapshot da semana {rep.semana_atual} congelado")

    # 3) dashboard
    _log("3/4 DASH", "gerando HTML...")
    D.build()

    # 4) abre
    if do_open:
        _log("4/4 OPEN", f"abrindo {D.HTML_OUT.name}...")
        try:
            os.startfile(str(D.HTML_OUT))          # Windows
        except AttributeError:
            webbrowser.open(D.HTML_OUT.as_uri())
    else:
        _log("4/4 OPEN", "pulado (--no-open)")

    print("\nVALIDAÇÃO (calculado vs relatório oficial):")
    _validacao(rep)
    print("\nOK. Fechamento gerado.")


if __name__ == "__main__":
    main()
