"""Headcount de fim de mês por local (animais + receptoras) — a série da aba
Plantel > Visualização de dados.

A conta é a do slide de headcount do comitê, sem cópia de regra: o fechamento do
mês (`_fechamento_do_mes`), as receptoras daquele fechamento
(`_receptoras_do_fechamento`) e `headcount_de` com o último dia do mês. É o mesmo
número que vai ao comitê e à semanal.

Saída: bases/plantel_headcount.json, publicado como `plantel.headcount.json` no
bucket privado. O prefixo `plantel` faz a policy de leitura do storage usar
hub_can('plantel'): quem vê a aba, vê a série — sem migration.

Mês já publicado não é recontado (passado fixo, como o comitê): o arquivo do
bucket é baixado antes e só entram os meses que faltam. `--refazer` reconta tudo.

Uso:
    python tools/build_headcount_mensal.py             # conta e grava o json local
    python tools/build_headcount_mensal.py --publicar  # + baixa o publicado antes e sobe
    python tools/build_headcount_mensal.py --refazer --publicar
"""
from __future__ import annotations

import json
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))
sys.path.insert(0, str(REPO / "scripts"))

SAIDA = REPO / "bases" / "plantel_headcount.json"
DESTINO = "plantel.headcount.json"
ORDEM = ("FAZENDA", "ARRENDAMENTO", "CTE", "CENTRAL", "SOCIO")


def _meses(hoje: date):
    """Do janeiro do ano anterior até o mês passado: o mês corrente ainda não fechou."""
    a, m = hoje.year - 1, 1
    while (a, m) < (hoje.year, hoje.month):
        yield a, m
        a, m = (a + 1, 1) if m == 12 else (a, m + 1)


def conta(ano: int, m: int) -> dict | None:
    import build_comite as bc
    from PGSemanalReport import headcount_de

    src = bc._fechamento_do_mes(ano, m)
    if src is None:
        return None
    rec = bc._receptoras_do_fechamento(src)
    fim = date(ano + m // 12, m % 12 + 1, 1) - timedelta(days=1)
    hc, _ = headcount_de(src, rec, fim)
    det = {k: v for k, v in (hc.get("detalhe") or {}).items() if k != "TOTAL GERAL"}
    locais = {k: {"animais": int(det[k]["animais"]), "receptoras": int(det[k]["receptoras"])}
              for k in [*ORDEM, *sorted(set(det) - set(ORDEM))] if k in det}
    return {"locais": locais, "plantel": src.name, "receptoras": rec.name if rec else None}


def main() -> int:
    args = set(sys.argv[1:])
    publicar, refazer = "--publicar" in args, "--refazer" in args

    atual = {"meses": {}}
    if publicar and not refazer:
        import publish_hub
        url, key = publish_hub.env()
        if publish_hub.baixa(url, key, DESTINO, SAIDA):
            atual = json.loads(SAIDA.read_text(encoding="utf-8"))
    elif SAIDA.exists() and not refazer:
        atual = json.loads(SAIDA.read_text(encoding="utf-8"))

    meses = dict(atual.get("meses") or {})
    novos, sem, erros = [], [], {}
    for ano, m in _meses(date.today()):
        chave = f"{ano}-{m:02d}"
        if chave in meses:
            continue
        try:
            r = conta(ano, m)
        except Exception as e:          # um mês ruim não derruba a série
            erros.setdefault(str(e), []).append(chave)
            continue
        if r is None:
            sem.append(chave)
            continue
        meses[chave] = r
        tot = sum(v["animais"] + v["receptoras"] for v in r["locais"].values())
        novos.append(chave)
        print(f"  [headcount] {chave}: {tot} ({r['plantel']})")

    if sem:
        print(f"  [headcount] sem fechamento do haras: {', '.join(sem)}")
    for msg, quais in erros.items():
        print(f"  [headcount] não consegui contar {quais[0]}"
              + (f" a {quais[-1]}" if len(quais) > 1 else "") + f": {msg}")
    if not meses:
        print("  [headcount] nenhum mês contado — nada a gravar")
        return 0

    SAIDA.parent.mkdir(parents=True, exist_ok=True)
    SAIDA.write_text(json.dumps({"atualizado_em": datetime.now().isoformat(timespec="seconds"),
                                 "meses": dict(sorted(meses.items()))},
                                ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"  [headcount] {len(meses)} mês(es) em {SAIDA.name}; novos: {', '.join(novos) or 'nenhum'}")

    if publicar:
        import publish_hub
        url, key = publish_hub.env()
        if not publish_hub.sobe(url, key, SAIDA, DESTINO, "application/json"):
            return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
