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
import re
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))
sys.path.insert(0, str(REPO / "scripts"))

SAIDA = REPO / "bases" / "plantel_headcount.json"
DESTINO = "plantel.headcount.json"
ORDEM = ("FAZENDA", "ARRENDAMENTO", "CTE", "CENTRAL", "SOCIO")


# Primeiro mês da série. Até jan/2024 o fechamento não traz o mês no nome do
# arquivo (`..._PAO_GRANDE_.xlsx`); ver _fechamento_sem_rotulo.
# Na Azure o espelho do Drive só tem as duas últimas estações: os meses antigos
# entram uma vez pelo notebook (--publicar) e ficam fixos no bucket.
INICIO = (2024, 1)


def _meses(hoje: date):
    """De INICIO até o mês passado: o mês corrente ainda não fechou."""
    a, m = INICIO
    while (a, m) < (hoje.year, hoje.month):
        yield a, m
        a, m = (a + 1, 1) if m == 12 else (a, m + 1)


RX_SUFIXO_MES = re.compile(r"_[A-Z]{3}_\d{2}$")


def _receptoras_para(src: Path):
    """Arquivo de receptoras daquele fechamento: o de prefixo mais recente que não
    passa do prefixo do fechamento — a regra do comitê (`_receptoras_do_fechamento`),
    com o nome aceito mais largo. Em 2025 o haras salvou variantes ('... USAR
    ESSA.xlsx', '... RECEPTORAS .xlsx') que o padrão do comitê não pega, e abr/mai/25
    saíam com as receptoras de março. Cópia de trabalho ('EDITAR', sufixo _MES_AA)
    continua fora. Para 2026 o resultado é o mesmo arquivo do slide."""
    import build_comite as bc
    pref = bc.RE_PREFIXO_DATA.match(src.name)
    if not pref:
        return bc._receptoras_do_fechamento(src)
    limite, cands = pref.group(1), []
    for d in bc._estacao_dirs():
        for f in d.glob("*PLANTEL ARRENDAMENTOS E RECEPTORAS*.xlsx"):
            mm = bc.RE_PREFIXO_DATA.match(f.name)
            nome = f.stem.strip().upper()
            if (f.name.startswith("~$") or not mm or mm.group(1) > limite
                    or "EDITAR" in nome or RX_SUFIXO_MES.search(nome)):
                continue
            cands.append((mm.group(1), f.stat().st_mtime, f))
    return max(cands)[2] if cands else None


def _fechamento_sem_rotulo(ano: int, m: int):
    """Até jan/2024 o haras salvava o fechamento sem o mês no nome
    ('240206_CONTROLE_DE_PLANTEL_PAO_GRANDE_.xlsx'), e `_fechamento_do_mes` só acha
    '..._JAN_24'. Sem arquivo rotulado, vale o primeiro sem rótulo gerado no mês
    seguinte: o de 06/02/2024 tem MOVIMENTAÇÕES até 02/02, o de 04/01 fecha dez/23."""
    import build_comite as bc
    a2, m2 = (ano + 1, 1) if m == 12 else (ano, m + 1)
    cands = [f for d in bc._estacao_dirs()
             for f in d.glob(f"{a2 % 100:02d}{m2:02d}??_CONTROLE_DE_PLANTEL_PAO_GRANDE_.xlsx")]
    return min(cands, key=lambda f: f.name) if cands else None


def conta(ano: int, m: int) -> dict | None:
    import build_comite as bc
    from PGSemanalReport import headcount_de

    src = bc._fechamento_do_mes(ano, m) or _fechamento_sem_rotulo(ano, m)
    if src is None:
        return None
    rec = _receptoras_para(src)
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
