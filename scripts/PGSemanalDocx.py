"""
PGSemanalDocx — parseia os relatórios semanais oficiais (Word) do Haras Pao Grande.

Os docx `Atualização semanal DD-MM-YY.docx` (pasta ATUALIZACAO SEMANAL) são a
VERDADE dos números. A data no nome = semana de referência. A janela da semana
vai do docx anterior até este.

Gera semanal_docx.json: lista de semanas (ordenadas) com os números de cada seção
+ as linhas de detalhe (produção). Usado como:
  - dados históricos reais do dashboard (calendário = datas dos docx)
  - alvo de validação do extractor automático (PGSemanalReport.py)
"""

from __future__ import annotations

import json
import re
from datetime import date
from pathlib import Path

import docx

BASE_DIR = Path(__file__).resolve().parent.parent  # raiz do projeto (scripts/ fica 1 nível abaixo)
DRIVE_DOCX = Path(
    r"G:\.shortcut-targets-by-id\1mBrSeztRwtBnMlkOMnq6aO4LQUkNjiTb"
    r"\PLANILHAS DE CONTROLE\ATUALIZACAO SEMANAL"
)
JSON_OUT = BASE_DIR / "bases" / "semanal_docx.json"

# só os relatórios completos (não os "da movimentação de animais").
# O ano vem com 2 ou 4 dígitos — o de 07-08-2026 veio com 4 e ficava de fora,
# derrubando a validação da semana.
NAME_RE = re.compile(r"^Atualização semanal (\d{2})-(\d{2})-(\d{2}|\d{4})\.docx$", re.IGNORECASE)


def _int(s):
    """Primeiro inteiro de um texto ('02', '01 (09 em transição)' -> 2, 1). None se '--'."""
    if s is None:
        return None
    m = re.search(r"-?\d+", s.replace(".", ""))
    return int(m.group(0)) if m else None


def _num(s):
    """Número decimal PT ('2,7' -> 2.7)."""
    if s is None:
        return None
    m = re.search(r"-?\d+(?:[.,]\d+)?", s)
    return float(m.group(0).replace(",", ".")) if m else None


def _after(line, label):
    """Texto após 'label:' na linha (case-insensitive)."""
    i = line.lower().find(label.lower())
    if i < 0:
        return None
    return line[i + len(label):].lstrip(" :").strip()


def parse_docx(path: Path, ref: date) -> dict:
    d = docx.Document(str(path))
    paras = [p.text.strip() for p in d.paragraphs if p.text.strip()]
    txt = "\n".join(paras)

    def find(label):
        for ln in paras:
            v = _after(ln, label)
            if v is not None:
                return v
        return None

    # detalhe de produção = linhas soltas entre os bullets (nascimento/aborto)
    def det_after(label, stop_prefixes):
        out = []
        grab = False
        for ln in paras:
            low = ln.lower()
            if label.lower() in low:
                grab = True
                continue
            if grab:
                if any(low.startswith(sp) or sp.lower() in low for sp in stop_prefixes):
                    break
                if ln.startswith(("•", "-", "*")) or ln[0].isdigit() and ")" in ln[:3]:
                    break
                out.append(ln)
        return out

    rep = {
        "ref": ref.isoformat(),
        "semana_txt": find("Semana"),
        "producao": {
            "confirmados_semana": _int(find("Embriões confirmados na semana")),
            "acumulado_estacao": _int(find("Acumulado na estação")),
            "acumulado_mes": _int(find("Acumulado no mês")),
            "nascimentos": _int(find("Nascimentos")),
            "abortos_obitos": _int(find("Abortos")),
        },
        "receptoras": {
            "total": _int(find("Total receptoras")),
            "prenhas": _int(find("Prenhas")),
            "vazias": _int(find("Vazias")),
            "indice": _num(find("Índice de eficiência")),
        },
        "headcount": {
            "total": _int(find("Total geral")),
            "fazenda_pg": _int(find("Fazenda Pao Grande") or find("Fazenda Pão Grande")),
            "arrendamento": _int(find("Arrendamento")),
            "cte": _int(find("Centro de Treinamento")),
            "socio": _int(find("Sócios")),
            "delta_txt": find("vs semana passada") or find("Δ"),
        },
        "terceiros": {
            "total": _int(find("Total terceiros")),
            "doadoras": _int(find("Doadoras terceiros")),
            "outros": _int(find("Outros terceiros")),
        },
        "saidas": {
            "saidas_semana": _int(find("Saídas na semana")),
            "vendidos_pendentes": _int(find("Vendidos pendentes de saída")),
            "sociedade_pendentes": _int(find("Em sociedade pendentes de saída")
                                         or find("Em sociedade pendentes")),
            "transferencias": _int(find("Transferências internas")),
            "entradas": _int(find("Entradas na semana")),
        },
        "detalhe": {
            "nascimentos": det_after("Nascimentos", ["Abortos", "Óbitos", "RECEPTORAS", "2)"]),
            "abortos_obitos": det_after("Abortos", ["RECEPTORAS", "2)"]),
            "confirmados": det_after("Embriões confirmados na semana", ["Acumulado"]),
        },
    }
    return rep


def _ref_from_name(name: str) -> date | None:
    m = NAME_RE.match(name)
    if not m:
        return None
    dd, mm, yy = (int(x) for x in m.groups())
    return date(yy if yy >= 100 else 2000 + yy, mm, dd)


def build_all() -> list:
    weeks = []
    for f in DRIVE_DOCX.glob("Atualização semanal *.docx"):
        if f.name.startswith("~$"):
            continue
        ref = _ref_from_name(f.name)
        if ref is None:
            continue  # pula "da movimentação de animais"
        try:
            weeks.append(parse_docx(f, ref))
        except Exception as exc:
            print(f"  ! erro em {f.name}: {exc}")
    weeks.sort(key=lambda w: w["ref"])
    return weeks


def main():
    import sys
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    weeks = build_all()
    JSON_OUT.parent.mkdir(parents=True, exist_ok=True)
    JSON_OUT.write_text(json.dumps(weeks, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"-> {JSON_OUT.name}: {len(weeks)} semanas")
    for w in weeks:
        p, r, h, s = w["producao"], w["receptoras"], w["headcount"], w["saidas"]
        print(f"  {w['ref']}  conf={p['confirmados_semana']} acum={p['acumulado_estacao']} "
              f"nasc={p['nascimentos']} | rec {r['total']}/{r['prenhas']}/{r['vazias']} idx={r['indice']} "
              f"| hc {h['total']} ({h['fazenda_pg']}/{h['arrendamento']}/{h['cte']}/{h['socio']}) "
              f"| saí={s['saidas_semana']} vpend={s['vendidos_pendentes']}")


if __name__ == "__main__":
    main()
