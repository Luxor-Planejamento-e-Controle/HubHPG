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


def _int_valor(s):
    """Inteiro do VALOR do bullet, não do rótulo.

    O rótulo passou a carregar a safra — 'Acumulado na estação 25/26: 61' — e aí o
    primeiro inteiro do resto da linha é 25 (a safra), não 61. O valor vem depois
    do último ':'. Foi o que fez o alvo de validação virar 25 em 07/08/2026.
    """
    if s is None:
        return None
    return _int(s.rsplit(":", 1)[-1] if ":" in s else s)


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

    # A semana declarada tem de sair da linha 'Semana: DD/MM a DD/MM'. find("Semana")
    # casava antes com o título '*ATUALIZAÇÃO SEMANAL*' (contém "semana"), e devolvia
    # "L*" em todos os relatórios — o campo era lixo e ninguém conferia a data.
    semana_ln = next((ln for ln in paras if ln.lower().startswith("semana:")), None)

    rep = {
        "ref": ref.isoformat(),
        "semana_txt": semana_ln,
        "producao": {
            "confirmados_semana": _int(find("Embriões confirmados na semana")),
            "acumulado_estacao": _int_valor(find("Acumulado na estação")),
            "acumulado_mes": _int_valor(find("Acumulado no mês")),
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
    # A data de referência vem do NOME do arquivo, mas o haras reaproveita o .docx da
    # semana passada como rascunho da nova: em 4 dos 25 relatórios a linha 'Semana:'
    # aponta outra semana que a do nome. O caso de 31/07/2026 é o pior — foi editado
    # em 07/08 e declara 'Semana: 03/08 a 07/08', com receptoras já da semana nova e
    # locais ainda da antiga. Validar contra um arquivo assim persegue fantasma.
    rep["semana_fim_declarada"] = _fim_declarado(semana_ln, ref.year)
    rep["ref_confere"] = (rep["semana_fim_declarada"] == ref.isoformat()
                          if rep["semana_fim_declarada"] else None)

    h = rep["headcount"]
    h["delta_net"] = _delta_net(h["delta_txt"])
    # O relatório é digitado à mão e já saiu com erro de digitação: em 31/07/2026 o
    # arrendamento foi escrito como 31 (era 41) e o total como 204, e a conta não
    # fecha. Quem valida contra ele precisa saber disso, senão persegue divergência
    # que não é do cálculo — foi o que fez o Δ da semana seguinte parecer errado.
    locais = [h["fazenda_pg"], h["arrendamento"], h["cte"], h["socio"]]
    h["soma_locais"] = sum(x for x in locais if x) if any(locais) else None
    h["coerente"] = (h["soma_locais"] == h["total"]) if (h["soma_locais"] and h["total"]) else None
    return rep


def _fim_declarado(txt, ano):
    """Fim da semana declarada em 'Semana: 03/08 a 07/08' -> '2026-08-07'. O ano não
    aparece na linha; vem da referência do nome do arquivo."""
    if not txt:
        return None
    pares = re.findall(r"(\d{1,2})/(\d{1,2})", txt)
    if not pares:
        return None
    dd, mm = (int(x) for x in pares[-1])
    try:
        return date(ano, mm, dd).isoformat()
    except ValueError:
        return None


def _delta_net(txt):
    """Δ líquido do texto '+01 / -07)' -> -6. None se não der pra ler."""
    if not txt:
        return None
    nums = re.findall(r"([+-])\s*(\d+)", txt)
    if not nums:
        return None
    return sum(int(v) if sinal == "+" else -int(v) for sinal, v in nums)


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
