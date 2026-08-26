"""Gera os dois PDFs de embriões que vão junto com a atualização semanal.

    EMBRIÕES PG <DD-MM-AAAA>.pdf                aba EMBRIÕES PAO GRANDE
    EMBRIÕES SÓCIOS-VENDIDOS <DD-MM-AAAA>.pdf   aba EMBRIOES SOCIOS - VENDIDOS

Hoje são exportados à mão da planilha `EMBRIÕES E MATRIZES`, toda semana, e é ela
que o haras confirmou ser a fonte. O PDF é retrato da aba: mesma ordem de colunas,
mesmas linhas, sem filtro nenhum — se a planilha tem, o PDF tem. Qualquer regra
aqui viraria divergência entre o que o sócio recebe e o que a planilha diz.

O que o script faz além de copiar: formata data como DD/MM/AAAA e cota como
percentual, que é como o PDF feito à mão sai, e mantém a linha em branco quando a
célula é vazia (várias linhas não têm comprador, e inventar '—' mudaria o
documento).

Uso:
    python tools/build_pdf_embrioes.py                  # data de hoje, saída no _cache
    python tools/build_pdf_embrioes.py 21/08/2026 ~/Downloads
"""
from __future__ import annotations

import sys
from datetime import date, datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))

from PGSemanalReport import _latest_emb_matrizes, _load, _s   # noqa: E402

LOGO = ROOT / "assets" / "pg-logo.png"
SAIDA_PADRAO = ROOT / "_cache" / "pdf"

# (aba, título impresso, nome do arquivo, primeira linha de dado, colunas)
# As colunas são as da planilha, na ordem em que o PDF feito à mão as mostra.
PAGINAS = [
    {
        "aba": "EMBRIÕES PAO GRANDE",
        "titulo": "EMBRIÕES PAO GRANDE - A NASCER",
        "arquivo": "EMBRIÕES PG {data}.pdf",
        "linha1": 4,
        "colunas": [("DOADORA", 1, "txt"), ("GARANHÃO", 2, "txt"), ("DATA", 3, "data"),
                    ("PREV PARTO", 4, "data"), ("RECEPTORA", 5, "txt"),
                    ("LOCAL", 6, "txt"), ("ESTACAO", 7, "txt"),
                    ("COTA PG", 8, "pct"), ("COMPRADOR/SOCIO", 9, "txt")],
        # mm; somam os 277 úteis do A4 paisagem. Sem isto o reportlab divide igual,
        # DOADORA e GARANHÃO quebram em duas linhas e a tabela vira duas páginas —
        # o PDF do haras cabe numa.
        "larguras": [38, 40, 17, 18, 20, 27, 18, 16, 83],
    },
    {
        "aba": "EMBRIOES SOCIOS - VENDIDOS",
        "titulo": "EMBRIÕES - ESTAÇÃO 2024/2025 e 2025/2026",
        "arquivo": "EMBRIÕES SÓCIOS-VENDIDOS {data}.pdf",
        "linha1": 4,
        "colunas": [("DOADORA", 1, "txt"), ("GARANHÃO", 2, "txt"), ("DATA", 3, "data"),
                    ("RECEPTORA", 4, "txt"), ("STATUS", 5, "txt"), ("LOCAL", 6, "txt"),
                    ("ESTAÇÃO", 7, "txt"), ("COMPRADOR/SOCIO", 8, "txt")],
        "larguras": [42, 42, 20, 26, 22, 28, 22, 75],
    },
]


def _fmt(valor, tipo: str) -> str:
    if valor is None or (isinstance(valor, str) and not valor.strip()):
        return ""
    if tipo == "data":
        if isinstance(valor, (datetime, date)):
            return valor.strftime("%d/%m/%Y")
        return _s(valor)
    if tipo == "pct":
        try:
            n = float(valor)
        except (TypeError, ValueError):
            return _s(valor)
        # 0.5 -> 50%, 1 -> 100%. Sai inteiro quando dá, que é como o PDF à mão mostra.
        p = n * 100
        return f"{p:.0f}%" if abs(p - round(p)) < 1e-6 else f"{p:.2f}%".replace(".", ",")
    return _s(valor)


def _linhas(ws, pagina) -> list[list[str]]:
    out = []
    for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if i < pagina["linha1"]:
            continue
        # a primeira coluna de dado vazia encerra a tabela: abaixo dela a planilha
        # tem só células de formatação
        if r[pagina["colunas"][0][1]] is None:
            continue
        out.append([_fmt(r[idx] if idx < len(r) else None, tipo)
                    for _, idx, tipo in pagina["colunas"]])
    return out


def _pdf(caminho: Path, titulo: str, cabecalho: list[str], linhas: list[list[str]],
         larguras: list[int] | None = None):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Table, TableStyle

    doc = SimpleDocTemplate(str(caminho), pagesize=landscape(A4),
                            leftMargin=10 * mm, rightMargin=10 * mm,
                            topMargin=10 * mm, bottomMargin=10 * mm,
                            title=titulo, author="Haras Pao Grande")
    # 6.4/7.6 com padding 2: é o que faz as ~32 linhas caberem numa página só,
    # como no PDF que o haras manda. Maior que isso empurra duas linhas pra página 2.
    cel = ParagraphStyle("cel", fontName="Helvetica", fontSize=6.4, leading=7.6)
    cab = ParagraphStyle("cab", fontName="Helvetica-Bold", fontSize=6.8, leading=8.4,
                         alignment=1)
    tit = ParagraphStyle("tit", fontName="Helvetica-Bold", fontSize=9.5, leading=12,
                         alignment=1)

    # cabeçalho do documento: logo à esquerda, título ocupando o resto — é o desenho
    # do PDF que o haras já manda
    # A célula do logo tem a MESMA largura da primeira coluna da tabela, e o cabeçalho
    # a mesma largura total. Sem isso as duas molduras não fecham no mesmo prumo e o
    # documento fica com a linha vertical do logo cortando a coluna DOADORA.
    logo = Image(str(LOGO), width=11 * mm, height=11 * mm) if LOGO.exists() else ""
    col_logo = (larguras[0] if larguras else 40) * mm
    total = sum(larguras) * mm if larguras else None
    topo = Table([[logo, Paragraph(titulo, tit)]],
                 colWidths=[col_logo, (total - col_logo) if total else None],
                 rowHeights=[14 * mm])
    topo.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.6, colors.HexColor("#1B2B44")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (0, 0), "CENTER"),
        ("SPAN", (1, 0), (-1, 0)),
    ]))

    dados = [[Paragraph(h, cab) for h in cabecalho]]
    dados += [[Paragraph(c, cel) for c in linha] for linha in linhas]
    tabela = Table(dados, repeatRows=1,
                   colWidths=[w * mm for w in larguras] if larguras else None)
    tabela.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.6, colors.HexColor("#1B2B44")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        # doadora e garanhão à esquerda; o resto centralizado, como no PDF do haras
        ("ALIGN", (2, 0), (-1, -1), "CENTER"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    doc.build([topo, tabela])


def run(quando: date | None = None, destino: Path | None = None) -> list[Path]:
    quando = quando or date.today()
    destino = Path(destino or SAIDA_PADRAO)
    destino.mkdir(parents=True, exist_ok=True)
    src = _latest_emb_matrizes()
    wb = _load(src)
    feitos = []
    for pagina in PAGINAS:
        if pagina["aba"] not in wb.sheetnames:
            print(f"[pdf] aba {pagina['aba']!r} não existe em {src.name} — pulada")
            continue
        linhas = _linhas(wb[pagina["aba"]], pagina)
        alvo = destino / pagina["arquivo"].format(data=quando.strftime("%d-%m-%Y"))
        _pdf(alvo, pagina["titulo"], [c[0] for c in pagina["colunas"]], linhas,
             pagina.get("larguras"))
        print(f"[pdf] {len(linhas):>3} linha(s) · {alvo.name}")
        feitos.append(alvo)
    wb.close()
    print(f"[pdf] fonte: {src.name}")
    return feitos


def main():
    args = sys.argv[1:]
    quando = None
    destino = None
    for a in args:
        if "/" in a and len(a) == 10:
            quando = datetime.strptime(a, "%d/%m/%Y").date()
        else:
            destino = Path(a).expanduser()
    run(quando, destino)


if __name__ == "__main__":
    main()
