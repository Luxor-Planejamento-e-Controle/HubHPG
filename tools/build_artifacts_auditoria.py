"""Recorta a auditoria gerada em duas páginas para publicar como Artifact.

Os dois artifacts de auditoria foram escritos à mão numa sessão anterior e
envelheceram: continuavam mostrando 5/19 e nomes de planilha sem caminho enquanto o
pipeline já publicava 23/23. Página que audita dado não pode ser a única coisa sem
fonte — foi o motivo de `build_auditoria.py` existir, e vale igual aqui.

Então nada é redigitado: este script lê `dashboards/auditoria_semanal.html`, o mesmo
arquivo que vai para o hub, e separa as duas seções em dois arquivos. Estilo,
cabeçalho e rodapé vêm de lá; some só o que o Artifact injeta por conta própria
(`<!doctype>`, `<html>`, `<head>`, `<body>`).

Saída (fora do repo, porque tem nome de animal e de comprador):
    <scratchpad>/artifact_auditoria_semanal.html
    <scratchpad>/artifact_auditoria_comite.html

Uso: python tools/build_artifacts_auditoria.py [destino]
"""
from __future__ import annotations

import os
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FONTE = ROOT / "dashboards" / "auditoria_semanal.html"

# Título de cada artifact. FIXO: o leitor acha o artifact pelo nome na galeria, e
# renomear a cada publicação faz parecer que é outra página.
TITULOS = {
    "semanal": "Auditoria de Fontes · 2026-08-21",
    "comite": "Auditoria de Fontes do Comitê",
}


def _bloco(html: str, tag: str) -> str:
    m = re.search(rf"<{tag}[^>]*>.*?</{tag}>", html, re.S)
    if not m:
        raise SystemExit(f"não achei <{tag}> em {FONTE.name}")
    return m.group(0)


def _secoes(html: str) -> list[str]:
    """As <section> de primeiro nível. Não há section dentro de section, então
    dividir pelo fechamento basta — e é mais robusto que regex aninhada."""
    partes = []
    for pedaco in html.split("<section>")[1:]:
        corpo, _, _ = pedaco.partition("</section>")
        partes.append(f"<section>{corpo}</section>")
    return partes


def _pagina(estilo: str, cabecalho: str, secao: str, rodape: str, titulo: str) -> str:
    # sem <!doctype>/<html>/<head>/<body>: o Artifact embrulha o arquivo
    return (f"<title>{titulo}</title>\n<style>{estilo}</style>\n"
            f'<div class="wrap">\n{cabecalho}\n{secao}\n{rodape}\n</div>\n')


def run(destino: Path | None = None) -> dict:
    if not FONTE.exists():
        raise SystemExit(f"rode tools/build_auditoria.py primeiro — falta {FONTE.name}")
    html = FONTE.read_text(encoding="utf-8")
    estilo = re.search(r"<style>(.*?)</style>", html, re.S).group(1)
    cabecalho = _bloco(html, "header")
    rodape = _bloco(html, "footer")
    secoes = _secoes(html)
    if not secoes:
        raise SystemExit("nenhuma seção encontrada na auditoria gerada")

    destino = destino or Path(
        os.getenv("CLAUDE_SCRATCHPAD")
        or os.getenv("TEMP") or "/tmp")
    destino.mkdir(parents=True, exist_ok=True)

    # Só a seção semanal. A do comitê tem página própria, com narrativa e tabela de
    # fontes mais rica — tools/build_artifact_comite.py. Os dois escreviam o mesmo
    # arquivo e o último a rodar apagava o outro.
    saidas = {}
    for chave, secao in (("semanal", secoes[0]),):
        # o cabeçalho da página fala das duas auditorias; em cada artifact sobra uma
        eyebrow = ("Haras Pao Grande &middot; fechamento semanal" if chave == "semanal"
                   else "Haras Pao Grande &middot; comitê mensal")
        cab = re.sub(r'(<span class="eyebrow">)[^<]*(</span>)',
                     rf"\g<1>{eyebrow}\g<2>", cabecalho)
        alvo = destino / f"artifact_auditoria_{chave}.html"
        alvo.write_text(_pagina(estilo, cab, secao, rodape, TITULOS[chave]),
                        encoding="utf-8")
        print(f"[artifact] {chave}: {alvo.stat().st_size // 1024} KB -> {alvo}")
        saidas[chave] = alvo
    return saidas


if __name__ == "__main__":
    run(Path(sys.argv[1]) if len(sys.argv) > 1 else None)
