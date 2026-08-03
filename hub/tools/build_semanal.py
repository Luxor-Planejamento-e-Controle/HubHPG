"""Leva o dashboard semanal (gerado por PGSemanalDashboard.py, neste repo) para a
aba Atualização Semanal do hub.

Diferente do que o LuxorP&CHub faz com os painéis do haras, aqui NÃO há re-skin:
o dashboard já nasce com a paleta do Haras Pao Grande, que é a mesma do hub. O
build só (a) confere que o arquivo continua autocontido e (b) esconde o cabeçalho
interno, que duplicaria o título que a topbar do hub já mostra.

A checagem de referência externa não é frescura: na fase gold o HTML vem do bucket
privado por `srcdoc`, sem base de URL própria, e qualquer CDN/imagem externa
simplesmente não carregaria — melhor falhar aqui do que publicar painel furado.

Saída: hub/assets/semanal/dashboard.html (gitignored — é dado do plantel).
Uso: python hub/tools/build_semanal.py
"""
import re
import sys
from pathlib import Path

HUB = Path(__file__).resolve().parent.parent
REPO = HUB.parent
SRC = REPO / "dashboards" / "dashboard_semanal.html"
OUTDIR = HUB / "assets" / "semanal"
OUT = OUTDIR / "dashboard.html"

EMBED = """
<style>
/* ===== ajustes de embed (o tema é o mesmo do hub, nada de re-skin) ===== */
/* a topbar do hub já mostra "Atualização Semanal" — o header interno duplicaria */
header{display:none !important}
body{padding-top:0 !important}
/* aproveita a largura toda do iframe */
.wrap{max-width:none !important}
/* scrollbars iguais às das outras abas */
::-webkit-scrollbar{height:9px;width:9px}
::-webkit-scrollbar-track{background:transparent}
::-webkit-scrollbar-thumb{background:#1B486B;border-radius:5px}
::-webkit-scrollbar-thumb:hover{background:#2A5F87}
::-webkit-scrollbar-corner{background:transparent}
</style>
"""


def run():
    if not SRC.exists():
        sys.exit(f"Fonte não encontrada: {SRC}\n"
                 f"Rode o fechamento primeiro: python PGSemanal.py")

    h = SRC.read_text(encoding="utf-8", errors="ignore")

    externos = sorted(set(re.findall(r'(?:src|href)=["\'](https?://[^"\']+)', h)))
    if externos:
        sys.exit("O dashboard ganhou referência externa — vendorize antes de publicar:\n  "
                 + "\n  ".join(externos))

    if "</head>" not in h:
        sys.exit("HTML sem </head> — o gerador mudou de forma; conferir PGSemanalDashboard.py.")
    h = h.replace("</head>", EMBED + "</head>", 1)

    OUTDIR.mkdir(parents=True, exist_ok=True)
    OUT.write_text(h, encoding="utf-8")
    print(f"[semanal] {len(h.encode('utf-8')) // 1024} KB -> assets/semanal/dashboard.html")


if __name__ == "__main__":
    run()
