"""Prepara o monograma do haras para a casca do hub.

O `pg_logo.png` da raiz é um quadrado 3375×3375 com o navy CHAPADO no fundo e a
marca ocupando só o miolo. Colado direto na sidebar isso vira um selo minúsculo
dentro de um retângulo de navy levemente diferente do fundo — foi o que
aconteceu. Aqui o fundo vira transparente e a imagem é cortada na marca, então
ela escala pela altura do texto e assenta no gradiente da sidebar.

O alfa sai da projeção da cor do pixel no eixo fundo→dourado, não de um corte
seco: assim a borda anti-serrilhada da fonte é preservada e a marca não fica
com degrau.

Uso: python tools/build_logo.py     (só precisa rodar se o logo mudar)
Saída: assets/pg-logo.png
"""
from pathlib import Path

import numpy as np
from PIL import Image

HUB = Path(__file__).resolve().parent.parent
SRC = HUB / "pg_logo.png"
OUT = HUB / "assets" / "pg-logo.png"

GOLD = np.array([202, 151, 3], dtype=float)   # --amber do tema
MARGEM = 12                                   # respiro em px, no tamanho original


def run():
    im = Image.open(SRC).convert("RGB")
    px = np.asarray(im, dtype=float)

    # fundo = cor do canto (o PNG é chapado nas bordas)
    bg = px[0, 0].astype(float)
    eixo = GOLD - bg
    t = ((px - bg) @ eixo) / float(eixo @ eixo)      # 0 = fundo, 1 = dourado
    alpha = np.clip(t, 0.0, 1.0)

    rgba = np.zeros(px.shape[:2] + (4,), dtype=np.uint8)
    rgba[..., :3] = GOLD.astype(np.uint8)            # marca em cor sólida
    rgba[..., 3] = (alpha * 255).round().astype(np.uint8)

    ys, xs = np.nonzero(rgba[..., 3] > 8)
    if not len(ys):
        raise SystemExit("Nenhum pixel de marca encontrado — o logo mudou de cor?")
    y0, y1 = max(0, ys.min() - MARGEM), min(rgba.shape[0], ys.max() + 1 + MARGEM)
    x0, x1 = max(0, xs.min() - MARGEM), min(rgba.shape[1], xs.max() + 1 + MARGEM)

    out = Image.fromarray(rgba[y0:y1, x0:x1], "RGBA")
    # 320 px de altura cobre retina em qualquer tamanho que a sidebar use
    escala = 320 / out.height
    out = out.resize((max(1, round(out.width * escala)), 320), Image.LANCZOS)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    out.save(OUT, optimize=True)
    print(f"[logo] {im.size[0]}×{im.size[1]} chapado -> {out.size[0]}×{out.size[1]} "
          f"transparente ({OUT.stat().st_size // 1024} KB)")


if __name__ == "__main__":
    run()
