"""Genera los recursos gráficos de Kardexpress.

  logo.png / kardexpress.ico  -> ícono propio: hoja de reporte con el semáforo, en los azules de Hidroobras.

Uso:  python tools/make_assets.py
"""
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "assets"

S = 1024                     # se dibuja grande y se reduce (antialiasing)
BLUE_DARK = (8, 74, 122)     # azules de la marca Hidroobras
BLUE_LIGHT = (38, 172, 226)
NAVY = (12, 44, 74)
AMBER, GREEN, RED = (246, 186, 46), (52, 190, 120), (236, 92, 80)
LINE = (203, 213, 225)


def _gradient(size, c1, c2):
    """Degradado diagonal de c1 (arriba-izquierda) a c2 (abajo-derecha)."""
    w, h = size
    grad = Image.new("RGB", (w, h))
    px = grad.load()
    for y in range(h):
        for x in range(w):
            t = (x + y) / (w + h - 2)
            px[x, y] = tuple(int(a + (b - a) * t) for a, b in zip(c1, c2))
    return grad


def _wave_mask(y0, amp, phase):
    import math
    m = Image.new("L", (S, S), 0)
    pts = [(x, y0 + amp * math.sin((x / S) * 2 * math.pi * 1.1 + phase)) for x in range(0, S + 1, 8)]
    ImageDraw.Draw(m).polygon(pts + [(S, S), (0, S)], fill=255)
    return m


def logo() -> Image.Image:
    small = 256
    grad = _gradient((small, small), BLUE_DARK, BLUE_LIGHT).resize((S, S), Image.BICUBIC)

    # Fondo: mosaico redondeado con degradado azul
    tile_mask = Image.new("L", (S, S), 0)
    ImageDraw.Draw(tile_mask).rounded_rectangle((0, 0, S - 1, S - 1), radius=int(S * 0.22), fill=255)
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    img.paste(grad, (0, 0), tile_mask)

    # Olas al pie (guiño a Hidroobras)
    for y0, amp, phase, alpha in ((S * 0.80, S * 0.035, 0.6, 46), (S * 0.86, S * 0.03, 2.2, 60)):
        wave = Image.new("RGBA", (S, S), (255, 255, 255, alpha))
        mask = Image.composite(_wave_mask(y0, amp, phase), Image.new("L", (S, S), 0), tile_mask)
        img.alpha_composite(Image.composite(wave, Image.new("RGBA", (S, S), (0, 0, 0, 0)), mask))

    # Hoja de reporte con sombra
    x0, y0, x1, y1 = S * 0.22, S * 0.14, S * 0.72, S * 0.80
    fold = S * 0.11
    sheet = [(x0, y0), (x1 - fold, y0), (x1, y0 + fold), (x1, y1), (x0, y1)]
    shadow = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    ImageDraw.Draw(shadow).polygon([(x + S * 0.015, y + S * 0.025) for x, y in sheet], fill=(4, 30, 55, 110))
    img.alpha_composite(shadow.filter(ImageFilter.GaussianBlur(S * 0.02)))
    d = ImageDraw.Draw(img)
    d.polygon(sheet, fill=(255, 255, 255))
    d.polygon([(x1 - fold, y0), (x1 - fold, y0 + fold), (x1, y0 + fold)], fill=(214, 232, 244))

    # Título del reporte
    r = S * 0.022
    d.rounded_rectangle((x0 + S * 0.07, S * 0.24, x0 + S * 0.27, S * 0.24 + S * 0.045), radius=r, fill=NAVY)

    # Filas del semáforo
    for i, (color, length) in enumerate(((AMBER, 0.24), (GREEN, 0.19), (RED, 0.22))):
        cy = S * (0.39 + i * 0.125)
        cx = x0 + S * 0.105
        rad = S * 0.04
        d.ellipse((cx - rad, cy - rad, cx + rad, cy + rad), fill=color)
        d.rounded_rectangle((cx + S * 0.07, cy - S * 0.018, cx + S * (0.07 + length), cy + S * 0.018),
                            radius=S * 0.018, fill=LINE)

    # Sello de verificado
    bx, by, br = S * 0.72, S * 0.74, S * 0.155
    ring = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    ImageDraw.Draw(ring).ellipse((bx - br + S * 0.01, by - br + S * 0.02, bx + br + S * 0.01, by + br + S * 0.02),
                                 fill=(4, 30, 55, 120))
    img.alpha_composite(ring.filter(ImageFilter.GaussianBlur(S * 0.015)))
    d = ImageDraw.Draw(img)
    d.ellipse((bx - br, by - br, bx + br, by + br), fill=(255, 255, 255))
    d.ellipse((bx - br * 0.80, by - br * 0.80, bx + br * 0.80, by + br * 0.80), fill=NAVY)
    w = int(S * 0.034)
    d.line([(bx - br * 0.40, by + br * 0.02), (bx - br * 0.10, by + br * 0.32), (bx + br * 0.44, by - br * 0.30)],
           fill=(255, 255, 255), width=w, joint="curve")
    for px, py in ((bx - br * 0.40, by + br * 0.02), (bx + br * 0.44, by - br * 0.30)):
        d.ellipse((px - w / 2, py - w / 2, px + w / 2, py + w / 2), fill=(255, 255, 255))

    return img.resize((256, 256), Image.LANCZOS)


def main() -> None:
    lg = logo()
    lg.save(OUT / "logo.png")
    lg.save(OUT / "kardexpress.ico", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    print("Recursos generados en", OUT)


if __name__ == "__main__":
    main()
