"""Generate the MarginIQ application icon (assets/marginiq.ico + marginiq.png).

The icon is a stylised DDR "eye diagram" (overlaid signal traces) with a cyan
margin-measurement arrow across the eye opening, on an Intel-blue rounded
tile. Run with the tool's own interpreter (Pillow is in requirements.txt):

    .venv\\Scripts\\python.exe scripts\\make_icon.py

The generated files are committed, so this only needs re-running when the
design changes.
"""
from __future__ import annotations

import math
import sys
from pathlib import Path

try:
    from PIL import Image, ImageDraw
except ImportError:  # pragma: no cover - Pillow is a pinned dependency
    print("Pillow is required: run setup.bat first.")
    sys.exit(1)

ROOT = Path(__file__).resolve().parent.parent
ASSETS = ROOT / "assets"
S = 1024  # master canvas; downsampled for every icon size (anti-aliasing)

TOP = (0x00, 0x8C, 0xE8)
BOTTOM = (0x00, 0x30, 0x5E)
CYAN = (0x00, 0xC7, 0xFD, 255)


def _gradient_tile() -> Image.Image:
    grad = Image.new("RGBA", (S, S))
    px = grad.load()
    for y in range(S):
        t = y / (S - 1)
        c = tuple(round(TOP[i] + (BOTTOM[i] - TOP[i]) * t) for i in range(3)) + (255,)
        for x in range(S):
            px[x, y] = c
    mask = Image.new("L", (S, S), 0)
    ImageDraw.Draw(mask).rounded_rectangle((24, 24, S - 24, S - 24), radius=210, fill=255)
    tile = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    tile.paste(grad, (0, 0), mask)
    return tile


def _eye_layer() -> Image.Image:
    layer = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    x0, x1, cy = 150, S - 150, S * 0.47
    for amp, alpha, width in ((250, 255, 30), (215, 170, 20), (282, 120, 16)):
        top, bot = [], []
        for i in range(121):
            t = i / 120
            x = x0 + (x1 - x0) * t
            dy = amp * math.sin(math.pi * t)
            top.append((x, cy - dy))
            bot.append((x, cy + dy))
        colour = (255, 255, 255, alpha)
        d.line(top, fill=colour, width=width, joint="curve")
        d.line(bot, fill=colour, width=width, joint="curve")
    # Flat rails outside the eye (the neighbouring unit intervals).
    for y in (cy - 250, cy + 250):
        d.line([(60, y), (175, y)], fill=(255, 255, 255, 200), width=26)
        d.line([(S - 175, y), (S - 60, y)], fill=(255, 255, 255, 200), width=26)
    # Cyan margin arrow across the eye opening.
    ax0, ax1, head = 300, S - 300, 70
    d.line([(ax0 + 20, cy), (ax1 - 20, cy)], fill=CYAN, width=44)
    d.polygon([(ax0 - 10, cy), (ax0 + head, cy - head * 0.8), (ax0 + head, cy + head * 0.8)], fill=CYAN)
    d.polygon([(ax1 + 10, cy), (ax1 - head, cy - head * 0.8), (ax1 - head, cy + head * 0.8)], fill=CYAN)
    # Baseline "measurement" ticks under the eye.
    by = S * 0.86
    d.line([(ax0, by), (ax1, by)], fill=(255, 255, 255, 150), width=18)
    for x in (ax0, (ax0 + ax1) / 2, ax1):
        d.line([(x, by - 34), (x, by + 34)], fill=(255, 255, 255, 190), width=18)
    return layer


def main() -> int:
    ASSETS.mkdir(exist_ok=True)
    master = Image.alpha_composite(_gradient_tile(), _eye_layer())
    sizes = [(n, n) for n in (16, 20, 24, 32, 40, 48, 64, 128, 256)]
    ico = ASSETS / "marginiq.ico"
    png = ASSETS / "marginiq.png"
    try:
        master.resize((256, 256), Image.LANCZOS).save(ico, format="ICO", sizes=sizes)
        master.resize((256, 256), Image.LANCZOS).save(png, format="PNG")
    except PermissionError as exc:
        print(f"Cannot write icon files: {exc}")
        return 1
    print(f"Wrote {ico}\nWrote {png}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
