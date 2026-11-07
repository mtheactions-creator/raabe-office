"""RaabePresent's textures (Fülleffekt > Struktur, Step 8b, D-078): our own, one for each of PowerPoint's 24 names,
drawn from periodic noise, so each tile repeats without a seam. No picture of anyone else is used. Writes
apps/raabe-present/public/textures/<key>.png (256 x 256).

    tools/feel-lab/.venv/Scripts/python.exe tools/textures/make_textures.py
"""
from __future__ import annotations

from pathlib import Path

import numpy as np
from PIL import Image

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "apps/raabe-present/public/textures"
N = 256
rng = np.random.default_rng(1107)
Y, X = np.mgrid[0:N, 0:N] / N


def noise(scale: float, ax: float = 1.0, ay: float = 1.0) -> np.ndarray:
    """Periodic noise: white noise filtered in the frequency domain (feature size `scale` pixels, stretched by ax/ay)."""
    f = np.fft.fftfreq(N)[:, None] * N, np.fft.fftfreq(N)[None, :] * N
    r = np.sqrt((f[0] * ay) ** 2 + (f[1] * ax) ** 2)
    spec = np.fft.fft2(rng.standard_normal((N, N))) * np.exp(-(r * scale / N) ** 2 * 4)
    out = np.real(np.fft.ifft2(spec))
    return (out - out.mean()) / (out.std() + 1e-9)


def mix(a, b, t):
    t = np.clip(t, 0, 1)[..., None]
    return np.array(a) * (1 - t) + np.array(b) * t


def save(name: str, rgb: np.ndarray) -> None:
    Image.fromarray(np.clip(rgb, 0, 255).astype(np.uint8), "RGB").save(OUT / f"{name}.png", optimize=True)


def marble(base, vein, k=6.0):
    turb = noise(40) * 0.6 + noise(12) * 0.25
    v = np.abs(np.sin((X + Y * 0.6) * 2 * np.pi * 2 + turb * k))
    return mix(vein, base, v ** 0.35) + noise(4)[..., None] * 4


def wood(light, dark, rings=9.0, stretch=8.0):
    d = noise(60, 1, stretch) * 0.12
    grain = np.sin((Y + d + noise(25, 1, stretch) * 0.03) * 2 * np.pi * rings) * 0.5 + 0.5
    fine = noise(3, 1, 10) * 0.08
    return mix(dark, light, grain ** 1.5 + fine)


def paper(base, fiber=6.0, amount=8.0):
    return np.array(base) + (noise(fiber) * amount + noise(1.5) * amount * 0.4)[..., None]


def textures() -> dict[str, np.ndarray]:
    t: dict[str, np.ndarray] = {}
    t["papyrus"] = mix((222, 204, 160), (196, 172, 120), 0.5 + noise(18, 1, 6) * 0.25 + noise(18, 6, 1) * 0.2)
    weave = (np.sin(X * 2 * np.pi * 64) * np.sin(Y * 2 * np.pi * 64)) * 0.5 + 0.5
    t["linen"] = mix((236, 232, 222), (214, 208, 196), weave * 0.6 + noise(2) * 0.1)
    twill = np.sin((X + Y) * 2 * np.pi * 48) * 0.5 + 0.5
    t["denim"] = mix((58, 86, 140), (32, 52, 96), twill * 0.7 + noise(2, 1, 4) * 0.15)
    bx = (np.floor(X * 8) + np.floor(Y * 8)) % 2
    sx = np.sin(Y * 2 * np.pi * 64) * 0.5 + 0.5
    sy = np.sin(X * 2 * np.pi * 64) * 0.5 + 0.5
    t["wovenmat"] = mix((206, 182, 132), (162, 132, 84), np.where(bx > 0, sx, sy) * 0.8 + noise(3) * 0.1)
    drops = np.zeros((N, N))
    for _ in range(60):
        cx, cy, r = rng.uniform(0, 1), rng.uniform(0, 1), rng.uniform(0.01, 0.04)
        dx = (X - cx + 0.5) % 1 - 0.5
        dy = (Y - cy + 0.5) % 1 - 0.5
        d = np.sqrt(dx * dx + dy * dy) / r
        drops = np.maximum(drops, np.where(d < 1, 1 - d ** 2, 0))
    t["waterdrops"] = mix((150, 178, 196), (226, 238, 246), drops * 0.9 + noise(30) * 0.05)
    t["paperbag"] = paper((176, 136, 92), 3, 10)
    stone = mix((214, 200, 170), (188, 170, 136), 0.5 + noise(20) * 0.3)
    ang = np.arctan2(Y - 0.5, X - 0.5)
    rad = np.sqrt((X - 0.5) ** 2 + (Y - 0.5) ** 2)
    spiral = np.exp(-((np.sin(ang * 1 + rad * 60) * 0.5 + 0.5) ** 8) * 0) * np.where(rad < 0.32, np.abs(np.sin(rad * 60 - ang)) ** 6, 0)
    t["fossil"] = stone - spiral[..., None] * 40
    t["sand"] = mix((222, 204, 168), (196, 174, 132), 0.5 + noise(1.2) * 0.35 + noise(20) * 0.15)
    t["greenmarble"] = marble((40, 96, 72), (180, 214, 190))
    t["whitemarble"] = marble((238, 236, 232), (150, 150, 156), 8)
    t["brownmarble"] = marble((110, 72, 46), (210, 176, 140))
    t["granite"] = mix((150, 146, 142), (70, 68, 66), (noise(1.5) > 0.9) * 0.9 + 0.3 + noise(8) * 0.15) + ((noise(1.2) > 1.8) * 60)[..., None]
    lines = ((np.floor(Y * 64) % 2 == 0) & (np.sin(X * 2 * np.pi * 9 + np.floor(Y * 64)) > -0.2)) * 1.0
    t["newsprint"] = mix((222, 220, 214), (150, 148, 146), lines * 0.45 + noise(2) * 0.08)
    t["recycledpaper"] = paper((216, 212, 204), 2, 6) - ((noise(1.5) > 2.2) * 50)[..., None]
    t["parchment"] = mix((240, 226, 186), (214, 190, 140), 0.4 + noise(40) * 0.35 + noise(6) * 0.1)
    t["stationery"] = paper((246, 246, 240), 4, 4)
    crinkle = np.abs(noise(10, 1, 3)) * 0.6 + np.abs(noise(10, 3, 1)) * 0.4
    t["bluetissuepaper"] = mix((206, 222, 240), (150, 180, 222), crinkle * 0.7)
    t["pinktissuepaper"] = mix((246, 214, 224), (226, 160, 184), crinkle * 0.7)
    net = np.minimum(np.abs(np.sin((X + Y) * np.pi * 16)), np.abs(np.sin((X - Y) * np.pi * 16)))
    t["purplemesh"] = mix((96, 52, 128), (206, 180, 226), (net > 0.2) * 1.0 + noise(3) * 0.05)
    bg = np.full((N, N, 3), (238, 232, 214), float)
    for _ in range(18):
        cx, cy, r = rng.uniform(0, 1), rng.uniform(0, 1), rng.uniform(0.03, 0.06)
        col = [(206, 90, 110), (230, 170, 70), (150, 110, 190), (220, 120, 150)][rng.integers(4)]
        dx = (X - cx + 0.5) % 1 - 0.5
        dy = (Y - cy + 0.5) % 1 - 0.5
        a = np.arctan2(dy, dx)
        petal = np.sqrt(dx * dx + dy * dy) < r * (0.65 + 0.35 * np.cos(5 * a))
        bg[petal] = col
        bg[np.sqrt(dx * dx + dy * dy) < r * 0.25] = (240, 210, 90)
    t["bouquet"] = bg
    t["cork"] = mix((196, 146, 92), (128, 86, 50), 0.4 + noise(2) * 0.3 + (noise(1.5) > 1.5) * 0.4)
    t["walnut"] = wood((128, 86, 54), (72, 44, 26))
    t["oak"] = wood((214, 176, 120), (168, 124, 76), 12)
    t["mediumwood"] = wood((176, 124, 76), (122, 78, 44), 7)
    return t


# PowerPoint's names (Fülleffekt > Struktur) for each file.
NAMES = {"papyrus": "Papyrus", "linen": "Leinen", "denim": "Blaue Baumwolle", "wovenmat": "Gewebe", "waterdrops": "Wassertropfen", "paperbag": "Papiertüte",
         "fossil": "Fischfossil", "sand": "Sand", "greenmarble": "Grüner Marmor", "whitemarble": "Weißer Marmor", "brownmarble": "Brauner Marmor", "granite": "Granit",
         "newsprint": "Zeitungspapier", "recycledpaper": "Recyclingpapier", "parchment": "Pergament", "stationery": "Briefpapier", "bluetissuepaper": "Blaues Seidenpapier",
         "pinktissuepaper": "Rosa Seidenpapier", "purplemesh": "Lila Netz", "bouquet": "Blumenstrauß", "cork": "Kork", "walnut": "Walnuss", "oak": "Eiche", "mediumwood": "Nussbaum"}


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    for k, img in textures().items():
        save(k, img)
    print(len(NAMES), "textures")


if __name__ == "__main__":
    main()
