"""RaabePresent's transition sounds (Step 8b, D-078): our own, synthesised from noise and tones, one for each of
PowerPoint's names (Übergänge > Sound). No recording of anyone else is used, so the files are ours under the
project's licence. Writes apps/raabe-present/public/sounds/<name>.wav (mono, 22050 Hz, 16 bit).

    tools/feel-lab/.venv/Scripts/python.exe tools/sounds/make_sounds.py
"""
from __future__ import annotations

import wave
from pathlib import Path

import numpy as np

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "apps/raabe-present/public/sounds"
SR = 22050
rng = np.random.default_rng(20261107)


def t(sec: float) -> np.ndarray:
    return np.arange(int(sec * SR)) / SR


def env(n: int, attack: float, decay: float) -> np.ndarray:
    """Attack then exponential decay (seconds)."""
    x = np.arange(n) / SR
    a = np.clip(x / max(attack, 1e-4), 0, 1)
    return a * np.exp(-np.maximum(0, x - attack) / max(decay, 1e-4))


def noise(sec: float) -> np.ndarray:
    return rng.standard_normal(int(sec * SR))


def lowpass(x: np.ndarray, cutoff: float) -> np.ndarray:
    """A second-order low-pass response applied in the frequency domain."""
    f = np.fft.rfftfreq(len(x), 1 / SR)
    return np.fft.irfft(np.fft.rfft(x) / (1 + (f / cutoff) ** 2), len(x))


def highpass(x: np.ndarray, cutoff: float) -> np.ndarray:
    return x - lowpass(x, cutoff)


def band(x: np.ndarray, lo: float, hi: float) -> np.ndarray:
    return highpass(lowpass(x, hi), lo)


def place(total: np.ndarray, part: np.ndarray, at: float, gain: float = 1.0) -> None:
    i = int(at * SR)
    j = min(len(total), i + len(part))
    if i < len(total):
        total[i:j] += part[: j - i] * gain


def clap() -> np.ndarray:
    n = band(noise(0.08), 800, 4000)
    return n * env(len(n), 0.001, 0.015)


def sweep(f0: float, f1: float, sec: float) -> np.ndarray:
    x = t(sec)
    f = f0 * (f1 / f0) ** (x / sec)
    return np.sin(2 * np.pi * np.cumsum(f) / SR)


def tone(f: float, sec: float, partials=((1, 1.0),)) -> np.ndarray:
    x = t(sec)
    return sum(g * np.sin(2 * np.pi * f * k * x) for k, g in partials)


def applause() -> np.ndarray:
    out = np.zeros(int(2.4 * SR))
    for _ in range(420):
        at = rng.uniform(0, 2.2)
        place(out, clap(), at, rng.uniform(0.2, 0.8) * min(1, at * 3) * min(1, (2.3 - at) * 2))
    return out


def bomb() -> np.ndarray:
    out = np.zeros(int(2.2 * SR))
    whistle = sweep(1800, 500, 1.0) * np.linspace(0.2, 0.5, int(1.0 * SR))
    place(out, whistle, 0)
    boom = lowpass(noise(1.2), 300) * env(int(1.2 * SR), 0.005, 0.35) * 6
    place(out, boom, 1.0)
    return out


def breeze() -> np.ndarray:
    n = band(noise(2.0), 200, 1200)
    return n * np.sin(np.pi * np.linspace(0, 1, len(n))) ** 2 * 0.8


def explosion() -> np.ndarray:
    n = lowpass(noise(1.6), 600)
    return n * env(len(n), 0.004, 0.45) * 5 + lowpass(noise(1.6), 90) * env(int(1.6 * SR), 0.01, 0.6) * 8


def bells() -> np.ndarray:
    out = np.zeros(int(2.4 * SR))
    for k, f in enumerate((880, 1108.7, 1318.5, 1760)):
        b = tone(f, 1.6, ((1, 1), (2.76, 0.4), (5.4, 0.2))) * env(int(1.6 * SR), 0.002, 0.5)
        place(out, b, k * 0.18, 0.4)
    return out


def hammer() -> np.ndarray:
    out = np.zeros(int(1.4 * SR))
    for k in range(3):
        hit = (band(noise(0.15), 200, 3000) * 2 + tone(180, 0.15) * 2) * env(int(0.15 * SR), 0.001, 0.03)
        place(out, hit, k * 0.4)
    return out


def camera() -> np.ndarray:
    out = np.zeros(int(0.6 * SR))
    for at in (0.0, 0.12):
        c = highpass(noise(0.05), 1500) * env(int(0.05 * SR), 0.0005, 0.008)
        place(out, c, at, 1.2)
    place(out, band(noise(0.12), 400, 2500) * env(int(0.12 * SR), 0.01, 0.04), 0.02, 0.4)
    return out


def cash() -> np.ndarray:
    out = np.zeros(int(1.6 * SR))
    for at in (0.0, 0.08, 0.16):
        place(out, highpass(noise(0.04), 1000) * env(int(0.04 * SR), 0.0005, 0.01), at, 0.8)
    place(out, band(noise(0.3), 300, 2000) * env(int(0.3 * SR), 0.01, 0.1), 0.3, 0.6)
    place(out, tone(2093, 1.0, ((1, 1), (2.4, 0.5))) * env(int(1.0 * SR), 0.002, 0.3), 0.55, 0.5)
    return out


def click() -> np.ndarray:
    c = highpass(noise(0.03), 2000) * env(int(0.03 * SR), 0.0003, 0.004)
    return np.concatenate([c * 1.5, np.zeros(int(0.1 * SR))])


def laser() -> np.ndarray:
    s = sweep(2400, 300, 0.45) * env(int(0.45 * SR), 0.003, 0.25)
    return np.concatenate([s, s * 0.6, np.zeros(int(0.1 * SR))]) * 0.6


def coin() -> np.ndarray:
    a = tone(988, 0.08, ((1, 1), (3, 0.2))) * env(int(0.08 * SR), 0.001, 0.05)
    b = tone(1319, 0.6, ((1, 1), (3, 0.2))) * env(int(0.6 * SR), 0.001, 0.2)
    return np.concatenate([a, b]) * 0.5


def arrow() -> np.ndarray:
    out = np.zeros(int(1.0 * SR))
    w = band(noise(0.45), 1500, 6000)
    place(out, w * np.linspace(0, 1, len(w)) ** 2 * 0.8, 0)
    place(out, (lowpass(noise(0.12), 400) * 3 + tone(120, 0.12)) * env(int(0.12 * SR), 0.001, 0.03), 0.45)
    return out


def typewriter() -> np.ndarray:
    out = np.zeros(int(2.0 * SR))
    at = 0.0
    while at < 1.5:
        place(out, band(noise(0.04), 1000, 6000) * env(int(0.04 * SR), 0.0005, 0.007), at, rng.uniform(0.6, 1.0))
        at += rng.uniform(0.07, 0.16)
    place(out, tone(2637, 0.4) * env(int(0.4 * SR), 0.001, 0.12), 1.6, 0.4)
    return out


def suction() -> np.ndarray:
    n = band(noise(1.0), 300, 3000)
    s = sweep(1200, 200, 1.0) * 0.2
    return (n * 0.6 + s) * np.linspace(1, 0, len(n)) ** 0.5 * np.linspace(0, 1, len(n)) ** 0.3


def suspense() -> np.ndarray:
    x = t(2.4)
    trem = 0.6 + 0.4 * np.sin(2 * np.pi * 9 * x)
    chord = sum(tone(f, 2.4, ((1, 1), (2, 0.4), (3, 0.2))) for f in (110, 116.5, 164.8))
    return chord * trem * np.minimum(1, x * 2) * np.minimum(1, (2.4 - x) * 3) * 0.25


def push() -> np.ndarray:
    n = band(noise(0.6), 300, 2500)
    return n * np.sin(np.pi * np.linspace(0, 1, len(n))) ** 3 * 0.9


def drumroll() -> np.ndarray:
    out = np.zeros(int(2.6 * SR))
    at = 0.0
    while at < 1.9:
        place(out, band(noise(0.05), 200, 5000) * env(int(0.05 * SR), 0.001, 0.02), at, 0.3 + 0.4 * at / 1.9)
        at += 0.035
    place(out, highpass(noise(0.9), 3000) * env(int(0.9 * SR), 0.002, 0.3), 2.0, 1.0)
    place(out, (lowpass(noise(0.3), 200) * 3) * env(int(0.3 * SR), 0.002, 0.08), 2.0, 1.0)
    return out


def wind() -> np.ndarray:
    x = t(2.6)
    n = band(noise(2.6), 150, 900)
    return n * (0.5 + 0.5 * np.sin(2 * np.pi * 0.6 * x - 1.5)) * np.minimum(1, x) * np.minimum(1, (2.6 - x) * 2) * 0.9


def whoosh() -> np.ndarray:
    n = highpass(noise(0.7), 2500)
    return n * np.sin(np.pi * np.linspace(0, 1, len(n))) ** 2 * 0.7


SOUNDS = {
    "applause": ("Applaus", applause), "bomb": ("Bombe", bomb), "breeze": ("Brise", breeze), "explode": ("Explosion", explosion),
    "chimes": ("Glocken", bells), "hammer": ("Hammer", hammer), "camera": ("Kamera", camera), "cashreg": ("Kasse", cash),
    "click": ("Klick", click), "laser": ("Laser", laser), "coin": ("Münze", coin), "arrow": ("Pfeil", arrow),
    "type": ("Schreibmaschine", typewriter), "suction": ("Sog", suction), "voltage": ("Spannung", suspense), "push": ("Stoß", push),
    "drumroll": ("Trommelwirbel", drumroll), "wind": ("Wind", wind), "whoosh": ("Zischen", whoosh),
}


def write(path: Path, x: np.ndarray) -> None:
    x = x / max(1e-9, float(np.max(np.abs(x)))) * 0.85
    data = (np.clip(x, -1, 1) * 32767).astype("<i2").tobytes()
    with wave.open(str(path), "wb") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(SR)
        w.writeframes(data)


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    for key, (label, fn) in SOUNDS.items():
        write(OUT / f"{key}.wav", fn())
        print(f"{key}.wav  {label}")


if __name__ == "__main__":
    main()
