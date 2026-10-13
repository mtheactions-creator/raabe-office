"""F6 text rendering: the same lines of a .docx shown by Word and by RaabeWrite at 100 %
zoom, captured from the screen and compared: sharpness (edge contrast), stroke darkness,
colour fringes (ClearType) and the horizontal glyph positions (ink profiles).

Needs the screen (both windows are shown); keyboard and mouse are not used. Crops are saved
to captures/rendering/ (local only)."""

from __future__ import annotations

import statistics
import time
from pathlib import Path

import numpy as np
from PIL import Image

from . import winapi
from .docx_fidelity import WordDocument, open_in_raabe
from .targets import WebViewTarget

OUT = Path(__file__).resolve().parents[1] / "captures" / "rendering"


def grab(x: int, y: int, w: int, h: int) -> np.ndarray:
    g = winapi.Grabber(x, y, w, h)
    try:
        raw = g.grab()
    finally:
        g.close()
    img = Image.frombuffer("RGBA", (w, h), raw, "raw", "BGRA", 0, 1).convert("RGB")
    return np.asarray(img).astype(np.float64)


def metrics(rgb: np.ndarray) -> dict:
    lum = rgb @ np.array([0.2126, 0.7152, 0.0722])
    ink = 255 - lum
    mask = ink > 40
    gx = np.abs(np.diff(lum, axis=1))[:-1, :]
    gy = np.abs(np.diff(lum, axis=0))[:, :-1]
    grad = np.hypot(gx, gy)
    edge = grad[mask[:-1, :-1]] if mask.any() else np.array([0.0])
    fringe = np.abs(rgb[..., 0] - rgb[..., 2])[mask] if mask.any() else np.array([0.0])
    return {
        "ink_pixels": int(mask.sum()),
        "sharpness": round(float(np.mean(edge)), 2),
        "stem_darkness": round(float(np.percentile(ink[mask], 95)) if mask.any() else 0.0, 1),
        "colour_fringe": round(float(np.mean(fringe)), 2),
    }


def profile(rgb: np.ndarray) -> np.ndarray:
    lum = rgb @ np.array([0.2126, 0.7152, 0.0722])
    return (255 - lum).sum(axis=0)


def compare_profiles(a: np.ndarray, b: np.ndarray, max_shift: int = 12) -> dict:
    """Best horizontal shift between two ink profiles and their correlation there."""
    n = min(len(a), len(b))
    a, b = a[:n] - a[:n].mean(), b[:n] - b[:n].mean()
    best = (-2.0, 0)
    for s in range(-max_shift, max_shift + 1):
        x = a[max(0, s): n + min(0, s)]
        y = b[max(0, -s): n - max(0, s)]
        if len(x) < 10 or not x.std() or not y.std():
            continue
        c = float(np.corrcoef(x, y)[0, 1])
        if c > best[0]:
            best = (c, s)
    return {"correlation": round(best[0], 4), "shift_px": best[1]}


def run(raabe: WebViewTarget, path: Path, lines: int = 6) -> dict:
    OUT.mkdir(parents=True, exist_ok=True)
    import ctypes
    dpi = ctypes.windll.user32.GetDpiForSystem() / 96
    # Word, shown, at 100 % in print layout, its window where RaabeWrite's is not.
    word = WordDocument(path)
    try:
        word.app.Visible = True
        word.win.WindowState = 0
        hwnd = word.win.Hwnd
        winapi.place(hwnd, 0, 0, 1900, 2000)
        word.win.View.Zoom.Percentage = 100
        # The caret goes out of the captured lines (it would count as ink).
        far = word.doc.Paragraphs(min(12, word.doc.Paragraphs.Count)).Range.Start
        word.win.Selection.SetRange(far, far)
        word.win.ScrollIntoView(word.doc.Range(0, 0))
        time.sleep(1.0)
        left, top, _w, h = word.win.GetPoint(0, 0, 0, 0, word.doc.Range(0, 0))
        winapi.activate(hwnd)
        time.sleep(0.8)
        w_px = int(560 * dpi)
        line_px = int(h)
        word_img = grab(int(left - 6 * dpi), int(top - 2 * dpi), w_px, int(line_px * lines + 4 * dpi))
    finally:
        word.close()
    # RaabeWrite, shown on the same screen.
    winapi.place(raabe.hwnd, 0, 0, 1900, 2000)
    open_in_raabe(raabe, path)
    raabe.js("__app.run('view.zoom100')")
    n = len(raabe.text())
    raabe.set_selection(n, n)
    raabe.js("__probe.reveal(0)")
    winapi.activate(raabe.hwnd)
    time.sleep(1.0)
    x, y, hh = raabe.point_at(0)
    ours_img = grab(int(x - 6 * dpi), int(y - 2 * dpi), w_px, int(line_px * lines + 4 * dpi))
    Image.fromarray(word_img.astype(np.uint8)).save(OUT / f"{path.stem}-word.png")
    Image.fromarray(ours_img.astype(np.uint8)).save(OUT / f"{path.stem}-raabe.png")
    mw, mo = metrics(word_img), metrics(ours_img)
    rows = []
    for k in range(lines):
        a, b = int(k * line_px + 2 * dpi), int((k + 1) * line_px + 2 * dpi)
        rows.append(compare_profiles(profile(word_img[a:b]), profile(ours_img[a:b])))
    return {
        "word": mw, "ours": mo,
        "sharpness_ratio": round(mo["sharpness"] / mw["sharpness"], 3) if mw["sharpness"] else None,
        "ink_ratio": round(mo["ink_pixels"] / mw["ink_pixels"], 3) if mw["ink_pixels"] else None,
        "lines": rows,
        "median_line_correlation": round(statistics.median(r["correlation"] for r in rows), 4),
    }
