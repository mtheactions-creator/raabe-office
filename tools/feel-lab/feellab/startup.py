"""F5 start time: from process launch until the window is visible and until document
text is actually painted. Word is started as a separate instance (/x) with the test
document and closed again with WM_CLOSE, so documents the user has open are never
touched."""

from __future__ import annotations

import ctypes
import statistics
import subprocess
import time
from pathlib import Path

from . import winapi
from .targets import PROTOTYPE_EXE, WINDOW_RECT, WordTarget

WINWORD = Path(r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE")
RUNS = 5
DOCX = Path(__file__).resolve().parents[1] / "captures" / "short.docx"


def _ensure_docx() -> None:
    """The "short" test document as a .docx, built once by Word itself."""
    if DOCX.exists():
        return
    DOCX.parent.mkdir(exist_ok=True)
    w = WordTarget("short")
    try:
        w.doc.SaveAs2(str(DOCX), 16)  # wdFormatDocumentDefault
    finally:
        w.close()


def _pid(hwnd: int) -> int:
    pid = ctypes.c_ulong()
    winapi.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return pid.value


def _text_painted(hwnd: int, evidence: Path, timeout: float = 30.0) -> float:
    """Waits until the middle of the window shows text: some near-black pixels on a
    mostly white page (a dark, not yet painted window does not count). The detected
    frame is saved as evidence."""
    left, top, right, bottom = winapi.window_rect(hwnd)
    w, h = right - left, bottom - top
    rx, ry, rw, rh = left + w // 5, top + int(h * 0.4), w * 2 // 5, int(h * 0.2)
    g = winapi.Grabber(rx, ry, rw, rh)
    centre = winapi.POINT(rx + rw // 2, ry + rh // 2)
    winapi.user32.WindowFromPoint.argtypes = (winapi.POINT,)
    winapi.user32.WindowFromPoint.restype = ctypes.c_void_p
    winapi.user32.GetAncestor.argtypes = (ctypes.c_void_p, ctypes.c_uint)
    winapi.user32.GetAncestor.restype = ctypes.c_void_p
    try:
        end = time.perf_counter() + timeout
        while time.perf_counter() < end:
            # Only pixels of the window under test count, never a window behind it.
            under = winapi.user32.WindowFromPoint(centre)
            if not under or winapi.user32.GetAncestor(under, 2) != hwnd:  # GA_ROOT
                time.sleep(0.002)
                continue
            raw = g.grab()
            greens = raw[1::28]  # green channel of every 7th pixel
            dark = sum(1 for v in greens if v < 90)
            light = sum(1 for v in greens if v > 230)
            if dark >= 20 and light > len(greens) // 2:
                t = time.perf_counter()
                from PIL import Image
                Image.frombuffer("RGBA", (rw, rh), raw, "raw", "BGRA", 0, 1).convert("RGB").save(evidence)
                return t
            time.sleep(0.002)
    finally:
        g.close()
    raise TimeoutError("no text painted")


def _summary(values: list[float]) -> dict:
    return {"median_ms": round(statistics.median(values)), "min_ms": round(min(values)),
            "max_ms": round(max(values)), "runs_ms": [round(v) for v in values]}


def word_once() -> tuple[float, float]:
    t0 = time.perf_counter()
    proc = subprocess.Popen([str(WINWORD), "/q", "/x", str(DOCX)])
    hwnd = winapi.wait_for_window(lambda h: winapi.window_class(h) == "OpusApp" and _pid(h) == proc.pid, timeout=60)
    shown = time.perf_counter()
    winapi.place(hwnd, *WINDOW_RECT)
    winapi.activate(hwnd, settle=0)
    painted = _text_painted(hwnd, DOCX.parent / "startup-word.png")
    winapi.user32.PostMessageW(hwnd, 0x0010, 0, 0)  # WM_CLOSE: document is unmodified
    proc.wait(30)
    return (shown - t0) * 1000, (painted - t0) * 1000


def prototype_once(engine: str) -> tuple[float, float]:
    t0 = time.perf_counter()
    proc = subprocess.Popen([str(PROTOTYPE_EXE), engine, "short"])
    title = f"Feel-Prototyp {engine.upper()}"
    hwnd = winapi.wait_for_window(lambda h: winapi.window_title(h) == title and _pid(h) == proc.pid)
    shown = time.perf_counter()
    winapi.place(hwnd, *WINDOW_RECT)
    winapi.activate(hwnd, settle=0)
    painted = _text_painted(hwnd, DOCX.parent / f"startup-proto-{engine}.png")
    proc.terminate()
    proc.wait(10)
    return (shown - t0) * 1000, (painted - t0) * 1000


def run() -> dict:
    _ensure_docx()
    out = {}
    for name, fn in (("word", word_once), ("proto-a", lambda: prototype_once("a")), ("proto-b", lambda: prototype_once("b"))):
        shown, painted = [], []
        for _ in range(RUNS):
            s, p = fn()
            shown.append(s)
            painted.append(p)
            time.sleep(1.5)
        out[name] = {"window_visible": _summary(shown), "text_painted": _summary(painted)}
    return out
