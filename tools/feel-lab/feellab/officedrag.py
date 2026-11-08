"""Real drags in our own Office instance for the drag test's new paths (D-127): modifier keys pressed mid-drag
(Umschalt, Strg, Alt; let go mid-drag or still held when the button goes up) and the pointer leaving the window
(coming back with the button down, or the button going up outside and the pointer coming back without it).
The probes (probe_excel_drag_mods.py, probe_word_drag_mods.py, probe_powerpoint_drag_mods.py) read the result
through COM. Points are screen pixels; the paths are dragtest.py's (MOD_POINTS, out_points) at the screen's scale.
"""
from __future__ import annotations

import time
from typing import Callable

from . import winapi
from .dragtest import MOD_POINTS

VARIANTS = ("plain", "shift-held", "ctrl-held", "alt-held", "shift-mid", "ctrl-mid", "alt-mid", "out-back", "out-release")


def front(hwnd: int, ours: Callable[[], bool], tries: int = 8) -> None:
    """Brings our Office window to the front; other sessions' windows may take it meanwhile (they are left alone)."""
    for _ in range(tries):
        try:
            winapi.activate(hwnd, settle=0.3)
        except RuntimeError:
            pass
        if ours():
            return
        time.sleep(0.6)
    fg = winapi.foreground()
    raise RuntimeError(f"our Office window is not in front (in front: a {winapi.window_class(fg)!r} window)")


def drag(p: tuple[float, float], variant: str, ours: Callable[[], bool], scale: float = 1.5,
         window: tuple[int, int, int, int] | None = None, direction: str = "right") -> None:
    """One real drag from screen point p. `window`: the Office window's screen rectangle (left, top, right, bottom),
    for the out-of-window variants; `direction`: "right" or "down"."""
    if not ours():
        raise RuntimeError("our Office window is not in front")
    winapi.move(*p)
    time.sleep(0.2)
    winapi.mouse_down(*p)
    time.sleep(0.08)
    held: list[str] = []
    if variant.startswith("out"):
        assert window is not None
        far = (window[2] + 90 - p[0]) if direction == "right" else (window[3] + 90 - p[1])
        n = max(4, int(far / 18))
        there = [far * i / n for i in range(1, n + 1)]
        back = [far - (far - 60) * i / n for i in range(1, n + 1)]
        up = False
        for j, v in enumerate(there + back):
            x, y = (p[0] + v, p[1]) if direction == "right" else (p[0], p[1] + v)
            winapi.move(x, y)
            time.sleep(0.016)
            if not up and not ours():
                _abort(x, y, [])
            if j == n - 1:
                time.sleep(0.8)
                if variant == "out-release":
                    winapi.mouse_up(x, y)
                    up = True
                    time.sleep(0.15)
        time.sleep(0.2)
        if not up:
            winapi.mouse_up(x, y)
        time.sleep(0.5)
        return
    key = variant.split("-")[0] if variant != "plain" else ""
    x, y = p
    for i, (dx, dy) in enumerate(MOD_POINTS):
        if key and i in (20, 30, 40) and not ours():
            # Another session's window took the front: no keys into it.
            winapi.mouse_up(x if i else p[0], y if i else p[1])
            for k in held:
                winapi.key_up(winapi.vk_for(k))
            raise RuntimeError("our Office window lost the front during the drag")
        if key and ((variant.endswith("-mid") and i == 20) or (variant.endswith("-held") and i == 30)):
            winapi.key_down(winapi.vk_for(key))
            held.append(key)
        if key and variant.endswith("-mid") and i == 40 and key in held:
            winapi.key_up(winapi.vk_for(key))
            held.remove(key)
        x, y = p[0] + dx * scale, p[1] + dy * scale
        winapi.move(x, y)
        if not ours():
            _abort(x, y, held)
        time.sleep(0.016)
    time.sleep(0.2)
    winapi.mouse_up(x, y)
    time.sleep(0.05)
    for k in held:
        winapi.key_up(winapi.vk_for(k))
    time.sleep(0.5)


def _abort(x: float, y: float, held: list[str]) -> None:
    """Another session's window came to the front during the drag (Windows takes the mouse capture then): the drag
    ends and is done again (run_variant)."""
    winapi.mouse_up(x, y)
    for k in held:
        winapi.key_up(winapi.vk_for(k))
    raise RuntimeError("our Office window lost the front during the drag")


def run_variant(attempt: Callable[[], dict], ours: Callable[[], bool], tries: int = 4) -> dict:
    """One variant, done again when another window took the front (up to `tries` times)."""
    for i in range(tries):
        try:
            return attempt()
        except RuntimeError as e:
            if ("front" not in str(e) and "activate" not in str(e)) or i == tries - 1:
                raise
            settle(ours)
            time.sleep(1.5)
    raise RuntimeError("unreachable")


def settle(ours: Callable[[], bool]) -> None:
    """After a drag (and reading its result): Esc twice, which also puts away the keytips Alt shows when let go."""
    for _ in range(2):
        if ours():
            winapi.press("esc")
            time.sleep(0.15)
