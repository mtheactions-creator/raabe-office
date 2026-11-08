"""The drag test (D-080 C): real mouse drags on an anchor of a RaabeX app (a handle, a border, the fill handle …) along many
paths, with the app's own state recorded at every displayed frame, then checked for jumps. The owner's first rule for
dragging: no sudden, unwanted jumps. Checked per drag:
- follow: the dragged value follows the pointer (|value - expected| within the anchor's tolerance, snapping allowed up to
  the anchor's snap distance);
- jump: no frame changes the value by more than the pointer moved since the frame before (plus the snap distance);
- direction: the value never moves against the pointer's motion;
- release: the value after the button goes up is the last one shown while dragging;
- escape / undo: Esc during a drag and Strg+Z after one bring back the value from before.

An anchor gives JavaScript that returns the start point (client CSS px) and the value(s), and a function for the expected
value from the pointer's offset. The pointer's offsets are taken from the app's own pointer events, so the test does not
depend on how Windows scales coordinates.
"""
from __future__ import annotations

import json
import math
import time
from dataclasses import dataclass, field
from typing import Callable

from . import winapi

RECORDER = r"""(() => {
  if (window.__drag) return true;
  const d = window.__drag = { samples: [], on: false, px: 0, py: 0, buttons: 0, probe: null };
  const pm = (e) => { d.px = e.clientX; d.py = e.clientY; d.buttons = e.buttons; };
  addEventListener("pointermove", pm, true); addEventListener("pointerdown", pm, true); addEventListener("pointerup", pm, true);
  d.start = (src) => {
    d.samples = []; d.probe = new Function("return (" + src + ")"); d.on = true;
    const tick = () => { if (!d.on) return; let v = null; try { v = d.probe(); } catch (e) { v = null; }
      d.samples.push([performance.now(), d.px, d.py, d.buttons, v]); requestAnimationFrame(tick); };
    requestAnimationFrame(tick);
  };
  d.stop = () => { d.on = false; return d.samples; };
  return true;
})()"""


@dataclass
class Anchor:
    """A draggable thing. `start` is JS giving {x, y} (client CSS px); `probe` JS giving a number (or {k: number});
    `expect(dx, dy, v0)` the value expected for a pointer offset in CSS px; `tol` the allowed difference, `snap` how far
    snapping may pull; `setup` JS run before each drag; `quantum`: the value moves in steps (cells), 0 for continuous."""
    name: str
    start: str
    probe: str
    expect: Callable[[float, float, float], float]
    tol: float = 1.5
    snap: float = 0.0
    setup: str = ""
    quantum: float = 0.0
    monotone: bool = True


@dataclass
class Path:
    name: str
    points: list[tuple[float, float]]  # offsets in CSS px from the start
    delay: float = 0.016
    keys: dict[int, str] = field(default_factory=dict)  # step index -> key combo held from there ("shift")


def paths() -> list[Path]:
    """The pointer paths every anchor is dragged along (offsets in CSS px)."""
    out = [
        Path("slow-right", [(i, 0) for i in range(1, 81)]),
        Path("slow-down", [(0, i) for i in range(1, 61)]),
        Path("fast-right", [(i * 20, 0) for i in range(1, 7)], delay=0.01),
        Path("fast-left", [(-i * 15, 0) for i in range(1, 5)], delay=0.01),
        Path("threshold", [(1, 0), (2, 1), (1, 1), (3, 2), (2, 2), (4, 3), (3, 1)], delay=0.03),
        Path("back-forth", [(i, 0) for i in range(1, 41)] + [(40 - i, 0) for i in range(1, 61)] + [(-20 + i, 0) for i in range(1, 41)]),
        Path("diagonal", [(i * 1.5, i) for i in range(1, 51)]),
        Path("zigzag", [(i * 2, 8 if i % 2 else -8) for i in range(1, 41)]),
        Path("jitter", [(30 + (1 if i % 3 == 0 else -1 if i % 3 == 1 else 0), (1 if i % 2 else 0)) for i in range(40)]),
        Path("long", [(i * 4, i * 2) for i in range(1, 101)], delay=0.008),
    ]
    return out


def screen_point(target, x: float, y: float) -> tuple[float, float]:
    """A client CSS px point of the app window as screen pixels."""
    ox, oy = winapi.client_origin(target.hwnd)
    s = winapi.dpi_scale(target.hwnd)
    return ox + x * s, oy + y * s


def run_drag(target, anchor: Anchor, path: Path, js: Callable[[str], object]) -> dict:
    """One drag: set up, record, press, move along the path, release; then Esc and undo checks are separate drags."""
    if anchor.setup:
        js(anchor.setup)
        time.sleep(0.2)
    js(RECORDER)
    st = js(anchor.start)
    if not st:
        return {"path": path.name, "error": "no start point"}
    x0, y0 = st["x"], st["y"]
    v0 = js(anchor.probe)
    sx, sy = screen_point(target, x0, y0)
    winapi.move(sx, sy)
    time.sleep(0.08)
    js(f"window.__drag.start({json.dumps(anchor.probe)})")
    winapi.mouse_down(sx, sy)
    time.sleep(0.05)
    held: list[str] = []
    for i, (dx, dy) in enumerate(path.points):
        if i in path.keys:
            k = path.keys[i]
            winapi.key_down(winapi.vk_for(k))
            held.append(k)
        px, py = screen_point(target, x0 + dx, y0 + dy)
        winapi.move(px, py)
        time.sleep(path.delay)
    time.sleep(0.12)
    last_px, last_py = screen_point(target, x0 + path.points[-1][0], y0 + path.points[-1][1])
    winapi.mouse_up(last_px, last_py)
    for k in held:
        winapi.key_up(winapi.vk_for(k))
    time.sleep(0.25)
    samples = js("window.__drag.stop()") or []
    after = js(anchor.probe)
    return {"path": path.name, "x0": x0, "y0": y0, "v0": v0, "after": after, "samples": samples}


def analyse(anchor: Anchor, rec: dict) -> dict:
    """The checks on one recorded drag."""
    if "error" in rec:
        return {"path": rec["path"], "ok": False, "error": rec["error"]}
    v0 = rec["v0"]
    if not isinstance(v0, (int, float)):
        return {"path": rec["path"], "ok": False, "error": f"probe gave {v0!r}"}
    x0, y0 = rec["x0"], rec["y0"]
    drag = [s for s in rec["samples"] if s[3] & 1 and isinstance(s[4], (int, float))]
    worst_follow = 0.0
    worst_jump = 0.0
    against = 0
    prev = None
    for t, px, py, _b, v in drag:
        dx, dy = px - x0, py - y0
        e = anchor.expect(dx, dy, v0)
        if anchor.quantum:
            dev = max(0.0, abs(v - e) - anchor.quantum)
        else:
            dev = abs(v - e)
        worst_follow = max(worst_follow, dev - anchor.snap if dev > anchor.tol else 0.0)
        if prev is not None:
            pv, ppx, ppy = prev
            moved = math.hypot(px - ppx, py - ppy)
            step = abs(v - pv)
            allowed = moved * 1.05 + anchor.tol + anchor.snap + anchor.quantum
            if step > allowed:
                worst_jump = max(worst_jump, step - allowed)
            if anchor.monotone and step > anchor.tol + anchor.quantum:
                want = anchor.expect(px - x0, py - y0, v0) - anchor.expect(ppx - x0, ppy - y0, v0)
                if want * (v - pv) < 0:
                    against += 1
        prev = (v, px, py)
    last = drag[-1][4] if drag else None
    after = rec["after"]
    release = abs(after - last) if isinstance(after, (int, float)) and isinstance(last, (int, float)) else None
    ok = worst_follow <= anchor.tol and worst_jump == 0 and against == 0 and (release is None or release <= anchor.tol + anchor.quantum)
    return {"path": rec["path"], "ok": ok, "frames": len(drag), "follow": round(worst_follow, 2), "jump": round(worst_jump, 2),
            "against": against, "release": None if release is None else round(release, 2), "v0": v0, "after": after}
