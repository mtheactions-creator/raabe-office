"""The drag test (D-080 C): real mouse drags on an anchor of a RaabeX app (a handle, a border, the fill handle …) along many
paths, with the app's own state recorded at every displayed frame, then checked for jumps. The owner's first rule for
dragging: no sudden, unwanted jumps. Checked per drag:
- follow: the dragged value follows the pointer (|value - expected| within the anchor's tolerance, snapping allowed up to
  the anchor's snap distance, cell-valued anchors one step);
- jump: no frame changes the value by more than the pointer moved since the frame before (plus tolerance, snap, step);
- direction: the value never moves against the pointer's motion;
- release: the value after the button goes up is the last one shown while dragging;
- escape / undo (own paths): Esc during a drag, and Strg+Z after one, bring back the value from before.

An anchor gives JavaScript that returns the start point (client CSS px, plus anything `expect` needs) and the value: a
number or {key: number} (two-dimensional anchors such as a corner handle), and a function for the expected value from the
pointer's offset. The pointer's offsets are taken from the app's own pointer events (and drag events, for HTML drag and
drop), so the test does not depend on how Windows scales coordinates. Values are best given in CSS px, so that "moved more
than the pointer" means what it says; `per_px` converts for other units (degrees).
"""
from __future__ import annotations

import inspect
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
  // HTML drag and drop (sheet tabs): no pointer events while the system's drag loop runs.
  const dm = (e) => { if (e.clientX || e.clientY) { d.px = e.clientX; d.py = e.clientY; d.buttons = 1; } };
  addEventListener("pointermove", pm, true); addEventListener("pointerup", pm, true);
  addEventListener("pointerdown", (e) => { pm(e); if (d.on && !d.down) d.down = [e.clientX, e.clientY]; }, true);
  addEventListener("dragover", dm, true); addEventListener("drag", dm, true);
  addEventListener("drop", () => { d.buttons = 0; }, true); addEventListener("dragend", () => { d.buttons = 0; }, true);
  d.start = (src) => {
    d.samples = []; d.down = null; d.probe = new Function("return (" + src + ")"); d.on = true;
    const tick = () => { if (!d.on) return; let v = null; try { v = d.probe(); } catch (e) { v = null; }
      d.samples.push([performance.now(), d.px, d.py, d.buttons, v, Date.now()]); requestAnimationFrame(tick); };
    requestAnimationFrame(tick);
  };
  d.stop = () => { d.on = false; return { samples: d.samples, down: d.down }; };
  return true;
})()"""


@dataclass
class Anchor:
    """A draggable thing. `start` is JS giving {x, y} (client CSS px, plus context for `expect`); `probe` JS giving a
    number or {k: number}; `expect(dx, dy, v0[, ctx])` the value expected for a pointer offset in CSS px (ctx: what
    `start` returned); `tol` the allowed difference, `snap` how far snapping may pull; `setup` JS run before each drag;
    `quantum`: the value moves in steps of this size (cells; a number or {k: size}), 0 for continuous; `per_px`: value
    units per CSS px of pointer motion; `wrap`: the value's period (angles: 360); `threshold`: the value may stay where it
    was until the pointer has moved this far (CSS px, the system's drag threshold); `esc` / `undo`: the anchor takes
    the Esc and Strg+Z checks; `paths`: the names of paths() to use (None: all)."""
    name: str
    start: str
    probe: str
    expect: Callable[..., object]
    tol: float = 1.5
    snap: float = 0.0
    setup: str = ""
    quantum: float | dict = 0.0
    monotone: bool = True
    per_px: float = 1.0
    wrap: float = 0.0
    threshold: float = 0.0
    esc: bool | str = True  # True: Esc during the drag puts the value back; "keeps": Esc changes nothing (Excel's cells)
    undo: bool = True
    paths: list[str] | None = None
    check: str = "slow-right"  # the path the Esc and Strg+Z drags follow
    after_release: float = 0.25
    note: str = ""
    click: str = ""        # JS giving {x, y}: a real click there before the drag (selects the object)
    undo_probe: str = ""   # JS for the undo check's value, when it is not `probe` (cell contents)
    planned: bool = False  # the page gets no pointer events while dragging (native scroll bar): use the moves sent
    lag: int = 0           # frames the value may trail the pointer (system drag loops: HTML drag and drop, scroll bars)


@dataclass
class Path:
    name: str
    points: list[tuple[float, float]]  # offsets in CSS px from the start
    delay: float = 0.016
    keys: dict[int, str] = field(default_factory=dict)  # step index -> key combo held from there ("shift")
    esc: bool = False   # Esc pressed after the last point, before the button goes up
    undo: bool = False  # Strg+Z pressed after the release


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


def check_paths(anchor: Anchor) -> list[Path]:
    """The Esc and Strg+Z drags of an anchor, along its first path."""
    base = next(p for p in paths() if p.name == anchor.check)
    pts = base.points
    out = []
    if anchor.esc:
        out.append(Path("escape", pts, delay=base.delay, esc=True))
    if anchor.undo:
        out.append(Path("undo", pts, delay=base.delay, undo=True))
    return out


def select_paths(anchor: Anchor) -> list[Path]:
    ps = paths()
    return [p for p in ps if anchor.paths is None or p.name in anchor.paths]


def screen_point(target, x: float, y: float) -> tuple[float, float]:
    """A client CSS px point of the app window as screen pixels."""
    ox, oy = winapi.client_origin(target.hwnd)
    s = winapi.dpi_scale(target.hwnd)
    return ox + x * s, oy + y * s


def _ours(target) -> bool:
    fg = winapi.foreground()
    return fg == target.hwnd or winapi.window_pid(fg) == winapi.window_pid(target.hwnd)


def run_drag(target, anchor: Anchor, path: Path, js: Callable[[str], object]) -> dict:
    """One drag: set up, record, press, move along the path, (Esc,) release (, Strg+Z)."""
    if anchor.setup:
        js(anchor.setup)
        time.sleep(0.25)
    if anchor.click:
        c = js(anchor.click)
        if not c:
            return {"path": path.name, "error": "nothing to click"}
        if not _ours(target):
            winapi.activate(target.hwnd)
        if not _ours(target):
            return {"path": path.name, "error": "app window not in front"}
        cx, cy = screen_point(target, c["x"], c["y"])
        winapi.click(cx, cy)
        time.sleep(0.35)
    js(RECORDER)
    st = js(anchor.start)
    if not st:
        return {"path": path.name, "error": "no start point"}
    x0, y0 = st["x"], st["y"]
    v0 = js(anchor.probe)
    u0 = js(anchor.undo_probe) if path.undo and anchor.undo_probe else None
    if not _ours(target):
        winapi.activate(target.hwnd)
        if not _ours(target):
            return {"path": path.name, "error": "app window not in front"}
    sx, sy = screen_point(target, x0, y0)
    winapi.move(sx, sy)
    time.sleep(0.08)
    js(f"window.__drag.start({json.dumps(anchor.probe)})")
    moves: list[tuple[float, float, float]] = []
    t_down = time.time() * 1000
    winapi.mouse_down(sx, sy)
    moves.append((t_down, 0.0, 0.0))
    time.sleep(0.05)
    held: list[str] = []
    for i, (dx, dy) in enumerate(path.points):
        if i in path.keys:
            k = path.keys[i]
            winapi.key_down(winapi.vk_for(k))
            held.append(k)
        px, py = screen_point(target, x0 + dx, y0 + dy)
        winapi.move(px, py)
        moves.append((time.time() * 1000, dx, dy))
        time.sleep(path.delay)
    time.sleep(0.12)
    if path.esc:
        winapi.press("esc")
        time.sleep(0.15)
    last_px, last_py = screen_point(target, x0 + path.points[-1][0], y0 + path.points[-1][1])
    t_up = time.time() * 1000
    winapi.mouse_up(last_px, last_py)
    for k in held:
        winapi.key_up(winapi.vk_for(k))
    time.sleep(anchor.after_release)
    got = js("window.__drag.stop()") or {}
    samples = got.get("samples") or []
    try:
        after = js(anchor.probe)
    except RuntimeError:
        after = None  # the dragged thing is gone (a tab stop dragged off the ruler)
    # The offsets count from where the press really was (screen pixels round the planned point).
    if got.get("down") and not anchor.planned:
        x0, y0 = got["down"]
    rec = {"path": path.name, "x0": x0, "y0": y0, "ctx": st, "v0": v0, "after": after, "samples": samples,
           "moves": moves, "t_down": t_down, "t_up": t_up}
    if path.undo:
        up = anchor.undo_probe or anchor.probe
        if u0 is not None:
            rec["v0"], rec["after"] = u0, js(up)
        if _ours(target):
            winapi.press("ctrl+z")
            time.sleep(0.4)
            rec["undone"] = js(up)
        else:
            rec["undone"] = None
    return rec


def _vals(v) -> dict | None:
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return {"v": float(v)}
    if isinstance(v, dict) and v and all(isinstance(x, (int, float)) and not isinstance(x, bool) for x in v.values()):
        return {k: float(x) for k, x in v.items()}
    return None


def _expect(anchor: Anchor, dx: float, dy: float, v0, ctx) -> dict:
    fn = anchor.expect
    try:
        n = len(inspect.signature(fn).parameters)
    except (TypeError, ValueError):
        n = 3
    e = fn(dx, dy, v0, ctx) if n >= 4 else fn(dx, dy, v0)
    return {"v": float(e)} if isinstance(e, (int, float)) else {k: float(x) for k, x in e.items()}


def analyse(anchor: Anchor, rec: dict) -> dict:
    """The checks on one recorded drag."""
    name = rec["path"]
    if "error" in rec:
        return {"path": name, "ok": False, "error": rec["error"]}
    v0 = _vals(rec["v0"])
    if v0 is None:
        return {"path": name, "ok": False, "error": f"probe gave {rec['v0']!r}"}
    raw0 = rec["v0"]
    # ctx's x, y: where the press really was, so that ctx["x"] + dx is the pointer.
    ctx = dict(rec.get("ctx") or {}, x=rec["x0"], y=rec["y0"])
    keys = list(v0)

    def diff(a: float, b: float) -> float:
        d = a - b
        if anchor.wrap:
            d = (d + anchor.wrap / 2) % anchor.wrap - anchor.wrap / 2
        return d

    def q(k: str) -> float:
        return float(anchor.quantum.get(k, 0.0)) if isinstance(anchor.quantum, dict) else float(anchor.quantum)

    after = _vals(rec["after"])
    if name == "escape" or name == "undo":
        back = after if name == "escape" else _vals(rec.get("undone"))
        if back is None:
            return {"path": name, "ok": False, "error": f"probe gave {rec.get('undone', rec['after'])!r}"}
        dev = max(abs(diff(back[k], v0[k])) for k in keys if k in back)
        ok = dev <= anchor.tol if anchor.esc is True or name == "undo" else dev > anchor.tol
        out = {"path": name, "ok": ok, "back": round(dev, 2), "v0": rec["v0"], "after": rec["after"]}
        if name == "escape" and anchor.esc == "keeps":
            out["expect"] = "the drag goes on (Esc changes nothing)"

        if name == "undo":
            moved = max(abs(diff(after[k], v0[k])) for k in keys if after and k in after) if after else 0.0
            out["changed"] = round(moved, 2)
            out["undone"] = rec.get("undone")
            if moved <= anchor.tol + max(q(k) for k in keys):
                out["ok"] = False
                out["error"] = "the drag changed nothing"
        return out

    x0, y0 = rec["x0"], rec["y0"]
    drag = []
    moves = rec.get("moves") or []
    recent: dict[float, list[tuple[float, float]]] = {}
    for t, px, py, b, v, *wall in rec["samples"]:
        if anchor.planned and wall and moves:
            # The pointer where the harness had sent it by then.
            now = wall[0]
            b = 1 if rec["t_down"] <= now < rec["t_up"] else 0
            mv = [m for m in moves if m[0] <= now]
            _t, mdx, mdy = mv[-1] if mv else moves[0]
            px, py = x0 + mdx, y0 + mdy
            # Every pointer place sent in the frames before: a sampled frame need not see each one.
            recent[t] = [(x0 + m[1], y0 + m[2]) for m in moves if now - anchor.lag * 17 <= m[0] <= now]
        vv = _vals(v)
        if b & 1 and vv is not None and all(k in vv for k in keys):
            drag.append((t, px, py, vv))
    worst_follow, worst_jump, against = 0.0, 0.0, 0
    wf = wj = None
    prev = None
    hist: list[tuple[dict, float, float]] = []
    left_threshold = False
    for t, px, py, v in drag:
        dx, dy = px - x0, py - y0
        e = _expect(anchor, dx, dy, raw0, ctx)
        inside = anchor.threshold and not left_threshold and abs(dx) + abs(dy) <= anchor.threshold
        # A value that trails the pointer by up to `lag` frames counts as following it.
        lagged = [_expect(anchor, qx - x0, qy - y0, raw0, ctx) for _v, qx, qy in hist[-anchor.lag:]] if anchor.lag else []
        lagged += [_expect(anchor, qx - x0, qy - y0, raw0, ctx) for qx, qy in recent.get(t, [])]
        for k in keys:
            dev = min([abs(diff(v[k], e[k]))] + [abs(diff(v[k], le[k])) for le in lagged])
            if q(k):
                dev = max(0.0, dev - q(k))
            if inside:
                dev = min(dev, abs(diff(v[k], v0[k])))
            f = dev - anchor.snap if dev > anchor.tol else 0.0
            if f > worst_follow:
                worst_follow = f
                wf = {"key": k, "dx": round(dx, 1), "dy": round(dy, 1), "v": round(v[k], 2), "want": round(e[k], 2)}
        if prev is not None:
            pv, ppx, ppy = prev
            moved = math.hypot(px - ppx, py - ppy)
            if anchor.lag:
                moved = max([moved] + [math.hypot(px - qx, py - qy) for _v, qx, qy in hist[-anchor.lag - 1:]])
            pe = _expect(anchor, ppx - x0, ppy - y0, raw0, ctx)
            for k in keys:
                step = abs(diff(v[k], pv[k]))
                allowed = moved * anchor.per_px * 1.05 + anchor.tol + anchor.snap + q(k)
                if anchor.threshold and not left_threshold:
                    allowed += anchor.threshold * anchor.per_px
                if step > allowed and step - allowed > worst_jump:
                    worst_jump = step - allowed
                    wj = {"key": k, "dx": round(px - x0, 1), "dy": round(py - y0, 1), "from": round(pv[k], 2), "to": round(v[k], 2), "moved": round(moved, 1)}
                if anchor.monotone and step > anchor.tol + q(k):
                    want = diff(e[k], pe[k])
                    if want * diff(v[k], pv[k]) < 0:
                        against += 1
        if anchor.threshold and abs(dx) + abs(dy) > anchor.threshold:
            left_threshold = True
        prev = (v, px, py)
        hist.append(prev)
    last = drag[-1][3] if drag else None
    release = None
    if after is not None and last is not None:
        release = max(abs(diff(after[k], last[k])) - q(k) for k in keys if k in after)
        release = max(0.0, release)
    ok = bool(drag) and worst_follow <= anchor.tol and worst_jump == 0 and against == 0 and (release is None or release <= anchor.tol)
    out = {"path": name, "ok": ok, "frames": len(drag), "follow": round(worst_follow, 2), "jump": round(worst_jump, 2),
           "against": against, "release": None if release is None else round(release, 2), "v0": rec["v0"], "after": rec["after"]}
    if not drag:
        out["error"] = "no frames with the button down"
    if wf and worst_follow > anchor.tol:
        out["worst_follow"] = wf
    if wj:
        out["worst_jump"] = wj
    return out


def summary(anchor: Anchor, results: list[dict]) -> dict:
    return {"anchor": anchor.name, "ok": sum(1 for r in results if r["ok"]), "of": len(results),
            **({"note": anchor.note} if anchor.note else {}), "drags": results}
