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
  if (window.__drag && window.__drag.v === 3) return true;
  const d = window.__drag = { v: 3, samples: [], on: false, px: 0, py: 0, buttons: 0, mods: 0, probe: null, scroll: null };
  // The modifier keys as the app's own pointer events see them (Umschalt 1, Strg 2, Alt 4).
  const mk = (e) => (e.shiftKey ? 1 : 0) | (e.ctrlKey ? 2 : 0) | (e.altKey ? 4 : 0);
  const pm = (e) => { d.px = e.clientX; d.py = e.clientY; d.buttons = e.buttons; d.mods = mk(e); };
  // HTML drag and drop (sheet tabs): no pointer events while the system's drag loop runs.
  const dm = (e) => { if (e.clientX || e.clientY) { d.px = e.clientX; d.py = e.clientY; d.buttons = 1; d.mods = mk(e); } };
  addEventListener("pointermove", pm, true); addEventListener("pointerup", (e) => { pm(e); d.ups++; }, true);
  addEventListener("pointerdown", (e) => { pm(e); if (d.on && !d.down) d.down = [e.clientX, e.clientY]; }, true);
  addEventListener("dragover", dm, true); addEventListener("drag", dm, true);
  // A modifier key pressed or let go counts at once (an app may apply it before the next pointer move, as Office does).
  const km = (e) => { if (e.key === "Shift" || e.key === "Control" || e.key === "Alt") d.mods = mk(e); };
  addEventListener("keydown", km, true); addEventListener("keyup", km, true);
  addEventListener("drop", () => { d.buttons = 0; }, true); addEventListener("dragend", () => { d.buttons = 0; }, true);
  d.start = (src, scroll) => {
    d.samples = []; d.down = null; d.ups = 0; d.probe = new Function("return (" + src + ")"); d.on = true;
    d.scroll = scroll ? new Function("return (" + scroll + ")") : null;
    const tick = () => { if (!d.on) return; let v = null, s = null; try { v = d.probe(); } catch (e) { v = null; }
      if (d.scroll) { try { s = d.scroll(); } catch (e) { s = null; } }
      d.samples.push([performance.now(), d.px, d.py, d.buttons, v, Date.now(), d.mods, s]); requestAnimationFrame(tick); };
    requestAnimationFrame(tick);
  };
  d.stop = () => { d.on = false; return { samples: d.samples, down: d.down, ups: d.ups }; };
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
    # D-127: the paths out of the window, modifier keys mid-drag, other zoom levels.
    out_dir: str = ""      # where the out-of-window paths leave the window ("right", "down"; default from `check`)
    out_follow: bool = False  # the value keeps following the pointer outside the window (sizes, borders)
    autoscroll: bool | None = None  # the view scrolls while the pointer is outside (Office's behaviour; None: not checked)
    scroll: str = ""       # JS giving the view's scroll position (px; summed axes): autoscroll and its allowance
    mods: dict = field(default_factory=dict)  # "shift"/"ctrl"/"alt" -> "same" | "skip" | "free" | expect(dx, dy, v0[, ctx])
    copies: dict = field(default_factory=dict)  # key -> objects added by a release with that key held (Office's), checked with `count`
    count: str = ""        # JS giving the number of objects (copies made by Strg)
    out_cancel: bool = False  # the button let go outside the window puts the value back (Excel's column and row borders)
    zoomed: bool = True    # the anchor is in the document (dragged at other zoom levels); False: the app's own chrome


@dataclass
class Path:
    name: str
    points: list[tuple[float, float]]  # offsets in CSS px from the start
    delay: float = 0.016
    keys: dict[int, str] = field(default_factory=dict)  # step index -> key combo held from there ("shift")
    esc: bool = False   # Esc pressed after the last point, before the button goes up
    undo: bool = False  # Strg+Z pressed after the release
    keys_up: dict[int, str] = field(default_factory=dict)  # step index -> key released there (before the button)
    out: str = ""       # "right"/"down": the points are built at run time to leave the window that way and come back
    release_outside: bool = False  # the button goes up outside; the way back is moved without a button
    family: str = ""    # "out", "mods", "zoom" (D-127)


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


# ----- D-127: out of the window, modifier keys mid-drag, other zoom levels -------------------------

MOD_POINTS = [(i * 1.5, i * 0.6) for i in range(1, 61)]
ZOOMS = (50, 75, 150, 200, 400)
ZOOM_PATHS = ("diagonal", "fast-right")


def out_dir(anchor: Anchor) -> str:
    return anchor.out_dir or ("down" if anchor.check == "slow-down" else "right")


def mod_paths() -> list[Path]:
    """A key pressed and released while the button is down (-mid: down at a third, up at two thirds), or pressed and
    still held when the button goes up (-held)."""
    out = []
    for k in ("shift", "ctrl", "alt"):
        out.append(Path(f"{k}-mid", MOD_POINTS, keys={20: k}, keys_up={40: k}, family="mods"))
        out.append(Path(f"{k}-held", MOD_POINTS, keys={30: k}, family="mods"))
    return out


def out_paths(anchor: Anchor) -> list[Path]:
    d = out_dir(anchor)
    return [Path("out-back", [], out=d, family="out"), Path("out-release", [], out=d, release_outside=True, family="out")]


def family_paths(anchor: Anchor, families: list[str]) -> list[Path]:
    out: list[Path] = []
    if "out" in families:
        out += out_paths(anchor)
    if "mods" in families:
        out += mod_paths()
    for f in families:
        # "zigzag*10": one of the base paths, again and again (flaky drags).
        if "*" in f:
            base, n = f.split("*")
            p = next(q for q in paths() if q.name == base)
            out += [Path(f"{base} #{i + 1}", p.points, delay=p.delay, family="repeat") for i in range(int(n))]
    return out


def zoom_paths() -> list[Path]:
    return [p for p in paths() if p.name in ZOOM_PATHS]


def out_points(direction: str, x0: float, y0: float, w: float, h: float, step: float = 12.0) -> tuple[list[tuple[float, float]], int]:
    """Offsets from the start out of the window (60 px past its edge) and back to 40 px from the start; and the index of
    the farthest point (where the pointer waits, and where the button goes up on the release-outside path)."""
    far = (w + 60 - x0) if direction == "right" else (h + 60 - y0)
    n = max(4, int(math.ceil(far / step)))
    there = [far * i / n for i in range(1, n + 1)]
    back = [far - (far - 40) * i / n for i in range(1, n + 1)]
    pts = there + back
    if direction == "right":
        return [(v, 0.0) for v in pts], n - 1
    return [(0.0, v) for v in pts], n - 1


def screen_point(target, x: float, y: float) -> tuple[float, float]:
    """A client CSS px point of the app window as screen pixels."""
    ox, oy = winapi.client_origin(target.hwnd)
    s = winapi.dpi_scale(target.hwnd)
    return ox + x * s, oy + y * s


def _ours(target) -> bool:
    if INPUT == "cdp":
        return True
    fg = winapi.foreground()
    return fg == target.hwnd or winapi.window_pid(fg) == winapi.window_pid(target.hwnd)


# Input: "real" (the OS mouse and keyboard, under the real-input lock) or "cdp" (DevTools' Input.dispatchMouseEvent and
# dispatchKeyEvent into the app's WebView2: trusted events that do not need the foreground; D-127). Set by
# RAABE_DRAG_INPUT=cdp or the drivers' --cdp.
INPUT = __import__("os").environ.get("RAABE_DRAG_INPUT", "real")
CDP_KEYS = {"shift": ("Shift", "ShiftLeft", 16, 8), "ctrl": ("Control", "ControlLeft", 17, 2), "alt": ("Alt", "AltLeft", 18, 1),
            "esc": ("Escape", "Escape", 27, 0), "z": ("z", "KeyZ", 90, 0)}


def cdp_mode() -> None:
    """Drags through DevTools only: the app starts in the background (off screen, never taking the focus), nothing is
    activated and the real mouse and keyboard stay untouched (no real-input lock needed)."""
    global INPUT
    import os
    INPUT = "cdp"
    os.environ["RAABE_FEEL_BACKGROUND"] = "1"
    winapi.place = lambda hwnd, x, y, w, h: winapi.place_quietly(hwnd, -4100 - x, y, w, h)
    winapi.activate = lambda *a, **k: None
    winapi.click = lambda *a, **k: None  # the drivers click through Input instead


class Input:
    """Mouse and keys for one drag, in client CSS px."""

    def __init__(self, target):
        self.t = target
        self.mods = 0
        self.down_ = False
        if INPUT == "cdp":
            t = getattr(target, "t", None)
            if t is not None and hasattr(t, "cdp"):
                self.send = t.cdp
            elif hasattr(target, "cdp"):
                self.send = target.cdp
            else:
                self.send = self._send

    def _send(self, method: str, params: dict | None = None):
        t = self.t
        t._id += 1
        t.ws.send(json.dumps({"id": t._id, "method": method, "params": params or {}}))
        while True:
            msg = json.loads(t.ws.recv())
            if msg.get("id") == t._id:
                if "error" in msg:
                    raise RuntimeError(msg["error"])
                return msg.get("result")

    def _mouse(self, kind: str, x: float, y: float) -> None:
        self.send("Input.dispatchMouseEvent", {"type": kind, "x": x, "y": y, "modifiers": self.mods, "clickCount": 1,
                                               "button": "left" if kind != "mouseMoved" or self.down_ else "none",
                                               "buttons": 1 if self.down_ else 0})

    def move(self, x: float, y: float) -> None:
        if INPUT == "cdp":
            self._mouse("mouseMoved", x, y)
        else:
            winapi.move(*screen_point(self.t, x, y))

    def down(self, x: float, y: float) -> None:
        self.down_ = True
        if INPUT == "cdp":
            self._mouse("mousePressed", x, y)
        else:
            winapi.mouse_down(*screen_point(self.t, x, y))

    def up(self, x: float, y: float) -> None:
        self.down_ = False
        if INPUT == "cdp":
            self._mouse("mouseReleased", x, y)
        else:
            winapi.mouse_up(*screen_point(self.t, x, y))

    def click(self, x: float, y: float) -> None:
        self.move(x, y)
        self.down(x, y)
        time.sleep(0.03)
        self.up(x, y)

    def key(self, name: str, down: bool) -> None:
        if INPUT != "cdp":
            (winapi.key_down if down else winapi.key_up)(winapi.vk_for(name))
            return
        key, code, vk, bit = CDP_KEYS[name]
        if down:
            self.mods |= bit
        self.send("Input.dispatchKeyEvent", {"type": "rawKeyDown" if down else "keyUp", "key": key, "code": code,
                                             "windowsVirtualKeyCode": vk, "nativeVirtualKeyCode": vk, "modifiers": self.mods})
        if not down:
            self.mods &= ~bit

    def press(self, combo: str) -> None:
        if INPUT != "cdp":
            winapi.press(combo)
            return
        keys = combo.split("+")
        for k in keys:
            self.key(k, True)
        for k in reversed(keys):
            self.key(k, False)


def run_drag(target, anchor: Anchor, path: Path, js: Callable[[str], object], before_start: str = "", reveal: str = "") -> dict:
    """One drag: set up, record, press, move along the path, (Esc,) release (, Strg+Z).
    `before_start`: JS run after the setup (a zoom level); `reveal`: JS with {dx} {dy} that scrolls the view so that the
    start point comes into the window when it is outside (other zoom levels)."""
    if anchor.setup:
        js(anchor.setup)
        time.sleep(0.25)
    if before_start:
        js(before_start)
        time.sleep(0.4)
    if anchor.click:
        c = js(anchor.click)
        if not c:
            return {"path": path.name, "error": "nothing to click"}
        if reveal:
            # The object is selected by a click first: it must be in the window for that (other zoom levels).
            cw, ch = js("[innerWidth, innerHeight]")
            if not (40 <= c["x"] <= cw - 60 and 150 <= c["y"] <= ch - 80):
                js(reveal.replace("{dx}", str(round(c["x"] - cw / 2))).replace("{dy}", str(round(c["y"] - ch / 2)))
                   .replace("{x}", str(round(c["x"]))).replace("{y}", str(round(c["y"]))))
                time.sleep(0.4)
                c = js(anchor.click)
        if not _ours(target):
            winapi.activate(target.hwnd)
        if not _ours(target):
            return {"path": path.name, "error": "app window not in front"}
        Input(target).click(c["x"], c["y"])
        time.sleep(0.35)
    js(RECORDER)
    w, h = js("[innerWidth, innerHeight]")
    st = js(anchor.start)
    if not st:
        return {"path": path.name, "error": "no start point"}
    if reveal and not (40 <= st["x"] <= w - 60 and 150 <= st["y"] <= h - 80):
        js(reveal.replace("{dx}", str(round(st["x"] - w / 2))).replace("{dy}", str(round(st["y"] - h / 2)))
           .replace("{x}", str(round(st["x"]))).replace("{y}", str(round(st["y"]))))
        time.sleep(0.4)
        st = js(anchor.start)
        if not st:
            return {"path": path.name, "error": "no start point"}
    x0, y0 = st["x"], st["y"]
    if not (0 <= x0 <= w and 0 <= y0 <= h):
        return {"path": path.name, "error": f"start point outside the window ({x0:.0f}, {y0:.0f})"}
    v0 = js(anchor.probe)
    u0 = js(anchor.undo_probe) if path.undo and anchor.undo_probe else None
    # Let go outside on an anchor whose drag Office drops there: the contents must be as before.
    c0 = js(anchor.undo_probe) if path.release_outside and anchor.out_cancel and anchor.undo_probe else None
    n0 = js(anchor.count) if anchor.count else None
    points, far = (out_points(path.out, x0, y0, w, h) if path.out else (path.points, -1))
    if not _ours(target):
        try:
            winapi.activate(target.hwnd)
        except RuntimeError:
            pass
        if not _ours(target):
            return {"path": path.name, "error": "app window not in front"}
    inp = Input(target)
    inp.move(x0, y0)
    time.sleep(0.08)
    js(f"window.__drag.start({json.dumps(anchor.probe)}, {json.dumps(anchor.scroll or None)})")
    moves: list[tuple[float, float, float]] = []
    t_down = time.time() * 1000
    inp.down(x0, y0)
    moves.append((t_down, 0.0, 0.0))
    time.sleep(0.05)
    held: list[str] = []
    t_up = None
    lost = False
    for i, (dx, dy) in enumerate(points):
        if (i in path.keys or i in path.keys_up) and not _ours(target):
            # Another session's window took the front: no keys into it.
            inp.up(x0 + dx, y0 + dy)
            for k in held:
                inp.key(k, False)
            js("window.__drag.stop()")
            return {"path": path.name, "error": "app window lost the front during the drag"}
        if i in path.keys:
            k = path.keys[i]
            inp.key(k, True)
            held.append(k)
        if i in path.keys_up and path.keys_up[i] in held:
            k = path.keys_up[i]
            inp.key(k, False)
            held.remove(k)
        inp.move(x0 + dx, y0 + dy)
        moves.append((time.time() * 1000, dx, dy))
        if INPUT != "cdp" and not lost and not _ours(target):
            # Another session's window came to the front during the drag: Windows takes the mouse capture
            # from the app then, so the drag is not the app's doing; it is done again (run_families).
            lost = True
        time.sleep(path.delay)
        if i == far:
            time.sleep(0.8)  # outside the window: the view may scroll meanwhile
            if path.release_outside:
                t_up = time.time() * 1000
                inp.up(x0 + dx, y0 + dy)
                moves.append((t_up, dx, dy))
                time.sleep(0.15)
    time.sleep(0.12)
    if path.esc:
        inp.press("esc")
        time.sleep(0.15)
    if t_up is None:
        t_up = time.time() * 1000
        inp.up(x0 + points[-1][0], y0 + points[-1][1])
    for k in held:
        inp.key(k, False)
    time.sleep(anchor.after_release)
    got = js("window.__drag ? window.__drag.stop() : { reloaded: true }") or {}
    if got.get("reloaded"):
        return {"path": path.name, "error": "the page was reloaded during the drag"}
    samples = got.get("samples") or []
    try:
        after = js(anchor.probe)
    except RuntimeError:
        after = None  # the dragged thing is gone (a tab stop dragged off the ruler)
    # The offsets count from where the press really was (screen pixels round the planned point).
    if got.get("down") and not anchor.planned:
        x0, y0 = got["down"]
    rec = {"path": path.name, "x0": x0, "y0": y0, "ctx": st, "v0": v0, "after": after, "samples": samples,
           "moves": moves, "t_down": t_down, "t_up": t_up, "w": w, "h": h, "ups": got.get("ups")}
    if lost:
        return {"path": path.name, "error": "another window came to the front during the drag"}
    if path.family:
        rec["family"] = path.family
        rec["release_outside"] = path.release_outside
    if n0 is not None:
        rec["count"] = [n0, js(anchor.count)]
    if c0 is not None:
        rec["content"] = [c0, js(anchor.undo_probe)]
    if path.undo:
        up = anchor.undo_probe or anchor.probe
        if u0 is not None:
            rec["v0"], rec["after"] = u0, js(up)
        if _ours(target):
            inp.press("ctrl+z")
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


MOD_BITS = (("shift", 1), ("ctrl", 2), ("alt", 4))


def mod_spec(anchor: Anchor, mods: int):
    """What the modifier keys held in a frame do to the anchor's expectation: "same", "skip", "free" or a function."""
    for k, bit in MOD_BITS:
        if mods & bit:
            spec = anchor.mods.get(k, "same")
            if spec != "same":
                return spec
    return "same"


def _expect(anchor: Anchor, dx: float, dy: float, v0, ctx, mods: int = 0) -> dict:
    fn = anchor.expect
    spec = mod_spec(anchor, mods) if mods else "same"
    if callable(spec):
        fn = spec
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
    fam = rec.get("family", "")
    w, h = rec.get("w") or 1e9, rec.get("h") or 1e9
    t_up = rec["t_up"]
    drag = []
    post = []   # frames after the button went up outside the window (the way back without a button)
    moves = rec.get("moves") or []
    recent: dict[float, list[tuple[float, float]]] = {}
    for t, px, py, b, v, *more in rec["samples"]:
        wall = more[0] if more else None
        mods = more[1] if len(more) > 1 and more[1] else 0
        scr = more[2] if len(more) > 2 else None
        if anchor.planned and wall and moves:
            # The pointer where the harness had sent it by then.
            now = wall
            b = 1 if rec["t_down"] <= now < t_up else 0
            mv = [m for m in moves if m[0] <= now]
            _t, mdx, mdy = mv[-1] if mv else moves[0]
            px, py = x0 + mdx, y0 + mdy
            # Every pointer place sent in the frames before: a sampled frame need not see each one.
            recent[t] = [(x0 + m[1], y0 + m[2]) for m in moves if now - anchor.lag * 17 <= m[0] <= now]
        vv = _vals(v)
        if vv is None or not all(k in vv for k in keys):
            continue
        if rec.get("release_outside") and wall and wall > t_up + 60:
            post.append((t, px, py, vv, scr))
            continue
        if b & 1 and (not wall or wall < t_up + 60 or not rec.get("release_outside")):
            drag.append((t, px, py, vv, mods, scr))
    worst_follow, worst_jump, against = 0.0, 0.0, 0
    wf = wj = None
    prev = None
    hist: list[tuple[dict, float, float]] = []
    left_threshold = False
    outside_frames = 0
    scrolled_outside = 0.0
    was_outside = False
    since_switch = 99
    for t, px, py, v, mods, scr in drag:
        since_switch = 0 if prev is not None and mods != prev[3] else since_switch + 1
        dx, dy = px - x0, py - y0
        outside = not (0 <= px <= w and 0 <= py <= h)
        if outside:
            outside_frames += 1
        spec = mod_spec(anchor, mods) if mods else "same"
        e = _expect(anchor, dx, dy, raw0, ctx, mods)
        inside = anchor.threshold and not left_threshold and abs(dx) + abs(dy) <= anchor.threshold
        # A value that trails the pointer by up to `lag` frames counts as following it.
        lagged = [_expect(anchor, qx - x0, qy - y0, raw0, ctx, mods) for _v, qx, qy in hist[-anchor.lag:]] if anchor.lag else []
        lagged += [_expect(anchor, qx - x0, qy - y0, raw0, ctx, mods) for qx, qy in recent.get(t, [])]
        check_follow = spec != "skip" and (not outside or anchor.out_follow)
        if since_switch < 2:
            # The frame a key went down or up and the one after: the app may apply it at once or with the next move.
            check_follow = False
        if fam == "out" and math.hypot(dx, dy) > 150:
            # The long way out crosses the limits an anchor has (an indent's, a margin's): what counts is that the
            # value is under the pointer again near the start, with no jump on the way (D-127).
            check_follow = False
        # Back in the window after the view scrolled: the dragged thing is under the pointer again within a frame or two.
        if was_outside and not outside and anchor.scroll:
            check_follow = False
        # ("snap", px): the key makes the value snap (Excel's Alt: to the cell grid) by up to px.
        extra = float(spec[1]) if isinstance(spec, tuple) else 0.0
        # "free": the key turns snapping off (Word, PowerPoint): only rounding is left.
        snap = min(anchor.snap, 1.0) if spec == "free" else anchor.snap + extra
        for k in keys:
            if not check_follow:
                break
            dev = min([abs(diff(v[k], e[k]))] + [abs(diff(v[k], le[k])) for le in lagged])
            if q(k):
                dev = max(0.0, dev - q(k))
            if inside:
                dev = min(dev, abs(diff(v[k], v0[k])))
            f = dev - snap if dev > anchor.tol else 0.0
            if f > worst_follow:
                worst_follow = f
                wf = {"key": k, "dx": round(dx, 1), "dy": round(dy, 1), "v": round(v[k], 2), "want": round(e[k], 2), "mods": mods}
        if prev is not None:
            pv, ppx, ppy, pmods, pscr = prev
            moved = math.hypot(px - ppx, py - ppy)
            if anchor.lag:
                moved = max([moved] + [math.hypot(px - qx, py - qy) for _v, qx, qy in hist[-anchor.lag - 1:]])
            # The view scrolled between the frames (auto-scroll): what is drawn moves with it.
            scrolled = abs(scr - pscr) if isinstance(scr, (int, float)) and isinstance(pscr, (int, float)) else 0.0
            if outside and scrolled:
                scrolled_outside += scrolled
            pe = _expect(anchor, ppx - x0, ppy - y0, raw0, ctx, mods)
            # A modifier key pressed or let go: the value may go where the new rule puts it (wanted), checked by follow.
            switched = since_switch < 2
            for k in keys:
                if switched:
                    break
                step = abs(diff(v[k], pv[k]))
                # As far as the pointer moved, or as far as the rule moves the value for that motion (Strg's sizing around
                # the middle moves a side twice as far); a cell-valued value whose axis may switch (the fill handle) two steps.
                allowed = max(moved * anchor.per_px, abs(diff(e[k], pe[k]))) * 1.05 + anchor.tol + anchor.snap + q(k) * (1 if anchor.monotone else 2) + scrolled + extra
                if anchor.threshold and not left_threshold:
                    allowed += anchor.threshold * anchor.per_px
                if step > allowed and step - allowed > worst_jump:
                    worst_jump = step - allowed
                    wj = {"key": k, "dx": round(px - x0, 1), "dy": round(py - y0, 1), "from": round(pv[k], 2), "to": round(v[k], 2), "moved": round(moved, 1), "mods": mods}
                if anchor.monotone and step > anchor.tol + q(k) and not scrolled and spec != "skip":
                    want = diff(e[k], pe[k])
                    if want * diff(v[k], pv[k]) < 0:
                        against += 1
        if anchor.threshold and abs(dx) + abs(dy) > anchor.threshold:
            left_threshold = True
        was_outside = outside
        prev = (v, px, py, mods, scr)
        hist.append((v, px, py))
    last = drag[-1][3] if drag else None
    release = None
    if after is not None and last is not None:
        release = max(abs(diff(after[k], last[k])) - q(k) for k in keys if k in after)
        release = max(0.0, release)
    if rec.get("release_outside") and anchor.out_cancel and "content" in rec:
        # Let go outside the window: nothing is moved (the contents compared, the view may have scrolled).
        release = 0.0 if rec["content"][0] == rec["content"][1] else 999.0
    elif rec.get("release_outside") and anchor.out_cancel and after is not None:
        # Let go outside the window: Office puts the value back.
        release = max(abs(diff(after[k], v0[k])) for k in keys if k in after)
    if drag and release is not None:
        # A key whose effect is not modelled ("skip") held at the release, or a copy made by Strg: the dragged value is
        # not where the drag ended (the original stays; Umschalt's insertion), so the release is not compared.
        lm = drag[-1][4]
        held_key = name.endswith("-held") and name.split("-")[0] or ""
        if (lm and mod_spec(anchor, lm) == "skip") or (held_key and anchor.copies.get(held_key)):
            release = None
    ok = bool(drag) and worst_follow <= anchor.tol and worst_jump == 0 and against == 0 and (release is None or release <= anchor.tol)
    out = {"path": name, "ok": ok, "frames": len(drag), "follow": round(worst_follow, 2), "jump": round(worst_jump, 2),
           "against": against, "release": None if release is None else round(release, 2), "v0": rec["v0"], "after": rec["after"]}
    if not drag:
        out["error"] = "no frames with the button down"
    if wf and worst_follow > anchor.tol:
        out["worst_follow"] = wf
    if wj:
        out["worst_jump"] = wj
    if fam == "out":
        out["outside_frames"] = outside_frames
        if not outside_frames:
            out["ok"] = False
            out["error"] = "the pointer never left the window (no frames outside)"
        if anchor.scroll:
            out["autoscrolled"] = round(scrolled_outside, 1)
            if anchor.autoscroll is not None and (scrolled_outside > 2) != anchor.autoscroll:
                out["ok"] = False
                out["error"] = "scrolled while outside" if scrolled_outside > 2 else "did not scroll while outside (Office does)"
        if rec.get("release_outside"):
            # No stuck drag: moving back without a button changes nothing.
            # The value changing between frames while the view did not scroll (a view that scrolls moves what is drawn).
            stuck = 0.0
            for (_t, _x, _y, a1, s1), (_t2, _x2, _y2, a2, s2) in zip(post, post[1:]):
                if s1 == s2:
                    stuck += max(abs(diff(a2[k], a1[k])) for k in keys if k in a1 and k in a2)
            stuck = max(0.0, stuck - max(q(k) for k in keys))
            out["stuck"] = round(stuck, 2)
            out["post_frames"] = len(post)
            if stuck > anchor.tol:
                out["ok"] = False
                out["error"] = "the value moved after the button went up outside (stuck drag)"
    if "count" in rec:
        n0, n1 = rec["count"]
        held = name.endswith("-held") and name.split("-")[0] or ""
        want = anchor.copies.get(held, 0) if held else 0
        out["copies"] = (n1 or 0) - (n0 or 0)
        if out["copies"] != want:
            out["ok"] = False
            out["error"] = f"{out['copies']} copies after the release, Office: {want}"
    return out


def summary(anchor: Anchor, results: list[dict]) -> dict:
    return {"anchor": anchor.name, "ok": sum(1 for r in results if r["ok"]), "of": len(results),
            **({"flaky": sum(1 for r in results if r.get("flaky"))} if any(r.get("flaky") for r in results) else {}),
            **({"note": anchor.note} if anchor.note else {}), "drags": results}


def zoomed(anchor: Anchor, z: int) -> Anchor:
    """The anchor at zoom z (%): snapping distances and steps that are in the document grow with it; an angle's
    threshold allowance grows when the handle comes nearer the middle (lower zoom)."""
    import dataclasses
    f = z / 100
    a = dataclasses.replace(anchor)
    if anchor.wrap:
        a.snap = anchor.snap * max(1.0, 1 / f)
        a.per_px = anchor.per_px * max(1.0, 1 / f)
        return a
    a.snap = anchor.snap * max(1.0, f)
    if isinstance(anchor.quantum, dict):
        a.quantum = {k: v * f for k, v in anchor.quantum.items()}
    elif anchor.quantum:
        a.quantum = anchor.quantum * f
    return a


def run_families(side, anchors: list[Anchor], js: Callable[[str], object], prepare: Callable[[Anchor], None],
                 families: list[str], zoom_js: Callable[[int], str] | None = None, reveal: str = "",
                 log: Callable[..., None] = print, touch: Callable[[], None] = lambda: None) -> list[dict]:
    """The new path families (D-127) for every anchor: "out" (out of the window and back, released outside), "mods"
    (Umschalt, Strg, Alt mid-drag) and "zoom" (diagonal and fast-right at ZOOMS, anchors in the document only).
    `prepare(anchor)` resets the app before each drag; `zoom_js(z)` gives JS that sets the zoom (100 for the others)."""
    report = []

    def safe(anchor: Anchor) -> None:
        # Resetting activates the window, which fails while another session's window holds the front: tried again.
        for i in range(6):
            try:
                prepare(anchor)
                return
            except RuntimeError as e:
                if "activate" not in str(e) or i == 5:
                    raise
                time.sleep(1.0)

    for anchor in anchors:
        plan: list[tuple[int | None, Path]] = [(None, p) for p in family_paths(anchor, families)]
        if "zoom" in families and anchor.zoomed and zoom_js:
            plan += [(z, p) for z in ZOOMS for p in zoom_paths() if anchor.paths is None or p.name in anchor.paths]
        res = []
        for z, p in plan:
            try:
                safe(anchor)
            except Exception as e:  # noqa: BLE001
                res.append({"path": p.name, "ok": False, "error": str(e)[:200]})
                continue
            if INPUT != "cdp":
                try:
                    winapi.activate(side.hwnd)
                except RuntimeError:
                    pass
            time.sleep(0.2)
            a = zoomed(anchor, z) if z else anchor
            try:
                rec = run_drag(side, a, p, js, before_start=zoom_js(z or 100) if zoom_js else "", reveal=reveal if z else "")
            except Exception as e:  # noqa: BLE001
                rec = {"path": p.name, "error": f"{type(e).__name__}: {e}"[:200]}
            for _retry in range(4):
                # Other sessions' windows may take the front now and then: the drag is done again.
                if "front" not in rec.get("error", "") and "activate" not in rec.get("error", ""):
                    break
                time.sleep(1.0)
                safe(anchor)
                try:
                    winapi.activate(side.hwnd)
                except RuntimeError:
                    pass
                rec = run_drag(side, a, p, js, before_start=zoom_js(z or 100) if zoom_js else "", reveal=reveal if z else "")
            if _ours(side):
                Input(side).press("esc")
            try:
                r = analyse(a, rec)
            except Exception as e:  # noqa: BLE001
                r = {"path": p.name, "ok": False, "error": f"analysis: {e}"[:200]}
            if not r["ok"] and not getattr(run_families, "no_retry", False):
                # Once more: a drag that fails once and passes then is reported as flaky (the PC is shared, other
                # sessions' windows and builds disturb real input), one that fails twice as failed.
                first = {k: v for k, v in r.items() if k not in ("v0", "after")}
                try:
                    safe(anchor)
                    time.sleep(0.3)
                    rec = run_drag(side, a, p, js, before_start=zoom_js(z or 100) if zoom_js else "", reveal=reveal if z else "")
                    r = analyse(a, rec)
                except Exception as e:  # noqa: BLE001
                    r = {"path": p.name, "ok": False, "error": f"{type(e).__name__}: {e}"[:200]}
                r["first_try"] = first
                if r["ok"]:
                    r["flaky"] = True
            if z:
                r["path"] = f"zoom {z}: {p.name}"
            res.append(r)
            log(anchor.name, r["path"], "ok" if r["ok"] else "FAIL", {k: v for k, v in r.items() if k not in ("path", "ok", "v0", "after")})
            touch()
        if zoom_js:
            js(zoom_js(100))
        report.append(summary(anchor, res))
    return report


def merge_report(path, report: dict) -> dict:
    """A partial run replaces only its anchors' drags (by anchor and path) in the results file."""
    if not path.exists():
        return report
    old = json.loads(path.read_text(encoding="utf-8"))
    new = {a["anchor"]: a for a in report["anchors"]}
    out = []
    for a in old.get("anchors", []):
        if a["anchor"] in new:
            n = new.pop(a["anchor"])
            got = {d["path"] for d in n["drags"]}
            drags = [d for d in a["drags"] if d["path"] not in got] + n["drags"]
            a = dict(n, drags=drags, ok=sum(1 for d in drags if d["ok"]), of=len(drags))
        out.append(a)
    out += list(new.values())
    return dict(old, **{k: v for k, v in report.items() if k != "anchors"}, anchors=out)
