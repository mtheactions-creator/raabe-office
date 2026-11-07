"""Behaviour catalogue for RaabePresent (D-070): every case runs with real keyboard and mouse in a
PowerPoint window of our own and in RaabePresent, and the outcome is compared: the slide's shapes
(place and size in points, rotation, text), the selection as PowerPoint's COM reports it (None,
Shapes with their order, Text with start and length, Slides), the number of slides and the current
slide. Takes over keyboard, mouse and screen; needs the real-input lock (CLAUDE.md).

A case is a setup (layout, shapes made through COM or the app's API) and steps:
    ("click", target, mods)       target: (shape, fx, fy) a point of a shape's box as fractions,
                                  ("slide", xpt, ypt) a point on the slide, ("thumb", n)
    ("dblclick", target)
    ("rclick", target)            right click, then Esc (the menu is not compared)
    ("drag", target, dxpt, dypt, mods)
    ("handle", shape, k, dxpt, dypt, mods)   k: nw n ne e se s sw w rot
    ("keys", combo, n)
    ("type", text)
Shapes are referred to by their index in the setup (0 = the first made; placeholders by "ph0").
"""
from __future__ import annotations

import json
import time
from dataclasses import dataclass, field

from . import winapi

SLIDE_W, SLIDE_H = 960.0, 540.0


@dataclass
class Case:
    name: str
    area: str
    shapes: list[dict]
    steps: list[tuple]
    layout: str = "Leer"
    slides: int = 1
    #: Compare text of these shapes only (others: place and size).
    note: str = ""
    tol: float = 1.5
    ignore: list[str] = field(default_factory=list)


LAYOUT_INDEX = {"Titelfolie": 1, "Titel und Inhalt": 2, "Nur Titel": 6, "Leer": 7}
ZOOM = 66


def _mods(m: str) -> str:
    return "+".join(x for x in (m or "").split("+") if x)


class Side:
    name = ""

    def slide_point(self, xpt: float, ypt: float) -> tuple[float, float]:
        raise NotImplementedError

    def shape_box(self, idx) -> tuple[float, float, float, float]:
        raise NotImplementedError

    def target(self, t) -> tuple[float, float]:
        if t[0] == "slide":
            return self.slide_point(t[1], t[2])
        if t[0] == "thumb":
            return self.thumb_point(t[1])
        x, y, w, h = self.shape_box(t[0])
        return self.slide_point(x + w * t[1], y + h * t[2])

    def thumb_point(self, n: int) -> tuple[float, float]:
        raise NotImplementedError

    def handle_point(self, idx, k: str) -> tuple[float, float]:
        raise NotImplementedError

    def run(self, case: Case) -> None:
        for st in case.steps:
            kind = st[0]
            if kind == "click":
                x, y = self.target(st[1])
                winapi.click(x, y, 1, _mods(st[2] if len(st) > 2 else ""))
            elif kind == "dblclick":
                x, y = self.target(st[1])
                winapi.click(x, y, 2)
            elif kind == "rclick":
                x, y = self.target(st[1])
                winapi.right_click(x, y)
                time.sleep(0.4)
                winapi.press("esc")
            elif kind in ("drag", "handle"):
                if kind == "drag":
                    x, y = self.target(st[1])
                    dx, dy, mods = st[2], st[3], st[4] if len(st) > 4 else ""
                else:
                    x, y = self.handle_point(st[1], st[2])
                    dx, dy, mods = st[3], st[4], st[5] if len(st) > 5 else ""
                x2, y2 = self.slide_point(0, 0)
                x3, y3 = self.slide_point(dx, dy)
                ex, ey = x + (x3 - x2), y + (y3 - y2)
                vks = [winapi.vk_for(m) for m in _mods(mods).split("+") if m]
                winapi.move(x, y)
                time.sleep(0.05)
                winapi.mouse_down(x, y)
                time.sleep(0.1)
                for vk in vks:
                    winapi.key_down(vk)
                for i in range(1, 21):
                    winapi.move(x + (ex - x) * i / 20, y + (ey - y) * i / 20)
                    time.sleep(0.015)
                time.sleep(0.15)
                winapi.mouse_up(ex, ey)
                time.sleep(0.05)
                for vk in reversed(vks):
                    winapi.key_up(vk)
            elif kind == "keys":
                for _ in range(st[2] if len(st) > 2 else 1):
                    winapi.press(st[1])
                    time.sleep(0.06)
            elif kind == "type":
                winapi.type_text(st[1])
            elif kind == "wait":
                time.sleep(st[1])
            time.sleep(0.25)


# ----- PowerPoint ----------------------------------------------------------------------------------

class PowerPointSide(Side):
    name = "powerpoint"

    def __init__(self):
        from .powerpoint import PowerPointTarget
        self.t = PowerPointTarget(visible=True)
        self.pres = None
        self.made: list = []

    @property
    def app(self):
        return self.t.app

    def reset(self, case: Case) -> None:
        if self.pres is None:
            from .powerpoint import THEME_2013
            self.pres = self.t.new_presentation(with_window=True, theme=THEME_2013, name="catalogue")
        pres = self.pres
        self.esc()
        while pres.Slides.Count:
            pres.Slides(1).Delete()
        lay = pres.SlideMaster.CustomLayouts(LAYOUT_INDEX.get(case.layout, 7))
        for i in range(case.slides):
            pres.Slides.AddSlide(i + 1, lay)
        sl = pres.Slides(1)
        self.made = []
        phs = [sl.Shapes(i) for i in range(1, sl.Shapes.Count + 1)]
        for sh in case.shapes:
            k, b = sh["kind"], sh.get("box", [0, 0, 100, 100])
            if k == "ph":
                s = phs[sh.get("ph", 0)]
                if sh.get("text"):
                    s.TextFrame.TextRange.Text = sh["text"]
                continue
            if k == "textbox":
                s = sl.Shapes.AddTextbox(1, *b)
                s.TextFrame.WordWrap = -1
            elif k == "table":
                s = sl.Shapes.AddTable(2, 2, b[0], b[1], b[2])
            elif k == "line":
                s = sl.Shapes.AddLine(b[0], b[1], b[0] + b[2], b[1] + b[3])
            else:
                s = sl.Shapes.AddShape(9 if k == "ellipse" else 1, *b)
            if sh.get("text") and k not in ("table", "line"):
                s.TextFrame.TextRange.Text = sh["text"]
            self.made.append(s)
        win = self.app.ActiveWindow
        win.ViewType = 9  # ppViewNormal
        win.View.GotoSlide(1)
        win.View.Zoom = ZOOM
        win.Selection.Unselect()
        winapi.activate(self.t.hwnd)
        # The slide area gets the focus (a click outside the slide on the grey area).
        x, y = self.slide_point(SLIDE_W + 20, SLIDE_H / 2)
        winapi.click(x, y)
        time.sleep(0.3)
        self.app.ActiveWindow.Selection.Unselect()

    def esc(self) -> None:
        for _ in range(3):
            try:
                _ = self.app.ActiveWindow.Selection.Type
                return
            except Exception:  # noqa: BLE001
                winapi.press("esc")
                time.sleep(0.3)

    def slide_point(self, xpt: float, ypt: float) -> tuple[float, float]:
        win = self.app.ActiveWindow
        return float(win.PointsToScreenPixelsX(xpt)), float(win.PointsToScreenPixelsY(ypt))

    def _shape(self, idx):
        sl = self.app.ActiveWindow.View.Slide
        if isinstance(idx, str) and idx.startswith("ph"):
            return sl.Shapes(int(idx[2:]) + 1)
        n_ph = sum(1 for i in range(1, sl.Shapes.Count + 1) if int(sl.Shapes(i).Type) == 14)
        return sl.Shapes(n_ph + idx + 1)

    def shape_box(self, idx):
        s = self._shape(idx)
        return float(s.Left), float(s.Top), float(s.Width), float(s.Height)

    def handle_point(self, idx, k):
        x, y, w, h = self.shape_box(idx)
        if k == "rot":
            px, py = self.slide_point(x + w / 2, y)
            # PowerPoint's rotation handle above the frame: measured by probe_powerpoint_rothandle.py (a drag from
            # 38 to 56 device px above the top edge rotates at 66 % zoom and 150 % scaling; the middle of that).
            return px, py - 45
        fx = {"nw": 0, "n": .5, "ne": 1, "e": 1, "se": 1, "s": .5, "sw": 0, "w": 0}[k]
        fy = {"nw": 0, "n": 0, "ne": 0, "e": .5, "se": 1, "s": 1, "sw": 1, "w": .5}[k]
        return self.slide_point(x + w * fx, y + h * fy)

    def thumb_point(self, n: int):
        raise NotImplementedError("thumbnails: through UI Automation (not built yet)")

    def state(self) -> dict:
        self.esc()
        win = self.app.ActiveWindow
        sl = win.View.Slide
        shapes = []
        for i in range(1, sl.Shapes.Count + 1):
            s = sl.Shapes(i)
            text = None
            try:
                if s.HasTextFrame:
                    text = str(s.TextFrame.TextRange.Text).replace("\x0b", "\v")
            except Exception:  # noqa: BLE001
                pass
            shapes.append({"id": int(s.Id), "box": [round(float(s.Left), 1), round(float(s.Top), 1), round(float(s.Width), 1), round(float(s.Height), 1)],
                           "rot": round(float(s.Rotation), 1), "text": text})
        sel = self.t.selection()
        names = {str(sl.Shapes(i).Name): int(sl.Shapes(i).Id) for i in range(1, sl.Shapes.Count + 1)}
        out_sel: dict = {"type": sel["type"]}
        if "shapes" in sel:
            out_sel["ids"] = [names.get(n, n) for n in sel["shapes"]]
        if sel["type"] == "Text":
            out_sel["start"], out_sel["length"] = sel.get("text_start"), sel.get("text_length")
        return {"slides": int(self.pres.Slides.Count), "slide": int(sl.SlideIndex), "shapes": shapes, "selection": out_sel}

    def close(self) -> None:
        self.t.close(kill_if_unreachable=True)


# ----- RaabePresent ---------------------------------------------------------------------------------

class PresentSide(Side):
    name = "present"

    def __init__(self, port: int = 9391):
        import os
        import subprocess
        import tempfile
        import urllib.request
        from pathlib import Path

        import websocket
        from .targets import ROOT, WINDOW_RECT, _pid
        exe = Path(os.environ.get("RAABE_PRESENT_EXE") or ROOT / "apps/raabe-present/src-tauri/target/release/raabe-present.exe")
        rec = Path(tempfile.gettempdir()) / "raabe-present-feel-recovery"
        for f in rec.glob("*") if rec.exists() else []:
            f.unlink(missing_ok=True)
        env = dict(os.environ, RAABE_FEEL_DEBUG_PORT=str(port))
        self.proc = subprocess.Popen([str(exe), "--feel", "present"], env=env)
        self.hwnd = winapi.wait_for_window(lambda h: winapi.window_title(h).endswith(" - RaabePresent") and _pid(h) == self.proc.pid)
        winapi.place(self.hwnd, *WINDOW_RECT)
        deadline = time.perf_counter() + 30
        while True:
            try:
                with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/list", timeout=1) as r:
                    pages = [p for p in json.load(r) if p.get("type") == "page"]
                if pages:
                    self.ws = websocket.create_connection(pages[0]["webSocketDebuggerUrl"], suppress_origin=True)
                    break
            except OSError:
                pass
            if time.perf_counter() > deadline:
                raise TimeoutError("DevTools endpoint not reachable")
            time.sleep(0.1)
        self._id = 0
        while not self.js("typeof window.__present === 'object'"):
            time.sleep(0.05)
        self.ids: list[int] = []

    def js(self, expression: str):
        self._id += 1
        self.ws.send(json.dumps({"id": self._id, "method": "Runtime.evaluate", "params": {"expression": expression, "returnByValue": True, "awaitPromise": True}}))
        while True:
            msg = json.loads(self.ws.recv())
            if msg.get("id") == self._id:
                res = msg["result"]
                if "exceptionDetails" in res:
                    raise RuntimeError(res["exceptionDetails"])
                return res["result"].get("value")

    def _origin(self):
        for child in winapi.child_windows(self.hwnd):
            if winapi.window_class(child) == "Chrome_RenderWidgetHostHWND":
                ox, oy = winapi.client_origin(child)
                return ox, oy, self.js("devicePixelRatio")
        ox, oy = winapi.client_origin(self.hwnd)
        return ox, oy, self.js("devicePixelRatio")

    def reset(self, case: Case) -> None:
        spec = {"layout": case.layout, "slides": case.slides, "shapes": case.shapes, "zoom": ZOOM}
        self.js(f"__present.setupCase({json.dumps(spec)})")
        time.sleep(0.2)
        winapi.activate(self.hwnd)
        self.origin = self._origin()
        st = self.js("__present.state()")
        self.all_ids = [s["id"] for s in st["shapes"]]
        self.phs = [s["id"] for s in st["shapes"] if s["ph"]]
        self.ids = [s["id"] for s in st["shapes"] if not s["ph"]]
        x, y = self.slide_point(SLIDE_W + 20, SLIDE_H / 2)
        winapi.click(x, y)
        time.sleep(0.2)

    def _client(self, x: float, y: float) -> tuple[float, float]:
        ox, oy, dpr = self.origin
        return ox + x * dpr, oy + y * dpr

    def slide_point(self, xpt: float, ypt: float):
        p = self.js(f"__present.slidePointToClient({xpt}, {ypt})")
        return self._client(p["x"], p["y"])

    def _sid(self, idx) -> int:
        if isinstance(idx, str) and idx.startswith("ph"):
            return self.phs[int(idx[2:])]
        # After grouping the setup's ids are gone: the shape at that place now (as PowerPoint's index).
        now = [s["id"] for s in self.js("__present.state()")["shapes"] if not s["ph"]]
        return self.ids[idx] if self.ids[idx] in now else now[idx]

    def shape_box(self, idx):
        sid = self._sid(idx)
        s = next(s for s in self.js("__present.state()")["shapes"] if s["id"] == sid)
        return s["x"] / 12700, s["y"] / 12700, s["w"] / 12700, s["h"] / 12700

    def handle_point(self, idx, k):
        p = self.js(f"__present.handlePoint({self._sid(idx)}, {json.dumps(k)})")
        if p is None and k != "rot":
            # No handle there (a short side): the same point on the frame as on PowerPoint's side.
            x, y, w, h = self.shape_box(idx)
            fx = {"nw": 0, "n": .5, "ne": 1, "e": 1, "se": 1, "s": .5, "sw": 0, "w": 0}[k]
            fy = {"nw": 0, "n": 0, "ne": 0, "e": .5, "se": 1, "s": 1, "sw": 1, "w": .5}[k]
            return self.slide_point(x + w * fx, y + h * fy)
        if p is None:
            raise RuntimeError(f"no handle {k}")
        return self._client(p["x"], p["y"])

    def thumb_point(self, n: int):
        r = self.js(f"__present.thumbRect({n})")
        return self._client(r["x"] + r["w"] / 2, r["y"] + r["h"] / 2)

    def state(self) -> dict:
        st = self.js("__present.state()")
        sel = self.js("__present.comSelection()")
        shapes = [{"id": s["id"], "box": [round(s["x"] / 12700, 1), round(s["y"] / 12700, 1), round(s["w"] / 12700, 1), round(s["h"] / 12700, 1)],
                   "rot": round(s["rot"], 1), "text": s["text"]} for s in st["shapes"]]
        out_sel = {"type": sel["type"]}
        if "ids" in sel:
            out_sel["ids"] = sel["ids"]
        if sel["type"] == "Text":
            out_sel["start"], out_sel["length"] = sel["start"], sel["length"]
        return {"slides": st["slides"], "slide": st["slide"], "shapes": shapes, "selection": out_sel}

    def close(self) -> None:
        try:
            self.ws.close()
        except Exception:  # noqa: BLE001
            pass
        self.proc.kill()


# ----- comparing ----------------------------------------------------------------------------------

def normalise(st: dict, tol: float) -> dict:
    """Ids become positions in z-order (the two programs number shapes alike only by chance)."""
    order = {s["id"]: i for i, s in enumerate(st["shapes"])}
    sel = dict(st["selection"])
    if "ids" in sel:
        sel["ids"] = [order.get(i, i) for i in sel["ids"]]
    return {"slides": st["slides"], "slide": st["slide"], "selection": sel,
            "shapes": [{"box": s["box"], "rot": s["rot"], "text": (s["text"] or "").replace("\n", "\v") if s["text"] is not None else None} for s in st["shapes"]]}


def differences(a: dict, b: dict, tol: float, ignore: list[str]) -> list[str]:
    out = []
    for k in ("slides", "slide"):
        if a[k] != b[k] and k not in ignore:
            out.append(f"{k}: {a[k]} vs {b[k]}")
    if a["selection"] != b["selection"] and "selection" not in ignore:
        out.append(f"selection: {a['selection']} vs {b['selection']}")
    if len(a["shapes"]) != len(b["shapes"]):
        out.append(f"shapes: {len(a['shapes'])} vs {len(b['shapes'])}")
    for i, (x, y) in enumerate(zip(a["shapes"], b["shapes"])):
        if "box" not in ignore and any(abs(p - q) > tol for p, q in zip(x["box"], y["box"])):
            out.append(f"shape {i} box: {x['box']} vs {y['box']}")
        if abs(x["rot"] - y["rot"]) > 0.5 and "rot" not in ignore:
            out.append(f"shape {i} rot: {x['rot']} vs {y['rot']}")
        if "text" not in ignore and (x["text"] or "") != (y["text"] or ""):
            out.append(f"shape {i} text: {x['text']!r} vs {y['text']!r}")
    return out
