"""Behaviour catalogue for RaabeCalc (D-061): every case runs with real keyboard and mouse in an
Excel window of our own and in RaabeCalc, and the outcome is compared: the cells (formula and
what they show), the selection, the active cell and the cell mode (Bereit, Eingeben,
Bearbeiten, Zeigen). Takes over keyboard, mouse and screen for the whole run.

A case is a setup (cells typed in through COM or the app's API, the start selection) and a
list of steps:
    ("type", "abc")            characters as typed
    ("keys", "ctrl+down", n)   a key combination, n times
    ("click", "B3", mods)      a click in the middle of a cell (mods: "shift", "ctrl")
    ("dblclick", "B3")
    ("drag", "A1", "C3")       press in A1, move to C3, release
    ("fill", "A1:A2", "A6")    drag the fill handle of A1:A2 to A6 (mods: "ctrl")
    ("border", "A1", "C3")     drag A1's selection border to C3 (moves the cells)
    ("colhead", "B", mods)     click a column header; ("rowhead", 3, mods)
    ("colborder", "B")         double-click B's right header border (AutoAnpassen)
"""

from __future__ import annotations

import re
import time
from dataclasses import dataclass, field

from . import uia, winapi


@dataclass
class Case:
    name: str
    setup: dict[str, str]
    steps: list[tuple]
    area: str
    start: str = "A1"
    #: Cells to compare besides the used range.
    check: list[str] = field(default_factory=list)
    #: Extra sheets to add before the case (names).
    sheets: int = 1
    #: Today's date or time in the result: compared by format only.
    today: bool = False
    #: Ranges merged before the steps (Verbinden), e.g. ["A1:C1"].
    merges: list[str] = field(default_factory=list)


def addr_rc(addr: str) -> tuple[int, int]:
    m = re.fullmatch(r"([A-Z]+)(\d+)", addr)
    if not m:
        raise ValueError(f"not a cell address: {addr!r}")
    col = 0
    for ch in m.group(1):
        col = col * 26 + ord(ch) - 64
    return int(m.group(2)), col


# ----- Excel ---------------------------------------------------------------------------------

class ExcelSide:
    """Runs cases in a visible Excel of our own (real input)."""

    def __init__(self):
        from .excel import ExcelTarget
        self.x = ExcelTarget(visible=True)
        wb = self.x.wb
        # The same default font as RaabeCalc (D-061): Calibri 11 gives Excel 2013-2022's grid.
        st = wb.Styles("Standard")
        st.Font.Name = "Calibri"
        st.Font.Size = 11
        self.x.app.ActiveWindow.Zoom = 100
        self.root = uia.element_for(self.x.hwnd)

    def reset(self, case: Case) -> None:
        app = self.x.app
        wb = self.x.wb
        while wb.Worksheets.Count > 1:
            app.DisplayAlerts = False
            wb.Worksheets(wb.Worksheets.Count).Delete()
        while wb.Worksheets.Count < case.sheets:
            wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        ws = wb.Worksheets(1)
        ws.Name = "Tabelle1"
        ws.Activate()
        ws.Cells.Clear()
        ws.Cells.UseStandardWidth = True
        ws.Cells.UseStandardHeight = True
        ws.Columns.Hidden = False
        ws.Rows.Hidden = False
        if ws.AutoFilterMode:
            ws.AutoFilterMode = False
        for addr, text in case.setup.items():
            ws.Range(addr).FormulaLocal = text
        for m in case.merges:
            ws.Range(m).Merge()
        ws.Range(case.start).Select()
        app.ActiveWindow.Zoom = 100
        app.ActiveWindow.ScrollRow = 1
        app.ActiveWindow.ScrollColumn = 1
        app.CutCopyMode = False

    def prepare(self, case: Case) -> None:
        """Cell positions are read before the steps: Excel refuses COM calls while a cell is edited."""
        self.cache = {}
        for st in case.steps:
            for a in st[1:]:
                if isinstance(a, str) and re.fullmatch(r"[A-Z]+\d+(:[A-Z]+\d+)?", a):
                    last = a.split(":")[-1]
                    for where in ("center", "br", "right", "bottom"):
                        self.cache[(last, where)] = self._point(last, where)
                        self.cache[(a.split(":")[0], where)] = self._point(a.split(":")[0], where)

    def point(self, addr: str, where: str = "center") -> tuple[float, float]:
        hit = getattr(self, "cache", {}).get((addr, where))
        return hit if hit else self._point(addr, where)

    def _point(self, addr: str, where: str = "center") -> tuple[float, float]:
        win = self.x.app.ActiveWindow
        r = self.x.sheet.Range(addr)
        left, top, w, hgt = r.Left, r.Top, r.Width, r.Height
        if where == "center":
            px, py = left + w / 2, top + hgt / 2
        elif where == "br":
            px, py = left + w, top + hgt
        elif where == "right":
            px, py = left + w, top + hgt / 2
        elif where == "bottom":
            px, py = left + w / 2, top + hgt
        else:
            px, py = left, top
        return win.PointsToScreenPixelsX(0) + _pt_to_px(px), win.PointsToScreenPixelsY(0) + _pt_to_px(py)

    def header_point(self, kind: str, which) -> tuple[float, float]:
        win = self.x.app.ActiveWindow
        if kind == "col":
            r = self.x.sheet.Range(f"{which}1")
            x = win.PointsToScreenPixelsX(0) + _pt_to_px(r.Left + r.Width / 2)
            y = win.PointsToScreenPixelsY(0) - 10 * 1.5
            return x, y
        if kind == "colborder":
            r = self.x.sheet.Range(f"{which}1")
            return win.PointsToScreenPixelsX(0) + _pt_to_px(r.Left + r.Width), win.PointsToScreenPixelsY(0) - 10 * 1.5
        r = self.x.sheet.Range(f"A{which}")
        return win.PointsToScreenPixelsX(0) - 14 * 1.5, win.PointsToScreenPixelsY(0) + _pt_to_px(r.Top + r.Height / 2)

    def ui_point(self, kind: str, arg=None) -> tuple[float, float]:
        if kind == "sheettab":
            el = uia.find(self.root, lambda d: d["type"] == "TabItem" and d["name"] == arg, max_depth=10)
        elif kind == "sheetadd":
            el = uia.find(self.root, lambda d: d["type"] == "Button" and d["name"] == "Blatt hinzufügen", max_depth=10)
        else:
            el = uia.find(self.root, lambda d: d["type"] == "Edit" and d["name"] == "Namenfeld", max_depth=10)
        r = uia.describe(el)["rect"]
        return r[0] + r[2] / 2, r[1] + r[3] / 2

    def status(self) -> list[str]:
        bar = uia.find(self.root, lambda d: d["type"] == "StatusBar", max_depth=6)
        out = []
        for c in uia.children(bar) if bar else []:
            n = uia.describe(c)["name"]
            for key in ("Mittelwert", "Anzahl", "Summe"):
                if n.startswith(key):
                    out.append(n.replace(" ", "").replace(":", ""))
        return out

    def flags(self) -> list[str]:
        """Texts the status bar shows for modes beside the cell mode (Ende-Modus, Markierung erweitern ...)."""
        bar = uia.find(self.root, lambda d: d["type"] == "StatusBar", max_depth=6)
        out = []
        for c in uia.children(bar) if bar else []:
            d = uia.describe(c)
            if d["type"] == "Text" and d["name"].startswith("Statusmeldung"):
                # Only that there is one: its text is Excel's own prose (D-058).
                out.append("Statusmeldung")
            elif d["type"] == "Text" and d["name"] and not d["name"].startswith(("Zellenmodus", "Zoom", "Mittelwert", "Anzahl", "Summe", "Numerische", "Min", "Max")):
                out.append(d["name"])
        return out

    def mode(self) -> str:
        el = uia.find(self.root, lambda d: d["type"] == "Text" and d["name"].startswith("Zellenmodus"), max_depth=8)
        return el.CurrentName.replace("Zellenmodus", "").strip() if el else "?"

    def bar_text(self) -> str:
        el = uia.find(self.root, lambda d: d["type"] == "Edit" and d["name"] == "Bearbeitungsleiste", max_depth=10)
        if not el:
            return ""
        try:
            p = el.GetCurrentPattern(uia.U.UIA_TextPatternId).QueryInterface(uia.U.IUIAutomationTextPattern)
            return p.DocumentRange.GetText(-1)
        except Exception:  # noqa: BLE001
            return ""

    def state(self, case: Case) -> dict:
        mode = self.mode()
        editing_text = None
        if mode not in ("Bereit", "?"):
            editing_text = self.bar_text()
            winapi.press("esc")
            time.sleep(0.2)
        app = self.x.app
        ws = app.ActiveSheet
        cells = {}
        used = ws.UsedRange
        addrs = set(case.check)
        for c in used.Cells:
            # In Markierung erweitern (F8) Excel's COM gives single cells as "A1:A1".
            ad = c.GetAddress(False, False).split(":")[0]
            if c.Formula != "" or ad in addrs:
                addrs.add(ad)
        for a in sorted(addrs, key=lambda s: addr_rc(s)):
            r = ws.Range(a)
            cells[a] = {"formula": r.FormulaLocal, "text": r.Text, "format": r.NumberFormatLocal,
                        "bold": bool(r.Font.Bold), "italic": bool(r.Font.Italic), "underline": r.Font.Underline != -4142,
                        "strike": bool(r.Font.Strikethrough), "wrap": bool(r.WrapText)}
        hidden_rows = [i for i in range(1, 30) if ws.Rows(i).Hidden]
        hidden_cols = [i for i in range(1, 15) if ws.Columns(i).Hidden]
        std = ws.StandardWidth
        widths = {c: round(ws.Columns(c).ColumnWidth, 2) for c in range(1, 8) if abs(ws.Columns(c).ColumnWidth - std) > 0.01}
        status = self.status()
        flags = self.flags()
        win = app.ActiveWindow
        return {
            "flags": flags, "status": status, "scroll": [win.ScrollRow, win.ScrollColumn], "zoom": win.Zoom,
            "mode": mode, "editing": editing_text, "cells": cells, "selection": app.Selection.GetAddress(False, False),
            "active": app.ActiveCell.GetAddress(False, False).split(":")[0], "sheet": ws.Name, "sheets": [s.Name for s in app.ActiveWorkbook.Worksheets],
            "hiddenRows": hidden_rows, "hiddenCols": hidden_cols, "widths": widths, "filter": bool(ws.AutoFilterMode),
        }

    def close(self) -> None:
        self.x.close()


def _pt_to_px(pt: float) -> float:
    return pt * 96 / 72 * 1.5


def wheel(x: float, y: float, notches: int, ctrl: bool = False, mod: str = "") -> None:
    """Real mouse wheel notches at a point (positive: away from the user, i.e. up)."""
    import ctypes
    winapi.move(x, y)
    time.sleep(0.05)
    if ctrl:
        mod = "ctrl"
    if mod:
        winapi.key_down(winapi.vk_for(mod))
    for _ in range(abs(notches)):
        inp = winapi.INPUT(type=winapi.INPUT_MOUSE, u=winapi._INPUTUNION(mi=winapi.MOUSEINPUT(dx=0, dy=0, mouseData=ctypes.c_ulong(120 if notches > 0 else (-120) & 0xFFFFFFFF).value, dwFlags=0x0800, time=0, dwExtraInfo=0)))
        winapi._send(inp)
        time.sleep(0.12)
    if mod:
        winapi.key_up(winapi.vk_for(mod))


def settle() -> None:
    """Between cases: the mouse away and a pause longer than the double-click time, so the next
    case's first click is never counted as the second click of the last one."""
    winapi.move(1900, 1950)
    time.sleep(0.6)


def run_step(side, step: tuple) -> None:
    kind = step[0]
    if kind == "type":
        winapi.type_text(step[1], gap=0.02)
    elif kind == "keys":
        for _ in range(step[2] if len(step) > 2 else 1):
            winapi.press(step[1])
            time.sleep(0.06)
    elif kind in ("click", "dblclick"):
        x, y = side.point(step[1])
        mods = step[2] if len(step) > 2 else ""
        winapi.click(x, y, 2 if kind == "dblclick" else 1, mods)
    elif kind == "drag":
        x1, y1 = side.point(step[1])
        x2, y2 = side.point(step[2])
        winapi.drag(x1, y1, x2, y2)
    elif kind == "fill":
        last = step[1].split(":")[-1]
        x1, y1 = side.point(last, "br")
        x2, y2 = side.point(step[2])
        mods = step[3] if len(step) > 3 else ""
        if mods:
            winapi.key_down(winapi.vk_for(mods))
        winapi.drag(x1 - 1, y1 - 1, x2, y2)
        if mods:
            winapi.key_up(winapi.vk_for(mods))
    elif kind == "fillclick":
        x1, y1 = side.point(step[1], "br")
        winapi.click(x1 - 1, y1 - 1, 2)
    elif kind == "border":
        x1, y1 = side.point(step[1], "right")
        # "A4@bottom": drop on that edge of the cell (Umschalt+drag picks the nearest gridline).
        addr, where = (step[2].split("@") + ["right"])[:2]
        x2, y2 = side.point(addr, where)
        if where == "bottom":
            x2 += 1
        mods = step[3] if len(step) > 3 else ""
        if mods:
            winapi.key_down(winapi.vk_for(mods))
        winapi.drag(x1 - 1, y1, x2 - 1, y2)
        if mods:
            winapi.key_up(winapi.vk_for(mods))
    elif kind in ("rclick", "rcolhead", "rrowhead"):
        # Right-click, then Esc closes the context menu; what is left is the selection it made.
        x, y = side.point(step[1]) if kind == "rclick" else side.header_point("col" if kind == "rcolhead" else "row", step[1])
        winapi.right_click(x, y)
        time.sleep(0.6)
        winapi.press("esc")
        time.sleep(0.3)
    elif kind == "colhead":
        x, y = side.header_point("col", step[1])
        winapi.click(x, y, 1, step[2] if len(step) > 2 else "")
    elif kind == "rowhead":
        x, y = side.header_point("row", step[1])
        winapi.click(x, y, 1, step[2] if len(step) > 2 else "")
    elif kind == "colborder":
        x, y = side.header_point("colborder", step[1])
        winapi.click(x, y, 2)
    elif kind == "colresize":
        x, y = side.header_point("colborder", step[1])
        winapi.drag(x - 1, y, x - 1 + step[2], y)
    elif kind in ("tabdbl", "tabclick"):
        x, y = side.ui_point("sheettab", step[1])
        winapi.click(x, y, 2 if kind == "tabdbl" else 1)
    elif kind == "sheetadd":
        x, y = side.ui_point("sheetadd")
        winapi.click(x, y)
    elif kind == "namebox":
        x, y = side.ui_point("namebox")
        winapi.click(x, y)
        time.sleep(0.15)
        winapi.press("ctrl+a")
        winapi.type_text(step[1], gap=0.02)
        winapi.press("enter")
    elif kind in ("wheel", "ctrlwheel", "shiftwheel"):
        x, y = side.point(step[1])
        wheel(x, y, step[2], kind == "ctrlwheel", "shift" if kind == "shiftwheel" else "")
    elif kind == "borderdbl":
        # Double-click on the selection's border towards a direction (Strg+arrow in Excel).
        x, y = side.point(step[1], step[2])
        winapi.click(x - (1 if step[2] == "right" else 0), y - (1 if step[2] == "bottom" else 0), 2)
    elif kind == "wait":
        time.sleep(step[1])
    else:
        raise ValueError(kind)
    time.sleep(0.12)


# ----- RaabeCalc -------------------------------------------------------------------------------

class CalcSide:
    """Runs cases in RaabeCalc (real input on its window; setup and readout through DevTools)."""

    def __init__(self, port: int = 9381):
        import json as _json
        import os
        import subprocess
        import urllib.request
        from pathlib import Path

        import websocket
        from .targets import ROOT, WINDOW_RECT, _pid
        exe = Path(os.environ.get("RAABE_CALC_EXE") or ROOT / "apps/raabe-calc/src-tauri/target/release/raabe-calc.exe")
        # Test runs end the app abruptly; their AutoRecover backups must not reach the next run.
        import tempfile
        rec = Path(tempfile.gettempdir()) / "raabe-calc-feel-recovery"
        for f in rec.glob("*") if rec.exists() else []:
            f.unlink(missing_ok=True)
        env = dict(os.environ, RAABE_FEEL_DEBUG_PORT=str(port))
        self.proc = subprocess.Popen([str(exe), "--feel", "calc"], env=env)
        self.hwnd = winapi.wait_for_window(lambda h: winapi.window_title(h).endswith(" - RaabeCalc") and _pid(h) == self.proc.pid)
        winapi.place(self.hwnd, *WINDOW_RECT)
        deadline = time.perf_counter() + 30
        while True:
            try:
                with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/list", timeout=1) as r:
                    pages = [p for p in _json.load(r) if p.get("type") == "page"]
                if pages:
                    self.ws = websocket.create_connection(pages[0]["webSocketDebuggerUrl"], suppress_origin=True)
                    break
            except OSError:
                pass
            if time.perf_counter() > deadline:
                raise TimeoutError("DevTools endpoint not reachable")
            time.sleep(0.1)
        self._id = 0
        while not self.js("typeof window.__calc === 'object'"):
            time.sleep(0.05)
        self.cache: dict = {}

    def js(self, expression: str):
        import json as _json
        self._id += 1
        self.ws.send(_json.dumps({"id": self._id, "method": "Runtime.evaluate",
                                  "params": {"expression": expression, "returnByValue": True, "awaitPromise": True}}))
        while True:
            msg = _json.loads(self.ws.recv())
            if msg.get("id") == self._id:
                res = msg["result"]
                if "exceptionDetails" in res:
                    raise RuntimeError(res["exceptionDetails"])
                return res["result"].get("value")

    def _origin(self) -> tuple[int, int, float]:
        for child in winapi.child_windows(self.hwnd):
            if winapi.window_class(child) == "Chrome_RenderWidgetHostHWND":
                ox, oy = winapi.client_origin(child)
                return ox, oy, self.js("devicePixelRatio")
        ox, oy = winapi.client_origin(self.hwnd)
        return ox, oy, self.js("devicePixelRatio")

    def reset(self, case: Case) -> None:
        import json as _json
        self.js(f"__calc.reset({_json.dumps(case.setup)}, {_json.dumps(case.start)}, {case.sheets}, {_json.dumps(case.merges)})")
        time.sleep(0.15)

    def prepare(self, case: Case) -> None:
        import json as _json
        ox, oy, dpr = self._origin()
        self.cache = {}
        for st in case.steps:
            for a in st[1:]:
                if isinstance(a, str) and re.fullmatch(r"[A-Z]+\d+(:[A-Z]+\d+)?(@\w+)?", a):
                    for one in a.split("@")[0].split(":"):
                        for where in ("center", "br", "right", "bottom"):
                            p = self.js(f"__calc.point({_json.dumps(one)}, {_json.dumps(where)})")
                            self.cache[(one, where)] = (ox + p["x"] * dpr, oy + p["y"] * dpr)
        self.origin = (ox, oy, dpr)

    def point(self, addr: str, where: str = "center") -> tuple[float, float]:
        return self.cache[(addr, where)]

    def header_point(self, kind: str, which) -> tuple[float, float]:
        import json as _json
        ox, oy, dpr = self.origin
        p = self.js(f"__calc.header({_json.dumps(kind)}, {_json.dumps(str(which))})")
        return ox + p["x"] * dpr, oy + p["y"] * dpr

    def ui_point(self, kind: str, arg=None) -> tuple[float, float]:
        import json as _json
        sel = {"sheettab": None, "sheetadd": ".sheetadd", "namebox": ".namebox"}[kind]
        if kind == "sheettab":
            js = f"[...document.querySelectorAll('.sheettab')].find(t => t.textContent === {_json.dumps(arg)}).getBoundingClientRect()"
        else:
            js = f"document.querySelector('{sel}').getBoundingClientRect()"
        r = self.js(f"(() => {{ const r = {js}; return {{x: r.left + r.width / 2, y: r.top + r.height / 2}}; }})()")
        ox, oy, dpr = self._origin()
        return ox + r["x"] * dpr, oy + r["y"] * dpr

    def state(self, case: Case) -> dict:
        import json as _json
        st = self.js(f"__calc.state({_json.dumps(case.check)})")
        st["status"] = self.js("[...document.querySelectorAll('.statusbar .stat')].map(e => e.textContent.replace(/[ :]/g, ''))")
        st["flags"] = self.js("[...document.querySelectorAll('.statusbar .flag:not([hidden])')].map(e => e.getAttribute('aria-label') || e.textContent)")
        # A status message stands where the cell mode was (Excel then has no Zellenmodus).
        if "Statusmeldung" in st["flags"]:
            st["mode"] = "?"
        st["scroll"] = self.js("[__calc.app.sheet.view.top + 1, __calc.app.sheet.view.left + 1]")
        st["zoom"] = self.js("__calc.app.sheet.view.zoom")
        return st

    def close(self) -> None:
        try:
            self.ws.close()
        finally:
            self.proc.terminate()
            self.proc.wait(5)


FIELDS = ("mode", "editing", "selection", "active", "sheet", "sheets", "hiddenRows", "hiddenCols", "filter", "status", "flags", "scroll", "zoom")
CELL_FIELDS = ("formula", "text", "format", "bold", "italic", "underline", "strike", "wrap")


def diff(excel: dict, calc: dict, today: bool = False) -> list[str]:
    out = []
    for f in FIELDS:
        if f not in excel:
            continue
        if excel.get(f) != calc.get(f):
            out.append(f"{f}: Excel {excel.get(f)!r}, RaabeCalc {calc.get(f)!r}")
    names = sorted(set(excel["cells"]) | set(calc["cells"]))
    for a in names:
        e, c = excel["cells"].get(a), calc["cells"].get(a)
        if e is None or c is None:
            if (e or c) and ((e or c)["formula"] != "" or (e or c)["text"] != ""):
                out.append(f"{a}: Excel {e and (e['formula'], e['text'])!r}, RaabeCalc {c and (c['formula'], c['text'])!r}")
            continue
        for f in CELL_FIELDS:
            if today and f in ("formula", "text"):
                continue
            if e.get(f) != c.get(f):
                out.append(f"{a}.{f}: Excel {e.get(f)!r}, RaabeCalc {c.get(f)!r}")
    # Column widths: only those Excel changed by itself (AutoAnpassen, formatted numbers).
    for col, w in excel.get("widths", {}).items():
        if abs(w - 8.45) < 0.02 or abs(w - 10.27) < 0.02:
            continue
        cw = calc.get("widths", {}).get(col)
        if cw is None or abs(cw - w) > 0.6:
            out.append(f"width {col}: Excel {w}, RaabeCalc {cw}")
    return out


def compare_calc(cases: list[Case], out, date: str) -> None:
    import json as _json
    import os
    # Excel's outcomes: all recordings, later ones replacing earlier ones (cases are recorded on different days).
    excel: dict = {}
    for f in sorted(out.glob("*-catalogue-calc-excel.json")):
        excel.update(_json.loads(f.read_text(encoding="utf-8")))
    side = CalcSide()
    results = {}
    try:
        for c in cases:
            if c.name not in excel:
                continue
            side.reset(c)
            side.prepare(c)
            settle()
            winapi.activate(side.hwnd)
            time.sleep(0.2)
            side.js("__calc.grid.focus()")
            for s in c.steps:
                run_step(side, s)
                if os.environ.get("CATDEBUG"):
                    print("   ", s, side.js("JSON.stringify([__calc.state([]).selection, __calc.app.grid.mode, document.activeElement.className])"))
            time.sleep(0.3)
            st = side.state(c)
            d = diff(excel[c.name], st, c.today)
            results[c.name] = {"area": c.area, "ok": not d, "diff": d}
            print(f"{'ok ' if not d else 'XX '} {c.area:10} {c.name:45} {'; '.join(d)[:160]}")
    finally:
        side.close()
    ok = sum(1 for r in results.values() if r["ok"])
    print(f"{ok}/{len(results)} cases as in Excel")
    path = out / f"{date}-catalogue-calc.json"
    path.write_text(_json.dumps({"summary": f"{ok}/{len(results)}", "cases": results}, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
