"""The chart Format pane with the real mouse and keyboard (D-113): what only OS input shows (the slider's
drag, the spin buttons, the colour palette's popup, the native dropdown, typing into a box, a gradient
stop dragged along its bar, Strg+Z from the keyboard). Runs only while holding %TEMP%\\raabe-real-input.lock;
before every input it checks that RaabeCalc's own window is in front.

    .venv/Scripts/python.exe check_calc_chartpane_real.py

Result: results/2026-11-09-calc-chartpane-real.json.
"""
from __future__ import annotations

import datetime
import json
import os
import tempfile
import time
from pathlib import Path

from feellab import winapi
from feellab.calc_catalogue import Case, CalcSide

HERE = Path(__file__).parent
LOCK = Path(tempfile.gettempdir()) / "raabe-real-input.lock"
DATA = {"A1": "Monat", "B1": "A", "C1": "B", "A2": "Jan", "B2": "3", "C2": "40", "A3": "Feb", "B3": "-1", "C3": "20", "A4": "Mrz", "B4": "5", "C4": "60",
        "A5": "Apr", "B5": "2", "C5": "30"}


class Real:
    def __init__(self, c: CalcSide):
        self.c = c
        self.results: list[dict] = []

    def ours(self) -> bool:
        return winapi.window_pid(winapi.foreground()) == self.c.proc.pid

    def front(self) -> None:
        if not self.ours():
            winapi.activate(self.c.hwnd)
        if not self.ours():
            raise RuntimeError("RaabeCalc is not in front; no input sent")

    def at(self, css: str, where: str = "center") -> tuple[float, float]:
        """Screen point of an element (a selector evaluated in the page), scrolled into view."""
        r = self.c.js(f"(() => {{ const e = {css}; if (!e) return null; e.scrollIntoView({{block: 'center'}}); const b = e.getBoundingClientRect(); return [b.left, b.top, b.width, b.height]; }})()")
        if not r:
            raise RuntimeError(f"not found: {css}")
        ox, oy, dpr = self.c._origin()
        x, y, w, h = r
        fx = {"center": 0.5, "left": 0.1, "right": 0.9}[where]
        return ox + (x + w * fx) * dpr, oy + (y + h / 2) * dpr

    def model(self) -> dict:
        return self.c.js("(() => { const m = __calc.app.sheet.charts[0]; const { raw, ...rest } = m; return JSON.parse(JSON.stringify(rest)); })()")

    def step(self, name: str, act, check) -> None:
        before = self.model()
        try:
            self.front()
            act()
            time.sleep(0.35)
            after = self.model()
            ok, detail = check(before, after)
        except Exception as e:  # noqa: BLE001
            ok, detail = False, f"error {str(e)[:120]}"
        self.results.append({"step": name, "ok": bool(ok), "detail": detail})
        print("ok " if ok else "BAD", name, detail, flush=True)


def line_of(m: dict, axis: str = "valAxis") -> dict:
    return ((m.get("format") or {}).get(axis) or {}).get("line") or {}


def control(label: str, sel: str) -> str:
    """The first control `sel` inside the pane line labelled `label`."""
    return (f"[...document.querySelectorAll('.taskpane .tp-line, .taskpane .tp-choice, .taskpane .tp-grad')]"
            f".find((l) => (l.querySelector('.tp-label')?.textContent ?? l.textContent).trim() === {json.dumps(label)})?.querySelector({json.dumps(sel)})")


def main() -> None:
    if LOCK.exists() and time.time() - LOCK.stat().st_mtime < 1800:
        print("real-input lock held by:", LOCK.read_text(encoding="utf-8", errors="replace")[:80], "- try later")
        return
    LOCK.unlink(missing_ok=True)
    with open(LOCK, "x", encoding="utf-8") as f:
        f.write(f"charts-pane real-input pass {datetime.datetime.now().isoformat(timespec='seconds')}")
    c = None
    try:
        c = CalcSide(9419)
        r = Real(c)
        c.reset(Case("chartpane-real", DATA, [], "A1"))
        c.js("__calc.app.grid.select({r0:0,c0:0,r1:4,c1:2}); __calc.app.charts.insert('column'); __calc.app.charts.formatSelection()")
        time.sleep(0.4)
        r.front()
        # The value axis through the pane's own element list (real clicks).
        winapi.click(*r.at("document.querySelector('.taskpane .tp-set')"))
        time.sleep(0.3)
        winapi.click(*r.at("[...document.querySelectorAll('.menu .item')].find((b) => b.textContent === 'Vertikal (Wert) Achse')"))
        time.sleep(0.3)
        el = c.js("__calc.app.charts.elementName()")
        r.results.append({"step": "element list: Vertikal (Wert) Achse", "ok": el == "Vertikal (Wert) Achse", "detail": el})
        winapi.click(*r.at("document.querySelector('.taskpane .tp-tab')"))
        time.sleep(0.2)
        r.step("radio Einfarbige Linie", lambda: winapi.click(*r.at(control("Einfarbige Linie", "input"))),
               lambda b, a: (line_of(a).get("type") == "solid", line_of(a)))
        r.step("colour palette: Farbe, a swatch", lambda: (winapi.click(*r.at(control("Farbe", ".tp-color"))), time.sleep(0.3),
                                                           winapi.click(*r.at("document.querySelectorAll('.menu.palette .swatches button')[13]"))),
               lambda b, a: (line_of(a).get("color") != line_of(b).get("color") and str(line_of(a).get("color", "")).startswith("#"), line_of(a).get("color")))
        r.step("slider Transparenz dragged", lambda: winapi.drag(*r.at(control("Transparenz", ".tp-slider"), "left"), *r.at(control("Transparenz", ".tp-slider"), "right")),
               lambda b, a: ((line_of(a).get("transparency") or 0) > 50, line_of(a).get("transparency")))
        r.step("spin button Stärke Höher", lambda: winapi.click(*r.at(control("Stärke", ".tp-spin button"))),
               lambda b, a: (abs((line_of(a).get("width") or 0) - (line_of(b).get("width") or 0) - 0.25) < 1e-6, line_of(a).get("width")))
        r.step("typed Stärke 3 + Enter", lambda: (winapi.click(*r.at(control("Stärke", "input.tp-num")), count=3), winapi.press("ctrl+a"), winapi.type_text("3"), winapi.press("enter")),
               lambda b, a: (line_of(a).get("width") == 3, line_of(a).get("width")))
        r.step("Strg+Z after typing", lambda: (c.js("document.activeElement.blur()"), winapi.press("ctrl+z")),
               lambda b, a: (line_of(b).get("width") == 3 and line_of(a).get("width") == 1, f"{line_of(b).get('width')} -> {line_of(a).get('width')}"))
        r.step("dropdown Strichtyp (opened, arrow down, Enter)", lambda: (winapi.click(*r.at(control("Strichtyp", "select"))), time.sleep(0.4), winapi.press("down"), winapi.press("enter")),
               lambda b, a: (line_of(a).get("dash") not in (None, "solid"), line_of(a).get("dash")))
        r.step("dropdown Haupttyp Außen (tick marks drawn)", lambda: (winapi.click(*r.at("[...document.querySelectorAll('.taskpane .tp-tab')].at(-1)")), time.sleep(0.3),
                                                                    winapi.click(*r.at("[...document.querySelectorAll('.taskpane .tp-sec-head')].find((h) => h.textContent.includes('Teilstriche'))")), time.sleep(0.3),
                                                                    winapi.click(*r.at(control("Haupttyp", "select"))), time.sleep(0.4), winapi.press("down"), winapi.press("down"), winapi.press("enter")),
               lambda b, a: (((a.get("format") or {}).get("valAxis") or {}).get("majorTick") == "out", ((a.get("format") or {}).get("valAxis") or {}).get("majorTick")))
        # A gradient stop dragged along its bar (Diagrammbereich, Füllung: Farbverlauf).
        winapi.click(*r.at("document.querySelector('.taskpane .tp-set')"))
        time.sleep(0.3)
        winapi.click(*r.at("[...document.querySelectorAll('.menu .item')].find((b) => b.textContent === 'Diagrammbereich')"))
        time.sleep(0.3)
        winapi.click(*r.at("document.querySelector('.taskpane .tp-tab')"))
        time.sleep(0.2)
        r.step("radio Farbverlauf", lambda: winapi.click(*r.at(control("Farbverlauf", "input"))),
               lambda b, a: (((a.get("format") or {}).get("chart") or {}).get("fill", {}).get("type") == "gradient", "ok"))
        def stops(m):
            return [s["pos"] for s in ((((m.get("format") or {}).get("chart") or {}).get("fill") or {}).get("gradient") or {}).get("stops", [])]
        r.step("gradient stop dragged", lambda: winapi.drag(*r.at("document.querySelectorAll('.taskpane .tp-grad-stop')[1]"), *r.at("document.querySelector('.taskpane .tp-grad-bar')", "left")),
               lambda b, a: (stops(a) != stops(b), f"{stops(b)} -> {stops(a)}"))
        r.step("Farbverlaufstopp hinzufügen (+)", lambda: winapi.click(*r.at("document.querySelector('.taskpane .tp-grad .tp-icon-btn')")),
               lambda b, a: (len(stops(a)) == len(stops(b)) + 1, f"{len(stops(b))} -> {len(stops(a))}"))
        out = HERE / "results" / "2026-11-09-calc-chartpane-real.json"
        out.write_text(json.dumps({"date": datetime.date.today().isoformat(), "passed": sum(1 for x in r.results if x["ok"]), "total": len(r.results), "steps": r.results},
                                  ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        print("passed", sum(1 for x in r.results if x["ok"]), "/", len(r.results))
    finally:
        if c:
            c.close()
        LOCK.unlink(missing_ok=True)


if __name__ == "__main__":
    os.environ.setdefault("PYTHONUTF8", "1")
    main()
