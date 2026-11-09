"""The chart buttons with the real mouse (D-115 - D-117): what only real input shows - hover (the side arrow, the
Livevorschau of a style, the filter's highlight), flyouts opening and closing with real clicks, and dragging a
coloured source range in the sheet. Runs the release exe on screen under the real-input lock; before every input
the window in front must be RaabeCalc's own. Readouts through DevTools.

    .venv/Scripts/python.exe check_calc_chartui_real.py

Result: results/2026-11-09-calc-chartui-real.json.
"""
from __future__ import annotations

import json
import sys
import time
from pathlib import Path

import os

from feellab import winapi
from feellab.calc_catalogue import CalcSide
from feellab.inputlock import real_input

_place = winapi.place
import check_calc_chartui as k  # noqa: E402

# The DevTools check runs off screen; this one needs the window on screen and in front.
winapi.place = _place
os.environ.pop("RAABE_FEEL_BACKGROUND", None)

HERE = Path(__file__).parent
RESULT = HERE / "results" / "2026-11-09-calc-chartui-real.json"
PORT = 9479


class Real:
    def __init__(self, c: CalcSide) -> None:
        self.c = c
        self.ui = k.Ui(c)

    def ours(self) -> bool:
        fg = winapi.foreground()
        return bool(fg) and winapi.window_pid(fg) == self.c.proc.pid

    def screen(self, x: float, y: float) -> tuple[float, float]:
        ox, oy, dpr = self.c._origin()
        return ox + x * dpr, oy + y * dpr

    def guard(self) -> None:
        if not self.ours():
            winapi.activate(self.c.hwnd, 0.3)
        if not self.ours():
            raise RuntimeError("RaabeCalc is not in front")

    def move(self, x: float, y: float) -> None:
        self.guard()
        winapi.move(*self.screen(x, y))
        time.sleep(0.35)

    def click(self, x: float, y: float) -> None:
        self.guard()
        winapi.click(*self.screen(x, y))
        time.sleep(0.35)

    def drag(self, x1: float, y1: float, x2: float, y2: float) -> None:
        self.guard()
        a, b = self.screen(x1, y1), self.screen(x2, y2)
        winapi.drag(a[0], a[1], b[0], b[1], steps=25)
        time.sleep(0.4)

    def at(self, sel: str, i: int = 0) -> tuple[float, float]:
        x, y, _, _ = self.ui.center(sel, i)
        return x, y


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    run = k.Run()
    with real_input("charts-buttons: RaabeCalc chart buttons, real mouse", wait=1800):
        c = CalcSide(PORT)
        r = Real(c)
        ui = r.ui
        try:
            winapi.place(c.hwnd, 40, 40, 2200, 1400)
            time.sleep(0.8)
            r.guard()
            k.fresh(ui)
            x, y, w, h = k.chart_center(ui)
            r.click(x - w * 0.45, y + h * 0.45)
            run.check("real click selects the chart", ui.js("!!__calc.app.charts.selected()"))
            # Diagrammelemente: hovering an entry shows its arrow; a click on it opens the side menu.
            r.click(*r.at('.chart-btn[data-which="elements"]'))
            run.check("real click opens Diagrammelemente", ui.js("document.querySelector('.cb-fly')?.dataset.which === 'elements'"))
            i = [t[0] for t in k.flyout_rows(ui)].index("Legende")
            rx, ry = r.at(".cb-fly .cb-row", i)
            r.move(rx - 30, ry)
            vis = ui.js(f"getComputedStyle(document.querySelectorAll('.cb-fly .cb-arrow')[{i}]).visibility")
            run.check("hovering Legende shows its arrow", vis == "visible", vis)
            r.click(*r.at(".cb-fly .cb-arrow", i))
            labels = ui.texts(".cb-sub button")
            run.check("the arrow opens Legende's side menu", labels[:4] == ["Rechts", "Oben", "Links", "Unten"], labels)
            r.click(*[p for p in [ui.js("(() => { const b = [...document.querySelectorAll('.cb-sub button')].find((x) => x.textContent.trim() === 'Oben').getBoundingClientRect(); return [b.left + b.width / 2, b.top + b.height / 2]; })()")]][0])
            run.check("Legende > Oben with the real mouse", ui.model()["legend"] == "t")
            # A click outside closes the flyout.
            p = ui.js("__calc.point('L25', 'center')")
            r.click(p["x"], p["y"])
            run.check("a click in the sheet closes the flyout and deselects", not ui.js("!!document.querySelector('.cb-fly')") and not ui.js("!!__calc.app.charts.selected()"))
            # Diagrammformatvorlagen: the style under the mouse is shown, leaving restores it.
            x, y, w, h = k.chart_center(ui)
            r.click(x - w * 0.45, y + h * 0.45)
            r.click(*r.at('.chart-btn[data-which="styles"]'))
            before = ui.js("document.querySelector('.chart-obj svg').outerHTML.length")
            r.move(*r.at(".cb-fly .cb-tile", 2))
            during = ui.js("document.querySelector('.chart-obj svg').outerHTML.length")
            r.move(*r.at(".cb-fly .cb-tab", 0))
            after = ui.js("document.querySelector('.chart-obj svg').outerHTML.length")
            run.check("Livevorschau with the real mouse (shown on hover, gone after)", during != before and after == before, [before, during, after])
            r.click(*r.at(".cb-fly .cb-tile", 2))
            run.check("a real click applies Formatvorlage 3", ui.js("__calc.app.charts.selected().style") == 203)
            ui.undo()
            # Diagrammfilter: hover a series row, the others fade; untick, Anwenden.
            if ui.js("document.querySelector('.cb-fly')?.dataset.which !== 'filter'"):
                r.click(*r.at('.chart-btn[data-which="filter"]'))
            r.move(*r.at(".cb-fly .cb-filter-row", 3))
            faded = ui.js("[...new Set([...document.querySelectorAll('.chart-obj svg [data-fade]')].map((e) => e.dataset.part.split(':')[1]))]")
            run.check("hovering West fades Nord and Süd", sorted(faded) == ["0", "1"], faded)
            r.click(*r.at(".cb-fly .cb-filter-row input", 3))
            r.click(*r.at(".cb-fly .cb-foot button"))
            run.check("untick West, Anwenden", [bool(s.get("filtered")) for s in ui.model()["series"]] == [False, False, True])
            ui.undo()
            # Drag the coloured values frame's corner in the sheet: D5 to C4.
            p = ui.js("__calc.point('L25', 'center')")
            r.click(p["x"], p["y"])
            x, y, w, h = k.chart_center(ui)
            r.click(x - w * 0.45, y + h * 0.45)
            cx, cy = r.at('.chart-range[data-kind="val"] .corner.se')
            q = ui.js("__calc.point('C4', 'center')")
            r.drag(cx, cy, q["x"], q["y"])
            m = ui.model()
            run.check("real drag of the values' corner to C4: two series, three months",
                      [s["val"]["ref"] for s in m["series"]] == ["Tabelle1!$B$2:$B$4", "Tabelle1!$C$2:$C$4"], [s["val"]["ref"] for s in m["series"]])
            ui.undo()
            # Datenquelle auswählen: point at cells with the real mouse while the dialog is open (shrunk).
            ui.js("__calc.app.charts.selectData()")
            time.sleep(0.5)
            r.click(*r.at(".dialog .rf button"))
            a = ui.js("__calc.point('A1', 'center')")
            b = ui.js("__calc.point('B5', 'center')")
            r.drag(a["x"], a["y"], b["x"], b["y"])
            val = ui.js("document.querySelector('.dialog .rf input').value")
            run.check("real drag over A1:B5 while the dialog is shrunk", val == "=Tabelle1!$A$1:$B$5", val)
            r.click(*r.at(".dialog .rf-bar .rf button"))
            ui.js("document.querySelector('.dialog .dlg-buttons button.primary').click()")
            time.sleep(0.4)
            run.check("OK: one series Nord", [s.get("name", {}).get("text") for s in ui.model()["series"]] == ["Nord"])
        except Exception as ex:  # noqa: BLE001
            import traceback
            run.check("run completed", False, traceback.format_exc()[-600:])
        finally:
            c.close()
    passed = sum(1 for x in run.cases if x["passed"])
    RESULT.write_text(json.dumps({"passed": passed, "total": len(run.cases), "cases": run.cases}, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(f"{passed}/{len(run.cases)} passed")


if __name__ == "__main__":
    main()
