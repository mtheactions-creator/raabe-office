"""What Excel does with modifier keys pressed mid-drag and with the pointer leaving the window (D-127, the drag test's
reference for RaabeCalc): our own visible Excel, real drags on a column border, a row border, the fill handle, a
selection's border, a range selection, a shape (body, corner handle, rotation handle) and a chart (body, corner), each
along feellab/officedrag.py's variants: plain; Umschalt, Strg, Alt pressed mid-drag and let go mid-drag (-mid) or held
over the release (-held); out of the window and back (out-back) and the button let go outside (out-release). Read back
through COM: the anchor's value, the number of shapes (copies) and the window's scroll position (auto-scroll).
Takes over mouse and keyboard (the real-input lock). Result: results/2026-11-09-excel-drag-mods.json.
"""
from __future__ import annotations

import json
import sys
import time
from pathlib import Path

from feellab import winapi
from feellab.calc_catalogue import _pt_to_px
from feellab.excel import ExcelTarget
from feellab.inputlock import real_input, touch
from feellab import officedrag
from feellab.catcher import catcher
from feellab.officedrag import VARIANTS, drag

HERE = Path(__file__).parent
RECT = (0, 0, 2000, 1700)


def main() -> None:
    want = sys.argv[1:]
    res: list[dict] = []
    with real_input("probe_excel_drag_mods.py (drag-paths agent)"):
        x = ExcelTarget(visible=True)
        guard = catcher(RECT)
        guard.__enter__()
        try:
            app = x.app
            winapi.place(x.hwnd, *RECT)
            app.ActiveWindow.Zoom = 100
            sh = x.sheet

            def ours() -> bool:
                fg = winapi.foreground()
                return fg == x.hwnd or winapi.window_pid(fg) == x.pid

            def front() -> None:
                officedrag.front(x.hwnd, ours)

            front()

            def sx(pt: float) -> float:
                return app.ActiveWindow.PointsToScreenPixelsX(0) + _pt_to_px(pt)

            def sy(pt: float) -> float:
                return app.ActiveWindow.PointsToScreenPixelsY(0) + _pt_to_px(pt)

            def cells() -> None:
                for s in list(sh.Shapes):
                    s.Delete()
                sh.Cells.Clear()
                sh.Columns.ColumnWidth = sh.StandardWidth
                sh.Rows.RowHeight = sh.StandardHeight
                app.ActiveWindow.ScrollRow = 1
                app.ActiveWindow.ScrollColumn = 1
                for a, v in {"A1": "Name", "B1": "Wert", "A3": "Ben", "B3": "7", "C3": "x1", "D4": "d4", "C5": "c5"}.items():
                    sh.Range(a).Value2 = v
                sh.Range("E8").Select()

            b = lambda: sh.Range("B1")
            r3 = lambda: sh.Range("A3")
            c3 = lambda: sh.Range("C3")

            def sel_c3() -> None:
                cells()
                c3().Select()

            def sel_b3() -> None:
                cells()
                sh.Range("B3").Select()

            def shape() -> None:
                cells()
                s = sh.Shapes.AddShape(1, 200, 120, 100, 80)
                s.Name = "Probe"
                s.Select()

            def chart() -> None:
                cells()
                co = sh.ChartObjects().Add(200, 120, 240, 160)
                co.Name = "Probe"
                co.Chart.SetSourceData(sh.Range("A1:B3"))
                co.Select()

            sp = lambda: sh.Shapes("Probe")
            geo = lambda: [round(sp().Left, 1), round(sp().Top, 1), round(sp().Width, 1), round(sp().Height, 1), round(sp().Rotation, 1)]
            grid = lambda: {a: sh.Range(a).Text for a in ("B3", "C3", "D3", "C4", "D4", "C5", "C6", "C7", "D5", "E5")}
            cases = [
                ("column border", cells, [lambda: (sx(b().Left + b().Width), sy(0) - 15)], "right", lambda: round(sh.Columns(2).ColumnWidth, 2)),
                ("row border", cells, [lambda: (sx(0) - 20, sy(r3().Top + r3().Height))], "down", lambda: round(sh.Rows(3).RowHeight, 2)),
                ("fill handle", sel_c3, [lambda: (sx(c3().Left + c3().Width), sy(c3().Top + c3().Height))], "down", grid),
                ("move selection by its border", sel_c3, [lambda: (sx(c3().Left) + 1, sy(c3().Top + c3().Height / 2))], "right", grid),
                ("select a range", sel_b3, [lambda: (sx(sh.Range("B3").Left + 20), sy(sh.Range("B3").Top + 8))], "down", lambda: app.Selection.Address),
                ("shape move", shape, [lambda: (sx(sp().Left + sp().Width / 2), sy(sp().Top + sp().Height / 2))], "right", geo),
                ("shape corner handle", shape, [lambda: (sx(sp().Left + sp().Width), sy(sp().Top + sp().Height))], "right", geo),
                ("shape rotation handle", shape, [(lambda k=k: (sx(sp().Left + sp().Width / 2), sy(sp().Top) - k)) for k in (20, 24, 28, 32, 36, 40, 16)], "right", geo),
                ("chart move", chart, [lambda: (sx(sp().Left + 6), sy(sp().Top + 6))], "right", geo),
                ("chart corner handle", chart, [(lambda k=k: (sx(sp().Left + sp().Width) - k, sy(sp().Top + sp().Height) - k)) for k in (2, 4, 6, 0, -2)], "right", geo),
            ]
            for name, setup, points, direction, read in cases:
                if want and not any(w in name for w in want):
                    continue
                out: dict = {"case": name, "out": direction}
                try:
                    point = None
                    for i, pt in enumerate(points):
                        setup()
                        time.sleep(0.3)
                        front()
                        before = read()
                        front()
                        officedrag.run_variant(lambda pt=pt: (setup(), time.sleep(0.3), front(), drag(pt(), "plain", ours), {})[-1], ours)
                        after = read()
                        winapi.press("esc")
                        if after != before:
                            point = pt
                            out["point"] = i
                            break
                    if point is None:
                        out["error"] = "the plain drag changed nothing"
                    else:
                        for v in VARIANTS:
                            def attempt(v=v) -> dict:
                                setup()
                                time.sleep(0.3)
                                front()
                                before, n0 = read(), sh.Shapes.Count
                                s0 = (app.ActiveWindow.ScrollRow, app.ActiveWindow.ScrollColumn)
                                front()
                                drag(point(), v, ours, window=winapi.window_rect(x.hwnd), direction=direction)
                                after, n1 = read(), sh.Shapes.Count
                                s1 = (app.ActiveWindow.ScrollRow, app.ActiveWindow.ScrollColumn)
                                return {"before": before, "after": after, "copies": n1 - n0, "scrolled": [s1[0] - s0[0], s1[1] - s0[1]]}
                            out[v] = officedrag.run_variant(attempt, ours)
                            officedrag.settle(ours)
                            time.sleep(0.2)
                            touch()
                except Exception as e:  # noqa: BLE001
                    out["error"] = str(e)[:200]
                res.append(out)
                print(json.dumps(out, ensure_ascii=False), flush=True)
                touch()
        finally:
            guard.__exit__(None, None, None)
            x.close()
    path = HERE / "results" / "2026-11-09-excel-drag-mods.json"
    if want and path.exists():
        old = json.loads(path.read_text(encoding="utf-8")).get("excel", [])
        names = {r["case"] for r in res}
        res = [r for r in old if r["case"] not in names] + res
    path.write_text(json.dumps({"excel": res}, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
