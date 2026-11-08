"""What Esc does during a mouse drag in Excel, and Strg+Z after one (D-096, the drag test's reference): our own visible
Excel, real drags on a column border, a row border, the fill handle, a selection's border, a shape (body, corner handle,
rotation handle) and a chart, each once with Esc pressed before the button goes up and once without (the control: the
drag itself works), then Strg+Z after the control drag. Read back through COM.
Takes over mouse and keyboard (the real-input lock). Result: results/<date>-excel-drag-esc.json.
"""
from __future__ import annotations

import datetime
import json
import time
from pathlib import Path

from feellab import winapi
from feellab.calc_catalogue import _pt_to_px
from feellab.excel import ExcelTarget
from feellab.inputlock import real_input, touch

HERE = Path(__file__).parent


def ours(x: ExcelTarget) -> bool:
    fg = winapi.foreground()
    return fg == x.hwnd or winapi.window_pid(fg) == x.pid


def sx(x: ExcelTarget, pt: float) -> float:
    return x.app.ActiveWindow.PointsToScreenPixelsX(0) + _pt_to_px(pt)


def sy(x: ExcelTarget, pt: float) -> float:
    return x.app.ActiveWindow.PointsToScreenPixelsY(0) + _pt_to_px(pt)


def drag(x: ExcelTarget, p: tuple[float, float], d: tuple[float, float], esc: bool) -> None:
    """A real drag from screen point p by d (screen px) in small steps; Esc before the release if asked."""
    winapi.activate(x.hwnd)
    if not ours(x):
        raise RuntimeError("Excel is not in front")
    winapi.move(*p)
    time.sleep(0.15)
    winapi.mouse_down(*p)
    time.sleep(0.08)
    n = 30
    for i in range(1, n + 1):
        winapi.move(p[0] + d[0] * i / n, p[1] + d[1] * i / n)
        time.sleep(0.015)
    time.sleep(0.2)
    if esc:
        if not ours(x):
            raise RuntimeError("Excel is not in front")
        winapi.press("esc")
        time.sleep(0.2)
    winapi.mouse_up(p[0] + d[0], p[1] + d[1])
    time.sleep(0.5)


def undo(x: ExcelTarget) -> None:
    if ours(x):
        winapi.press("ctrl+z")
        time.sleep(0.5)


def case(x: ExcelTarget, name: str, setup, point, delta, read) -> dict:
    try:
        return _case(x, name, setup, point, delta, read)
    except Exception as e:  # noqa: BLE001
        print(name, "error:", e, flush=True)
        return {"case": name, "error": str(e)[:200], "esc_cancels": None}


def _case(x: ExcelTarget, name: str, setup, point, delta, read) -> dict:
    out: dict = {"case": name}
    if isinstance(point, list):
        # Candidate press points (a handle's exact place is not in the object model): the first one whose plain
        # drag changes something is used.
        for i, pt in enumerate(point):
            setup()
            time.sleep(0.3)
            before = read()
            drag(x, pt(), delta, False)
            winapi.press("esc")
            if read() != before:
                out["point"] = i
                point = pt
                break
        else:
            return {"case": name, "error": "no candidate point changed anything", "esc_cancels": None}
    for esc in (True, False):
        setup()
        time.sleep(0.3)
        before = read()
        drag(x, point(), delta, esc)
        after = read()
        rec = {"before": before, "after": after, "changed": after != before}
        if not esc:
            undo(x)
            rec["undone"] = read()
            rec["undo_restores"] = rec["undone"] == before
        out["esc" if esc else "plain"] = rec
        winapi.press("esc")
        touch()
    out["esc_cancels"] = not out["esc"]["changed"] and out["plain"]["changed"]
    return out


def main() -> None:
    res: list[dict] = []
    with real_input("probe_excel_drag_esc.py"):
        x = ExcelTarget(visible=True)
        try:
            app = x.app
            app.ActiveWindow.Zoom = 100
            sh = x.sheet
            winapi.activate(x.hwnd)

            def cells() -> None:
                for s in list(sh.Shapes):
                    s.Delete()
                sh.Cells.Clear()
                sh.Columns.ColumnWidth = sh.StandardWidth
                sh.Rows.RowHeight = sh.StandardHeight
                for a, v in {"A1": "Name", "B1": "Wert", "A3": "Ben", "B3": "7", "C3": "x1"}.items():
                    sh.Range(a).Value2 = v
                sh.Range("E8").Select()

            b = lambda: sh.Range("B1")
            res.append(case(x, "column border", cells, lambda: (sx(x, b().Left + b().Width), sy(x, 0) - 15), (60, 0),
                            lambda: round(sh.Columns(2).ColumnWidth, 2)))
            r3 = lambda: sh.Range("A3")
            res.append(case(x, "row border", cells, lambda: (sx(x, 0) - 20, sy(x, r3().Top + r3().Height)), (0, 30),
                            lambda: round(sh.Rows(3).RowHeight, 2)))
            c3 = lambda: sh.Range("C3")

            def sel_c3() -> None:
                cells()
                c3().Select()

            res.append(case(x, "fill handle", sel_c3, lambda: (sx(x, c3().Left + c3().Width), sy(x, c3().Top + c3().Height)), (0, 60),
                            lambda: [sh.Range(a).Text for a in ("C3", "C4", "C5", "C6")]))
            res.append(case(x, "move selection by its border", sel_c3, lambda: (sx(x, c3().Left) + 1, sy(x, c3().Top + c3().Height / 2)), (0, 60),
                            lambda: [sh.Range(a).Text for a in ("C3", "C4", "C5", "C6")]))

            def shape() -> None:
                cells()
                s = sh.Shapes.AddShape(1, 200, 120, 100, 80)
                s.Name = "Probe"
                s.Select()

            sp = lambda: sh.Shapes("Probe")
            geo = lambda: [round(sp().Left, 1), round(sp().Top, 1), round(sp().Width, 1), round(sp().Height, 1), round(sp().Rotation, 1)]
            res.append(case(x, "shape move", shape, lambda: (sx(x, sp().Left + sp().Width / 2), sy(x, sp().Top + sp().Height / 2)), (60, 30), geo))
            res.append(case(x, "shape corner handle", shape, lambda: (sx(x, sp().Left + sp().Width), sy(x, sp().Top + sp().Height)), (60, 30), geo))
            # The rotation handle sits above the top edge's middle.
            res.append(case(x, "shape rotation handle", shape, [(lambda k=k: (sx(x, sp().Left + sp().Width / 2), sy(x, sp().Top) - k)) for k in (20, 24, 28, 32, 36, 40, 16)], (60, 0), geo))

            def chart() -> None:
                cells()
                co = sh.ChartObjects().Add(200, 120, 240, 160)
                co.Name = "Probe"
                co.Chart.SetSourceData(sh.Range("A1:B3"))
                co.Select()

            co = lambda: sh.Shapes("Probe")
            cgeo = lambda: [round(co().Left, 1), round(co().Top, 1), round(co().Width, 1), round(co().Height, 1)]
            res.append(case(x, "chart move", chart, lambda: (sx(x, co().Left + 6), sy(x, co().Top + 6)), (60, 30), cgeo))
            res.append(case(x, "chart corner handle", chart, [(lambda k=k: (sx(x, co().Left + co().Width) - k, sy(x, co().Top + co().Height) - k)) for k in (2, 4, 6, 0, -2)], (60, 30), cgeo))
        finally:
            x.close()
    out = HERE / "results" / f"{datetime.date.today().isoformat()}-excel-drag-esc.json"
    out.write_text(json.dumps({"excel": res}, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    for r in res:
        print(r["case"], {True: "esc cancels", False: "esc does NOT cancel", None: "error"}[r["esc_cancels"]], "| undo restores:", r.get("plain", {}).get("undo_restores"), r)


if __name__ == "__main__":
    main()
