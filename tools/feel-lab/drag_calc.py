"""The drag test in RaabeCalc (D-080 C): every anchor dragged along feellab/dragtest.py's paths with real input, checked for
jumps. Takes mouse and keyboard (the real-input lock).

    .venv/Scripts/python.exe drag_calc.py [anchor ...]

Result: results/<date>-drag-calc.json.
"""
from __future__ import annotations

import datetime
import json
import sys
import time
from pathlib import Path

from feellab import winapi
from feellab.calc_catalogue import Case, CalcSide
from feellab.dragtest import Anchor, analyse, paths, run_drag
from feellab.inputlock import real_input, touch

HERE = Path(__file__).parent
G = "(() => { const g = __calc.app.grid; const c = document.querySelector('canvas.grid-canvas').getBoundingClientRect(); return { g, c }; })()"


def col_border(col: int) -> Anchor:
    return Anchor(
        f"column border {chr(65 + col)}",
        start=f"(() => {{ const g = __calc.app.grid; const c = document.querySelector('canvas.grid-canvas').getBoundingClientRect(); const r = g.rectOnCanvas({{ r0: 0, c0: {col}, r1: 0, c1: {col} }}); return {{ x: c.left + r.x + r.w, y: c.top + r.y - 8 }}; }})()",
        probe=f"__calc.app.grid.rectOnCanvas({{ r0: 0, c0: {col}, r1: 0, c1: {col} }}).w",
        expect=lambda dx, dy, v0: max(0.0, v0 + dx),
        tol=1.5, setup="__calc.app.grid.contentChanged && 0")


def row_border(row: int) -> Anchor:
    return Anchor(
        f"row border {row + 1}",
        start=f"(() => {{ const g = __calc.app.grid; const c = document.querySelector('canvas.grid-canvas').getBoundingClientRect(); const r = g.rectOnCanvas({{ r0: {row}, c0: 0, r1: {row}, c1: 0 }}); return {{ x: c.left + r.x - 14, y: c.top + r.y + r.h }}; }})()",
        probe=f"__calc.app.grid.rectOnCanvas({{ r0: {row}, c0: 0, r1: {row}, c1: 0 }}).h",
        expect=lambda dx, dy, v0: max(0.0, v0 + dy), tol=1.5)


ANCHORS = {a.name: a for a in [col_border(1), row_border(2)]}


def main() -> None:
    names = sys.argv[1:] or list(ANCHORS)
    report: dict = {"anchors": []}
    with real_input("drag_calc.py (main session)"):
        side = CalcSide()
        try:
            for name in names:
                anchor = ANCHORS[name]
                res = []
                for p in paths():
                    side.reset(Case("drag", {"A1": "Name", "B1": "Wert", "A2": "Anna", "B2": "12"}, [], "D5", sheets=1))
                    winapi.activate(side.hwnd)
                    time.sleep(0.2)
                    rec = run_drag(side, anchor, p, side.js)
                    a = analyse(anchor, rec)
                    res.append(a)
                    print(name, p.name, "ok" if a["ok"] else "FAIL", {k: v for k, v in a.items() if k not in ("path", "ok")}, flush=True)
                    touch()
                report["anchors"].append({"anchor": name, "ok": sum(1 for r in res if r["ok"]), "of": len(res), "drags": res})
        finally:
            side.close()
    (HERE / "results" / f"{datetime.date.today().isoformat()}-drag-calc.json").write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    for a in report["anchors"]:
        print(a["anchor"], a["ok"], "/", a["of"])


if __name__ == "__main__":
    main()
