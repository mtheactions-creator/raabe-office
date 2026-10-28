"""Probe: Excel's Zielwertsuche (Range.GoalSeek) for school functions: where it ends (the changing
cell's value) for linear, quadratic, exponential and trigonometric formulas, start values on
either side, and a case without a solution. Own invisible Excel, no real input.
Result: results/<date>-excel-goalseek.json.
"""
from __future__ import annotations

import datetime
import json
from pathlib import Path

from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
# (formula in B1 with A1 as the variable, goal, start value in A1)
CASES = [
    ("=3*A1+2", 20, 0), ("=A1^2", 2, 1), ("=A1^2", 2, -1), ("=A1^2-4*A1+3", 0, 10), ("=EXP(A1)", 10, 0),
    ("=1000*(1+A1)^10", 2000, 0.05), ("=SIN(A1)", 0.5, 0), ("=A1^2", -1, 1), ("=1/A1", 4, 1), ("=A1^3-A1", 0, 0.5),
]


def main() -> None:
    x = ExcelTarget(visible=False)
    out = []
    try:
        ws = x.sheet
        for f, goal, start in CASES:
            ws.Cells.Clear()
            ws.Range("A1").Value = start
            ws.Range("B1").Formula = f
            ok = bool(ws.Range("B1").GoalSeek(goal, ws.Range("A1")))
            out.append({"formula": f, "goal": goal, "start": start, "found": ok, "x": ws.Range("A1").Value, "y": ws.Range("B1").Value})
            print(out[-1])
    finally:
        x.close()
    (HERE / "results" / f"{datetime.date.today().isoformat()}-excel-goalseek.json").write_text(json.dumps(out, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
