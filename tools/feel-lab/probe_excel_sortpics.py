"""Probe: which pictures (shapes) Excel moves when it sorts a range: inside one row, over two rows,
in a column outside the sorted range, and with each placement (Von Zellposition und -größe
abhängig, Nur von Zellposition abhängig, Von Zellposition und -größe unabhängig).
Own invisible Excel, no real input. Result: results/<date>-excel-sortpics.json.
"""
from __future__ import annotations

import datetime
import json
from pathlib import Path

from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
# (name, top-left cell, rows tall, placement: 1 move and size, 2 move only, 3 free)
SHAPES = [
    ("in row 1", "B1", 1, 1), ("in row 2 move only", "B2", 1, 2), ("in row 3 free", "B3", 1, 3),
    ("over rows 4-5", "B4", 2, 1), ("outside in D1", "D1", 1, 1),
]


def main() -> None:
    x = ExcelTarget(visible=False)
    out = []
    try:
        ws = x.sheet
        for i, v in enumerate([5, 4, 3, 2, 1]):
            ws.Cells(i + 1, 1).Value = v
            ws.Cells(i + 1, 2).Value = f"r{i + 1}"
        for name, at, tall, place in SHAPES:
            c = ws.Range(at)
            # A shape a little smaller than its cells, so it lies inside them.
            s = ws.Shapes.AddShape(1, c.Left + 2, c.Top + 2, c.Width - 4, c.Height * tall - 4)
            s.Name = name
            s.Placement = place
        before = {s.Name: [s.TopLeftCell.GetAddress(False, False), s.BottomRightCell.GetAddress(False, False)] for s in ws.Shapes}
        # Sort A1:B5 ascending by column A (rows reverse).
        so = ws.Sort
        so.SortFields.Clear()
        so.SortFields.Add(ws.Range("A1:A5"), 0, 1)
        so.SetRange(ws.Range("A1:B5"))
        so.Header = 2
        so.Apply()
        after = {s.Name: [s.TopLeftCell.GetAddress(False, False), s.BottomRightCell.GetAddress(False, False)] for s in ws.Shapes}
        rows = [ws.Cells(r, 2).Value for r in range(1, 6)]
        out = {"rows after sort": rows, "before": before, "after": after}
        print(json.dumps(out, indent=1))
    finally:
        x.close()
    (HERE / "results" / f"{datetime.date.today().isoformat()}-excel-sortpics.json").write_text(json.dumps(out, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
