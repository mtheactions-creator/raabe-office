"""Probe: Nach Farbe filtern / sortieren in Excel through COM (own invisible Excel): which rows
stay visible for a fill colour, Keine Füllung, a font colour and Automatisch, and the order after
sorting by a fill colour (rows with it on top). Result: results/<date>-excel-colorfilter.json.
"""
from __future__ import annotations

import datetime
import json
from pathlib import Path

from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
VALUES = ["Note", 2, 3, 1, 5, 4, 2, 6]
FILLS = {3: 0x00FFFF, 5: 0x0000FF, 7: 0x00FFFF}   # rows (1-based): yellow, red, yellow (BGR)
FONTS = {6: 0x0000FF}                               # row 6: red text


def setup(ws) -> None:
    ws.Cells.Clear()
    for i, v in enumerate(VALUES):
        ws.Cells(i + 1, 1).Value = v
    for r, c in FILLS.items():
        ws.Cells(r, 1).Interior.Color = c
    for r, c in FONTS.items():
        ws.Cells(r, 1).Font.Color = c


def main() -> None:
    x = ExcelTarget(visible=False)
    out: dict = {"filters": [], "sort": None}
    try:
        ws = x.sheet
        # xlFilterCellColor 8, xlFilterFontColor 9, xlFilterNoFill 12, xlFilterAutomaticFontColor 13
        for name, crit, op in (("fill yellow", 0x00FFFF, 8), ("fill red", 0x0000FF, 8), ("no fill", None, 12), ("font red", 0x0000FF, 9), ("automatic font", None, 13)):
            setup(ws)
            kw = {"Field": 1, "Operator": op}
            if crit is not None:
                kw["Criteria1"] = crit
            ws.Range("A1:A8").AutoFilter(**kw)
            shown = [r for r in range(2, 9) if not ws.Rows(r).Hidden]
            out["filters"].append({"name": name, "shown": shown})
            print(name, shown)
            ws.AutoFilterMode = False
        setup(ws)
        s = ws.Sort
        s.SortFields.Clear()
        f = s.SortFields.Add(Key=ws.Range("A2:A8"), SortOn=1, Order=1)  # xlSortOnCellColor, xlAscending = on top
        f.SortOnValue.Color = 0x00FFFF
        s.SetRange(ws.Range("A2:A8"))
        s.Header = 2
        s.Apply()
        out["sort"] = [ws.Cells(r, 1).Value for r in range(2, 9)]
        print("sort yellow on top", out["sort"])
    finally:
        x.close()
    (HERE / "results" / f"{datetime.date.today().isoformat()}-excel-colorfilter.json").write_text(json.dumps(out, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
