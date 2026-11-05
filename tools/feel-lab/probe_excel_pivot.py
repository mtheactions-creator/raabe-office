"""Probe: the reports Excel's PivotTables produce (D-066): for a small school data set and several
configurations (rows, columns, values with each function, filters, the three report layouts), the
values of the whole report (TableRange2) and where it starts. Own invisible Excel, COM only.
Result: results/<date>-excel-pivot.json (used by packages/sheet-engine/test/pivot.test.ts).
"""
from __future__ import annotations

import datetime
import json
from pathlib import Path

from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
DATA = [
    ["Name", "Klasse", "Fach", "Note", "Geschlecht"],
    ["Anna", "8a", "Mathe", 2, "w"], ["Ben", "8a", "Mathe", 3, "m"], ["Cem", "8b", "Mathe", 1, "m"], ["Dora", "8b", "Mathe", 4, "w"],
    ["Anna", "8a", "Deutsch", 1, "w"], ["Ben", "8a", "Deutsch", 2, "m"], ["Cem", "8b", "Deutsch", 3, "m"], ["Dora", "8b", "Deutsch", 2, "w"],
    ["Emil", "8c", "Mathe", 5, "m"], ["Emil", "8c", "Deutsch", 4, "m"],
]
FUNC = {"sum": -4157, "count": -4112, "average": -4106, "max": -4136, "min": -4139}
CONFIGS = {
    "rows-sum": {"rows": ["Klasse"], "values": [("Note", "sum")]},
    "rows2-avg": {"rows": ["Klasse", "Fach"], "values": [("Note", "average")]},
    "rows-cols-count": {"rows": ["Klasse"], "cols": ["Fach"], "values": [("Name", "count")]},
    "two-values": {"rows": ["Fach"], "values": [("Note", "sum"), ("Note", "count")]},
    "filter-max": {"rows": ["Klasse"], "values": [("Note", "max")], "filter": ("Geschlecht", "w")},
    "tabular": {"rows": ["Klasse", "Fach"], "values": [("Note", "sum")], "layout": 1},
    "outline": {"rows": ["Klasse", "Fach"], "values": [("Note", "min")], "layout": 2},
    "cols-only": {"cols": ["Klasse"], "values": [("Note", "sum")]},
}


def main() -> None:
    x = ExcelTarget(visible=False)
    out: dict = {"data": DATA, "configs": {}}
    try:
        ws = x.sheet
        for r, row in enumerate(DATA, start=1):
            for c, v in enumerate(row, start=1):
                ws.Cells(r, c).Value = v
        src = ws.Range(ws.Cells(1, 1), ws.Cells(len(DATA), len(DATA[0])))
        for name, cfg in CONFIGS.items():
            dest = x.wb.Worksheets.Add()
            pc = x.wb.PivotCaches().Create(1, src)
            pt = pc.CreatePivotTable(dest.Range("A3"), "P" + name.replace("-", ""))
            for f in cfg.get("rows", []):
                pt.PivotFields(f).Orientation = 1
            for f in cfg.get("cols", []):
                pt.PivotFields(f).Orientation = 2
            if "filter" in cfg:
                fld = pt.PivotFields(cfg["filter"][0])
                fld.Orientation = 3
                fld.CurrentPage = cfg["filter"][1]
            for f, fn in cfg["values"]:
                pt.AddDataField(pt.PivotFields(f), None, FUNC[fn])
            if "layout" in cfg:
                pt.RowAxisLayout(cfg["layout"])
            rng = pt.TableRange2
            vals = rng.Value
            out["configs"][name] = {
                "cfg": cfg, "top": rng.Row, "left": rng.Column,
                "cells": [[v for v in row] for row in (vals if isinstance(vals, tuple) else ((vals,),))],
            }
            print(name, rng.Address, flush=True)
    finally:
        x.close()
    (HERE / "results" / f"{datetime.date.today().isoformat()}-excel-pivot.json").write_text(json.dumps(out, ensure_ascii=False, indent=1, default=str), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
