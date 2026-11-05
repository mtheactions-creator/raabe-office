"""Probe: Excel's Diagrammformatvorlagen (Diagrammentwurf > Diagrammformatvorlagen, Chart.ChartStyle
201 …) per chart type (D-066). Our own invisible Excel, COM only, no input. For each type it makes one
chart per style on its own sheet and saves the workbook to TEMP (excel-chartstyles-<type>.xlsx);
RaabeCalc's chart reader turns those charts into style data (packages/charts/test/gen-chartstyles.test.ts).
Also records which style numbers each type offers and what AddChart2 starts with.
Result: results/<date>-excel-chartstyles.json.
"""
from __future__ import annotations

import datetime
import json
import tempfile
from pathlib import Path

from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
TMP = Path(tempfile.gettempdir())
TYPES = {"column": 51, "columnStacked": 52, "bar": 57, "line": 4, "lineMarkers": 65, "pie": 5, "doughnut": -4120,
         "scatter": -4169, "area": 1, "radar": -4151}
DATA = [("Monat", "A", "B"), ("Jan", 3, 4), ("Feb", 5, 2), ("Mrz", 2, 6), ("Apr", 4, 3)]


def main() -> None:
    out: dict = {"types": {}}
    for name, t in TYPES.items():
        x = ExcelTarget(visible=False)
        try:
            wb = x.wb
            data_ws = wb.Worksheets(1)
            data_ws.Name = "Daten"
            for r, row in enumerate(DATA, start=1):
                for c, v in enumerate(row, start=1):
                    data_ws.Cells(r, c).Value = v
            src = data_ws.Range("A1:B5") if t in (5, -4120) else data_ws.Range("A1:C5")
            styles: list[int] = []
            misses = 0
            first = None
            for st in range(201, 400):
                ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
                sh = ws.Shapes.AddChart2(-1, t, 20, 20, 360, 216)
                ch = sh.Chart
                ch.SetSourceData(src)
                if first is None:
                    first = ch.ChartStyle
                try:
                    ch.ChartStyle = st
                    ok = ch.ChartStyle == st
                except Exception:  # noqa: BLE001
                    ok = False
                if ok:
                    ws.Name = f"S{st}"
                    styles.append(st)
                    misses = 0
                else:
                    x.app.DisplayAlerts = False
                    ws.Delete()
                    misses += 1
                    if misses >= 12 and styles:
                        break
            f = TMP / f"excel-chartstyles-{name}.xlsx"
            wb.SaveAs(str(f), 51)
            out["types"][name] = {"xlType": t, "default": first, "styles": styles}
            print(name, first, styles, flush=True)
        finally:
            x.close()
    (HERE / "results" / f"{datetime.date.today().isoformat()}-excel-chartstyles.json").write_text(json.dumps(out, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
