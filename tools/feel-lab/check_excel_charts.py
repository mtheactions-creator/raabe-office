"""Check: charts as RaabeCalc writes them open in Excel without repair (Step 7 package E, D-061).

Runs packages/charts/test/excel-open.test.ts, which writes two workbooks with every chart kind
(plain, and with titles, axis titles, legend on the right, data labels and trendlines) into
TEMP/raabe-charts-check, then opens each in an invisible Excel of our own. A workbook Excel would
have to repair does not open through COM with alerts off, so opening is the test; then the chart
types and series formulas are compared with what was written and every chart is exported as PNG
next to the workbooks for a look. Runs invisibly, no input; nothing is saved to results/.

    .venv/Scripts/python.exe check_excel_charts.py
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

from feellab.excel import ExcelTarget

ROOT = Path(__file__).resolve().parents[2]
OUT = Path(tempfile.gettempdir()) / "raabe-charts-check"
XL_TYPES = {
    "column": 51, "columnStacked": 52, "bar": 57, "barStacked": 58, "line": 4, "lineMarkers": 65, "area": 1,
    "pie": 5, "doughnut": -4120, "scatter": -4169, "scatterSmooth": 73, "scatterLines": 74, "histogram": 118,
}


def main() -> int:
    OUT.mkdir(exist_ok=True)
    env = dict(os.environ, RAABE_CHARTS_XLSX=str(OUT))
    r = subprocess.run("npx vitest run packages/charts/test/excel-open.test.ts", cwd=ROOT, env=env, shell=True,
                       capture_output=True, text=True, errors="replace")
    if r.returncode:
        print(r.stdout[-2000:], r.stderr[-2000:])
        return 1
    expected = json.loads((OUT / "expected.json").read_text(encoding="utf-8"))
    failures = 0
    x = ExcelTarget(visible=False)
    try:
        for name, charts in expected.items():
            path = OUT / name
            try:
                wb = x.app.Workbooks.Open(str(path))
            except Exception as e:  # noqa: BLE001 - Excel refuses files it would have to repair
                print(f"{name}: does not open ({e})")
                failures += 1
                continue
            try:
                ws = wb.Worksheets(1)
                found = ws.ChartObjects().Count
                ok = found == len(charts)
                print(f"{name}: opens, {found}/{len(charts)} charts")
                for i, want in enumerate(charts, 1):
                    ch = ws.ChartObjects(i).Chart
                    got_type = ch.ChartType
                    series = []
                    if want["series"]:
                        sc = ch.SeriesCollection()
                        series = [sc.Item(k).Formula for k in range(1, sc.Count + 1)]
                    good = got_type == XL_TYPES[want["kind"]] and series == want["series"]
                    ch.Export(str(OUT / f"{path.stem}-{i:02d}-{want['kind']}.png"))
                    if not good:
                        ok = False
                        print(f"  {want['kind']}: type {got_type} (want {XL_TYPES[want['kind']]}), series {series} (want {want['series']})")
                if not ok:
                    failures += 1
            finally:
                wb.Close(False)
    finally:
        x.close()
    print("all charts open in Excel as written" if not failures else f"{failures} workbook(s) with differences")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
