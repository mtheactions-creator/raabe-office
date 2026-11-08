"""Records how Excel turns column widths into pixels for different standard fonts (the font of the Standard cell
style): for each font and size, the standard width, the width in points of columns set to 1, 5, 8.43, 10 and 20
characters, and the default row height. As the reference for RaabeCalc's widths per workbook (D-089).

    .venv/Scripts/python.exe probe_excel_mdw.py

Writes tools/feel-lab/results/<date>-excel-mdw.json. Our own invisible Excel; no real input."""

from __future__ import annotations

import datetime
import json
from pathlib import Path

from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
FONTS = [("Aptos Narrow", 11), ("Calibri", 11), ("Arial", 10), ("Arial", 11), ("Aptos", 11), ("Calibri", 12), ("Times New Roman", 12),
         ("Cambria", 11), ("Verdana", 10), ("Segoe UI", 9), ("Calibri", 10), ("Aptos Narrow", 12), ("Tahoma", 10), ("Consolas", 11)]


def main() -> None:
    xl = ExcelTarget(visible=False)
    out = []
    try:
        for name, size in FONTS:
            wb = xl.app.Workbooks.Add()
            try:
                st = wb.Styles("Standard") if True else None
                st.Font.Name = name
                st.Font.Size = size
                ws = wb.Worksheets(1)
                row = {"font": name, "size": size, "standard_width": ws.StandardWidth, "standard_points": ws.Columns(30).Width,
                       "row_height": ws.Rows(30).Height, "widths": {}}
                for i, w in enumerate([1, 5, 8.43, 10, 20, 0.5]):
                    c = ws.Columns(i + 1)
                    c.ColumnWidth = w
                    row["widths"][str(w)] = {"chars": c.ColumnWidth, "points": c.Width}
                out.append(row)
                print(row)
            finally:
                wb.Close(False)
    finally:
        xl.close()
    (HERE / "results" / f"{datetime.date.today().isoformat()}-excel-mdw.json").write_text(json.dumps(out, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
