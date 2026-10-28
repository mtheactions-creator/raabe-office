"""Probe: what Excel's fill kinds from the right-drag menu produce (Range.AutoFill with
xlFillDays 5, xlFillWeekdays 6, xlFillMonths 7, xlFillYears 8, xlLinearTrend 9, xlGrowthTrend
10), for RaabeCalc's fillDates/fillTrend. Own invisible Excel, no real input.
Result: results/<date>-excel-autofill-types.json.
"""
from __future__ import annotations

import datetime
import json
from pathlib import Path

from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
TYPES = {"days": 5, "workdays": 6, "months": 7, "years": 8, "linear": 9, "growth": 10}
# (kind, source values as serials or numbers, is date, how many cells to fill in all)
# Date kinds are left out: through COM, AutoFill copies dates whatever the type, and DataSeries
# steps them by hours; their rules (days, weekdays, months as EDATUM) are plain.
CASES = [
    ("linear", [1, 3], False, 6), ("linear", [2, 3, 7], False, 6), ("linear", [5], False, 4),
    ("growth", [2, 4], False, 6), ("growth", [1, 2, 5], False, 6), ("growth", [3], False, 4),
]


def main() -> None:
    x = ExcelTarget(visible=False)
    out = []
    try:
        ws = x.sheet
        for kind, src, is_date, total in CASES:
            ws.Cells.Clear()
            if is_date:
                ws.Range(ws.Cells(1, 1), ws.Cells(len(src), 1)).NumberFormat = "dd.mm.yyyy"
            for i, v in enumerate(src):
                if is_date:
                    # Typed as a date (FormulaLocal), the way Excel recognises one.
                    d = datetime.date(1899, 12, 30) + datetime.timedelta(days=v)
                    ws.Cells(i + 1, 1).FormulaLocal = d.strftime("%d.%m.%Y")
                else:
                    ws.Cells(i + 1, 1).Value = v
            if is_date:
                # COM's AutoFill copies dates whatever the type; Reihe (DataSeries, chronological,
                # step 1) is the same filling for one start date.
                if len(src) > 1:
                    continue
                unit = {"days": 1, "workdays": 2, "months": 3, "years": 4}[kind]
                ws.Range(ws.Cells(1, 1), ws.Cells(total, 1)).DataSeries(Rowcol=2, Type=4, Date=unit, Step=1)
            else:
                ws.Range(ws.Cells(1, 1), ws.Cells(len(src), 1)).AutoFill(ws.Range(ws.Cells(1, 1), ws.Cells(total, 1)), TYPES[kind])
            vals = [ws.Cells(i, 1).Value2 for i in range(1, total + 1)]
            out.append({"kind": kind, "source": src, "date": is_date, "values": vals})
            print(kind, src, vals)
    finally:
        x.close()
    (HERE / "results" / f"{datetime.date.today().isoformat()}-excel-autofill-types.json").write_text(json.dumps(out, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
