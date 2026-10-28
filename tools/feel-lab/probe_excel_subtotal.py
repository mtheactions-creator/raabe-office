"""Probe: Excel's Teilergebnis (Range.Subtotal) on a sorted list: which rows it inserts, their
labels and formulas, the Gesamtergebnis row, and the outline levels and their summary rows.
Own invisible Excel, no real input. Result: results/<date>-excel-subtotal.json.
"""
from __future__ import annotations

import datetime
import json
from pathlib import Path

from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
DATA = [["Klasse", "Name", "Note", "Punkte"], ["7a", "Anna", 2, 11], ["7a", "Ben", 3, 8], ["7b", "Cem", 1, 14], ["7b", "Dora", 4, 6], ["7b", "Emil", 2, 10], ["8a", "Finn", 3, 9]]
# (function: xlSum -4157, xlAverage -4106, xlCount -4112, xlMax -4136, xlMin -4139), total columns (1-based)
CASES = {"Summe Punkte": (-4157, [4]), "Mittelwert Note und Punkte": (-4106, [3, 4]), "Anzahl Note": (-4112, [3]), "Max Punkte": (-4136, [4]),
         "Min Punkte": (-4139, [4]), "Produkt Note": (-4149, [3]), "Anzahl Zahlen": (-4113, [3]), "StdAbw": (-4155, [4]), "StdAbwN": (-4156, [4]),
         "Varianz": (-4164, [4]), "VarianzN": (-4165, [4])}


def main() -> None:
    x = ExcelTarget(visible=False)
    out = {}
    try:
        ws = x.sheet
        for name, (fn, cols) in CASES.items():
            ws.Cells.Clear()
            ws.Cells.ClearOutline()
            for r, row in enumerate(DATA):
                for c, v in enumerate(row):
                    ws.Cells(r + 1, c + 1).Value = v
            ws.Range("A1:D7").Subtotal(GroupBy=1, Function=fn, TotalList=cols, Replace=True, PageBreaks=False, SummaryBelowData=1)
            used = ws.UsedRange
            rows = []
            for r in range(1, used.Rows.Count + 1):
                rows.append({"cells": [ws.Cells(r, c).FormulaLocal for c in range(1, 5)], "bold": [bool(ws.Cells(r, c).Font.Bold) for c in range(1, 5)],
                             "level": ws.Rows(r).OutlineLevel, "text": [ws.Cells(r, c).Text for c in range(1, 5)]})
            out[name] = rows
            print(name, [r["cells"][0] for r in rows if "TEIL" in "".join(r["cells"])][-2:], [c for r in rows[-1:] for c in r["cells"] if "TEIL" in c])
    finally:
        x.close()
    (HERE / "results" / f"{datetime.date.today().isoformat()}-excel-subtotal.json").write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
