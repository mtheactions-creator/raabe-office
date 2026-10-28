"""Probe: what Excel's Blitzvorschau (Range.FlashFill) fills in for typical school data: parts of
names, initials, changed case, e-mail parts, reordered dates, joined columns. One or two example
rows are typed, Excel fills the rest of the column. Own invisible Excel, no real input.
Result: results/<date>-excel-flashfill.json.
"""
from __future__ import annotations

import datetime
import json
from pathlib import Path

from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
NAMES = ["Max Mustermann", "Erika Musterfrau", "Jonas Weber", "Lea Sophie Schmidt", "Ali Yilmaz", "Anna-Lena Krüger"]
MAILS = ["max.mustermann@raabeschule.de", "erika.musterfrau@raabeschule.de", "j.weber@schule.de", "lea@example.org", "ali.yilmaz@schule.de", "a.krueger@raabeschule.de"]
CLASSES = ["Klasse 7b", "Klasse 10a", "Klasse 5c", "Klasse 12", "Klasse 8d", "Klasse 6a"]
CODES = ["2026-10-28", "2025-01-05", "2024-12-31", "2026-03-01", "2025-07-14", "2023-11-09"]

# name: (columns of source texts, examples for the first rows)
CASES = {
    "first name": ([NAMES], ["Max"]),
    "last name": ([NAMES], ["Mustermann"]),
    "initial and last": ([NAMES], ["M. Mustermann"]),
    "last, first": ([NAMES], ["Mustermann, Max"]),
    "upper last": ([NAMES], ["MUSTERMANN"]),
    "lower first": ([NAMES], ["max"]),
    "mail user": ([MAILS], ["max.mustermann"]),
    "mail domain": ([MAILS], ["raabeschule.de"]),
    "class": ([CLASSES], ["7b"]),
    "date reordered": ([CODES], ["28.10.2026"]),
    "two columns joined": ([["Max", "Erika", "Jonas", "Lea", "Ali", "Anna"], ["Mustermann", "Musterfrau", "Weber", "Schmidt", "Yilmaz", "Krüger"]], ["Max Mustermann"]),
    "two columns initials": ([["Max", "Erika", "Jonas", "Lea", "Ali", "Anna"], ["Mustermann", "Musterfrau", "Weber", "Schmidt", "Yilmaz", "Krüger"]], ["MM"]),
    "two examples": ([NAMES], ["Max M.", "Erika M."]),
}


def main() -> None:
    x = ExcelTarget(visible=False)
    out = {}
    try:
        ws = x.sheet
        for name, (cols, examples) in CASES.items():
            ws.Cells.Clear()
            n = len(cols[0])
            for j, col in enumerate(cols):
                for i, v in enumerate(col):
                    ws.Cells(i + 1, j + 1).NumberFormat = "@"
                    ws.Cells(i + 1, j + 1).Value = v
            target = len(cols) + 1
            ws.Range(ws.Cells(1, target), ws.Cells(n, target)).NumberFormat = "@"
            for i, e in enumerate(examples):
                ws.Cells(i + 1, target).Value = e
            try:
                ws.Range(ws.Cells(1, target), ws.Cells(n, target)).FlashFill()
                result = [ws.Cells(i + 1, target).Text for i in range(n)]
            except Exception as err:  # noqa: BLE001
                result = f"error: {err.__class__.__name__}"
            out[name] = {"sources": cols, "examples": examples, "result": result}
            print(name, result)
    finally:
        x.close()
    (HERE / "results" / f"{datetime.date.today().isoformat()}-excel-flashfill.json").write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
