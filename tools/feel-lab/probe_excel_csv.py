"""Probe: CSV in German Excel (Step 7, D-061), in an invisible Excel of our own.

1. Saving: a sheet with numbers, formatted numbers, dates, texts with separators, quotes and
   line breaks, booleans, errors and gaps (data starting at B2) is saved as "CSV UTF-8" and
   as "CSV (Trennzeichen-getrennt)", both as the user interface does (Local=True); the bytes
   are recorded (BOM, separator, quoting, line ends, how each value is written).
2. Opening: CSV files written here (UTF-8 with and without BOM, Windows-1252) are opened as
   a double-click does (Local=True); for every field Excel's Value2, Formula, NumberFormat
   (English) and Text are recorded.

Result: packages/sheet-engine/test/fixtures/excel-csv.json (the CSV tests replay it).

    .venv/Scripts/python.exe probe_excel_csv.py
"""

from __future__ import annotations

import datetime
import json
import tempfile
from pathlib import Path

import pythoncom

from feellab.excel import ExcelTarget

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "packages" / "sheet-engine" / "test" / "fixtures" / "excel-csv.json"


def english(obj, name: str, value=None):
    o = obj._oleobj_
    if value is None:
        return o.Invoke(o.GetIDsOfNames(name), 1033, pythoncom.DISPATCH_PROPERTYGET, 1)
    o.Invoke(o.GetIDsOfNames(name), 1033, pythoncom.DISPATCH_PROPERTYPUT, 0, value)
    return None


# Cells for saving: (address, how to set, value, English number format or None).
SAVE = [
    ("B2", "v", "Text", None), ("C2", "v", 1234.5, None), ("D2", "v", -0.25, None), ("E2", "v", 1e-7, None),
    ("F2", "v", 123456789012345.0, None), ("G2", "v", 1 / 3, None),
    ("B3", "v", 0.05, "0%"), ("C3", "v", 1234.5, '#,##0.00 "€"'), ("D3", "v", 46322, "m/d/yyyy"), ("E3", "v", 0.5729166666666666, "h:mm"),
    ("F3", "v", 46322.5729166667, "m/d/yyyy h:mm"), ("G3", "v", 1234.5678, "0.00"), ("H3", "v", 1234567.891, "#,##0.00"),
    ("B4", "v", "Semikolon;darin", None), ("C4", "v", 'Anführungs"zeichen', None), ("D4", "v", "Zeile 1\nZeile 2", None),
    ("E4", "v", "  Leerzeichen  ", None), ("F4", "v", "Grüße ß €", None), ("G4", "v", "Komma, darin", None), ("H4", "v", "'0123", None),
    ("B5", "f", "=1/0", None), ("C5", "v", True, None), ("D5", "f", "=1=2", None), ("E5", "f", '=""', None), ("F5", "f", "=2*21", None),
    ("H6", "v", "rechts außen", None),
    ("B7", "v", 12, None),
]

# Fields for opening (the same line in every file; one field per column).
OPEN = [
    "1,5", "1.000", "1.000,5", "1.5", "1,234.5", "-3", "(3)", "3-", "+4", "5%", "5 %", "1e3", "1E3", "1,5E2", "0123", "  7 ",
    "27.10.2026", "27.10.26", "27.10.", "2026-10-27", "10/27/2026", "13:45", "13:45:30", "27.10.2026 13:45", "Okt 26", "1/2",
    "WAHR", "falsch", "TRUE", "#NV", "#DIV/0!", "=SUMME(1;2)", "=1+1", "5 €", "€ 5", "5,00 €", "$5", "abc", "", '"quoted;semi"',
    '"say ""hi"""', "Grüße", "1 000", "1'000", "0,5", ",5", "-0", "3,14159265358979323846",
]


def save_part(x) -> dict:
    wb = x.app.Workbooks.Add()
    ws = wb.Worksheets(1)
    for addr, how, value, fmt in SAVE:
        c = ws.Range(addr)
        if how == "v":
            c.Value2 = value
        else:
            c.Formula = value
        if fmt:
            english(c, "NumberFormat", fmt)
    ws.Range("D4").WrapText = True
    out = {"cells": {addr: {"set": value, "format": fmt, "text": ws.Range(addr).Text} for addr, how, value, fmt in SAVE}}
    with tempfile.TemporaryDirectory() as td:
        kinds = (("csv_utf8", 62), ("csv", 6))
        for label, fmt_id in kinds:
            wb.SaveAs(str(Path(td) / f"{label}.csv"), fmt_id, Local=True)
        wb.Close(False)
        for label, fmt_id in kinds:
            data = (Path(td) / f"{label}.csv").read_bytes()
            out[label] = {
                "bom": data.startswith(b"\xef\xbb\xbf"),
                "text": data.decode("utf-8-sig") if fmt_id == 62 else data.decode("cp1252"),
                "encoding": "utf-8" if fmt_id == 62 else "cp1252",
            }
    return out


def open_part(x) -> dict:
    line = ";".join(OPEN)
    out: dict = {"fields": OPEN, "files": {}}
    with tempfile.TemporaryDirectory() as td:
        variants = {
            "utf8_bom": b"\xef\xbb\xbf" + line.encode("utf-8"),
            "utf8": line.encode("utf-8"),
            "cp1252": line.encode("cp1252", errors="replace"),
            "comma": ",".join(f.replace(",", ".") if not f.startswith('"') else f for f in OPEN[:6]).encode("utf-8"),
            "multiline": b'Kopf;"zwei\r\nZeilen";Ende\r\n1;2;3\r\n',
        }
        for label, data in variants.items():
            p = Path(td) / f"{label}.csv"
            p.write_bytes(data)
            wb = x.app.Workbooks.Open(str(p), Local=True)
            ws = wb.Worksheets(1)
            used = ws.UsedRange
            cells = []
            for r in range(1, used.Row + used.Rows.Count):
                row = []
                for c in range(1, used.Column + used.Columns.Count):
                    cell = ws.Cells(r, c)
                    v = cell.Value2
                    if isinstance(v, int) and v < 0 and cell.Text.startswith("#"):
                        v = {"error": cell.Text}
                    row.append({"value": v, "formula": cell.Formula, "format": english(cell, "NumberFormat"), "text": cell.Text})
                cells.append(row)
            out["files"][label] = {"rows": cells}
            wb.Close(False)
    return out


def main() -> None:
    x = ExcelTarget(visible=False)
    try:
        result = {
            "date": datetime.date.today().isoformat(),
            "excel": str(x.app.Version),
            "list_separator": x.app.International[4], "decimal": x.app.International[2], "thousands": x.app.International[3],
            "save": save_part(x),
            "open": open_part(x),
        }
    finally:
        x.close()
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(result, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print("written", OUT.name)


if __name__ == "__main__":
    main()
