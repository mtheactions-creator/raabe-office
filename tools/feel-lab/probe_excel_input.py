"""Probe: what German Excel makes of typed cell input (Step 7, D-061). In an invisible Excel of our
own each input is assigned with `Range.FormulaLocal` (the German form, as typed) to a cell that is
empty and formatted as General, and again to cells formatted beforehand (Text, Prozent, Datum,
...); recorded are the stored value (Value2 and its type), whether it became a formula and which,
the number format Excel assigned (file form through the English LCID, and German), and the text
shown. FormulaLocal parses like typing in nearly every respect; the exceptions known from Excel's
documentation (automatic percent entry in percent cells, Lotus-style "+1+2" formulas) are left
to the tests. Also recorded: the edit text of numbers (FormulaLocal of a constant) and Excel's
date functions on edge cases (DATE, YEAR … SECOND, WEEKDAY, EDATE, EOMONTH; 1900 and 1904).
Runs invisibly, no input.
Result: packages/sheet-engine/test/fixtures/excel-input.json (the tests replay it).

    .venv/Scripts/python.exe probe_excel_input.py
"""
from __future__ import annotations

import datetime
import json
import sys
from pathlib import Path

import pythoncom

from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
OUT = HERE.parent.parent / "packages" / "sheet-engine" / "test" / "fixtures" / "excel-input.json"

EN = 0x409
GET = pythoncom.DISPATCH_PROPERTYGET
PUT = pythoncom.DISPATCH_PROPERTYPUT


def put(obj, name: str, value, lcid: int = EN) -> None:
    o = obj._oleobj_
    o.Invoke(o.GetIDsOfNames(name), lcid, PUT, 0, value)


def get(obj, name: str, lcid: int = EN):
    o = obj._oleobj_
    return o.Invoke(o.GetIDsOfNames(name), lcid, GET, 1)


NUMBERS = [
    "1", "-1", "+1", "0", "-0", "00", "007", "1,5", "-1,5", "+1,5", ",5", "-,5", "+,5", "1,", "1,50", "1,5000", "1.234",
    "1.234,5", "1.234.567", "1.234.567,89", "-1.234,5", "12.34", "1.2", "1.23", "1,234.5", "1.234.56", "1 234", "1 234,5",
    "1E3", "1e3", "1,5E-3", "1,5e+3", "-1E3", "1E", "E3", "1E400", "1E-400", "1E+300", "1,5E3,5", "1.000E3", "1E03",
    "1.000,00", "(5)", "(5,5)", "(-5)", "5-", "5 -", "- 5", "--5", "+-5", "-+5", "1.5", "1.50", "12.2026", "1,5,5",
    "123456789012345678", "1234567890123456", "123456789012345", "12345678901", "123456789012", "0,1234567890123",
    "1,23456789012345678", "0,000001", "0,0000001", "99999999999", "999999999999", "1000000", "-0,0", "1 000 000",
    " 5", "5 ", " 5 ", "5,5 ", "1.0", "1.000.0", "1..2", "1,,2", ".5", "0.5", "1.", "1.000.",
]
PERCENT = ["100%", "15%", "15 %", "1,5%", "-15%", "%15", "15%%", "1.000%", "0,5%", "12,345%", "1E2%", "(15%)",
           "15% ", "+15%", "15,0%", "15,00%", "1.234,5%", "% 15", "15 %%"]
CURRENCY = ["5 €", "5€", "€5", "€ 5", "-5 €", "5,50 €", "1.234,50 €", "1234,5€", "5 EUR", "$5", "5 $", "5,00 €",
            "-1.234 €", "(5 €)", "5 € ", "-€5", "€-5", "5,5€", "1.234 €", "1.234,567 €", "5 €€", "€", "5 £", "5,123 €",
            "1E3 €", "15% €", "5 € 5", "0 €", "-0,5 €", "€1.234,50", "+5 €"]
FRACTIONS = ["1 1/2", "0 3/4", "3/4", "1/2", "1 1/2 kg", "2 3/8", "-1 1/2", "12/7", "1/0", "0 1/3", "1 2/3", "0 0/1",
             "1 3/2", "10 1/16", "0 1/100", "0 11/100", "1  1/2", "+1 1/2", "1 -1/2", "0 1/1000", "5/4", "1 1/2,5", "0 2/4"]
DATES = [
    "1.1.2026", "01.01.2026", "1.1.26", "1.1.1", "1.1.", "1.1", "31.12.", "31.12.99", "31.12.29", "1.1.30", "1.1.29",
    "29.2.1900", "28.2.1900", "1.1.1900", "31.12.1899", "1.1.10000", "31.12.9999", "29.2.2023", "29.2.2024", "32.1.2026",
    "1.13.2026", "0.1.2026", "1. Jan", "1. Januar", "1 Jan", "1-Jan", "1-Jan-26", "1. Jan 2026", "1. Januar 2026",
    "Jan 26", "Jan 2026", "Januar 2026", "Jan. 26", "Mrz 26", "März 2026", "Mär 26", "Mai 26", "Okt 26", "Dez 26",
    "Oct 26", "Dec 26", "May 26", "2026-01-01", "2026-1-1", "2026/01/01", "1/1/2026", "1/1/26", "1/1", "1-1-2026", "1-1",
    "26-1-1", "13/1/2026", "1/13/2026", "Mo, 1.1.2026", "Montag, 5. Januar 2026", "5. Januar", "Januar", "Jan", "1.1.2026.",
    "1. 1. 2026", "1 .1.2026", "1.1.26 ", " 1.1.26", "2026-13-01", "12.2026", "1.2026", "1.26", "13.26", "12.99",
    "5.2.", "15.3", "3.15", "1-Feb", "Feb-1", "Feb 1", "1 Feb 26", "1 Februar 2026", "1.Jan", "1.Jan.26", "1.Jan.2026",
    "1/Jan/2026", "Jan 1, 2026", "1 Jan.", "1. Jan.", "1. Mrz.", "1. Mär.", "Sep 26", "Sept 26", "Juni 26", "Jun 26",
    "Juli 26", "Jul 26", "1. März", "1. Maerz", "jan 26", "JAN 26", "Jan 0", "Jan 100", "Jan 1900", "Jan 1899",
    "1.1.0", "1.1.00", "1.1.100", "1.1.999", "1.1.1000", "1.1.2", "2.29.2024",
]
TIMES = ["12:30", "12:30:15", "12:30:15,5", "0:00", "24:00", "25:00", "100:00", "12:60", "1:2", "12:30:61", ":30", "12:",
         "8:00 PM", "8:00 pm", "8 PM", "8:00PM", "8:00 AM", "12:00 AM", "12:00 PM", "8:00 a", "8:00 p", "30:15,5",
         "0:30,5", "1:30:45,123", "13:00 PM", "9999:00", "10000:00", "-1:00", "12:30:15 PM", "1:2:3", "01:02:03", "1:60",
         "0:59:60", "23:59:59,9999", "8 Uhr", "12.30", "12:30 Uhr", "12h30", "8 AM", "8 am", "0:0", "00:00:00",
         "12:30,5", "1:30,25", "100:30:30"]
DATETIMES = ["1.1.2026 12:30", "1.1.2026 12:30:15", "1.1.26 8:00", "1.1. 12:30", "2026-01-01 12:30", "1.1.2026 8 PM",
             "12:30 1.1.2026", "1.1.2026 25:00", "1. Jan 2026 12:30", "Jan 26 12:30", "1.1.2026  12:30", "1.1.2026 8:00 PM",
             "1/1/2026 12:30", "2026-01-01T12:30", "1.1.2026, 12:30"]
BOOLS = ["wahr", "WAHR", "Wahr", "falsch", "FALSCH", "true", "TRUE", "false", " wahr", "wahr ", "WAHR()", "ja"]
TEXTS = ["abc", "a1", "Hallo Welt", "'123", "'=1+2", "'", "''x", "1a", "a 1", " ", "1,5 kg", "#NV", "#WERT!", "#DIV/0!",
         "#N/A", "#BEZUG!", "#NAME?", "#ZAHL!", "#NULL!", "#ÜBERLAUF!", "x%", "€€", "- ", "+", "-", ".", ",", "1-2",
         "1+2", "1*2", "2^3", "A1", "SUMME(1)", "1:2:3:4", "1.1.1.1", "1/2/3/4", "€ 5 €", "5 kg", "+ 5", "(5", "5)",
         "'  x", "'1.1.2026", "´5", "0x10", "1_000", "١٢"]
FORMULAS = ["=1+2", "=1,5+1", "+1+2", "-1+2", "+5", "-5", "=SUMME(1;2)", "=A1", "+A1", "-A1", "=wahr", "=1/0", "=",
            "==1", "=\"a\"", "=1.5", "=1;2", "+", "=1+", "=SUMME(", "=summe(1;2)", "=A1*2", "+-1", "=-1", "= 1", "=1 ",
            "@SUMME(1;2)", "=15%", "=1E3", "=HEUTE()", "=DATUM(2026;1;1)", "=\"1\"+1", "=A1:B2", "=1,5", "=1.000",
            "-A1+1", "+1,5", "-1,5", "+1.5", "=Tabelle1!A1", "=JETZT()", "*5", "/5", "=1 + 2"]

GENERAL = NUMBERS + PERCENT + CURRENCY + FRACTIONS + DATES + TIMES + DATETIMES + BOOLS + TEXTS + FORMULAS

# Inputs typed into cells formatted beforehand (file form of the format; German in the comment).
PRE_FORMATS = ["@", "0%", "0.00%", "0.00", "#,##0", '#,##0.00\\ "€"', "m/d/yyyy", "dd/mm/yyyy", "h:mm", "[h]:mm:ss",
               "0.00E+00", "# ?/?", "m/d/yyyy h:mm", "mmmm yyyy", "0", '_-* #,##0.00\\ "€"_-;\\-* #,##0.00\\ "€"_-;_-* "-"??\\ "€"_-;_-@_-',
               "d-mmm", "General", ";;;", "[Red]0.00"]
PRE_INPUTS = ["15", "0,5", "1,5", "-1", "0", "15%", "1.1.2026", "1.1.", "12:30", "1.1.2026 12:30", "5 €", "abc", "1 1/2",
              "1E3", "wahr", "=1+2", "-5", "1.234", "1.234,5", "1000", "0,05", "100", "1,0", "0,999", "1", "45000",
              "Jan 26", "25:00", "'15", "3/4", "12,5", "-0,5", "1,5E3", "(5)"]


# Excel's date functions on edge cases, for the serial arithmetic in dates.ts.
DATE_Y = [0, 1, 99, 100, 1899, 1900, 1901, 1904, 2000, 2024, 2026, 9999, 10000, -1]
DATE_M = [-13, -1, 0, 1, 2, 3, 12, 13, 25]
DATE_D = [-1, 0, 1, 28, 29, 30, 31, 32, 60, 400]
SERIALS = [0, 1, 59, 60, 61, 366, 1462, 36526, 45000.5, 45000.99999, 46022.75, 2958465, 2958465.99999, 0.999994, 1.5]
WEEKDAY_TYPES = [1, 2, 3, 11, 12, 13, 14, 15, 16, 17]
MONTH_STEPS = [-25, -1, 0, 1, 11, 12, 13, 120]


def date_functions(sh) -> dict:
    """DATE over a grid, and YEAR/MONTH/DAY/WEEKDAY/HOUR/MINUTE/SECOND/EDATE/EOMONTH of serials."""
    cases = []
    for y in DATE_Y:
        for m in DATE_M:
            for d in DATE_D:
                cases.append(("DATE", [y, m, d], f"=DATE({y},{m},{d})"))
    for s in SERIALS:
        for fn in ("YEAR", "MONTH", "DAY", "HOUR", "MINUTE", "SECOND"):
            cases.append((fn, [s], f"={fn}({s})"))
        for t in WEEKDAY_TYPES:
            cases.append(("WEEKDAY", [s, t], f"=WEEKDAY({s},{t})"))
        for n in MONTH_STEPS:
            cases.append(("EDATE", [s, n], f"=EDATE({s},{n})"))
            cases.append(("EOMONTH", [s, n], f"=EOMONTH({s},{n})"))
    out = []
    for k, (fn, args, f) in enumerate(cases):
        c = sh.Range(f"A{k + 1}")
        put(c, "Formula", f)
        v = c.Value2
        if isinstance(v, int) and v < -2000000000:
            v = "error"
        out.append({"fn": fn, "args": args, "value": v})
    sh.Cells.Clear()
    return out


# The text Excel shows in the formula bar (and on F2) for a stored value: FormulaLocal of a cell
# holding a constant, per number format.
EDIT_VALUES = [0, 1, -1, 1 / 3, 0.15, 1234.5, 45000.5, 45658.75, 0.5208333333333334, 1e20, 1e-20, 123456789012345,
               2 / 3 * 1e10, 1.5, 46023, 46023.520833333336, 2958465.5, -0.25, 12.345, 0.000123, 1.0416666666666667,
               0.999994, 59, 60, 61, 0.1 + 0.2, 1234567.891, -1234.5678, 0.5, 25.5, 100, 1e15, 1e16, 99999999999999.9,
               1e21, 1e22, 1.5e25, 1.23456789e30, 1e100, -1e100, 1e300, 1e-5, 1.5e-6, 1e-9, 1.23e-10, 1.234567e-7,
               1.23456e-8, 0.0001, 0.00001234, 123456789012345678, 1234567890123456789012, 2 / 3, -2 / 3 * 1e-5,
               9.99999999999999e14, 1.1e-15, 3.14159265358979e-12]
# FormulaLocal shows the number the same way whatever the format (recorded for three formats).
EDIT_FORMATS = ["General", "m/d/yyyy", "0%"]


def edit_texts(sh) -> list:
    out = []
    n = len(EDIT_VALUES)
    for j, fmt in enumerate(EDIT_FORMATS):
        texts = []
        for i, v in enumerate(EDIT_VALUES):
            c = sh.Range(f"A{i + 1}")
            c.ClearContents()
            put(c, "NumberFormat", fmt)
            c.Value2 = v
            texts.append(c.FormulaLocal)
        out.append({"code": fmt, "edit": texts})
        sh.Range(f"A1:A{n}").Clear()
    return out


def record(cell) -> dict:
    v = cell.Value2
    rec = {"value": v if not isinstance(v, (datetime.datetime,)) else str(v), "type": type(v).__name__,
           "formula": bool(cell.HasFormula), "en": get(cell, "NumberFormat"), "local": cell.NumberFormatLocal,
           "text": cell.Text}
    if isinstance(v, int) and not isinstance(v, bool) and v < -2000000000:
        rec["type"] = "error"
        rec["error"] = v
    if cell.HasFormula:
        rec["f"] = get(cell, "Formula")
        rec["fl"] = cell.FormulaLocal
    rec["edit"] = cell.FormulaLocal
    if cell.PrefixCharacter:
        rec["prefix"] = cell.PrefixCharacter
    return rec


def run_inputs(sh, inputs: list, fmt: str | None) -> list:
    out = []
    for k, t in enumerate(inputs):
        c = sh.Range(f"A{k + 1}")
        if fmt is not None:
            put(c, "NumberFormat", fmt)
        rec: dict = {"in": t}
        try:
            c.FormulaLocal = t
            rec.update(record(c))
        except Exception as e:  # noqa: BLE001
            rec["rejected"] = True
            rec["error"] = str(e)[:80]
        out.append(rec)
    sh.Cells.Clear()
    return out


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    x = ExcelTarget(visible=False)
    out: dict = {"date": str(datetime.date.today()), "excel": x.app.Version}
    try:
        app = x.app
        out["options"] = {"autoPercentEntry": bool(app.AutoPercentEntry), "date1904": bool(x.wb.Date1904),
                          "twoDigitYearCutoff": "Windows setting (Excel uses 1930-2029 by default)"}
        sh = x.wb.Worksheets(1)
        sh.Cells.ColumnWidth = 60
        sh.Range("A1:A1000").ColumnWidth = 60
        print("general", len(GENERAL))
        out["general"] = run_inputs(sh, GENERAL, None)
        out["formatted"] = {}
        for fmt in PRE_FORMATS:
            sh.Cells.ColumnWidth = 60
            print("format", fmt)
            out["formatted"][fmt] = run_inputs(sh, PRE_INPUTS, fmt)
        print("edit texts")
        out["editValues"] = EDIT_VALUES
        out["editTexts"] = edit_texts(sh)
        print("date functions")
        sh.Cells.ColumnWidth = 20
        out["dateFunctions"] = date_functions(sh)
        x.wb.Date1904 = True
        out["dateFunctions1904"] = date_functions(sh)
        x.wb.Date1904 = False
    finally:
        x.close()
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))
    print("written", OUT.name)


if __name__ == "__main__":
    main()
