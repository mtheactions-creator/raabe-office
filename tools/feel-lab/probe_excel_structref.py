"""Records Excel's structured references (D-080) as the reference for
packages/sheet-engine/test/structref-excel.test.ts: values, the German form Excel shows
(FormulaLocal), the English form (Formula) and the text the .xlsx file stores, for formulas
outside and inside a table, a calculated column and the Ergebniszeile.

Starts an invisible Excel of its own (`ExcelTarget(visible=False)`); saves one workbook to TEMP.

    .venv/Scripts/python.exe probe_excel_structref.py

Writes packages/sheet-engine/test/fixtures/excel-structref.json. No real input."""

from __future__ import annotations

import json
import re
import tempfile
import time
import zipfile
from pathlib import Path

from feellab.excel import ExcelTarget

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "packages" / "sheet-engine" / "test" / "fixtures" / "excel-structref.json"
ERRORS = {-2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?", -2146826288: "#NULL!",
          -2146826252: "#NUM!", -2146826265: "#REF!", -2146826273: "#VALUE!", -2146826215: "#SPILL!", -2146826201: "#CALC!"}

DATA = [["Name", "Note", "Punkte"], ["Ada", 2, 12], ["Ben", 3, 9], ["Cem", 1, 14], ["Dora", 4, 7]]

# Outside the table (column G and on), German as typed.
OUTSIDE = [
    "=SUMME(Noten[Note])",
    "=ZEILEN(Noten[#Alle])",
    "=ANZAHL2(Noten[[#Kopfzeilen];[Name]])",
    "=SUMME(Noten[[Note]:[Punkte]])",
    "=SUMME(Noten)",
    "=ZEILEN(Noten[#Daten])",
    "=SUMME(Noten[[#Daten];[#Ergebnisse];[Punkte]])",
    "=MITTELWERT(Noten[Punkte])",
    "=SVERWEIS(\"Ben\";Noten;3;FALSCH)",
    "=INDEX(Noten[Name];2)",
    "=ANZAHL(Noten[[#Alle];[Note]])",
    "=SPALTEN(Noten[#Kopfzeilen])",
    "=Noten[[#Kopfzeilen];[Punkte]]",
    "=Noten[@Punkte]*2",
    "=Noten[@Name]",
    "=SUMME(Noten[[#Diese Zeile];[Note]:[Punkte]])",
    "=ZÄHLENWENN(Noten[Note];\"<3\")",
    "=SUMMEWENN(Noten[Name];\"Ben\";Noten[Punkte])",
    "=Noten[Fehlt]",
    "=ZEILEN(Noten[[#Kopfzeilen];[#Daten]])",
    "=Noten[#Ergebnisse]",
]


def val(v):
    if isinstance(v, int) and v in ERRORS:
        return {"error": ERRORS[v]}
    if isinstance(v, tuple):
        return [[val(x) for x in row] for row in v]
    return v


def main() -> None:
    xl = ExcelTarget(visible=False)
    out: dict = {"data": DATA, "outside": [], "inside": [], "file": {}}
    try:
        sh = xl.sheet
        for i, row in enumerate(DATA):
            for k, v in enumerate(row):
                sh.Cells(i + 1, k + 1).Value = v
        lo = sh.ListObjects.Add(1, sh.Range("A1:C5"), None, 1)
        lo.Name = "Noten"
        lo.ShowTotals = True
        lo.ListColumns(3).TotalsCalculation = 1  # Summe
        out["totals_formula"] = {"local": sh.Range("C6").FormulaLocal, "formula": sh.Range("C6").Formula, "value": val(sh.Range("C6").Value2)}
        for i, f in enumerate(OUTSIDE):
            r = sh.Range(f"G{i + 2}")
            try:
                r.Formula2Local = f
            except Exception:  # noqa: BLE001 - Excel refuses the formula (its "Problem with this formula")
                out["outside"].append({"cell": f"G{i + 2}", "typed": f, "rejected": True})
                continue
            time.sleep(0.02)
            out["outside"].append({"cell": f"G{i + 2}", "typed": f, "local": r.Formula2Local, "formula": r.Formula2, "value": val(r.Value2)})
        # A calculated column: a fourth column, one formula typed in its first data cell.
        lc = lo.ListColumns.Add()
        lc.Name = "Doppelt"
        sh.Range("D2").Formula2Local = "=[@Punkte]*2"
        for row in range(2, 6):
            r = sh.Range(f"D{row}")
            out["inside"].append({"cell": f"D{row}", "local": r.Formula2Local, "formula": r.Formula2, "value": val(r.Value2)})
        # The cells right of the table moved one column (rows 1 to 6) and see four columns now.
        for o in out["outside"]:
            if o.get("rejected"):
                continue
            row = int(o["cell"][1:])
            o["after_cell"] = ("H" if row <= 6 else "G") + str(row)
            o["value_after"] = val(sh.Range(o["after_cell"]).Value2)
        out["table_after"] = {"ref": lo.Range.Address.replace("$", ""), "columns": [c.Name for c in lo.ListColumns]}
        path = Path(tempfile.gettempdir()) / "raabe-structref.xlsx"
        if path.exists():
            path.unlink()
        xl.wb.SaveAs(str(path), 51)
        with zipfile.ZipFile(path) as z:
            sheet = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
            table = next(z.read(n).decode("utf-8") for n in z.namelist() if n.startswith("xl/tables/table"))
        cells = {}
        for m in re.finditer(r'<c r="([A-Z]+\d+)"[^>]*?(?:/>|>(.*?)</c>)', sheet):
            f = re.search(r"<f[^>]*>([^<]*)</f>", m.group(2) or "")
            if f:
                cells[m.group(1)] = f.group(1).replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&").replace("&quot;", '"')
        out["file"]["cells"] = cells
        out["file"]["calculatedColumnFormula"] = re.findall(r"<calculatedColumnFormula>([^<]*)</calculatedColumnFormula>", table)
        out["file"]["totalsRowFunction"] = re.findall(r'totalsRowFunction="([^"]*)"', table)
    finally:
        xl.close()
    OUT.write_text(json.dumps(out, indent=1, ensure_ascii=False) + "\n", encoding="utf-8", newline="\n")
    print(json.dumps(out, indent=1, ensure_ascii=False)[:4000])


if __name__ == "__main__":
    main()
