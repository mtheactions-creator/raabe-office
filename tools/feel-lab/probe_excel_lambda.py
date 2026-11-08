"""Records Excel's LET, LAMBDA and the LAMBDA helpers (MAP, REDUCE, SCAN, NACHZEILE, NACHSPALTE,
MATRIXERSTELLEN, WURDEAUSGELASSEN), defined names holding a LAMBDA and the text the .xlsx file stores
(D-080), as the reference for packages/sheet-engine/test/lambda-excel.test.ts.

Starts an invisible Excel of its own; saves one workbook to TEMP.

    .venv/Scripts/python.exe probe_excel_lambda.py

Writes packages/sheet-engine/test/fixtures/excel-lambda.json. No real input."""

from __future__ import annotations

import json
import re
import tempfile
import time
import zipfile
from pathlib import Path

from feellab.excel import ExcelTarget

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "packages" / "sheet-engine" / "test" / "fixtures" / "excel-lambda.json"
ERRORS = {-2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?", -2146826288: "#NULL!",
          -2146826252: "#NUM!", -2146826265: "#REF!", -2146826273: "#VALUE!", -2146826215: "#SPILL!", -2146826238: "#CALC!"}

DATA = {"A1": 1, "A2": 2, "A3": 3, "B1": 10, "B2": 20, "B3": 30}
NAMES = [("Verdoppeln", "=LAMBDA(x,x*2)"), ("Fakul", "=LAMBDA(n,IF(n<=1,1,n*Fakul(n-1)))"), ("Plus", "=LAMBDA(a,b,IF(ISOMITTED(b),a,a+b))")]
CASES = [
    "=LET(x;2;x*3)",
    "=LET(x;2;y;x+1;x*y)",
    "=LET(a;A1:A3;SUMME(a))",
    "=LET(a;A1:A3;a*2)",
    "=LAMBDA(x;x+1)(2)",
    "=LAMBDA(x;y;x*y)(3;4)",
    "=LAMBDA(x;x)",
    "=LET(f;LAMBDA(x;x*2);f(5))",
    "=MAP(A1:A3;LAMBDA(x;x*2))",
    "=MAP(A1:A3;B1:B3;LAMBDA(a;b;a+b))",
    "=REDUCE(0;A1:A3;LAMBDA(a;b;a+b))",
    "=REDUCE(1;A1:B3;LAMBDA(a;b;a*b))",
    "=SCAN(0;A1:A3;LAMBDA(a;b;a+b))",
    "=SCAN(\"\";A1:A3;LAMBDA(a;b;a&b))",
    "=NACHZEILE(A1:B3;LAMBDA(r;SUMME(r)))",
    "=NACHSPALTE(A1:B3;LAMBDA(c;MAX(c)))",
    "=MATRIXERSTELLEN(2;3;LAMBDA(r;c;r*c))",
    "=LAMBDA(x;y;WENN(WURDEAUSGELASSEN(y);x;x+y))(1;)",
    "=LAMBDA(x;y;WENN(WURDEAUSGELASSEN(y);x;x+y))(1;2)",
    "=LAMBDA(x;x+1)(1;2)",
    "=SUMME(LAMBDA(x;x))",
    "=MAP(A1:A3;LAMBDA(x;{1.2}))",
    "=NACHZEILE(A1:B3;LAMBDA(r;r))",
    "=Verdoppeln(4)",
    "=Fakul(5)",
    "=Plus(3)",
    "=Plus(3;4)",
    "=MAP(A1:A3;Verdoppeln)",
    "=LET(x;5;LAMBDA(y;x+y))(1)",
    "=LET(x;1;x;2;x)",
    "=MATRIXERSTELLEN(0;2;LAMBDA(r;c;r))",
]


def val(v):
    if isinstance(v, int) and v in ERRORS:
        return {"error": ERRORS[v]}
    if isinstance(v, tuple):
        return [[val(x) for x in row] for row in v]
    return v


def main() -> None:
    xl = ExcelTarget(visible=False)
    out: dict = {"data": DATA, "names": NAMES, "cases": [], "file": {}}
    try:
        sh = xl.sheet
        for a, v in DATA.items():
            sh.Range(a).Value = v
        out["names_rejected"] = []
        for n, f in NAMES:
            try:
                xl.wb.Names.Add(Name=n, RefersToLocal=f.replace(",", ";").replace("IF(", "WENN(").replace("ISOMITTED(", "WURDEAUSGELASSEN("))
            except Exception:  # noqa: BLE001
                out["names_rejected"].append(n)
        out["names_local"] = {n: xl.wb.Names(n).RefersToLocal for n, _ in NAMES if n not in out["names_rejected"]}
        for i, f in enumerate(CASES):
            r = sh.Range(f"D{3 * i + 1}")
            try:
                r.Formula2Local = f
            except Exception:  # noqa: BLE001 - Excel refuses the formula
                out["cases"].append({"cell": f"D{3 * i + 1}", "typed": f, "rejected": True})
                continue
            time.sleep(0.02)
            spill = r.HasSpill
            area = r.SpillingToRange if spill else r
            out["cases"].append({"cell": f"D{3 * i + 1}", "typed": f, "local": r.Formula2Local, "formula": r.Formula2,
                                 "value": val(area.Value2), "spill": area.Address.replace("$", "") if spill else None})
        path = Path(tempfile.gettempdir()) / "raabe-lambda.xlsx"
        if path.exists():
            path.unlink()
        xl.wb.SaveAs(str(path), 51)
        with zipfile.ZipFile(path) as z:
            sheet = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
            book = z.read("xl/workbook.xml").decode("utf-8")
        cells = {}
        for m in re.finditer(r'<c r="([A-Z]+\d+)"[^>]*?(?:/>|>(.*?)</c>)', sheet):
            f = re.search(r"<f[^>]*>([^<]*)</f>", m.group(2) or "")
            if f:
                cells[m.group(1)] = f.group(1).replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&").replace("&quot;", '"')
        out["file"]["cells"] = cells
        out["file"]["names"] = {m.group(1): m.group(2).replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&")
                                for m in re.finditer(r'<definedName name="([^"]+)"[^>]*>([^<]*)</definedName>', book)}
    finally:
        xl.close()
    OUT.write_text(json.dumps(out, indent=1, ensure_ascii=False) + "\n", encoding="utf-8", newline="\n")
    for c in out["cases"]:
        print(c.get("local") or c["typed"], "->", c.get("value"), c.get("spill") or "", "REJECTED" if c.get("rejected") else "")
    print(json.dumps(out["file"], ensure_ascii=False, indent=1)[:3000])
    print(out["names_local"])


if __name__ == "__main__":
    main()
