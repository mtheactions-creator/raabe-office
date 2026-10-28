"""Probe: Excel's Datenüberprüfung (data validation) through COM, in an invisible Excel of our own,
for RaabeCalc's validation (apps/raabe-calc/src/validation.ts). No real input, no windows.

1. The Validation object's defaults after `Validation.Add` (titles, messages, flags) and how
   Formula1/Formula2 read back (German local form).
2. Excel's verdict on entries: each case puts a validation on B2, enters a text through
   FormulaLocal (German parsing, as typed) and reads `Validation.Value` (valid or not).
3. The file form: each kind of validation saved as .xlsx, its <dataValidation> XML (and the x14
   form Excel uses for lists on other sheets).

    .venv/Scripts/python.exe probe_excel_validation.py

Result: results/<date>-excel-validation.json (labels, values and XML only).
"""

from __future__ import annotations

import datetime
import json
import re
import tempfile
import zipfile
from pathlib import Path

from feellab.excel import ExcelTarget

HERE = Path(__file__).parent

# Excel's constants (XlDVType, XlDVAlertStyle, XlFormatConditionOperator).
TYPES = {"any": 0, "whole": 1, "decimal": 2, "list": 3, "date": 4, "time": 5, "textLength": 6, "custom": 7}
OPS = {"between": 1, "notBetween": 2, "equal": 3, "notEqual": 4, "greaterThan": 5, "lessThan": 6, "greaterThanOrEqual": 7, "lessThanOrEqual": 8}
STYLES = {"stop": 1, "warning": 2, "information": 3}

# (name, type, operator, formula1, formula2, ignoreBlank, context cells, typed entries)
CASES = [
    ("whole between 1 6", "whole", "between", "1", "6", True, {}, ["3", "3,5", "0", "7", "6", "1", "abc", "WAHR", "=2+2", "=2,5", "6,0", "50%", "1.1.2026", "-1", " 3", "'3"]),
    ("whole notBetween 1 6", "whole", "notBetween", "1", "6", True, {}, ["0", "3", "7", "7,5", "abc"]),
    ("whole equal 5", "whole", "equal", "5", None, True, {}, ["5", "4", "5,0", "abc"]),
    ("whole notEqual 5", "whole", "notEqual", "5", None, True, {}, ["5", "4", "4,5"]),
    ("whole greaterThan 5", "whole", "greaterThan", "5", None, True, {}, ["5", "6", "100000"]),
    ("whole lessThan 5", "whole", "lessThan", "5", None, True, {}, ["5", "4", "-100"]),
    ("whole greaterThanOrEqual 5", "whole", "greaterThanOrEqual", "5", None, True, {}, ["5", "4"]),
    ("whole lessThanOrEqual 5", "whole", "lessThanOrEqual", "5", None, True, {}, ["5", "6"]),
    ("decimal between 1,5 2,5", "decimal", "between", "1,5", "2,5", True, {}, ["2", "1,5", "2,6", "abc", "WAHR", "=1+1", "15.3.2026"]),
    ("decimal greaterThan ref", "decimal", "greaterThan", "=$D$1", None, True, {"D1": "10"}, ["11", "9", "10"]),
    ("decimal between blank ref, ignore", "decimal", "between", "=$D$1", "=$D$2", True, {"D1": "1"}, ["5", "0", "abc"]),
    ("decimal between blank ref, no ignore", "decimal", "between", "=$D$1", "=$D$2", False, {"D1": "1"}, ["5", "0", "abc"]),
    ("list literal", "list", None, "Ja;Nein;Vielleicht", None, True, {}, ["Ja", "ja", "JA", "Nein", "Nein ", "x", "1"]),
    ("list literal numbers", "list", None, "1;2;2,5", None, True, {}, ["1", "2", "2,5", "2,50", "3", "abc"]),
    ("list range", "list", None, "=$D$1:$D$4", None, True, {"D1": "Rot", "D2": "Grün", "D3": "1,5", "D4": "15.03.2026"}, ["Rot", "rot", "Grün", "Blau", "1,5", "1,50", "15.03.2026", "46096"]),
    ("list range with blank, ignore", "list", None, "=$D$1:$D$3", None, True, {"D1": "Rot", "D3": "Blau"}, ["Rot", "Blau", "Gelb"]),
    ("list range with blank, no ignore", "list", None, "=$D$1:$D$3", None, False, {"D1": "Rot", "D3": "Blau"}, ["Rot", "Blau", "Gelb"]),
    ("date between", "date", "between", "01.01.2026", "31.12.2026", True, {}, ["15.03.2026", "1.1.2025", "31.12.2026", "abc", "46100", "15.03.2026 12:00", "31.12.2026 12:00", "WAHR"]),
    ("date equal", "date", "equal", "15.03.2026", None, True, {}, ["15.03.2026", "15.03.2026 12:00", "16.03.2026"]),
    ("date greaterThan", "date", "greaterThan", "15.03.2026", None, True, {}, ["15.03.2026", "15.03.2026 00:01", "16.03.2026"]),
    ("time between", "time", "between", "08:00", "12:00", True, {}, ["09:30", "13:00", "0,4", "abc", "1.1.2026 09:00", "08:00", "12:00:01", "WAHR"]),
    ("time lessThan", "time", "lessThan", "12:00", None, True, {}, ["11:59", "12:00", "0"]),
    ("textLength between 2 5", "textLength", "between", "2", "5", True, {}, ["a", "ab", "abcdef", "12345", "123456", "1,5", "15.03.2026", "=WIEDERHOLEN(\"x\";7)", "WAHR", "1234,5678", "0,1"]),
    ("textLength lessThanOrEqual 3", "textLength", "lessThanOrEqual", "3", None, True, {}, ["abc", "abcd", "äöü", "123", "1000"]),
    ("custom isnumber", "custom", None, "=ISTZAHL(B2)", None, True, {}, ["5", "abc", "=1/0", "WAHR"]),
    ("custom compare", "custom", None, "=B2>C2", None, True, {"C2": "3"}, ["5", "3", "abc"]),
    ("custom len", "custom", None, "=LÄNGE(B2)<4", None, True, {}, ["abc", "abcd"]),
    ("custom text result", "custom", None, "=\"x\"", None, True, {}, ["5"]),
    ("custom number result", "custom", None, "=2", None, True, {}, ["5"]),
    ("custom zero result", "custom", None, "=0", None, True, {}, ["5"]),
    ("custom error result", "custom", None, "=1/0", None, True, {}, ["5"]),
    ("custom blank ref, ignore", "custom", None, "=B2>D9", None, True, {}, ["5", "-5"]),
    ("custom blank ref, no ignore", "custom", None, "=B2>D9", None, False, {}, ["5", "-5"]),
    ("list literal spaces", "list", None, "Ja; Nein ;Vielleicht", None, True, {}, ["Ja", "Nein", " Nein", "Nein ", "Vielleicht"]),
    ("list range numbers formatted", "list", None, "=$D$1:$D$2", None, True, {"D1": "1,50 €", "D2": "50%"}, ["1,5", "1,50 €", "0,5", "50%", "Rot"]),
    ("whole formula text", "whole", "between", "1", "6", True, {}, ["=\"3\"", "=WAHR"]),
]


def text_of(v):
    return v if isinstance(v, (str, int, float, bool)) or v is None else str(v)


def props(v) -> dict:
    out = {}
    for name in ("Type", "AlertStyle", "Operator", "Formula1", "Formula2", "IgnoreBlank", "InCellDropdown", "ShowInput", "ShowError",
                 "InputTitle", "InputMessage", "ErrorTitle", "ErrorMessage", "Value"):
        try:
            out[name] = text_of(getattr(v, name))
        except Exception as e:  # noqa: BLE001 - some properties fail for some types
            out[name] = f"<error {str(e)[:60]}>"
    return out


def add(ws, addr: str, kind: str, op: str | None, f1: str | None, f2: str | None, style: str = "stop"):
    v = ws.Range(addr).Validation
    v.Delete()
    args = [TYPES[kind], STYLES[style], OPS[op] if op else 1]
    if f1 is not None:
        args.append(f1)
        if f2 is not None:
            args.append(f2)
    v.Add(*args)
    return v


def sheet_xml(path: Path) -> tuple[list[str], list[str]]:
    with zipfile.ZipFile(path) as z:
        xml = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
    main = re.findall(r"<dataValidation\b.*?(?:</dataValidation>|/>)", xml, re.S)
    ext = re.findall(r"<ext\b[^>]*\{CCE6A557[^>]*>.*?</ext>", xml, re.S)
    return main, ext


def main() -> None:
    out: dict = {"date": datetime.date.today().isoformat(), "defaults": {}, "cases": [], "files": {}}
    x = ExcelTarget(visible=False)
    try:
        app = x.app
        wb = x.wb
        ws = wb.Worksheets(1)
        intl = app.International  # all settings as a tuple (xlListSeparator 5, xlDecimalSeparator 3)
        out["listSeparator"] = intl[4]
        out["decimalSeparator"] = intl[2]
        # 1. Defaults of a fresh validation of each type.
        for kind in TYPES:
            f1 = {"any": None, "list": "Ja;Nein", "date": "01.01.2026", "time": "08:00", "custom": "=ISTZAHL(A1)"}.get(kind, "1")
            f2 = {"whole": "6", "decimal": "6", "date": "31.12.2026", "time": "12:00", "textLength": "6"}.get(kind)
            out["defaults"][kind] = props(add(ws, "A1", kind, "between", f1, f2))
        ws.Range("A1").Validation.Delete()
        print("defaults", json.dumps(out["defaults"]["whole"], ensure_ascii=False))
        # 2. Excel's verdicts.
        for name, kind, op, f1, f2, ignore, ctx, entries in CASES:
            ws.Cells.Clear()
            ws.Cells.Validation.Delete()
            for a, t in ctx.items():
                ws.Range(a).FormulaLocal = t
            try:
                v = add(ws, "B2", kind, op, f1, f2)
            except Exception as e:  # noqa: BLE001 - Excel refuses some formulas (an error result)
                out["cases"].append({"name": name, "type": kind, "formula1": f1, "refused": str(e)[:80]})
                print(name, "refused by Excel")
                continue
            v.IgnoreBlank = ignore
            res = []
            for e in entries:
                c = ws.Range("B2")
                c.Clear()
                c.Validation.Delete()
                v = add(ws, "B2", kind, op, f1, f2)
                v.IgnoreBlank = ignore
                c.FormulaLocal = e
                ok = c.Validation.Value
                res.append({"entry": e, "valid": bool(ok), "value": text_of(c.Value2), "text": c.Text, "formula": c.Formula})
            out["cases"].append({"name": name, "type": kind, "operator": op, "formula1": f1, "formula2": f2, "ignoreBlank": ignore, "context": ctx,
                                 "read": {"Formula1": text_of(v.Formula1), "Formula2": text_of(safe(lambda: v.Formula2))}, "results": res})
            print(name, " ".join(f"{r['entry']}={'ok' if r['valid'] else 'X'}" for r in res))
        # 3. The file form of each kind.
        with tempfile.TemporaryDirectory() as td:
            kinds = [
                ("list literal", "list", None, "Ja;Nein", None, {}),
                ("list literal numbers", "list", None, "1;2;2,5", None, {}),
                ("list literal comma", "list", None, "a,b;c", None, {}),
                ("list literal quote", "list", None, 'sagt "hallo";x', None, {}),
                ("list range", "list", None, "=$D$1:$D$3", None, {}),
                ("list other sheet", "list", None, "=Listen!$A$1:$A$3", None, {}),
                ("list name", "list", None, "=Faecher", None, {}),
                ("whole between", "whole", "between", "1", "6", {}),
                ("decimal greaterThan", "decimal", "greaterThan", "1,5", None, {"style": "warning"}),
                ("date between", "date", "between", "01.01.2026", "31.12.2026", {"style": "information"}),
                ("time lessThan", "time", "lessThan", "12:00", None, {}),
                ("textLength lessThanOrEqual", "textLength", "lessThanOrEqual", "10", None, {}),
                ("custom", "custom", None, "=ISTZAHL(B2)", None, {}),
                ("custom other sheet", "custom", None, "=B2<=Listen!$B$1", None, {}),
                ("any with message", "any", None, None, None, {"input": ("Hinweis", "Bitte eine Zahl")}),
                ("hidden dropdown, no error", "list", None, "Ja;Nein", None, {"dropdown": False, "showError": False}),
                ("messages", "whole", "between", "1", "6", {"input": ("Note", "1 bis 6"), "error": ("Falsch", "Nur 1 bis 6")}),
            ]
            for name, kind, op, f1, f2, opt in kinds:
                b = app.Workbooks.Add()
                s = b.Worksheets(1)
                lst = b.Worksheets.Add(After=s)
                lst.Name = "Listen"
                for i, t in enumerate(["Deutsch", "Mathe", "Physik"]):
                    lst.Cells(i + 1, 1).Value2 = t
                lst.Range("B1").Value2 = 10
                b.Names.Add("Faecher", "=Listen!$A$1:$A$3")
                s.Activate()
                v = add(s, "B2:B5", kind, op, f1, f2, opt.get("style", "stop"))
                if "input" in opt:
                    v.InputTitle, v.InputMessage = opt["input"]
                if "error" in opt:
                    v.ErrorTitle, v.ErrorMessage = opt["error"]
                if "dropdown" in opt:
                    v.InCellDropdown = opt["dropdown"]
                if "showError" in opt:
                    v.ShowError = opt["showError"]
                p = Path(td) / "v.xlsx"
                if p.exists():
                    p.unlink()
                b.SaveAs(str(p), 51)
                b.Close(False)
                main_xml, ext = sheet_xml(p)
                # How Excel reads its own file back (German form).
                b = app.Workbooks.Open(str(p))
                back = b.Worksheets(1).Range("B2").Validation
                reread = {"Formula1": text_of(safe(lambda: back.Formula1)), "Formula2": text_of(safe(lambda: back.Formula2))}
                b.Close(False)
                # The uid is random per save: not part of the record.
                strip = lambda t: re.sub(r' xr:uid="\{[^}]*\}"', "", t)  # noqa: E731
                out["files"][name] = {"main": [strip(m) for m in main_xml], "x14": [strip(e) for e in ext], "reread": reread}
                print(name, out["files"][name])
    finally:
        x.close()
    (HERE / "results" / f"{out['date']}-excel-validation.json").write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


def safe(fn):
    try:
        return fn()
    except Exception:  # noqa: BLE001
        return None


if __name__ == "__main__":
    main()
