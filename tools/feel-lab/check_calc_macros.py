"""Makros in RaabeCalc against Excel (D-080 B, D-091), without real input. For every case of macro_cases_excel.py:
1. RaabeCalc (off screen, never focused) builds the case's workbook, gets the macro as a VBA module and saves an .xlsm
   (so Excel also checks that it opens our VBA project), then runs the macro itself with its own interpreter;
2. our own invisible Excel opens RaabeCalc's .xlsm and runs the same macro (Application.Run; Excel started by
   automation runs macros without any change to the Trust Center). MsgBox and InputBox would wait for a click in the
   hidden Excel, so Excel's copy of the project gets a second module that answers them from the case's list (VBA
   takes a project's own MsgBox before the built-in one); RaabeCalc runs the original code in its test mode.
3. both workbooks are compared: sheets (names, order, visibility, the active one), every cell's FormulaLocal and
   shown text, fill and font where the case asks, hidden rows and columns, the messages shown, errors.
Result: results/<date>-calc-macros.json (cases and differences only).

    .venv/Scripts/python.exe check_calc_macros.py [case ...]
"""
from __future__ import annotations

import datetime
import json
import os
import re
import subprocess
import sys
import threading
import tempfile
from pathlib import Path

from feellab import calc_catalogue, winapi
from feellab.calc_catalogue import CalcSide
from feellab.excel import ExcelTarget
from macro_cases_excel import CASES

HERE = Path(__file__).parent
TMP = Path(tempfile.gettempdir()) / "raabe-calc-macros"
TMP.mkdir(exist_ok=True)

os.environ["RAABE_FEEL_BACKGROUND"] = "1"
calc_catalogue.winapi.place = lambda hwnd, x, y, w, h: winapi.place_quietly(hwnd, -4000, 0, w, h)

NL = chr(10)
# Answers MsgBox and InputBox in Excel's copy from a list, records their prompts (chr(1) between them).
SHIM = NL.join([
    "Public ShimAnswers As Variant, ShimPos As Long, ShimLog As String, ShimErr As String",
    "Sub ShimSet(a As Variant)",
    "    ShimAnswers = a: ShimPos = 0: ShimLog = \"\"",
    "End Sub",
    "Sub ShimRun()",
    "    ShimErr = \"\"",
    "    On Error GoTo f",
    "    MACRO",
    "    Exit Sub",
    "f:",
    "    ShimErr = Err.Number & \" \" & Err.Description",
    "End Sub",
    "Function ShimGetErr() As String",
    "    ShimGetErr = ShimErr",
    "End Function",
    "Function ShimGetLog() As String",
    "    ShimGetLog = ShimLog",
    "End Function",
    "Private Function ShimNext(d As Variant) As Variant",
    "    ShimNext = d",
    "    If IsArray(ShimAnswers) Then",
    "        If ShimPos <= UBound(ShimAnswers) Then ShimNext = ShimAnswers(ShimPos): ShimPos = ShimPos + 1",
    "    End If",
    "End Function",
    "Function MsgBox(Prompt As Variant, Optional Buttons As Long = 0, Optional Title As Variant, Optional HelpFile As Variant, Optional Context As Variant) As Long",
    "    ShimLog = ShimLog & Prompt & Chr(1)",
    "    Dim d As Long",
    "    Select Case Buttons And 7",
    "        Case 2: d = 3",
    "        Case 3, 4: d = 6",
    "        Case 5: d = 4",
    "        Case Else: d = 1",
    "    End Select",
    "    MsgBox = ShimNext(d)",
    "End Function",
    "Function InputBox(Prompt As Variant, Optional Title As Variant, Optional Default As Variant = \"\", Optional XPos As Variant, Optional YPos As Variant, Optional HelpFile As Variant, Optional Context As Variant) As String",
    "    ShimLog = ShimLog & Prompt & Chr(1)",
    "    InputBox = CStr(ShimNext(Default))",
    "End Function",
    "Function XInputBox(Prompt As Variant, Optional Title As Variant, Optional Default As Variant = \"\", Optional Left As Variant, Optional Top As Variant, Optional HelpFile As Variant, Optional HelpContextID As Variant, Optional Kind As Variant = 2) As Variant",
    "    ShimLog = ShimLog & Prompt & Chr(1)",
    "    Dim t As Variant: t = ShimNext(Default)",
    "    If Kind = 1 Then",
    "        XInputBox = CDbl(t)",
    "    ElseIf Kind = 8 Then",
    "        Set XInputBox = Range(t)",
    "    Else",
    "        XInputBox = CStr(t)",
    "    End If",
    "End Function",
])

ERRORS = {-2146826281: "#DIV/0!", -2146826246: "#NV", -2146826259: "#NAME?", -2146826288: "#NULL!",
          -2146826252: "#ZAHL!", -2146826265: "#BEZUG!", -2146826273: "#WERT!"}


def bgr_hex(n) -> str | None:
    if n is None:
        return None
    n = int(n)
    return "#%02X%02X%02X" % (n & 255, (n >> 8) & 255, (n >> 16) & 255)


def setup_of(case: dict) -> dict:
    out = {}
    for k, v in case["setup"].items():
        a = k.split("!")[-1]
        out[a] = ("WAHR" if v else "FALSCH") if isinstance(v, bool) else str(v).replace(".", ",") if isinstance(v, float) else str(v)
    return out


def calc_side(c: CalcSide, case: dict, xlsm: Path) -> dict:
    setup = json.dumps(setup_of(case))
    excel_code = re.sub(r"Application\.InputBox\s*\(", "XInputBox(", case["code"], flags=re.I)
    # "Type" is a reserved word in VBA, so the stand-in calls Application.InputBox's parameter Kind.
    excel_code = re.sub(r"\bType:=", "Kind:=", excel_code)
    c.js(f"__calc.reset({setup}, 'A1', 1)")
    c.js(f"__calc.setMacros({json.dumps([{'name': 'Modul1', 'code': excel_code}, {'name': 'Antworten', 'code': SHIM.replace('    MACRO', '    ' + case['macro'])}])})")
    if xlsm.exists():
        xlsm.unlink()
    c.js(f"__calc.save({json.dumps(str(xlsm))})")
    c.js(f"__calc.reset({setup}, 'A1', 1)")
    c.js(f"__calc.setMacros({json.dumps(case['code'])})")
    run = c.js(f"__calc.runMacro({json.dumps(case['macro'])}, {json.dumps(case.get('answers', []))})")
    st = c.js("__calc.macroState()")
    st["shown"] = run.get("shown", [])
    st["error"] = run.get("error")
    return st


def run_us(app, *args):
    """Application.Run with the US locale, as VBA itself calls Excel. Through COM with the German locale Excel takes
    the macro's Formula and NumberFormat as local text (=SUM gives #NAME?), which VBA in Excel never sees."""
    import pythoncom
    dispid = app._oleobj_.GetIDsOfNames("Run")
    return app._oleobj_.Invoke(dispid, 1033, pythoncom.DISPATCH_METHOD, True, *args)


def excel_side(x: ExcelTarget, case: dict, xlsm: Path) -> dict:
    app = x.app
    app.DisplayAlerts = False
    wb = app.Workbooks.Open(str(xlsm))
    try:
        if wb.HasVBProject is False:
            return {"open_error": "Excel finds no VBA project in RaabeCalc's file"}
        run_us(app, f"'{wb.Name}'!ShimSet", list(case.get("answers", [])) or "")
        # The macro runs inside VBA's own error handler: an error is reported, not shown in a dialog nobody sees.
        # Called by name inside the module (an error in a macro started with Application.Run is not trappable).
        run_us(app, f"'{wb.Name}'!ShimRun")
        err = run_us(app, f"'{wb.Name}'!ShimGetErr") or ""
        error = {"message": err} if err else None
        log = run_us(app, f"'{wb.Name}'!ShimGetLog") or ""
        sheets = []
        for i in range(1, wb.Worksheets.Count + 1):
            ws = wb.Worksheets(i)
            cells = {}
            ur = ws.UsedRange
            r0, c0 = ur.Row, ur.Column
            for r in range(r0, r0 + min(ur.Rows.Count, 80)):
                for col in range(c0, c0 + min(ur.Columns.Count, 30)):
                    cell = ws.Cells(r, col)
                    f = cell.FormulaLocal
                    fill = None if cell.Interior.ColorIndex == -4142 else bgr_hex(cell.Interior.Color)
                    fc = bgr_hex(cell.Font.Color)
                    bold, italic = bool(cell.Font.Bold), bool(cell.Font.Italic)
                    if f == "" and fill is None and not bold and fc == "#000000":
                        continue
                    cells[cell.GetAddress(False, False)] = {"formula": f, "text": cell.Text, "fill": fill, "fontColor": None if fc == "#000000" else fc,
                                                          "bold": bold, "italic": italic}
            last_r = r0 + ur.Rows.Count + 5
            hidden_rows = [r for r in range(1, last_r) if ws.Rows(r).Hidden]
            hidden_cols = [k for k in range(1, c0 + ur.Columns.Count + 5) if ws.Columns(k).Hidden]
            sheets.append({"name": ws.Name, "visible": int(ws.Visible), "cells": cells, "hiddenRows": hidden_rows, "hiddenCols": hidden_cols,
                           "protected": bool(ws.ProtectContents)})
        return {"sheets": sheets, "active": wb.ActiveSheet.Name, "shown": [s for s in log.split(chr(1)) if s], "error": error}
    finally:
        wb.Close(False)


def compare(case: dict, ex: dict, ca: dict) -> list[str]:
    if "open_error" in ex:
        return [ex["open_error"]]
    out = []
    checks = set(case.get("check", []))
    if bool(ex.get("error")) != bool(ca.get("error")):
        out.append(f"error: Excel {ex.get('error')} RaabeCalc {ca.get('error')}")
    if [s["name"] for s in ex["sheets"]] != [s["name"] for s in ca["sheets"]]:
        out.append(f"sheets: Excel {[s['name'] for s in ex['sheets']]} RaabeCalc {[s['name'] for s in ca['sheets']]}")
    if ex["active"] != ca["active"]:
        out.append(f"active sheet: Excel {ex['active']} RaabeCalc {ca['active']}")
    if ex["shown"] != ca["shown"]:
        out.append(f"messages: Excel {ex['shown']} RaabeCalc {ca['shown']}")
    for es, cs in zip(ex["sheets"], ca["sheets"]):
        n = es["name"]
        for k in ("visible", "protected"):
            if es[k] != cs[k]:
                out.append(f"{n} {k}: Excel {es[k]} RaabeCalc {cs[k]}")
        if "hidden" in checks or es["hiddenRows"] or es["hiddenCols"] or cs["hiddenRows"] or cs["hiddenCols"]:
            for k in ("hiddenRows", "hiddenCols"):
                if es[k] != cs[k]:
                    out.append(f"{n} {k}: Excel {es[k]} RaabeCalc {cs[k]}")
        addrs = sorted(set(es["cells"]) | set(cs["cells"]), key=lambda a: (re.sub(r"\d", "", a).rjust(3), int(re.sub(r"\D", "", a))))
        blank = {"formula": "", "text": "", "fill": None, "fontColor": None, "bold": False, "italic": False}
        for a in addrs:
            e, c = es["cells"].get(a, blank), cs["cells"].get(a, blank)
            fields = ["formula", "text"] + (["fill"] if "fill" in checks else []) + (["fontColor", "bold", "italic"] if "font" in checks else [])
            for f in fields:
                # Automatic and black font colours look the same.
                if f == "fontColor" and (e.get(f) or "#000000") == (c.get(f) or "#000000"):
                    continue
                if (e.get(f) or None) != (c.get(f) or None):
                    out.append(f"{n}!{a} {f}: Excel {e.get(f)!r} RaabeCalc {c.get(f)!r}")
    return out


def main() -> None:
    want = set(sys.argv[1:])
    cases = [c for c in CASES if not want or c["name"] in want]
    results = []
    c = CalcSide()
    try:
        ours = {}
        for case in cases:
            xlsm = TMP / f"{case['name']}.xlsm"
            try:
                ours[case["name"]] = calc_side(c, case, xlsm)
            except Exception as e:  # noqa: BLE001
                ours[case["name"]] = {"failed": str(e)[:300]}
    finally:
        c.close()
    x = ExcelTarget(visible=False)
    try:
        for case in cases:
            # A dialog in the hidden Excel (a compile error, an unhandled runtime error) would wait forever: after a
            # minute only our own Excel process is ended and a new one started for the next case.
            hung = []
            timer = threading.Timer(60, lambda pid=x.pid: (hung.append(1), subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True)))
            ca = ours[case["name"]]
            xlsm = TMP / f"{case['name']}.xlsm"
            if "failed" in ca:
                results.append({"case": case["name"], "ok": False, "diff": [f"RaabeCalc: {ca['failed']}"]})
                continue
            timer.start()
            try:
                ex = excel_side(x, case, xlsm)
            except Exception as e:  # noqa: BLE001
                results.append({"case": case["name"], "ok": False, "diff": ["Excel waited on a dialog (ended)" if hung else f"Excel: {str(e)[:300]}"]})
                if hung:
                    x = ExcelTarget(visible=False)
                continue
            finally:
                timer.cancel()
            d = compare(case, ex, ca)
            results.append({"case": case["name"], "ok": not d, "diff": d[:12]})
    finally:
        x.close()
    ok = sum(r["ok"] for r in results)
    for r in results:
        print(("OK  " if r["ok"] else "DIFF") + " " + r["case"])
        for line in r["diff"]:
            print("      " + line)
    print(f"{ok}/{len(results)} macros leave the same workbook")
    if not want:
        path = HERE / "results" / f"{datetime.date.today().isoformat()}-calc-macros.json"
        path.write_text(json.dumps({"cases": len(results), "same": ok, "results": results}, ensure_ascii=False, indent=1), encoding="utf-8", newline=NL)


if __name__ == "__main__":
    main()
