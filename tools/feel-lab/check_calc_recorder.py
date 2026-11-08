"""Makro aufzeichnen in RaabeCalc against Excel (D-119), without real input. For every case of recorder_cases.py:
1. "By hand" in the real exe (off screen, trusted DevTools input): Alt, F, M, R opens Makro aufzeichnen (Relative Verweise
   verwenden first by Alt, F, M, V when the case asks), the name, shortcut and description are typed, OK clicked; then the
   case's keys and clicks; the status bar's button ends the recording. Compared with Excel's recording by hand
   (results/2026-11-09-excel-recorder.json, probe_excel_recorder.py): the macro's text line by line, and the workbook
   each left (cells, formulas, shown text, formats, sheets, selection).
2. Excel's recorded macro run in RaabeCalc's interpreter on the case's start: the same workbook as Excel by hand?
3. RaabeCalc's recorded macro run in our own invisible Excel (RaabeCalc saves it as .xlsm with the start cells; Excel
   opens it and runs it with Application.Run): the same workbook as Excel by hand?
4. The macro's shortcut (Strg+letter) runs it in RaabeCalc.

    .venv/Scripts/python.exe check_calc_recorder.py [case ...]

Result: results/2026-11-09-calc-recorder.json.
"""
from __future__ import annotations

import json
import os
import re
import sys
import tempfile
import time
from pathlib import Path

from feellab import calc_catalogue, winapi
from feellab.calc_catalogue import CalcSide
from feellab.cdpinput import Cdp
from feellab.excel import ExcelTarget
from recorder_cases import CASES

HERE = Path(__file__).parent
EXCEL = json.loads((HERE / "results" / "2026-11-09-excel-recorder.json").read_text(encoding="utf-8"))
RESULT = HERE / "results" / "2026-11-09-calc-recorder.json"
TMP = Path(tempfile.gettempdir()) / "raabe-recorder"
TMP.mkdir(exist_ok=True)
NL = chr(10)

os.environ["RAABE_FEEL_BACKGROUND"] = "1"
calc_catalogue.winapi.place = lambda hwnd, x, y, w, h: winapi.place_quietly(hwnd, -4000, 0, w, h)

EXTRA_JS = """(() => {
  const app = __calc.app, sheet = app.sheet, out = {};
  for (const a of Object.keys(__calc.state([]).cells)) {
    const r = __calc.parse(a); const c = sheet.get(r.r0, r.c0); const st = app.wb.styles.get(c?.s);
    const hex = st.fill ? parseInt(st.fill.slice(1), 16) : null;
    out[a] = { fill: hex === null ? null : ((hex >> 16) & 255) | (hex & 0xff00) | ((hex & 255) << 16), size: st.font.size,
      h: ({ left: -4131, center: -4108, right: -4152 })[st.h] ?? 1 };
  }
  const hidden = []; for (let r = 0; r < 11; r++) if (sheet.rows.get(r)?.hidden) hidden.push(r + 1);
  return { extra: out, rowHidden: hidden };
})()"""


def norm_code(code: str) -> list[str]:
    return [l.rstrip() for l in code.replace(chr(13), "").split(NL) if l.strip() != ""]


def calc_state(c: CalcSide) -> dict:
    st = c.js("__calc.state([])")
    st.update(c.js(EXTRA_JS))
    return st


def compare_state(ex: dict, ca: dict) -> list[str]:
    out = []
    for k in ("sheets", "sheet", "selection", "active", "rowHidden"):
        if k in ex and ex.get(k) != ca.get(k):
            out.append(f"{k}: Excel {ex.get(k)} RaabeCalc {ca.get(k)}")
    addrs = sorted(set(ex.get("cells", {})) | set(ca.get("cells", {})))
    for a in addrs:
        e, k = ex["cells"].get(a, {}), ca.get("cells", {}).get(a, {})
        for f in ("formula", "text", "bold", "italic"):
            if (e.get(f) or None) != (k.get(f) or None):
                out.append(f"{a} {f}: Excel {e.get(f)!r} RaabeCalc {k.get(f)!r}")
        ee, kk = ex.get("extra", {}).get(a, {}), ca.get("extra", {}).get(a, {})
        for f in ("fill", "size", "h"):
            if f in ee and ee.get(f) != kk.get(f):
                out.append(f"{a} {f}: Excel {ee.get(f)!r} RaabeCalc {kk.get(f)!r}")
    return out


def record_by_hand(c: CalcSide, inp: Cdp, case: dict) -> dict:
    c.js(f"__calc.reset({json.dumps(case['setup'])}, {json.dumps(case['start'])}, 1)")
    time.sleep(0.2)
    c.js("__calc.grid.focus()")
    if bool(case.get("relative")) != c.js("__calc.app.recorder.relative"):
        inp.keytips("F M V")
    rel = c.js("__calc.app.recorder.relative")
    inp.keytips("F M R")
    time.sleep(0.4)
    if not c.js("[...document.querySelectorAll('.dialog')].some(d => d.textContent.includes('Makroname'))"):
        return {"failed": "Alt, F, M, R did not open Makro aufzeichnen"}
    default_name = c.js("document.querySelector('.dialog input[aria-label=Makroname]').value")
    inp.key("ctrl+a")
    inp.type(f"Makro_{case['name']}")
    inp.key("tab")
    if case.get("shortcut"):
        inp.type(case["shortcut"])
    inp.key("tab")
    inp.key("tab")
    if case.get("description"):
        inp.type(case["description"])
    inp.click_el("[...document.querySelectorAll('.dialog button')].find(b => b.textContent === 'OK')")
    time.sleep(0.3)
    if not c.js("__calc.app.recorder.recording"):
        return {"failed": "not recording after OK"}
    c.js("__calc.grid.focus()")
    for st in case["steps"]:
        inp.step(st)
    time.sleep(0.3)
    inp.click_el("document.querySelector('.statusbar .recbtn')")
    time.sleep(0.3)
    mods = c.js("__calc.macroModules()")
    code = NL.join(m["code"].replace(chr(13), "") for m in mods)
    m = re.search(r"^Sub Makro_" + case["name"] + r"\(\)$.*?^End Sub$", code, re.S | re.M)
    state = calc_state(c)
    return {"code": m.group(0) if m else None, "state": state, "relative": rel, "defaultName": default_name, "recording": c.js("__calc.app.recorder.recording")}


def run_in_calc(c: CalcSide, case: dict, code: str) -> dict:
    c.js(f"__calc.reset({json.dumps(case['setup'])}, {json.dumps(case['start'])}, 1)")
    c.js(f"__calc.setMacros({json.dumps(code)})")
    run = c.js(f"__calc.runMacro('Makro_{case['name']}', [])")
    st = calc_state(c)
    st["error"] = run.get("error")
    return st


def shortcut_runs(c: CalcSide, inp: Cdp, case: dict, code: str) -> bool | None:
    """Strg+letter runs the macro: the workbook afterwards is the one the macro gives."""
    if not case.get("shortcut"):
        return None
    want = run_in_calc(c, case, code)
    c.js(f"__calc.reset({json.dumps(case['setup'])}, {json.dumps(case['start'])}, 1)")
    c.js(f"__calc.setMacros({json.dumps(code)})")
    c.js("__calc.grid.focus()")
    inp.key(f"ctrl+{case['shortcut']}")
    time.sleep(0.8)
    got = calc_state(c)
    return not compare_state(want, got)


def save_for_excel(c: CalcSide, case: dict, code: str, path: Path) -> None:
    c.js(f"__calc.reset({json.dumps(case['setup'])}, {json.dumps(case['start'])}, 1)")
    c.js(f"__calc.setMacros({json.dumps(code)})")
    if path.exists():
        path.unlink()
    c.js(f"__calc.save({json.dumps(str(path))})")


def run_in_excel(x: ExcelTarget, case: dict, path: Path) -> dict:
    from check_calc_macros import run_us
    app = x.app
    app.DisplayAlerts = False
    wb = app.Workbooks.Open(str(path))
    try:
        ws = wb.Worksheets(1)
        ws.Activate()
        ws.Range(case["start"]).Select()
        err = None
        try:
            run_us(app, f"'{wb.Name}'!Makro_{case['name']}")
        except Exception as e:  # noqa: BLE001
            err = str(e)[:200]
        ws = app.ActiveSheet
        cells, extra = {}, {}
        for cell in ws.UsedRange.Cells:
            if cell.Formula == "":
                continue
            a = cell.GetAddress(False, False)
            cells[a] = {"formula": cell.FormulaLocal, "text": cell.Text, "bold": bool(cell.Font.Bold), "italic": bool(cell.Font.Italic)}
            extra[a] = {"fill": None if cell.Interior.ColorIndex == -4142 else int(cell.Interior.Color), "size": cell.Font.Size, "h": cell.HorizontalAlignment}
        return {"cells": cells, "extra": extra, "sheets": [s.Name for s in wb.Worksheets], "sheet": ws.Name,
                "selection": app.Selection.GetAddress(False, False), "active": app.ActiveCell.GetAddress(False, False),
                "rowHidden": [k for k in range(1, 12) if ws.Rows(k).Hidden], "error": err}
    finally:
        wb.Close(False)


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    want = set(sys.argv[1:])
    cases = [k for k in CASES if (not want or k["name"] in want) and "code" in EXCEL["cases"].get(k["name"], {})]
    results: dict = {}
    c = CalcSide()
    inp = Cdp(c)
    try:
        for case in cases:
            ex = EXCEL["cases"][case["name"]]
            r: dict = {}
            try:
                hand = record_by_hand(c, inp, case)
            except Exception as e:  # noqa: BLE001
                hand = {"failed": str(e)[:300]}
            r["byHand"] = hand
            if "failed" not in hand and hand.get("code"):
                ours, theirs = norm_code(hand["code"]), norm_code(ex["code"])
                r["codeSame"] = ours == theirs
                if ours != theirs:
                    import difflib
                    r["codeDiff"] = [l for l in difflib.unified_diff(theirs, ours, "Excel", "RaabeCalc", lineterm="", n=0)][2:40]
                r["stateDiff"] = compare_state(ex["state"], hand["state"])
            # Excel's own macro in RaabeCalc.
            try:
                st = run_in_calc(c, case, ex["code"])
                r["excelMacroInCalc"] = compare_state(ex["state"], st) + ([f"error {st['error']}"] if st.get("error") else [])
            except Exception as e:  # noqa: BLE001
                r["excelMacroInCalc"] = [f"failed: {str(e)[:200]}"]
            if hand.get("code"):
                try:
                    r["shortcut"] = shortcut_runs(c, inp, case, hand["code"])
                except Exception as e:  # noqa: BLE001
                    r["shortcut"] = f"failed: {str(e)[:200]}"
                save_for_excel(c, case, hand["code"], TMP / f"ours-{case['name']}.xlsm")
            results[case["name"]] = r
            print(case["name"], "code", r.get("codeSame"), "state", len(r.get("stateDiff", [])) if "stateDiff" in r else "-",
                  "excel->calc", len(r["excelMacroInCalc"]), "shortcut", r.get("shortcut"), flush=True)
    finally:
        c.close()
    x = ExcelTarget(visible=False)
    try:
        for case in cases:
            r = results[case["name"]]
            path = TMP / f"ours-{case['name']}.xlsm"
            if not path.exists():
                continue
            try:
                st = run_in_excel(x, case, path)
                r["oursInExcel"] = compare_state(EXCEL["cases"][case["name"]]["state"], st) + ([f"error {st['error']}"] if st.get("error") else [])
            except Exception as e:  # noqa: BLE001
                r["oursInExcel"] = [f"failed: {str(e)[:200]}"]
    finally:
        x.close()
    counts = {
        "cases": len(cases),
        "codeSame": sum(bool(r.get("codeSame")) for r in results.values()),
        "byHandSameWorkbook": sum(r.get("stateDiff") == [] for r in results.values()),
        "excelMacroInCalc": sum(r.get("excelMacroInCalc") == [] for r in results.values()),
        "oursInExcel": sum(r.get("oursInExcel") == [] for r in results.values()),
        "shortcut": f"{sum(r.get('shortcut') is True for r in results.values())}/{sum(r.get('shortcut') is not None for r in results.values())}",
    }
    for name, r in results.items():
        print("=====", name)
        for k in ("codeDiff", "stateDiff", "excelMacroInCalc", "oursInExcel"):
            for line in r.get(k) or []:
                print(f"  {k}: {line}")
        if "failed" in r.get("byHand", {}):
            print("  byHand failed:", r["byHand"]["failed"])
    print(counts)
    if not want:
        RESULT.write_text(json.dumps({"counts": counts, "results": results}, ensure_ascii=False, indent=1), encoding="utf-8", newline=NL)


if __name__ == "__main__":
    main()
