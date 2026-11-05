"""Datentabellen (Was-wäre-wenn-Analyse) against Excel (D-066), both ways, without real input:
1. RaabeCalc (off screen) makes a one-variable table down a column, one across a row and a
   two-variable table, and saves; our own invisible Excel opens the file, recalculates, and must see
   each as a data table (=TABLE(...)) with the same results.
2. Excel makes the same tables (Range.Table) and saves; RaabeCalc opens the file and must show the
   same results and the formula {=MEHRFACHOPERATION(...)}.
Result: results/<date>-datatable.json.
"""
from __future__ import annotations

import datetime
import json
import tempfile
from pathlib import Path

from feellab.calc_catalogue import Case, CalcSide
from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
TMP = Path(tempfile.gettempdir())
# Zins in A1, Betrag in A2; B4 = A2 * (1 + A1); tables below.
SETUP = {"A1": "0,05", "A2": "100",
         # column table: values D2:D5, formula in E1 (=A2*(1+A1)), results E2:E5
         "D2": "0,01", "D3": "0,02", "D4": "0,03", "D5": "0,04", "E1": "=A2*(1+A1)",
         # row table: values H2:K2... use row 8: values B8:E8, formula in A9, results B9:E9
         "B8": "0,01", "A9": "=A2*(1+A1)", "C8": "0,02", "D8": "0,03", "E8": "0,04",
         # two-variable: corner G12 formula, rates H12:J12 (row), amounts G13:G15 (column)
         "G12": "=A2*(1+A1)", "H12": "0,01", "I12": "0,02", "J12": "0,03", "G13": "100", "G14": "200", "G15": "300"}
TABLES = [
    {"name": "column", "sel": "D1:E5", "row": None, "col": "A1", "ref": "E2:E5"},
    {"name": "row", "sel": "A8:E9", "row": "A1", "col": None, "ref": "B9:E9"},
    {"name": "two", "sel": "G12:J15", "row": "A1", "col": "A2", "ref": "H13:J15"},
]


def norm(v):
    return round(v, 9) if isinstance(v, float) else v


def main() -> None:
    raabe_file, excel_file = TMP / "raabe-datatable.xlsx", TMP / "excel-datatable.xlsx"
    report: dict = {"raabe->excel": [], "excel->raabe": []}
    c = CalcSide()
    ours: dict = {}
    try:
        c.reset(Case("dt", SETUP, [], "A1", sheets=1))
        # The row table needs its formula in the first column: A9.
        c.js("(() => { const app = __calc.app; })()")
        for t in TABLES:
            js = f"""(() => {{ const app = __calc.app; const p = (a) => {{ if (!a) return undefined; const r = __calc.parse(a); return {{ row: r.r0, col: r.c0 }}; }};
              app.whatIf.create(__calc.parse({json.dumps(t['sel'])}), p({json.dumps(t['row'])}), p({json.dumps(t['col'])})); }})()"""
            c.js(js)
        for t in TABLES:
            ours[t["name"]] = c.js(f"""(() => {{ const s = __calc.app.sheet; const r = __calc.parse({json.dumps(t['ref'])}); const out = [];
              for (let i = r.r0; i <= r.r1; i++) {{ const row = []; for (let k = r.c0; k <= r.c1; k++) row.push(s.value(i, k)); out.push(row); }} return out; }})()""")
        c.js(f"__calc.save({json.dumps(str(raabe_file))})")
    finally:
        c.close()
    x = ExcelTarget(visible=False)
    try:
        wb = x.app.Workbooks.Open(str(raabe_file))
        ws = wb.Worksheets(1)
        x.app.Calculate()
        for t in TABLES:
            rng = ws.Range(t["ref"])
            vals = rng.Value
            got = [[norm(v) for v in row] for row in vals]
            want = [[norm(v) for v in row] for row in ours[t["name"]]]
            f = ws.Range(t["ref"].split(":")[0]).Formula
            ok = got == want and ("TABLE(" in f.upper() or "MEHRFACHOPERATION(" in f.upper())
            report["raabe->excel"].append({"table": t["name"], "ok": ok, **({} if ok else {"excel": got, "raabe": want, "formula": f})})
            print("R->E", t["name"], ok, f, flush=True)
        wb.Close(False)
        wb = x.app.Workbooks.Add()
        ws = wb.Worksheets(1)
        for a, v in SETUP.items():
            ws.Range(a).FormulaLocal = v
        excel_vals = {}
        for t in TABLES:
            sel = ws.Range(t["sel"])
            row = ws.Range(t["row"]) if t["row"] else None
            col = ws.Range(t["col"]) if t["col"] else None
            if row is not None and col is not None:
                sel.Table(row, col)
            elif row is not None:
                sel.Table(row)
            else:
                sel.Table(None, col)
            excel_vals[t["name"]] = [[norm(v) for v in r] for r in ws.Range(t["ref"]).Value]
        wb.SaveAs(str(excel_file), 51)
        wb.Close(False)
    finally:
        x.close()
    c = CalcSide()
    try:
        c.js(f"__calc.open({json.dumps(str(excel_file))})")
        for t in TABLES:
            got = c.js(f"""(() => {{ const s = __calc.app.sheet; const r = __calc.parse({json.dumps(t['ref'])}); const out = [];
              for (let i = r.r0; i <= r.r1; i++) {{ const row = []; for (let k = r.c0; k <= r.c1; k++) row.push(s.value(i, k)); out.push(row); }} return out; }})()""")
            dt = c.js(f"__calc.app.sheet.get(__calc.parse({json.dumps(t['ref'])}).r0, __calc.parse({json.dumps(t['ref'])}).c0)?.dt ?? null")
            got = [[norm(v) for v in row] for row in got]
            ok = got == excel_vals[t["name"]] and bool(dt)
            report["excel->raabe"].append({"table": t["name"], "ok": ok, **({} if ok else {"raabe": got, "excel": excel_vals[t["name"]], "dt": dt})})
            print("E->R", t["name"], ok, dt, flush=True)
    finally:
        c.close()
    (HERE / "results" / f"{datetime.date.today().isoformat()}-datatable.json").write_text(json.dumps(report, ensure_ascii=False, indent=1, default=str), encoding="utf-8", newline="\n")
    for k, v in report.items():
        print(k, sum(1 for r in v if r["ok"]), "/", len(v))


if __name__ == "__main__":
    main()
