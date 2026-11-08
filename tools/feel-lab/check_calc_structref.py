"""Structured references, berechnete Spalten and growing tables against Excel (D-080), both ways,
without real input:
1. RaabeCalc (off screen) makes a table, types a header next to it (the table grows by a column), a
   formula with [@Note] into the new column (a berechnete Spalte), a name below the table (it grows by
   a row) and formulas outside the table; saves. Our own invisible Excel opens the file, recalculates
   and must show the same table, the same formulas (Formula2Local) and the same values.
2. RaabeCalc opens the workbook Excel saved in probe_excel_structref.py (TEMP/raabe-structref.xlsx)
   and must show Excel's formulas and values (fixture excel-structref.json).
Result: results/<date>-calc-structref.json.
"""
from __future__ import annotations

import datetime
import json
import tempfile
from pathlib import Path

from feellab.calc_catalogue import Case, CalcSide
from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
ROOT = HERE.parents[1]
TMP = Path(tempfile.gettempdir())
DATA = {"A1": "Name", "B1": "Fach", "C1": "Note", "A2": "Anna", "B2": "Mathe", "C2": "2", "A3": "Ben", "B3": "Deutsch", "C3": "3",
        "A4": "Cem", "B4": "Mathe", "C4": "1", "A5": "Dora", "B5": "Englisch", "C5": "2"}
# Typed one after another into the table's surroundings (German, as a user types).
TYPED = [("D1", "Doppelt"), ("D2", "=[@Note]*2"), ("A6", "Emil"), ("C6", "4"),
         ("G1", "=SUMME(Noten[Doppelt])"), ("G2", "=ZEILEN(Noten)"), ("G3", "=SVERWEIS(\"Cem\";Noten;4;FALSCH)"),
         ("G4", "=ANZAHL2(Noten[[#Alle];[Name]])"), ("G5", "=Noten[@Note]")]
ERRORS = {-2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?", -2146826288: "#NULL!",
          -2146826252: "#NUM!", -2146826265: "#REF!", -2146826273: "#VALUE!"}


def norm(v):
    if isinstance(v, int) and v in ERRORS:
        return ERRORS[v]
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def main() -> None:
    raabe_file = TMP / "raabe-structref-ours.xlsx"
    report: dict = {"raabe->excel": [], "excel->raabe": []}
    c = CalcSide()
    try:
        c.reset(Case("structref", DATA, [], "A1", sheets=1))
        c.js("""(() => { const app = __calc.app; const t = app.tables.create(__calc.parse("A1:C5"), true); t.name = "Noten"; })()""")
        for a, text in TYPED:
            c.js(f"""(() => {{ const app = __calc.app; const r = __calc.parse({json.dumps(a)}); app.grid.selectCell(r.r0, r.c0); app.gridHost().commit(r, {json.dumps(text)}, {{}}); }})()""")
        ours = c.js("""(() => { const app = __calc.app; const sh = app.sheet; const t = sh.tables[0];
          const cells = {}; for (const a of ["D2","D3","D4","D5","D6","G1","G2","G3","G4","G5"]) { const r = __calc.parse(a); const cell = sh.get(r.r0, r.c0);
            const v = cell?.v; cells[a] = { f: cell?.f ?? null, v: v && typeof v === "object" ? String(v) : v ?? null }; }
          return { ref: t.ref, columns: t.columns.map((x) => [x.name, x.formula ?? null]), cells }; })()""")
        c.js(f"__calc.save({json.dumps(str(raabe_file))})")
    finally:
        c.close()
    report["ours"] = ours
    x = ExcelTarget(visible=False)
    try:
        wb = x.app.Workbooks.Open(str(raabe_file))
        ws = wb.Worksheets(1)
        x.app.CalculateFull()
        lo = ws.ListObjects(1) if ws.ListObjects.Count else None
        excel = {"name": lo.Name if lo else None, "range": lo.Range.Address.replace("$", "") if lo else None,
                 "columns": [lo.ListColumns(i).Name for i in range(1, lo.ListColumns.Count + 1)] if lo else [],
                 "cells": {a: {"local": ws.Range(a).Formula2Local, "value": norm(ws.Range(a).Value2)} for a in ours["cells"]}}
        wb.Close(False)
    finally:
        x.close()
    report["excel"] = excel
    checks = [("table name", excel["name"] == "Noten"), ("range grew to A1:D6", excel["range"] == "A1:D6"),
              ("columns", excel["columns"] == ["Name", "Fach", "Note", "Doppelt"])]
    shown = {"D2": "=[@Note]*2", "D3": "=[@Note]*2", "D4": "=[@Note]*2", "D5": "=[@Note]*2", "D6": "=[@Note]*2",
             "G1": "=SUMME(Noten[Doppelt])", "G2": "=ZEILEN(Noten)", "G3": "=SVERWEIS(\"Cem\";Noten;4;FALSCH)",
             "G4": "=ANZAHL2(Noten[[#Alle];[Name]])", "G5": "=Noten[@Note]"}
    for a, want in shown.items():
        e = excel["cells"][a]
        o = ours["cells"][a]
        checks.append((f"{a} formula as Excel shows it", e["local"] == want))
        checks.append((f"{a} value the same", norm(o["v"]) == e["value"]))
    for name, ok in checks:
        report["raabe->excel"].append({"check": name, "ok": ok})
        print("R->E", name, ok, flush=True)

    # 2. Excel's own file.
    fixture = json.loads((ROOT / "packages" / "sheet-engine" / "test" / "fixtures" / "excel-structref.json").read_text(encoding="utf-8"))
    excel_file = TMP / "raabe-structref.xlsx"
    if excel_file.exists():
        c = CalcSide()
        try:
            c.js(f"__calc.open({json.dumps(str(excel_file))})")
            texts = c.js("""(() => { const app = __calc.app; const sh = app.sheet; const out = {};
              for (const a of %s) { const r = __calc.parse(a); app.grid.selectCell(r.r0, r.c0); out[a] = { bar: document.querySelector('.fx-input')?.value ?? null, v: (() => { const v = sh.get(r.r0, r.c0)?.v; return v && typeof v === "object" ? String(v) : v ?? null; })() }; }
              return out; })()""" % json.dumps(["D2", "D3", "D4", "D5", "C6"] + [o["cell"].replace("G", "H") if int(o["cell"][1:]) <= 6 else o["cell"] for o in fixture["outside"] if not o.get("rejected")]))
        finally:
            c.close()
        want = {d["cell"]: d for d in fixture["inside"]}
        want["C6"] = {"local": fixture["totals_formula"]["local"], "value": fixture["totals_formula"]["value"]}
        for o in fixture["outside"]:
            if o.get("rejected"):
                continue
            cell = o["cell"].replace("G", "H") if int(o["cell"][1:]) <= 6 else o["cell"]
            want[cell] = {**o, "value": o.get("value_after", o["value"])}
        for a, w in want.items():
            t = texts.get(a, {})
            wv = w["value"]["error"] if isinstance(w["value"], dict) else norm(w["value"])
            ok_f = t.get("bar") == w["local"]
            ok_v = norm(t.get("v")) == wv
            report["excel->raabe"].append({"cell": a, "ok": ok_f and ok_v, **({"bar": t.get("bar"), "want": w["local"], "v": t.get("v"), "want_v": wv} if not (ok_f and ok_v) else {})})
            print("E->R", a, ok_f, ok_v, t.get("bar"), w["local"], flush=True)
    (HERE / "results" / f"{datetime.date.today().isoformat()}-calc-structref.json").write_text(json.dumps(report, ensure_ascii=False, indent=1, default=str), encoding="utf-8", newline="\n")
    for key in ("raabe->excel", "excel->raabe"):
        v = report[key]
        print(key, sum(1 for r in v if r["ok"]), "/", len(v))


if __name__ == "__main__":
    main()
