"""Szenario-Manager and Konsolidieren against Excel (D-084), both ways, without real input:
1. RaabeCalc (off screen, never focused) adds five scenarios (contiguous and separate cells, a text
   value, a German decimal, Schutz flags), shows one, makes the Szenariobericht, consolidates two sheets by
   labels and with Verknüpfungen mit Quelldaten, and saves. Our own invisible Excel opens the file: the
   scenarios must be ours (names, cells, values, comments, flags), Excel must show them with its own
   calculation, the report's cells and outline must be there, Excel's own Konsolidieren of the same sources
   must give our cells, the link rows must be Excel's formulas.
2. Excel adds a scenario of its own to that workbook, shows another and saves; RaabeCalc opens Excel's file
   and must hold every scenario as Excel reports it, show Excel's scenario with the same values, and keep
   Konsolidieren's settings. RaabeCalc also opens Excel's fixture (xlsx-scenarios.xlsx) and must hold its
   scenarios as COM read them (excel-scenarios.json).
Result: results/<date>-calc-scenarios.json (checks only, no paths or names).

    .venv/Scripts/python.exe check_calc_scenarios.py
"""
from __future__ import annotations

import datetime
import json
import os
import tempfile
from pathlib import Path

from feellab import calc_catalogue, winapi
from feellab.calc_catalogue import Case, CalcSide
from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
ROOT = HERE.parents[1]
TMP = Path(tempfile.gettempdir())

# RaabeCalc runs off screen and never takes the focus (someone may be working at the PC).
os.environ["RAABE_FEEL_BACKGROUND"] = "1"
calc_catalogue.winapi.place = lambda hwnd, x, y, w, h: winapi.place_quietly(hwnd, -4000, 0, w, h)

SETUP = {"A1": "Preis", "B1": "10", "A2": "Menge", "B2": "100", "A3": "Kosten", "B3": "500", "A4": "Gewinn", "B4": "=B1*B2-B3",
         "A5": "Steuer", "B5": "=B4*0,3", "D1": "Rabatt", "E1": "0,05", "E2": "=E1*2"}
NORD = {"B1": "Jan", "C1": "Feb", "A2": "Äpfel", "B2": "10", "C2": "20", "A3": "Birnen", "B3": "5", "C3": "7", "A4": "Kirschen", "B4": "3", "C4": "x"}
SUED = {"B1": "Feb", "C1": "Mär", "A2": "birnen", "B2": "1", "C2": "2", "A3": "Äpfel", "B3": "4", "A4": "Pflaumen", "B4": "6", "C4": "8", "A5": "Äpfel", "B5": "100", "C5": "200"}
USER = "Raabe Office Test"
SCEN = [  # name, cells, values, comment, locked, hidden
    ("Gut", "B1:B2", ["12", "120"], "Hohe Preise", True, False),
    ("Schlecht", "B1:B2", ["8", "80"], "", False, True),
    ("Mittel", "B1;B3", ["10", "400"], "Mittelweg", True, False),
    ("Text", "E1", ["abc"], "Text als Wert", True, False),
    ("Dezimal", "B2", ["0,25"], "Komma", True, False),
]
ERRORS = {-2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?", -2146826288: "#NULL!",
          -2146826252: "#NUM!", -2146826265: "#REF!", -2146826273: "#VALUE!"}

BUILD = r"""(async () => {
  const app = __calc.app, P = __calc.parse;
  const type = (si, cells) => { app.activateSheet(si); for (const [a, t] of Object.entries(cells)) { const r = P(a); app.grid.selectCell(r.r0, r.c0); app.gridHost().commit(r, t, {}); } };
  ["Tabelle1", "Nord", "Süd", "Gesamt", "Links"].forEach((n, i) => (app.wb.sheets[i].name = n));
  type(1, %(nord)s); type(2, %(sued)s); app.activateSheet(0);
  const sheet = app.sheet;
  const cells = (t) => t.split(";").flatMap((x) => { const r = P(x); const out = []; for (let row = r.r0; row <= r.r1; row++) for (let col = r.c0; col <= r.c1; col++) out.push({ row, col }); return out; });
  for (const [name, c, vals, comment, locked, hidden] of %(scen)s) {
    app.scenarios.save("Szenario hinzufügen", sheet, (s) => s.list.push({ name, cells: cells(c).map((p, i) => ({ ...p, val: vals[i] })), user: %(user)s,
      ...(comment ? { comment } : {}), ...(locked ? { locked: true } : {}), ...(hidden ? { hidden: true } : {}) }));
  }
  app.scenarios.show(sheet, 2);
  app.scenarios.makeReport(sheet, [P("B4:B5")]);
  // Konsolidieren on Gesamt (labels) and Links (Verknüpfungen mit Quelldaten).
  app.activateSheet(app.wb.sheets.findIndex((s) => s.name === "Gesamt"));
  await app.consolidation.run(app.sheet, P("A1"), ["Nord!$A$1:$C$4", "Süd!$A$1:$C$5"], { fn: "sum", top: true, left: true, link: false });
  app.activateSheet(app.wb.sheets.findIndex((s) => s.name === "Links"));
  await app.consolidation.run(app.sheet, P("A1"), ["Nord!$A$1:$C$4", "Süd!$A$1:$C$5"], { fn: "average", top: true, left: true, link: true });
  const dump = (s) => { const out = {}; for (const [k, c] of s.cells) { const row = Math.floor(k / 16384), col = k %% 16384; const v = c.v; out[__calc.cell(row, col)] = { v: v && typeof v === "object" ? String(v) : v, f: c.f ?? null }; } return out; };
  const rows = (s) => { const out = {}; for (const [r, i] of s.rows) if (i.outline || i.hidden) out[r + 1] = [i.outline ?? 0, !!i.hidden]; return out; };
  const sh = (n) => app.wb.sheetByName(n);
  return { sheets: app.wb.sheets.map((s) => s.name), scenarios: sh("Tabelle1").scenarios, t1: dump(sh("Tabelle1")),
    report: dump(sh("Szenariobericht")), reportRows: rows(sh("Szenariobericht")), cons: dump(sh("Gesamt")), links: dump(sh("Links")), linkRows: rows(sh("Links")),
    consSettings: sh("Gesamt").consolidate };
})()"""


def norm(v):
    if isinstance(v, int) and v in ERRORS:
        return ERRORS[v]
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def same(a, b) -> bool:
    a, b = norm(a), norm(b)
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) <= 1e-9 * max(1, abs(a), abs(b))
    return a == b


def comment_of(s) -> str:
    """A scenario without a comment: COM refuses to read the property (Excel's own files too)."""
    try:
        return s.Comment
    except Exception:  # noqa: BLE001
        return ""


def excel_scenarios(ws) -> list[dict]:
    out = []
    for i in range(1, ws.Scenarios().Count + 1):
        s = ws.Scenarios(i)
        vals = s.Values
        vals = [v[0] if isinstance(v, tuple) else v for v in (vals if isinstance(vals, tuple) else [vals])]
        out.append({"name": s.Name, "cells": s.ChangingCells.Address.replace(",", ";"), "values": [str(v) for v in vals], "comment": comment_of(s),
                    "locked": bool(s.Locked), "hidden": bool(s.Hidden)})
    return out


def addr(a: str) -> str:
    """Our "B1:B2" / "B1;B3" as Excel's absolute address."""
    import re
    return re.sub(r"([A-Z]+)(\d+)", r"$\1$\2", a)


def main() -> None:
    ours_file = TMP / "raabe-scen-ours.xlsx"
    excel_file = TMP / "raabe-scen-excel.xlsx"
    for f in (ours_file, excel_file):
        if f.exists():
            f.unlink()
    report: dict = {"raabe->excel": [], "excel->raabe": []}
    checks = report["raabe->excel"]

    def check(name: str, ok: bool, **info) -> None:
        checks.append({"check": name, "ok": bool(ok), **({} if ok else info)})
        print("R->E" if checks is report["raabe->excel"] else "E->R", name, ok, info if not ok else "", flush=True)

    c = CalcSide()
    try:
        c.reset(Case("scenarios", SETUP, [], "A1", sheets=5))
        c.js("window.__calc.cell = (r, c) => { let s = ''; for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s; return s + (r + 1); }")
        ours = c.js(BUILD % {"nord": json.dumps(NORD), "sued": json.dumps(SUED), "scen": json.dumps(SCEN), "user": json.dumps(USER)})
        c.js(f"__calc.save({json.dumps(str(ours_file))})")
    finally:
        c.close()

    x = ExcelTarget(visible=False)
    try:
        wb = x.app.Workbooks.Open(str(ours_file))
        x.app.CalculateFull()
        names = [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]
        check("sheets in our order (report before Tabelle1)", names == ours["sheets"], excel=names, ours=ours["sheets"])
        ws = wb.Worksheets("Tabelle1")
        got = excel_scenarios(ws)
        check("five scenarios", len(got) == len(SCEN), n=len(got))
        for (name, cells, vals, comment, locked, hidden), g in zip(SCEN, got):
            check(f"{name}: name", g["name"] == name, got=g["name"])
            check(f"{name}: cells", g["cells"] == addr(cells), got=g["cells"])
            # COM gives the values as the German text the file holds.
            check(f"{name}: values", g["values"] == vals, got=g["values"])
            check(f"{name}: comment", g["comment"] == comment, got=g["comment"])
            check(f"{name}: Schutz", (g["locked"], g["hidden"]) == (locked, hidden), got=[g["locked"], g["hidden"]])
        check("Mittel shown: B1, B3, B4 as Excel computes", [norm(ws.Range(a).Value2) for a in ("B1", "B3", "B4")] == [10, 400, 600],
              got=[norm(ws.Range(a).Value2) for a in ("B1", "B3", "B4")])
        ws.Scenarios("Gut").Show()
        # B3 is still Mittel's 400: 12*120-400.
        check("Excel shows our scenario Gut: B1, B2, B4", [norm(ws.Range(a).Value2) for a in ("B1", "B2", "B4")] == [12, 120, 1040],
              got=[norm(ws.Range(a).Value2) for a in ("B1", "B2", "B4")])
        ws.Scenarios("Dezimal").Show()
        check("Excel shows our German decimal: B2 = 0.25", same(ws.Range("B2").Value2, 0.25), got=ws.Range("B2").Value2)
        # The report sheet.
        rep = wb.Worksheets("Szenariobericht")
        bad = [a for a, v in ours["report"].items() if not same(rep.Range(a).Value2 if rep.Range(a).Value2 is not None else None, v["v"])]
        check(f"report: {len(ours['report'])} cells read back", not bad, cells=bad[:8])
        rows = {str(r): [int(rep.Rows(r).OutlineLevel) - 1, bool(rep.Rows(r).Hidden)] for r in range(1, 20) if rep.Rows(r).OutlineLevel > 1 or rep.Rows(r).Hidden}
        check("report: outline and the hidden comments row", rows == {str(k): v for k, v in ours["reportRows"].items()}, excel=rows, ours=ours["reportRows"])
        check("report: Excel's labels", [rep.Range(a).Value2 for a in ("B2", "D3", "B5", "B10")] == ["Szenariobericht", "Aktuelle Werte:", "Veränderbare Zellen:", "Ergebniszellen:"])
        # Konsolidieren: Excel's own result for the same sources on a sheet of its own.
        t4 = wb.Worksheets("Gesamt")
        bad = [a for a, v in ours["cons"].items() if not same(t4.Range(a).Value2, v["v"])]
        check(f"Konsolidieren: {len(ours['cons'])} cells read back", not bad, cells=bad[:8])
        fresh = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        fresh.Range("A1").Consolidate(["Nord!Z1S1:Z4S3", "Süd!Z1S1:Z5S3"], -4157, True, True, False)
        used = fresh.UsedRange
        theirs = {}
        for r in range(1, used.Row + used.Rows.Count):
            for col in range(1, used.Column + used.Columns.Count):
                v = fresh.Cells(r, col).Value2
                if v is not None:
                    theirs[fresh.Cells(r, col).Address.replace("$", "")] = norm(v)
        mine = {a: norm(v["v"]) for a, v in ours["cons"].items() if v["v"] is not None}
        check("Konsolidieren by labels: Excel's own result equals ours", theirs == mine, excel=theirs, ours=mine)
        fresh.Delete()
        t5 = wb.Worksheets("Links")
        bad = []
        for a, v in ours["links"].items():
            if v["f"] is not None:
                if not t5.Range(a).HasFormula or not same(t5.Range(a).Value2, v["v"]):
                    bad.append(a)
            elif not same(t5.Range(a).Value2, v["v"]):
                bad.append(a)
        check(f"Verknüpfungen: {len(ours['links'])} cells, formulas and values", not bad, cells=bad[:8])
        wrong = {a: t5.Range(a).Formula2Local for a, v in ours["links"].items() if v["f"] is not None and t5.Range(a).Formula2Local != "=" + v["f"].replace("AVERAGE(", "MITTELWERT(")}
        n = sum(1 for v in ours["links"].values() if v["f"] is not None)
        check(f"Verknüpfungen: {n} formulas as Excel shows them (=Süd!$B$4, =MITTELWERT(D2:D3))", not wrong, cells=wrong)
        lrows = {str(r): [int(t5.Rows(r).OutlineLevel) - 1, bool(t5.Rows(r).Hidden)] for r in range(1, 30) if t5.Rows(r).OutlineLevel > 1 or t5.Rows(r).Hidden}
        check("Verknüpfungen: outline", lrows == {str(k): v for k, v in ours["linkRows"].items()}, excel=lrows, ours=ours["linkRows"])
        # 2. Excel adds a scenario of its own and saves.
        ws.Activate()
        ws.Scenarios().Add("Excel", ws.Range("B3"), (250,), "Von Excel", False, False)
        ws.Scenarios("Text").Show()
        excel_list = excel_scenarios(ws)
        wb.SaveAs(str(excel_file), 51)
        wb.Close(False)
    finally:
        x.close()

    checks = report["excel->raabe"]
    c = CalcSide()
    try:
        c.js(f"__calc.open({json.dumps(str(excel_file))})")
        back = c.js("""(() => { const app = __calc.app; const s = app.wb.sheetByName("Tabelle1"); return { list: s.scenarios.list, current: s.scenarios.current, show: s.scenarios.show,
          cons: app.wb.sheetByName("Gesamt").consolidate, e1: s.get(0, 4)?.v ?? null }; })()""")
        check("six scenarios", len(back["list"]) == len(excel_list), n=len(back["list"]))
        for g, b in zip(excel_list, back["list"]):
            cells = ";".join(f"${chr(65 + c['col'])}${c['row'] + 1}" for c in b["cells"])
            ok_cells = cells == g["cells"] or g["cells"] == "$B$1:$B$2" and cells == "$B$1;$B$2"
            check(f"{g['name']}: as Excel reports it", b["name"] == g["name"] and ok_cells and [c["val"] for c in b["cells"]] == g["values"] and (b.get("comment") or "") == g["comment"]
                  and (bool(b.get("locked")), bool(b.get("hidden"))) == (g["locked"], g["hidden"]), ours=[b["name"], cells, [c["val"] for c in b["cells"]]], excel=[g["name"], g["cells"], g["values"]])
        check("Text shown last (E1 = abc)", back["e1"] == "abc", got=back["e1"])
        idx = next(i for i, s in enumerate(back["list"]) if s["name"] == "Excel")
        shown = c.js(f"""(() => {{ const app = __calc.app; app.activateSheet(app.wb.sheets.findIndex((s) => s.name === "Tabelle1")); app.scenarios.show(app.sheet, {idx});
          return [app.sheet.get(2, 1)?.v, app.calc.value(app.wb.active, 3, 1)]; }})()""")
        # Saved by Excel after Gut (B1 12), Dezimal (B2 0,25) and Mittel (B3 400, shown in RaabeCalc): Excel's B3 250 gives 12*0.25-250.
        check("Excel's scenario shown in RaabeCalc: B3 = 250, B4 = -247", shown[0] == 250 and same(shown[1], -247), got=shown)
        # Excel drops leftLabels when it reads a file, its own files too (checked: Excel's fixture opened and saved again).
        want_cons = {k: v for k, v in ours["consSettings"].items() if k != "left"}
        check("Konsolidieren's settings as Excel keeps them (without Linker Spalte, Excel's glitch)", back["cons"] == want_cons, ours=want_cons, back=back["cons"])
        # Excel's fixture.
        fixture = json.loads((ROOT / "packages" / "sheet-engine" / "test" / "fixtures" / "excel-scenarios.json").read_text(encoding="utf-8"))
        c.js(f"__calc.open({json.dumps(str(ROOT / 'packages' / 'ooxml' / 'test' / 'fixtures' / 'xlsx-scenarios.xlsx'))})")
        fx = c.js("__calc.app.wb.sheets[0].scenarios.list")
        for want, b in zip(fixture["final"], fx):
            check(f"fixture {want['name']}", b["name"] == want["name"] and [cc["val"] for cc in b["cells"]] == [str(v) for v in want["values"]] and bool(b.get("hidden")) == want["hidden"])
    finally:
        c.close()
    (HERE / "results" / f"{datetime.date.today().isoformat()}-calc-scenarios.json").write_text(json.dumps(report, ensure_ascii=False, indent=1, default=str), encoding="utf-8", newline="\n")
    for key in ("raabe->excel", "excel->raabe"):
        v = report[key]
        print(key, sum(1 for r in v if r["ok"]), "/", len(v))


if __name__ == "__main__":
    main()
