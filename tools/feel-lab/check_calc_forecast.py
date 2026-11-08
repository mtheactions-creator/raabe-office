"""Prognoseblatt against Excel (D-092), without real input; RaabeCalc runs off screen and never takes the focus.

RaabeCalc makes the two sheets Excel made in probe_excel_forecast.py from the same data (24 months of Besucher):
(1) the dialog's defaults (Liniendiagramm, Prognoseende 02.12.2026, 95 %), (2) Säulendiagramm with Saisonalität
manuell 6 and Prognosestatistiken. Compared with Excel's recording (results/2026-11-08-excel-forecast.json): the
table's columns, every cell's formula as Excel shows it, the dates, the number formats, the chart's type and series;
the values within the tolerances of ets-excel.test.ts. Then our own invisible Excel opens RaabeCalc's file: it finds
the tables and the chart and computes the formulas to the values it computed itself.
Result: results/<date>-calc-forecast.json.
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
TMP = Path(tempfile.gettempdir())
PORT = 9475
os.environ["RAABE_FEEL_BACKGROUND"] = "1"
calc_catalogue.winapi.place = lambda hwnd, x, y, w, h: winapi.place_quietly(hwnd, -4000, 0, w, h)
REC = json.loads((HERE / "results" / "2026-11-08-excel-forecast.json").read_text(encoding="utf-8"))


def de_date(serial: float) -> str:
    d = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(serial))
    return d.strftime("%d.%m.%Y")


def main() -> None:
    data = REC["data"]
    setup = {"A1": data[0][0], "B1": data[0][1]}
    for i, (t, v) in enumerate(data[1:], start=2):
        setup[f"A{i}"] = de_date(t)
        setup[f"B{i}"] = str(v).replace(".", ",")
    n = len(data) - 1
    runs = {
        "default": {"chart": "line", "season": 1, "stats": False},
        "options": {"chart": "column", "season": 6, "stats": True},
    }
    ours_file = TMP / "raabe-forecast-ours.xlsx"
    report: dict = {"checks": [], "runs": {}}
    c = CalcSide(PORT)
    try:
        c.reset(Case("forecast", setup, [], "A1", sheets=1))
        c.js("""(() => { const app = __calc.app; app.sheet.name = "Werte"; app.grid.refreshLayout(); })()""")
        for name, o in runs.items():
            res = c.js(f"""(() => {{ const app = __calc.app, P = __calc.parse;
              const src = app.wb.sheets.find((s) => s.name === "Werte"); app.wb.active = app.wb.sheets.indexOf(src); app.grid.refreshLayout();
              const made = app.forecast.build({{ sheet: src, timeline: P("A1:A{n + 1}"), values: P("B1:B{n + 1}"), header: true, chart: {json.dumps(o["chart"])},
                end: 46358, confidence: 0.95, season: {o["season"]}, stats: {json.dumps(o["stats"])}, fill: 1, agg: 1 }});
              const s = app.sheet; const out = {{ sheet: made, index: app.wb.sheets.indexOf(s) + 1, cells: {{}}, tables: (s.tables ?? []).map((t) => ({{ name: t.name, ref: t.ref, style: t.style, columns: t.columns.map((x) => x.name) }})),
                charts: (s.charts ?? []).map((m) => ({{ kind: m.kind, series: m.series.map((x) => ({{ name: x.name?.ref ?? null, cat: x.cat?.ref ?? null, val: x.val.ref ?? null }})), from: m.from, to: m.to, legend: m.legend }})) }};
              for (const [k, cell] of s.cells) {{ const row = Math.floor(k / 16384), col = k % 16384; if (row > 40 || col > 9) continue;
                app.grid.selectCell(row, col); const a = String.fromCharCode(65 + col) + (row + 1);
                const v = cell.v; out.cells[a] = {{ bar: document.querySelector('.fx-input')?.value ?? null, v: v && typeof v === "object" ? String(v) : v ?? null, fmt: app.wb.styles.get(cell.s).numFmt }}; }}
              return out; }})()""")
            report["runs"][name] = res
        c.js(f"__calc.save({json.dumps(str(ours_file))})")
    finally:
        c.close()

    def add(name, ok, detail=None):
        report["checks"].append({"check": name, "ok": bool(ok), **({"detail": detail} if not ok and detail is not None else {})})

    for name in runs:
        ex, us = REC["runs"][name], report["runs"][name]
        add(f"{name}: the new sheet stands before the data", us["index"] == ex["index"], [us["index"], ex["index"]])
        add(f"{name}: table columns as Excel's", [t["columns"] for t in us["tables"]] == [t["columns"] for t in ex["tables"]], [us["tables"], ex["tables"]])
        add(f"{name}: table style", all(t["style"] == "TableStyleMedium2" for t in us["tables"]))
        sheet_ex = ex["sheet"]
        f_same = v_close = total = 0
        bad = []
        for a, e in ex["cells"].items():
            u = us["cells"].get(a)
            total += 1
            if not u:
                bad.append([a, "missing"])
                continue
            want = e["f"].replace(sheet_ex + "!", "")
            have = (u["bar"] or "")
            if e["f"].startswith("="):
                if have == want:
                    f_same += 1
                else:
                    bad.append([a, have, want])
                ev, uv = e["v"], u["v"]
                if isinstance(ev, (int, float)) and isinstance(uv, (int, float)):
                    # Excel's alpha, beta and gamma are not matched (D-092): statistics are not compared, intervals loosely.
                    tol = None if "ETS.STAT" in e["f"] else 0.03 if "KONFINT" in e["f"] and e["f"].startswith("=C") else 0.5 if "KONFINT" in e["f"] else 0.025
                    if tol is None or abs(uv - ev) <= tol * max(1e-9, abs(ev)):
                        v_close += 1
                    else:
                        bad.append([a, "value", uv, ev])
            else:
                same = have.replace(".", ",") == e["text"] or have == e["f"] or str(u["v"]).replace(".", ",") == e["f"]
                f_same += 1 if same else 0
                v_close += 1
                if not same:
                    bad.append([a, have, e["f"], e["text"]])
        report["runs"][name]["summary"] = {"cells": total, "same text/formula": f_same, "values within tolerance": v_close}
        add(f"{name}: every cell's formula or value as Excel shows it ({f_same}/{total})", f_same == total, bad[:12])
        add(f"{name}: values within tolerance ({v_close}/{total})", v_close == total, bad[:12])
        ch_ex, ch_us = ex["charts"][0], us["charts"][0] if us["charts"] else None
        add(f"{name}: chart type", ch_us and ch_us["kind"] == ("line" if ex["charts"][0]["type"] == 4 else "column"), ch_us)
        add(f"{name}: chart series", ch_us and len(ch_us["series"]) == len(ch_ex["series"]), [ch_us, ch_ex])
    # Excel opens RaabeCalc's file.
    x = ExcelTarget(visible=False)
    try:
        wb = x.app.Workbooks.Open(str(ours_file))
        x.app.CalculateFull()
        for name in runs:
            us = report["runs"][name]
            ws = wb.Worksheets(us["sheet"])
            ex = REC["runs"][name]
            los = [lo.Name for lo in ws.ListObjects]
            add(f"{name}: Excel finds the tables", len(los) == len(ex["tables"]), los)
            add(f"{name}: Excel finds the chart", ws.ChartObjects().Count == 1)
            if ws.ChartObjects().Count:
                ch = ws.ChartObjects(1).Chart
                add(f"{name}: Excel's chart type and series count", (ch.ChartType, ch.SeriesCollection().Count) == (ex["charts"][0]["type"], len(ex["charts"][0]["series"])), [ch.ChartType, ch.SeriesCollection().Count])
            worst = 0.0
            for a, e in ex["cells"].items():
                if not e["f"].startswith("="):
                    continue
                v = ws.Range(a).Value2
                if isinstance(v, (int, float)) and isinstance(e["v"], (int, float)):
                    worst = max(worst, abs(v - e["v"]) / max(1e-9, abs(e["v"])))
                local = ws.Range(a).Formula2Local
                if local != e["f"]:
                    add(f"{name} {a}: Excel shows RaabeCalc's formula as its own", False, [local, e["f"]])
            add(f"{name}: Excel computes RaabeCalc's formulas to its own values (max rel. diff {worst:.2e})", worst < 1e-9)
        wb.Close(False)
    finally:
        x.close()
    ok = sum(1 for r in report["checks"] if r["ok"])
    print("checks", ok, "/", len(report["checks"]))
    for r in report["checks"]:
        if not r["ok"]:
            print("   FAIL", json.dumps(r, ensure_ascii=False, default=str)[:500])
    (HERE / "results" / f"{datetime.date.today().isoformat()}-calc-forecast.json").write_text(json.dumps(report, ensure_ascii=False, indent=1, default=str), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
