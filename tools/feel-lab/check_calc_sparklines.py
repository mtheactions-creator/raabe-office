"""Sparklines against Excel (D-085), both ways, without real input:
1. RaabeCalc (off screen, never focused) makes sparkline groups with every option of the Sparkline tab
   (its commands and the functions behind its menus: type, Anzeigen, a gallery style, Sparklinefarbe from
   the design colours, Stärke, Achse, empty and hidden cells, date axis, Gruppieren, Gruppierung aufheben,
   Löschen) and saves; our own invisible Excel opens the file and every option is read back through COM.
2. RaabeCalc opens the workbook Excel saved in probe_excel_sparklines.py (TEMP/raabe-sparklines-excel.xlsx)
   and must hold every group as COM read it there (fixture excel-sparklines.json); saved again unchanged, the
   sparkline XML must be Excel's byte for byte, and Excel must open that file with the same groups.
Result: results/<date>-calc-sparklines.json.

    .venv/Scripts/python.exe check_calc_sparklines.py
"""
from __future__ import annotations

import datetime
import json
import os
import re
import tempfile
import zipfile
from pathlib import Path

from feellab import calc_catalogue, winapi
from feellab.calc_catalogue import Case, CalcSide
from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
ROOT = HERE.parents[1]
TMP = Path(tempfile.gettempdir())
FIXTURE = ROOT / "packages" / "ooxml" / "test" / "fixtures" / "excel-sparklines.json"

# RaabeCalc runs off screen and never takes the focus (someone may be working at the PC).
os.environ["RAABE_FEEL_BACKGROUND"] = "1"
calc_catalogue.winapi.place = lambda hwnd, x, y, w, h: winapi.place_quietly(hwnd, -4000, 0, w, h)

DATA = {"A1": "3", "B1": "5", "C1": "-2", "D1": "8", "E1": "4",
        "A2": "1", "C2": "4", "D2": "-3", "E2": "6",
        "A3": "2", "B3": "7", "C3": "7", "D3": "1", "E3": "5",
        "A8": "01.01.2026", "B8": "02.01.2026", "C8": "05.01.2026", "D8": "06.01.2026", "E8": "12.01.2026"}

# Built in RaabeCalc: (location column, what is done) through the tab's commands and functions.
BUILD = r"""(async () => {
  const app = __calc.app, sp = app.sparklines, P = __calc.parse, cmd = (id) => app.commands.run(id);
  const sel = (a, act) => { const r = P(a); const at = act ? P(act) : r; app.grid.select(r, { row: at.r0, col: at.c0 }); };
  const add = (t, d, p) => sp.add(t, P(d), P(p));
  // F: a line group with every Anzeigen box, a gallery style, Stärke, axis, empty/hidden cells, vertical axis.
  add("line", "A1:E3", "F1:F3"); sel("F1");
  for (const c of ["spark.high", "spark.low", "spark.negative", "spark.first", "spark.last", "spark.markers"]) await cmd(c);
  sp.applyStyle(1);
  sp.edit("Stärke", (g) => { g.lineWeight = 2.25; });
  sp.edit("Achse", (g) => { g.displayXAxis = true; g.rightToLeft = true; });
  sp.edit("Ausgeblendete und leere Zellen", (g) => { g.displayEmptyCellsAs = "span"; g.displayHidden = true; });
  sp.edit("Achse", (g) => { g.minAxisType = "custom"; g.manualMin = -5; });
  sp.axisType("max", "group");
  // G: column, switched to Gewinn/Verlust and back to Säule, Sparklinefarbe from the palette (Akzent 2, dunkler 25 %).
  add("column", "A1:E3", "G1:G3"); sel("G1");
  await cmd("spark.typeWinLoss"); await cmd("spark.typeColumn"); await cmd("spark.high");
  sp.setColor("series", "accent2,75000", "Sparklinefarbe");
  sp.setColor("high", "FF0000", "Datenpunktfarbe");
  // H: Gewinn/Verlust with a date axis and Null for empty cells.
  add("stacked", "A2:E2", "H2"); sel("H2");
  sp.edit("Datumsachsentyp", (g) => { g.dateRange = "Tabelle1!A8:E8"; g.displayEmptyCellsAs = "zero"; });
  sp.axisType("min", "group");
  // I: Gruppierung aufheben of I1:I3, then Gruppieren of I1:I2 led by I2.
  add("line", "A1:E3", "I1:I3"); sel("I1:I3"); await cmd("spark.ungroup");
  sel("I1:I2", "I2");
  await cmd("spark.group");
  // J: Löschen of J2 alone; K: Löschen of the group touched by K2.
  add("line", "A1:E3", "J1:J3"); sel("J2"); await cmd("spark.clearSelected");
  add("column", "A1:E3", "K1:K3"); sel("K2"); await cmd("spark.clearGroups");
  return JSON.parse(JSON.stringify(app.sheet.sparklines));
})()"""

XL_TYPE = {"line": 1, "column": 2, "stacked": 3}
XL_BLANKS = {"gap": 1, "zero": 2, "span": 3}
XL_AXIS = {"individual": 2, "group": 1, "custom": 3}
POINTS = {"Highpoint": "high", "Lowpoint": "low", "Firstpoint": "first", "Lastpoint": "last", "Negative": "negative", "Markers": "markers"}


def bgr(v) -> str:
    v = int(v)
    return "#%02X%02X%02X" % (v & 0xFF, (v >> 8) & 0xFF, (v >> 16) & 0xFF)


def com_groups(ws) -> list[dict]:
    out = []
    gs = ws.Cells.SparklineGroups
    for i in range(1, gs.Count + 1):
        g = gs(i)
        p, ax = g.Points, g.Axes
        d = {"location": g.Location.Address.replace("$", ""), "count": g.Count, "type": g.Type, "lineWeight": g.LineWeight,
             "blanks": g.DisplayBlanksAs, "hidden": bool(g.DisplayHidden), "series": bgr(g.SeriesColor.Color),
             "axis": bool(ax.Horizontal.Axis.Visible), "axisColor": bgr(ax.Horizontal.Axis.Color.Color), "rtl": bool(ax.Horizontal.RightToLeftPlotOrder),
             "date": bool(ax.Horizontal.IsDateAxis), "minType": ax.Vertical.MinScaleType, "maxType": ax.Vertical.MaxScaleType,
             "sources": [g.Item(k).SourceData for k in range(1, g.Count + 1)]}
        if d["minType"] == 3:
            d["min"] = ax.Vertical.CustomMinScaleValue
        if d["maxType"] == 3:
            d["max"] = ax.Vertical.CustomMaxScaleValue
        if d["date"]:
            d["dateRange"] = g.DateRange
        for k, m in POINTS.items():
            d[m] = [bool(getattr(p, k).Visible), bgr(getattr(p, k).Color.Color)]
        out.append(d)
    return out


def cell(row: int, col: int) -> str:
    s, c = "", col + 1
    while c:
        c, m = divmod(c - 1, 26)
        s = chr(65 + m) + s
    return f"{s}{row + 1}"


def expect_of(g: dict) -> dict:
    """What COM must report for one of RaabeCalc's groups."""
    want = {"count": len(g["items"]), "type": XL_TYPE[g["type"]], "lineWeight": g.get("lineWeight", 0.75), "blanks": XL_BLANKS[g["displayEmptyCellsAs"]],
            "hidden": bool(g.get("displayHidden")), "series": g["colors"]["series"]["rgb"].upper(), "axis": bool(g.get("displayXAxis")),
            "axisColor": g["colors"]["axis"]["rgb"].upper(), "rtl": bool(g.get("rightToLeft")), "date": bool(g.get("dateRange")),
            "minType": XL_AXIS[g.get("minAxisType", "individual")], "maxType": XL_AXIS[g.get("maxAxisType", "individual")],
            "sources": [i["f"].split("!")[1] for i in g["items"]]}
    if want["minType"] == 3:
        want["min"] = g["manualMin"]
    if want["maxType"] == 3:
        want["max"] = g["manualMax"]
    if want["date"]:
        want["dateRange"] = g["dateRange"].split("!")[1]
    for m in POINTS.values():
        want[m] = [bool(g.get(m)), g["colors"][m]["rgb"].upper()]
    return want


def close_rgb(a: str, b: str) -> bool:
    """Colours from tints may differ by one step of rounding."""
    return all(abs(int(a[i:i + 2], 16) - int(b[i:i + 2], 16)) <= 1 for i in (1, 3, 5))


def compare(want: dict, got: dict) -> list[str]:
    bad = []
    for k, v in want.items():
        g = got.get(k)
        if isinstance(v, str) and v.startswith("#"):
            ok = isinstance(g, str) and close_rgb(v, g)
        elif isinstance(v, list) and len(v) == 2 and isinstance(v[1], str) and v[1].startswith("#"):
            ok = isinstance(g, list) and g[0] == v[0] and close_rgb(v[1], g[1])
        elif isinstance(v, float) or isinstance(g, float):
            ok = g is not None and abs(float(g) - float(v)) < 1e-9
        else:
            ok = g == v
        if not ok:
            bad.append(f"{k}: want {v!r}, Excel {g!r}")
    return bad


def main() -> None:
    report: dict = {"raabe->excel": [], "excel->raabe": [], "verbatim": [], "excel-reopen": []}
    ours_file = TMP / "raabe-sparklines-ours.xlsx"
    c = CalcSide()
    try:
        c.reset(Case("sparklines", DATA, [], "A1", sheets=1))
        ours = c.js(BUILD)
        c.js(f"__calc.save({json.dumps(str(ours_file))})")
    finally:
        c.close()
    report["ours"] = ours
    x = ExcelTarget(visible=False)
    try:
        wb = x.app.Workbooks.Open(str(ours_file))
        got = com_groups(wb.Worksheets(1))
        wb.Close(False)
    finally:
        x.close()
    for g in ours:
        locs = [cell(i["row"], i["col"]) for i in g["items"]]
        match = next((e for e in got if e["location"].replace(",", ":").split(":")[0] == locs[0]), None)
        bad = ["no group in Excel"] if match is None else compare(expect_of(g), match)
        report["raabe->excel"].append({"group": ",".join(locs), "ok": not bad, **({"diff": bad} if bad else {})})
        print("R->E", ",".join(locs), "ok" if not bad else bad, flush=True)
    n_ours = len(ours)
    report["raabe->excel"].append({"group": "number of groups", "ok": len(got) == n_ours, **({"diff": [f"{n_ours} vs {len(got)}"]} if len(got) != n_ours else {})})

    # 2. Excel's own file.
    excel_file = TMP / "raabe-sparklines-excel.xlsx"
    fixture = json.loads(FIXTURE.read_text(encoding="utf-8"))
    if not excel_file.exists():
        print("run probe_excel_sparklines.py first")
    else:
        resaved = TMP / "raabe-sparklines-resaved.xlsx"
        c = CalcSide()
        try:
            c.js(f"__calc.open({json.dumps(str(excel_file))})")
            sheets = c.js("""(() => Object.fromEntries(__calc.app.wb.sheets.map((s) => [s.name, JSON.parse(JSON.stringify(s.sparklines ?? []))])))()""")
            c.js(f"__calc.save({json.dumps(str(resaved))})")
        finally:
            c.close()
        for name, case in fixture["cases"].items():
            if "groups" not in case:
                continue
            ours_g = sheets.get(name[:31], [])
            bad = []
            if len(ours_g) != len(case["groups"]):
                bad.append(f"{len(ours_g)} groups, Excel {len(case['groups'])}")
            for com in case["groups"]:
                first = re.split("[:,]", com["location"])[0]
                g = next((o for o in ours_g if any(cell(i["row"], i["col"]) == first for i in o["items"])), None)
                if g is None:
                    bad.append(f"{first}: missing")
                    continue
                got = {"count": com["count"], "type": com["type"], "lineWeight": com["lineWeight"], "blanks": com["blanks"], "hidden": com["hidden"],
                       "series": com["series"]["rgb"], "axis": com["axis"]["visible"], "axisColor": com["axis"]["color"]["rgb"], "rtl": com["axis"]["rtl"],
                       "date": com["axis"]["date"], "minType": com["axis"]["minType"], "maxType": com["axis"]["maxType"]}
                if com["axis"]["minType"] == 3:
                    got["min"] = com["axis"]["min"]
                if com["axis"]["maxType"] == 3:
                    got["max"] = com["axis"]["max"]
                if com["axis"]["date"]:
                    got["dateRange"] = com["dateRange"]
                for k, m in POINTS.items():
                    got[m] = [com[k]["visible"], com[k]["color"]["rgb"]]
                want = expect_of(g)
                want.pop("sources")
                bad += [f"{first} {d}" for d in compare(want, got)]
            report["excel->raabe"].append({"sheet": name, "ok": not bad, **({"diff": bad} if bad else {})})
            print("E->R", name, "ok" if not bad else bad, flush=True)

        # Saved again unchanged: Excel's sparkline XML byte for byte.
        def exts(path: Path) -> dict:
            out = {}
            with zipfile.ZipFile(path) as z:
                wbx = z.read("xl/workbook.xml").decode("utf-8")
                rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
                targets = {m.group(2): m.group(1) for m in re.finditer(r'Target="([^"]+)"[^>]*Id="(\w+)"', rels)}
                targets.update({m.group(1): m.group(2) for m in re.finditer(r'Id="(\w+)"[^>]*Target="([^"]+)"', rels)})
                for sname, rid in re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(\w+)"', wbx):
                    t = targets[rid].lstrip("/")
                    xml = z.read(t if t.startswith("xl/") else "xl/" + t).decode("utf-8")
                    m = re.search(r"<x14:sparklineGroups\b[\s\S]*?</x14:sparklineGroups>", xml)
                    out[sname] = m.group(0) if m else None
            return out
        before, after = exts(excel_file), exts(resaved)
        for name, xml in before.items():
            if xml is None:
                continue
            same = after.get(name) == xml
            report["verbatim"].append({"sheet": name, "ok": same})
        print("verbatim", sum(r["ok"] for r in report["verbatim"]), "/", len(report["verbatim"]), flush=True)
        x = ExcelTarget(visible=False)
        try:
            wb = x.app.Workbooks.Open(str(resaved))
            for name, case in fixture["cases"].items():
                if "groups" not in case:
                    continue
                got = com_groups(wb.Worksheets(name[:31]))
                ok = len(got) == len(case["groups"]) and sorted(g["location"] for g in got) == sorted(g["location"] for g in case["groups"])
                report["excel-reopen"].append({"sheet": name, "ok": ok})
            wb.Close(False)
        finally:
            x.close()
        print("excel-reopen", sum(r["ok"] for r in report["excel-reopen"]), "/", len(report["excel-reopen"]), flush=True)

    out = HERE / "results" / f"{datetime.date.today().isoformat()}-calc-sparklines.json"
    out.write_text(json.dumps(report, ensure_ascii=False, indent=1, default=str), encoding="utf-8", newline="\n")
    for key in ("raabe->excel", "excel->raabe", "verbatim", "excel-reopen"):
        v = report[key]
        print(key, sum(1 for r in v if r["ok"]), "/", len(v))


if __name__ == "__main__":
    main()
