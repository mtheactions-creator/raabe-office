"""Chart formats against Excel (D-066), both ways, without real input:
1. RaabeCalc (off screen) makes charts with the Format pane's settings and saves them; our own
   invisible Excel opens the file and reads each setting through COM.
2. Excel makes the same charts with the settings through COM and saves them; RaabeCalc opens the file
   and its chart models must hold the same settings.
Result: results/<date>-chart-format.json; prints each difference.
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
DATA = {"A1": "Monat", "B1": "A", "C1": "B", "A2": "Jan", "B2": "3", "C2": "4", "A3": "Feb", "B3": "-1", "C3": "2", "A4": "Mrz", "B4": "5", "C4": "6",
        "A5": "Apr", "B5": "2", "C5": "3"}

# What RaabeCalc sets on each chart (model patches) and what Excel's COM must then report.
CHARTS = [
    ("column", {
        "gapWidth": 60, "overlap": 15, "dataLabels": True,
        "format": {
            "chart": {"fill": {"type": "solid", "color": "#FFF2CC"}, "line": {"type": "solid", "color": "#C00000", "width": 2, "dash": "dash"}},
            "plot": {"fill": {"type": "solid", "color": "#DDEBF7"}},
            "title": {"text": {"size": 18, "bold": True}},
            "valAxis": {"min": -2, "max": 8, "majorUnit": 2, "minorUnit": 0.5, "reverse": True, "numFmt": "0.0", "majorTick": "out", "labelPos": "low"},
            "catAxis": {"reverse": True},
            "majorGrid": {"line": {"type": "solid", "color": "#A6A6A6", "width": 1.5, "dash": "sysDot"}},
            "minorGrid": {"show": True},
            "legend": {"overlay": True},
            "dataLabels": {"showCat": True, "showSer": True, "pos": "inEnd", "separator": ", "},
        },
        "series": [{"fill": {"type": "solid", "color": "#70AD47"}, "invertIfNegative": True}, {"trendline": {"type": "poly", "order": 3, "showEq": True, "showR2": True, "forward": 1, "name": "Kurve"}}],
    }),
    ("lineMarkers", {
        "series": [{"marker": {"symbol": "diamond", "size": 9}, "smooth": True, "line": {"type": "solid", "color": "#7030A0", "width": 3, "dash": "lgDash"}},
                   {"errorBars": {"dir": "plus", "type": "fixedVal", "value": 1, "noEndCap": True}}],
    }),
    ("doughnut", {"firstSliceAngle": 90, "holeSize": 40}),
]

# MsoLineDashStyle (msoLineSysDash 10 and msoLineSysDot 11 are DrawingML's sysDash and sysDot).
DASH = {"solid": 1, "sysDash": 10, "sysDot": 11, "dash": 4, "dashDot": 5, "lgDash": 7, "lgDashDot": 8, "lgDashDotDot": 9}
TICK = {"none": -4142, "in": 2, "out": 3, "cross": 4}
LBLPOS = {"nextTo": 4, "high": -4127, "low": -4134, "none": -4142}
DLPOS = {"ctr": -4108, "inEnd": 3, "inBase": 4, "outEnd": 2, "bestFit": 5, "l": -4131, "r": -4152, "t": 0, "b": 1}
MARKER = {"square": 1, "diamond": 2, "triangle": 3, "star": 5, "circle": 8, "plus": 9, "x": -4168, "dot": -4118, "dash": -4115, "none": -4142}
TREND = {"linear": -4132, "log": -4133, "poly": 3, "power": 4, "exp": 5, "movingAvg": 6}


def rgb(hex_: str) -> int:
    r, g, b = int(hex_[1:3], 16), int(hex_[3:5], 16), int(hex_[5:7], 16)
    return r + (g << 8) + (b << 16)


def expect_of(kind: str, p: dict) -> list[tuple[str, object]]:
    """(COM expression on the chart `ch`, expected value) for every setting in the patch."""
    out: list[tuple[str, object]] = []
    f = p.get("format", {})
    if "gapWidth" in p: out.append(("ch.ChartGroups(1).GapWidth", p["gapWidth"]))
    if "overlap" in p: out.append(("ch.ChartGroups(1).Overlap", p["overlap"]))
    if "firstSliceAngle" in p: out.append(("ch.ChartGroups(1).FirstSliceAngle", p["firstSliceAngle"]))
    if "holeSize" in p: out.append(("ch.ChartGroups(1).DoughnutHoleSize", p["holeSize"]))
    ca = f.get("chart", {})
    if "fill" in ca: out.append(("ch.ChartArea.Format.Fill.ForeColor.RGB", rgb(ca["fill"]["color"])))
    if "line" in ca:
        out += [("ch.ChartArea.Format.Line.ForeColor.RGB", rgb(ca["line"]["color"])), ("ch.ChartArea.Format.Line.Weight", ca["line"]["width"]),
                ("ch.ChartArea.Format.Line.DashStyle", DASH[ca["line"]["dash"]])]
    if "plot" in f: out.append(("ch.PlotArea.Format.Fill.ForeColor.RGB", rgb(f["plot"]["fill"]["color"])))
    t = f.get("title", {}).get("text", {})
    if "size" in t: out.append(("ch.ChartTitle.Format.TextFrame2.TextRange.Font.Size", t["size"]))
    if t.get("bold"): out.append(("ch.ChartTitle.Format.TextFrame2.TextRange.Font.Bold", -1))
    va = f.get("valAxis", {})
    for k, expr in [("min", "MinimumScale"), ("max", "MaximumScale"), ("majorUnit", "MajorUnit"), ("minorUnit", "MinorUnit")]:
        if k in va: out.append((f"ch.Axes(2).{expr}", va[k]))
    if va.get("reverse"): out.append(("ch.Axes(2).ReversePlotOrder", True))
    if "numFmt" in va: out.append(("ch.Axes(2).TickLabels.NumberFormat", va["numFmt"]))
    if "majorTick" in va: out.append(("ch.Axes(2).MajorTickMark", TICK[va["majorTick"]]))
    if "labelPos" in va: out.append(("ch.Axes(2).TickLabelPosition", LBLPOS[va["labelPos"]]))
    if f.get("catAxis", {}).get("reverse"): out.append(("ch.Axes(1).ReversePlotOrder", True))
    g = f.get("majorGrid", {}).get("line")
    if g: out += [("ch.Axes(2).MajorGridlines.Format.Line.ForeColor.RGB", rgb(g["color"])), ("ch.Axes(2).MajorGridlines.Format.Line.Weight", g["width"]),
                  ("ch.Axes(2).MajorGridlines.Format.Line.DashStyle", DASH[g["dash"]])]
    if f.get("minorGrid", {}).get("show"): out.append(("ch.Axes(2).HasMinorGridlines", True))
    if f.get("legend", {}).get("overlay"): out.append(("ch.Legend.IncludeInLayout", False))
    dl = f.get("dataLabels")
    if dl:
        out += [("ch.SeriesCollection(1).DataLabels().ShowCategoryName", True), ("ch.SeriesCollection(1).DataLabels().ShowSeriesName", True),
                ("ch.SeriesCollection(1).DataLabels().Position", DLPOS[dl["pos"]]), ("ch.SeriesCollection(1).DataLabels().Separator", dl["separator"])]
    for i, s in enumerate(p.get("series", []), start=1):
        sc = f"ch.SeriesCollection({i})"
        if "fill" in s: out.append((f"{sc}.Format.Fill.ForeColor.RGB", rgb(s["fill"]["color"])))
        if s.get("invertIfNegative"): out.append((f"{sc}.InvertIfNegative", True))
        if "line" in s:
            out += [(f"{sc}.Format.Line.ForeColor.RGB", rgb(s["line"]["color"])), (f"{sc}.Format.Line.Weight", s["line"]["width"]), (f"{sc}.Format.Line.DashStyle", DASH[s["line"]["dash"]])]
        if "marker" in s: out += [(f"{sc}.MarkerStyle", MARKER[s["marker"]["symbol"]]), (f"{sc}.MarkerSize", s["marker"]["size"])]
        if s.get("smooth"): out.append((f"{sc}.Smooth", True))
        tl = s.get("trendline")
        if tl:
            out += [(f"{sc}.Trendlines(1).Type", TREND[tl["type"]]), (f"{sc}.Trendlines(1).DisplayEquation", tl["showEq"]), (f"{sc}.Trendlines(1).DisplayRSquared", tl["showR2"])]
            if "order" in tl: out.append((f"{sc}.Trendlines(1).Order", tl["order"]))
            if "forward" in tl: out.append((f"{sc}.Trendlines(1).Forward", tl["forward"]))
            if "name" in tl: out.append((f"{sc}.Trendlines(1).Name", tl["name"]))
        eb = s.get("errorBars")
        if eb: out.append((f"{sc}.ErrorBars.EndStyle", 2 if eb.get("noEndCap") else 1))
    return out


def close(a, b) -> bool:
    if isinstance(a, (int, float)) and isinstance(b, (int, float)) and not isinstance(a, bool):
        return abs(float(a) - float(b)) < 0.01
    return a == b


def main() -> None:
    raabe_file = TMP / "raabe-chart-format.xlsx"
    excel_file = TMP / "excel-chart-format.xlsx"
    report: dict = {"raabe->excel": [], "excel->raabe": []}
    # 1. RaabeCalc writes.
    c = CalcSide()
    try:
        c.reset(Case("charts", DATA, [], "x"))
        for i, (kind, patch) in enumerate(CHARTS):
            c.js(f"__calc.app.grid.select({{r0:0,c0:0,r1:4,c1:2}}); __calc.app.charts.insert({json.dumps(kind)})")
            c.js(f"""(() => {{ const m = __calc.app.sheet.charts[{i}]; const p = {json.dumps(patch)};
              for (const [k, v] of Object.entries(p)) if (k !== 'series') m[k] = v;
              (p.series || []).forEach((s, j) => Object.assign(m.series[j], s)); delete m.raw; }})()""")
        c.js(f"__calc.save({json.dumps(str(raabe_file))})")
    finally:
        c.close()
    x = ExcelTarget(visible=False)
    try:
        wb = x.app.Workbooks.Open(str(raabe_file))
        ws = wb.Worksheets(1)
        for i, (kind, patch) in enumerate(CHARTS, start=1):
            ch = ws.ChartObjects(i).Chart  # noqa: F841 (used by eval)
            for expr, want in expect_of(kind, patch):
                try:
                    got = eval(expr)
                except Exception as e:  # noqa: BLE001
                    got = f"error {str(e)[:60]}"
                ok = close(got, want)
                report["raabe->excel"].append({"chart": kind, "what": expr, "want": want, "got": got, "ok": ok})
                if not ok:
                    print("R->E", kind, expr, "want", want, "got", got)
        wb.Close(False)
        # 2. Excel writes the same settings.
        wb = x.app.Workbooks.Add()
        ws = wb.Worksheets(1)
        for a, v in DATA.items():
            ws.Range(a).FormulaLocal = v
        for i, (kind, patch) in enumerate(CHARTS, start=1):
            xl = {"column": 51, "lineMarkers": 65, "doughnut": -4120}[kind]
            ch = ws.Shapes.AddChart2(-1, xl, 300, 20 + 260 * (i - 1), 360, 216).Chart
            ch.SetSourceData(ws.Range("A1:C5"))
            if patch.get("dataLabels"):
                ch.SetElement(202)
            for expr, want in expect_of(kind, patch):
                obj, prop = expr.rsplit(".", 1)
                try:
                    if prop == "HasMinorGridlines" or prop.startswith("Display") or prop in ("Smooth", "InvertIfNegative", "ReversePlotOrder", "IncludeInLayout"):
                        setattr(eval(obj), prop, want)
                    elif expr.endswith("Trendlines(1).Type"):
                        eval(obj.rsplit(".", 1)[0]).Trendlines().Add(want)
                    elif prop == "EndStyle":
                        # ErrorBar(Direction xlY, Include plus, Type fixed, Amount 1), then the end style.
                        ser = eval(obj.rsplit(".", 1)[0])
                        ser.ErrorBar(1, 2, 1, 1)
                        ser.ErrorBars.EndStyle = want
                    elif prop == "NumberFormat":
                        # COM's NumberFormat is read in the local form here: "0,0" is the file's "0.0".
                        eval(obj).NumberFormatLocal = want.replace(".", ",")
                    else:
                        setattr(eval(obj), prop, want)
                except Exception as e:  # noqa: BLE001
                    print("set", expr, "failed", str(e)[:80])
        wb.SaveAs(str(excel_file), 51)
        wb.Close(False)
    finally:
        x.close()
    c = CalcSide()
    try:
        c.js(f"__calc.open({json.dumps(str(excel_file))})")
        models = c.js("JSON.parse(JSON.stringify(__calc.app.sheet.charts.map((m) => { const { raw, ...rest } = m; return rest; })))")
    finally:
        c.close()
    for (kind, patch), m in zip(CHARTS, models):
        def cmp(path: str, want, got):
            ok = close(got, want) if not isinstance(want, dict) else all(close(got.get(k) if isinstance(got, dict) else None, v) for k, v in want.items())
            # A series' solid fill is kept as its colour (the same in the file and on screen).
            if not ok and path.endswith(".fill") and isinstance(want, dict) and want.get("type") == "solid":
                ok = m["series"][int(path.split("[")[1].split("]")[0])].get("color") == want.get("color")
            report["excel->raabe"].append({"chart": kind, "what": path, "want": want, "got": got, "ok": ok})
            if not ok:
                print("E->R", kind, path, "want", want, "got", got)
        for k in ("gapWidth", "overlap", "firstSliceAngle", "holeSize", "dataLabels"):
            if k in patch:
                cmp(k, patch[k], m.get(k))
        for k, v in patch.get("format", {}).items():
            got = m.get("format", {}).get(k, {})
            for kk, vv in v.items():
                cmp(f"format.{k}.{kk}", vv, got.get(kk))
        for i, s in enumerate(patch.get("series", [])):
            for kk, vv in s.items():
                cmp(f"series[{i}].{kk}", vv, m["series"][i].get(kk))
    out = HERE / "results" / f"{datetime.date.today().isoformat()}-chart-format.json"
    out.write_text(json.dumps(report, ensure_ascii=False, indent=1, default=str), encoding="utf-8", newline="\n")
    for k, v in report.items():
        print(k, sum(1 for r in v if r["ok"]), "/", len(v))


if __name__ == "__main__":
    main()
