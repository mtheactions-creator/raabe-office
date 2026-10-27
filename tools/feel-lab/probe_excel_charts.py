"""Probe: Excel's charts (Step 7 package E, D-061). In an invisible Excel of our own with the
Office 2013-2022 design (Calibri, the accent colours RaabeCalc uses) charts are made through COM
(`Shapes.AddChart2` with the default style, as Einfügen > Diagramm does) and recorded:
- the automatic value-axis scaling (MinimumScale, MaximumScale, MajorUnit) for many data sets and
  several chart sizes, with the plot area's inner size the scaling depends on;
- how Excel splits a selection into series (rows or columns, names, categories) and titles it;
- the default layout of every chart kind (chart, plot area, title, legend, gap width, overlap) and
  the drawn result as SVG (`Chart.Export(..., "SVG")`), summarised as text positions and shapes;
- the histogram's automatic bins (labels and counts read from the exported SVG);
- linear trendlines: the equation and R² labels as Excel writes them;
- where a new chart is placed.
The exported pictures stay in TEMP (raabe-charts/) for visual comparison; they are not committed.
Runs invisibly, no input.
Result: results/2026-10-27-excel-charts.json (the chart engine's tests replay it).

    .venv/Scripts/python.exe probe_excel_charts.py [section ...]
"""
from __future__ import annotations

import json
import math
import os
import random
import re
import sys
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path

from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
OUT = HERE / "results" / "2026-10-27-excel-charts.json"
PICS = Path(tempfile.gettempdir()) / "raabe-charts"
THEME = r"C:\Program Files\Microsoft Office\root\Document Themes 16\Office 2013 - 2022 Theme.thmx"

KINDS = {
    "column": 51, "columnStacked": 52, "bar": 57, "barStacked": 58, "line": 4, "lineMarkers": 65,
    "area": 1, "pie": 5, "doughnut": -4120, "scatter": -4169, "scatterSmooth": 72, "scatterLines": 74,
    "histogram": 118,
}
X_VALUE, X_CATEGORY = 2, 1


# ----- helpers -----------------------------------------------------------------------------------

def fill(ws, rows: list[list]) -> str:
    """Writes rows from A1 on; returns the range address."""
    ws.Cells.Clear()
    for r, row in enumerate(rows):
        for c, v in enumerate(row):
            if v is not None and v != "":
                ws.Cells(r + 1, c + 1).Value2 = v
    return ws.Range(ws.Cells(1, 1), ws.Cells(len(rows), max(len(r) for r in rows))).Address


def column(values: list[float]) -> list[list]:
    return [[v] for v in values]


def make(ws, kind: str, addr: str, w: float = -1, h: float = -1):
    """A chart of `kind` from the selection `addr`, as Einfügen > Diagramm makes it."""
    for co in list(ws.ChartObjects()):
        co.Delete()
    ws.Range(addr).Select()
    if w > 0:
        shp = ws.Shapes.AddChart2(-1, KINDS[kind], 20, 20, w, h)
    else:
        shp = ws.Shapes.AddChart2(-1, KINDS[kind])
    return shp


def scale(ax) -> dict:
    return {"min": ax.MinimumScale, "max": ax.MaximumScale, "major": ax.MajorUnit}


def inner(ch) -> dict:
    pa = ch.PlotArea
    return {"left": r2(pa.InsideLeft), "top": r2(pa.InsideTop), "width": r2(pa.InsideWidth), "height": r2(pa.InsideHeight)}


def r2(v) -> float:
    return round(float(v), 2)


def export_svg(ch, name: str) -> dict:
    PICS.mkdir(exist_ok=True)
    p = PICS / f"{name}.svg"
    ch.Export(str(p), "SVG")
    ch.Export(str(PICS / f"{name}.png"))
    return summarise_svg(p.read_text(encoding="utf-8"))


def summarise_svg(svg: str) -> dict:
    """Texts, rectangles and lines of an exported chart in points from its top left (the SVG is
    drawn at 2 px per point)."""
    root = ET.fromstring(svg)
    ns = "{http://www.w3.org/2000/svg}"
    k = 0.5
    ox = oy = 0.0
    g = root.find(f"{ns}g")
    if g is not None and g.get("transform"):
        m = re.match(r"translate\(([-\d.]+) ([-\d.]+)\)", g.get("transform"))
        if m:
            ox, oy = float(m.group(1)), float(m.group(2))
    texts, rects, lines = [], [], []
    for el in root.iter():
        tag = el.tag.replace(ns, "")
        if tag == "text":
            m = re.match(r"translate\(([-\d.]+) ([-\d.]+)\)", el.get("transform", ""))
            if m:
                texts.append({"text": "".join(el.itertext()), "x": r2((float(m.group(1)) + ox) * k), "y": r2((float(m.group(2)) + oy) * k),
                              "size": r2(float(el.get("font-size", "0")) * k), "fill": el.get("fill")})
        elif tag == "rect" and el.get("fill") not in (None, "none", "#FFFFFF"):
            rects.append({"x": r2((float(el.get("x")) + ox) * k), "y": r2((float(el.get("y")) + oy) * k),
                          "w": r2(float(el.get("width")) * k), "h": r2(float(el.get("height")) * k), "fill": el.get("fill")})
        elif tag == "path" and el.get("d"):
            nums = [float(n) for n in re.findall(r"-?\d+(?:\.\d+)?", el.get("d"))]
            if len(nums) == 4 and el.get("stroke"):
                lines.append({"x0": r2((nums[0] + ox) * k), "y0": r2((nums[1] + oy) * k), "x1": r2((nums[2] + ox) * k), "y1": r2((nums[3] + oy) * k),
                              "stroke": el.get("stroke")})
    return {"texts": texts, "rects": rects[:60], "lines": lines[:40]}


# ----- sections ------------------------------------------------------------------------------------

def axis_cases() -> list[list[float]]:
    """Data sets for the automatic value-axis scaling: maxima across decades, minima close to the
    maximum (where Excel stops starting at 0), negatives, decimals, constants, random sets."""
    cases: list[list[float]] = []
    for mx in [0.07, 0.3, 0.0123, 1, 1.5, 2, 2.5, 3, 4, 4.8, 5, 6, 7, 8, 9, 9.5, 10, 11, 12, 13, 15, 17, 19, 20, 23, 25, 28, 30,
               37, 45, 50, 55, 60, 75, 80, 88, 95, 99, 100, 101, 105, 120, 150, 180, 200, 250, 333, 480, 500, 720, 950, 1000,
               1234, 2500, 5000, 9999, 45678, 1234567]:
        cases.append([0, mx / 3, mx])
    for mn in [0, 10, 50, 70, 75, 80, 82, 83, 84, 85, 87, 90, 95, 98, 99, 99.9]:
        cases.append([mn, (mn + 100) / 2, 100])
    for mn, mx in [(1000, 1010), (1000, 1100), (500, 520), (12.3, 12.9), (40, 47), (3, 3.2), (0.5, 0.52)]:
        cases.append([mn, mx])
    for mn, mx in [(-10, 10), (-3, 7), (-100, 20), (-5, -1), (-50, -40), (-0.5, 0.25), (-1234, 567), (-17, 0), (-8, 2)]:
        cases.append([mn, mx])
    for v in [0, 5, -5, 100, 0.5]:
        cases.append([v, v, v])
    rnd = random.Random(61)
    for _ in range(40):
        mag = 10 ** rnd.uniform(-1, 4)
        a = rnd.uniform(-0.3, 1) * mag
        b = rnd.uniform(0, 1) * mag
        vals = [round(rnd.uniform(min(a, b), max(a, b)), 3) for _ in range(5)] + [round(a, 3), round(b, 3)]
        cases.append(vals)
    return cases


def section_axes(x) -> dict:
    ws = x.sheet
    out = []
    cases = axis_cases()
    sizes = [(360, 216), (300, 150), (360, 300), (480, 400)]
    for i, vals in enumerate(cases):
        addr = fill(ws, column(vals))
        for w, h in (sizes if i % 3 == 0 else sizes[:1]):
            shp = make(ws, "column", addr, w, h)
            ch = shp.Chart
            rec = {"values": vals, "kind": "column", "size": [w, h], "axis": scale(ch.Axes(X_VALUE)), "inner": inner(ch)}
            out.append(rec)
        if i % 4 == 0:
            for kind in ["line", "bar"]:
                shp = make(ws, kind, addr)
                ch = shp.Chart
                out.append({"values": vals, "kind": kind, "size": [360, 216], "axis": scale(ch.Axes(X_VALUE)), "inner": inner(ch)})
    # Scatter: both axes are value axes.
    scat = []
    rnd = random.Random(7)
    for i in range(16):
        xs = sorted(round(rnd.uniform(-5, 40) * (1 + i), 2) for _ in range(6))
        ys = [round(rnd.uniform(-2, 30) * (1 + i % 5), 2) for _ in range(6)]
        if i % 3 == 0:
            xs = [abs(v) for v in xs]
            ys = [abs(v) for v in ys]
        addr = fill(ws, [[a, b] for a, b in zip(xs, ys)])
        shp = make(ws, "scatter", addr)
        ch = shp.Chart
        scat.append({"x": xs, "y": ys, "xAxis": scale(ch.Axes(X_CATEGORY)), "yAxis": scale(ch.Axes(X_VALUE)), "inner": inner(ch)})
    # Stacked columns scale on the sums.
    stacked = []
    for rows in [[[1, 2], [3, 4], [5, 6]], [[10, 20, 5], [3, 40, 2]], [[-3, 4], [5, -6], [2, 2]]]:
        addr = fill(ws, rows)
        shp = make(ws, "columnStacked", addr)
        stacked.append({"rows": rows, "axis": scale(shp.Chart.Axes(X_VALUE))})
    return {"axes": out, "scatter": scat, "stacked": stacked}


def series_info(ch) -> list[str]:
    sc = ch.SeriesCollection()
    return [sc.Item(i).Formula for i in range(1, sc.Count + 1)]


def section_series(x) -> list:
    """How the selection becomes series; title and legend Excel gives."""
    ws = x.sheet
    cases = {
        "3x5 texts": [[None, "Q1", "Q2", "Q3", "Q4"], ["Nord", 1, 2, 3, 4], ["Süd", 5, 6, 7, 8]],
        "5x3 texts": [[None, "A", "B"], ["Jan", 3, 5], ["Feb", 7, 2], ["Mär", 4, 9], ["Apr", 8, 6]],
        "4x4 texts": [[None, "A", "B", "C"], ["x", 1, 2, 3], ["y", 4, 5, 6], ["z", 7, 8, 9]],
        "3x3 numbers": [[1, 2, 3], [4, 5, 6], [7, 8, 9]],
        "2x6 numbers": [[1, 2, 3, 4, 5, 6], [6, 5, 4, 3, 2, 1]],
        "6x2 numbers": [[1, 6], [2, 5], [3, 4], [4, 3], [5, 2], [6, 1]],
        "one column header": [["Punkte"], [3], [5], [2], [8]],
        "one column": [[3], [5], [2], [8]],
        "one row header": [["Punkte", 3, 5, 2, 8]],
        "labels and values": [["Jan", 3], ["Feb", 5], ["Mär", 2], ["Apr", 8]],
        "labels, values, header": [["Monat", "Umsatz"], ["Jan", 3], ["Feb", 5], ["Mär", 2]],
        "years as numbers": [["Jahr", "Umsatz"], [2020, 3], [2021, 5], [2022, 2], [2023, 8]],
        "years, empty corner": [[None, "Umsatz"], [2020, 3], [2021, 5], [2022, 2]],
        "year header row": [[None, 2020, 2021, 2022], ["A", 1, 2, 3], ["B", 4, 5, 6]],
        "header no corner": [["A", "B", "C"], [1, 2, 3], [4, 5, 6], [7, 8, 9], [1, 1, 1]],
        "header with corner text": [["Name", "A", "B"], ["x", 1, 2], ["y", 3, 4], ["z", 5, 6]],
    }
    out = []
    for name, rows in cases.items():
        addr = fill(ws, rows)
        for kind in ["column", "pie", "scatter"]:
            shp = make(ws, kind, addr)
            ch = shp.Chart
            out.append({"case": name, "rows": rows, "kind": kind, "plotBy": ch.PlotBy, "series": series_info(ch),
                        "title": ch.ChartTitle.Text if ch.HasTitle else None, "legend": ch.Legend.Position if ch.HasLegend else None})
    # Active cell only: Excel takes the current region.
    fill(ws, [[None, "A", "B"], ["Jan", 3, 5], ["Feb", 7, 2], ["Mär", 4, 9]])
    ws.Range("B2").Select()
    shp = ws.Shapes.AddChart2(-1, KINDS["column"])
    out.append({"case": "active cell in region", "kind": "column", "series": series_info(shp.Chart)})
    return out


LAYOUT_DATA = [[None, "Reihe A", "Reihe B", "Reihe C"], ["Jan", 4.3, 2.4, 2], ["Feb", 2.5, 4.4, 2], ["Mär", 3.5, 1.8, 3], ["Apr", 4.5, 2.8, 5]]


def section_layout(x) -> dict:
    ws = x.sheet
    out = {}
    for kind in KINDS:
        if kind == "histogram":
            rows = column([1, 2, 2, 3, 3, 3, 4, 4, 5, 7, 8, 12, 15, 15, 16, 20])
        elif kind.startswith("scatter"):
            rows = [["x", "y"], [1, 2.2], [2, 3.9], [3, 6.1], [4, 8.2], [5, 9.7]]
        elif kind in ("pie", "doughnut"):
            rows = [[None, "Anteil"], ["Jan", 8.2], ["Feb", 3.2], ["Mär", 1.4], ["Apr", 1.2]]
        else:
            rows = LAYOUT_DATA
        addr = fill(ws, rows)
        shp = make(ws, kind, addr)
        ch = shp.Chart
        rec: dict = {"rows": rows, "object": {"width": shp.Width, "height": shp.Height}}
        try:
            rec["inner"] = inner(ch)
            pa = ch.PlotArea
            rec["plot"] = {"left": r2(pa.Left), "top": r2(pa.Top), "width": r2(pa.Width), "height": r2(pa.Height)}
        except Exception:  # noqa: BLE001 - histograms have no classic plot area
            pass
        try:
            if ch.HasTitle:
                t = ch.ChartTitle
                rec["title"] = {"text": t.Text, "left": r2(t.Left), "top": r2(t.Top), "size": t.Font.Size, "font": t.Font.Name, "color": t.Font.Color}
            if ch.HasLegend:
                lg = ch.Legend
                rec["legend"] = {"position": lg.Position, "left": r2(lg.Left), "top": r2(lg.Top), "width": r2(lg.Width), "height": r2(lg.Height), "size": lg.Font.Size}
            g = ch.ChartGroups(1)
            for p in ["GapWidth", "Overlap", "FirstSliceAngle", "DoughnutHoleSize"]:
                try:
                    rec[p] = getattr(g, p)
                except Exception:  # noqa: BLE001
                    pass
            s = ch.SeriesCollection(1)
            try:
                rec["marker"] = {"style": s.MarkerStyle, "size": s.MarkerSize}
            except Exception:  # noqa: BLE001
                pass
            try:
                rec["lineWeight"] = s.Format.Line.Weight
            except Exception:  # noqa: BLE001
                pass
        except Exception as e:  # noqa: BLE001
            rec["error"] = str(e)
        rec["svg"] = export_svg(ch, f"layout-{kind}")
        out[kind] = rec
    # With axis titles, data labels and a legend on the right, as the Diagrammelemente menu sets them.
    addr = fill(ws, LAYOUT_DATA)
    shp = make(ws, "column", addr)
    ch = shp.Chart
    ch.SetElement(301)  # msoElementPrimaryCategoryAxisTitleAdjacentToAxis
    ch.SetElement(311)  # msoElementPrimaryValueAxisTitleRotated
    ch.SetElement(101)  # msoElementLegendRight
    ch.SetElement(202)  # msoElementDataLabelShow (default position)
    rec = {"inner": inner(ch), "catTitle": ch.Axes(1).AxisTitle.Text, "valTitle": ch.Axes(2).AxisTitle.Text,
           "legend": {"left": r2(ch.Legend.Left), "top": r2(ch.Legend.Top), "width": r2(ch.Legend.Width), "height": r2(ch.Legend.Height)},
           "svg": export_svg(ch, "layout-column-elements")}
    out["columnElements"] = rec
    return out


def section_histogram(x) -> list:
    ws = x.sheet
    rnd = random.Random(118)
    sets: list[list[float]] = [
        [1, 2, 2, 3, 3, 3, 4, 4, 5, 7, 8, 12, 15, 15, 16, 20],
        list(range(1, 11)), list(range(1, 101)), [5, 5, 5, 5], [1, 2], [3.5, 4.25, 7.75, 1.5, 2.0, 9.5, 6.25],
        [0.1, 0.25, 0.33, 0.47, 0.52, 0.6, 0.61, 0.75, 0.9], [100, 250, 400, 120, 330, 800, 760, 50, 999, 640],
        [-5, -3, 0, 2, 4, 4, 6, 9, 12], [1, 1, 1, 2, 50],
    ]
    for n in [5, 12, 30, 50, 100, 200]:
        sets.append([round(rnd.gauss(50, 12), 1) for _ in range(n)])
        sets.append([rnd.randint(1, 6) for _ in range(n)])
        sets.append([round(rnd.uniform(0, 1000)) for _ in range(n)])
    out = []
    for i, vals in enumerate(sets):
        addr = fill(ws, column(vals))
        shp = make(ws, "histogram", addr)
        svg = export_svg(shp.Chart, f"hist-{i:02d}")
        bins = bin_labels(svg)
        bars = [b for b in svg["rects"] if b["fill"] == "#4472C4"]
        out.append({"values": vals, "bins": bins, "bars": [{"x": b["x"], "h": b["h"]} for b in bars],
                    "valueLabels": [t["text"] for t in svg["texts"] if not t["text"].startswith(("[", "(")) and t["text"] != "Diagrammtitel"]})
    return out


def bin_labels(svg: dict) -> list[str]:
    """The histogram's bin labels; Excel draws a minus sign as a run of its own ("[", "-", "5, 4,2]")."""
    rows: dict[float, list[dict]] = {}
    for t in svg["texts"]:
        rows.setdefault(t["y"], []).append(t)
    out: list[tuple[float, str]] = []
    for runs in rows.values():
        runs.sort(key=lambda t: t["x"])
        if not runs[0]["text"].startswith(("[", "(")):
            continue
        cur, x0 = "", 0.0
        for r in runs:
            if not cur:
                x0 = r["x"]
            cur += r["text"]
            if cur.endswith("]"):
                out.append((x0, cur))
                cur = ""
    # Crowded labels stand in two rows: back into the bins' order.
    return [t for _, t in sorted(out)]


def section_trend(x) -> list:
    ws = x.sheet
    rnd = random.Random(3)
    out = []
    sets = [
        ([1, 2, 3, 4, 5], [2.2, 3.9, 6.1, 8.2, 9.7]),
        ([1, 2, 3, 4], [3, 5, 7, 9]),
        ([1, 2, 3, 4], [9, 7, 5, 3]),
        ([0, 10, 20, 30], [1.234, 2.5, 3.1, 5.0]),
        ([1, 2, 3, 4, 5, 6], [100, 210, 290, 420, 480, 610]),
        ([1, 2, 3], [0.001, 0.002, 0.0035]),
        ([1, 2, 3, 4], [5, 5, 5, 5]),
        ([2000, 2005, 2010, 2015, 2020], [6.1, 6.5, 6.9, 7.3, 7.8]),
        ([1, 2, 3, 4, 5], [-2, -4.5, -6, -8.25, -10]),
        ([1, 2, 3, 4, 5], [12345, 23456, 34567, 45678, 56789]),
    ]
    for _ in range(6):
        xs = sorted(round(rnd.uniform(0, 50), 1) for _ in range(8))
        k, d = rnd.uniform(-3, 3), rnd.uniform(-20, 20)
        sets.append((xs, [round(k * v + d + rnd.gauss(0, 4), 2) for v in xs]))
    for i, (xs, ys) in enumerate(sets):
        addr = fill(ws, [[a, b] for a, b in zip(xs, ys)])
        shp = make(ws, "scatter", addr)
        s = shp.Chart.SeriesCollection(1)
        t = s.Trendlines().Add(-4132)  # xlLinear
        t.DisplayEquation = True
        t.DisplayRSquared = True
        # The label's Text stays empty through COM until it is drawn: read it from the picture.
        svg = export_svg(shp.Chart, f"trend-{i:02d}")
        label = trend_lines(svg)
        out.append({"x": xs, "y": ys, "label": label, "inner": inner(shp.Chart)})
    # Trendline on a column chart (categories 1..n).
    addr = fill(ws, column([3, 5, 4, 8, 9]))
    shp = make(ws, "column", addr)
    t = shp.Chart.SeriesCollection(1).Trendlines().Add(-4132)
    t.DisplayEquation = True
    t.DisplayRSquared = True
    svg = export_svg(shp.Chart, "trend-column")
    out.append({"column": [3, 5, 4, 8, 9], "label": trend_lines(svg)})
    return out


def trend_lines(svg: dict) -> list[dict]:
    """The trendline label's lines: Excel draws a line in several runs (a minus sign apart)."""
    rows: dict[float, list[dict]] = {}
    for t in svg["texts"]:
        rows.setdefault(t["y"], []).append(t)
    out = []
    for y, runs in rows.items():
        runs.sort(key=lambda t: t["x"])
        text = "".join(t["text"] for t in runs)
        if text.startswith(("y =", "R")):
            out.append({"text": text, "x": runs[0]["x"], "y": y, "size": runs[0]["size"]})
    return out


def section_limits(x) -> list:
    """How many intervals fit an axis of a given length: data 0..m with m stepping so that the
    span per unit (1.05 m with unit 1) crosses the limit; for vertical (column) and horizontal
    (bar) value axes and three label widths (1, 1000, 0.1 steps)."""
    ws = x.sheet
    out = []
    ms = [round(2.5 + 0.25 * i, 2) for i in range(29)]
    for kind, lengths in [("column", range(80, 300, 20)), ("bar", range(140, 560, 40))]:
        for factor in [1, 1000, 0.1]:
            for length in lengths:
                w, h = (360, length) if kind == "column" else (length, 216)
                for m in ms:
                    addr = fill(ws, column([0, round(m * factor, 6)]))
                    shp = make(ws, kind, addr, w, h)
                    ch = shp.Chart
                    a = scale(ch.Axes(X_VALUE))
                    out.append({"kind": kind, "factor": factor, "size": [w, h], "m": m, "axis": a, "inner": inner(ch)})
    return out


def section_place(x) -> dict:
    ws = x.sheet
    addr = fill(ws, LAYOUT_DATA)
    shp = make(ws, "column", addr)
    win = x.app.ActiveWindow
    vr = win.VisibleRange
    return {"left": shp.Left, "top": shp.Top, "width": shp.Width, "height": shp.Height,
            "visible": {"left": vr.Left, "top": vr.Top, "width": vr.Width, "height": vr.Height},
            "usable": {"width": win.UsableWidth, "height": win.UsableHeight}}


SECTIONS = {"axes": section_axes, "series": section_series, "layout": section_layout, "histogram": section_histogram,
            "trend": section_trend, "place": section_place, "limits": section_limits}


def main() -> None:
    names = sys.argv[1:] or list(SECTIONS)
    data = json.loads(OUT.read_text(encoding="utf-8")) if OUT.exists() else {}
    x = ExcelTarget(visible=False)
    try:
        x.wb.ApplyTheme(THEME)
        data["excel"] = {"version": x.app.Version, "theme": "Office 2013 - 2022", "units": "points (1 pt = 4/3 px)"}
        for n in names:
            print("section", n, flush=True)
            data[n] = SECTIONS[n](x)
    finally:
        x.close()
    text = json.dumps(data, ensure_ascii=False, indent=0)
    if re.search(r"C:\\\\Users|\\\\Users\\\\", text):
        raise SystemExit("private path in the results")
    OUT.write_text(text, encoding="utf-8", newline=chr(10))
    print("wrote", OUT.name)


if __name__ == "__main__":
    main()
