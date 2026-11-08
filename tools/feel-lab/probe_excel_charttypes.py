"""Probe: every chart type of Excel 365 except the map (D-086). In an invisible Excel of our own with
the Office 2013-2022 design, one chart per type is made through COM (`Shapes.AddChart2` with the
default style, as Einfügen > Diagramme does) from a fixed data block on a sheet of its own, then
recorded: Chart.ChartType, the series (formula, type, axis group), title, legend, the axes and their
scaling, the chart group's options (gap width, split of Kreis aus Kreis, bubble scale, hi-lo lines,
up/down bars), the 3-D view (Elevation, Rotation, Perspective, RightAngleAxes, DepthPercent) and the
plot area. Each chart is exported as PNG (TEMP/raabe-charttypes/<id>.png) for the visual comparison
(compare_calc_charttypes.py); the workbook is saved in TEMP and copied, without the user name and
the save path (anonymize_office.py), to packages/ooxml/test/fixtures/excel-charttypes.xlsx (the unit tests read it).
Runs invisibly, no input.
Result: results/2026-11-08-excel-charttypes.json and packages/ooxml/test/fixtures/excel-charttypes.json.

    .venv/Scripts/python.exe probe_excel_charttypes.py
"""
from __future__ import annotations

import json
import os
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
ROOT = HERE.parents[1]
OUT = HERE / "results" / "2026-11-08-excel-charttypes.json"
FIX = ROOT / "packages" / "ooxml" / "test" / "fixtures"
PICS = Path(tempfile.gettempdir()) / "raabe-charttypes"
THEME = r"C:\Program Files\Microsoft Office\root\Document Themes 16\Office 2013 - 2022 Theme.thmx"
NEUTRAL = "Raabe Office"

CAT = [["", "Reihe A", "Reihe B", "Reihe C"], ["Jan", 4.3, 2.4, 2], ["Feb", 2.5, 4.4, 2], ["Mrz", 3.5, 1.8, 3], ["Apr", 4.5, 2.8, 5], ["Mai", 3.2, 3.6, 4]]
PIE = [["Region", "Umsatz"], ["Nord", 35], ["Süd", 25], ["Ost", 20], ["West", 12], ["Mitte", 8]]
PIE7 = [["Posten", "Anteil"], ["A", 30], ["B", 25], ["C", 18], ["D", 12], ["E", 7], ["F", 5], ["G", 3]]
BUBBLE = [["X", "Y", "Größe"], [1, 3, 10], [2, 5, 4], [3, 2, 8], [4, 6, 2], [5, 4, 6]]
DATES = [46082 + i for i in range(5)]  # 1.3.2026 …
HLC = [["Datum", "Hoch", "Tief", "Schluss"]] + [[d, h, l, c] for d, h, l, c in zip(DATES, [55, 57, 56, 60, 58], [50, 52, 51, 54, 53], [53, 56, 52, 58, 54])]
OHLC = [["Datum", "Eröffnung", "Hoch", "Tief", "Schluss"]] + [[d, o, h, l, c] for d, o, h, l, c in zip(DATES, [51, 53, 55, 53, 58], [55, 57, 56, 60, 58], [50, 52, 51, 52, 53], [53, 56, 52, 58, 54])]
VHLC = [["Datum", "Volumen", "Hoch", "Tief", "Schluss"]] + [[d, v, h, l, c] for d, v, h, l, c in zip(DATES, [1200, 900, 1500, 1100, 800], [55, 57, 56, 60, 58], [50, 52, 51, 54, 53], [53, 56, 52, 58, 54])]
VOHLC = [["Datum", "Volumen", "Eröffnung", "Hoch", "Tief", "Schluss"]] + [[d, v, o, h, l, c] for d, v, o, h, l, c in zip(DATES, [1200, 900, 1500, 1100, 800], [51, 53, 55, 53, 58], [55, 57, 56, 60, 58], [50, 52, 51, 52, 53], [53, 56, 52, 58, 54])]
SURFACE = [["", "x1", "x2", "x3", "x4", "x5"]] + [[f"y{r}"] + [round(50 + 40 * ((c - 3) ** 2 - (r - 3) ** 2) / 4 + 5 * r, 1) for c in range(1, 6)] for r in range(1, 6)]
HIER = [["Bereich", "Gruppe", "Wert"], ["Obst", "Äpfel", 30], ["Obst", "Birnen", 20], ["Obst", "Kirschen", 10], ["Gemüse", "Möhren", 25],
        ["Gemüse", "Kohl", 15], ["Brot", "Vollkorn", 18], ["Brot", "Weißbrot", 12]]
WATERFALL = [["Posten", "Betrag"], ["Start", 100], ["Verkauf", 40], ["Miete", -30], ["Material", -15], ["Zuschuss", 25], ["Ende", 120]]
FUNNEL = [["Stufe", "Anzahl"], ["Besucher", 1000], ["Interessenten", 600], ["Angebote", 300], ["Verhandlungen", 150], ["Abschlüsse", 80]]
BOX = [["Klasse 8a", "Klasse 8b", "Klasse 8c"], [3, 2, 4], [5, 3, 4], [4, 4, 5], [6, 3, 3], [2, 5, 4], [4, 4, 6], [5, 2, 4], [3, 3, 5], [4, 6, 4], [5, 3, 2], [4, 4, 4], [12, 3, 5]]
PARETO = [["Fehler", "Anzahl"], ["Kratzer", 12], ["Delle", 30], ["Farbe", 7], ["Riss", 22], ["Sonstiges", 4]]
HIST = [["Punkte"]] + [[v] for v in [12, 15, 18, 22, 25, 27, 28, 30, 31, 33, 35, 36, 38, 40, 41, 44, 47, 52, 55, 61]]

# id: (xlChartType, data, after-make tweak)
TYPES: dict[str, tuple[int, list, str | None]] = {
    "columnStacked100": (53, CAT, None), "barStacked100": (59, CAT, None),
    "lineStacked": (63, CAT, None), "lineStacked100": (64, CAT, None), "lineMarkersStacked": (66, CAT, None), "lineMarkersStacked100": (67, CAT, None),
    "areaStacked": (76, CAT, None), "areaStacked100": (77, CAT, None),
    "column3DClustered": (54, CAT, None), "column3DStacked": (55, CAT, None), "column3DStacked100": (56, CAT, None), "column3D": (-4100, CAT, None),
    "bar3DClustered": (60, CAT, None), "bar3DStacked": (61, CAT, None), "bar3DStacked100": (62, CAT, None),
    "line3D": (-4101, CAT, None), "area3D": (-4098, CAT, None), "area3DStacked": (78, CAT, None), "area3DStacked100": (79, CAT, None),
    "pie3D": (-4102, PIE, None), "pieOfPie": (68, PIE7, None), "barOfPie": (71, PIE7, None),
    "radar": (-4151, CAT, None), "radarMarkers": (81, CAT, None), "radarFilled": (82, CAT, None),
    "bubble": (15, BUBBLE, None), "bubble3D": (87, BUBBLE, None),
    "scatterSmoothMarkers": (72, BUBBLE, None), "scatterLinesNoMarkers": (75, BUBBLE, None),
    "stockHLC": (88, HLC, None), "stockOHLC": (89, OHLC, None), "stockVHLC": (90, VHLC, None), "stockVOHLC": (91, VOHLC, None),
    "surface3D": (83, SURFACE, None), "surfaceWireframe": (84, SURFACE, None), "surfaceContour": (85, SURFACE, None), "surfaceContourWireframe": (86, SURFACE, None),
    "comboColumnLine": (51, CAT, "comboLine"), "comboColumnLineSecondary": (51, CAT, "comboLineSecondary"), "comboAreaColumn": (76, CAT, "comboAreaColumn"),
    "waterfall": (119, WATERFALL, "waterfallTotal"), "funnel": (123, FUNNEL, None), "treemap": (117, HIER, None), "sunburst": (120, HIER, None),
    "boxWhisker": (121, BOX, None), "pareto": (122, PARETO, None), "histogram": (118, HIST, None),
    # Kinds RaabeCalc had before, for the same comparison.
    "column": (51, CAT, None), "columnStacked": (52, CAT, None), "bar": (57, CAT, None), "barStacked": (58, CAT, None), "line": (4, CAT, None),
    "lineMarkers": (65, CAT, None), "area": (1, CAT, None), "pie": (5, PIE, None), "doughnut": (-4120, PIE, None),
}


def r2(v) -> float | None:
    try:
        return round(float(v), 2)
    except Exception:  # noqa: BLE001
        return None


def tryget(fn, default=None):
    try:
        return fn()
    except Exception:  # noqa: BLE001
        return default


def fill(ws, rows: list[list]) -> str:
    for r, row in enumerate(rows):
        for c, v in enumerate(row):
            if v is not None and v != "":
                ws.Cells(r + 1, c + 1).Value2 = v
    if rows[1] and isinstance(rows[1][0], int) and rows[1][0] > 40000:
        ws.Range(ws.Cells(2, 1), ws.Cells(len(rows), 1)).NumberFormatLocal = "TT.MM.JJJJ"
    return ws.Range(ws.Cells(1, 1), ws.Cells(len(rows), max(len(r) for r in rows))).Address


def tweak(ch, what: str | None) -> None:
    if what is None:
        return
    sc = ch.SeriesCollection()
    if what in ("comboLine", "comboLineSecondary"):
        s = sc.Item(sc.Count)
        s.ChartType = 4
        if what == "comboLineSecondary":
            s.AxisGroup = 2
    elif what == "comboAreaColumn":
        sc.Item(sc.Count).ChartType = 51
    elif what == "waterfallTotal":
        s = sc.Item(1)
        s.Points(s.Points().Count).IsTotal = True


def axes(ch) -> list[dict]:
    out = []
    for group in (1, 2):
        for t in (1, 2, 3):
            if not tryget(lambda: ch.HasAxis(t, group), False):
                continue
            ax = tryget(lambda: ch.Axes(t, group))
            if ax is None:
                continue
            rec = {"type": t, "group": group, "min": r2(tryget(lambda: ax.MinimumScale)), "max": r2(tryget(lambda: ax.MaximumScale)),
                   "major": r2(tryget(lambda: ax.MajorUnit)), "hasTitle": tryget(lambda: bool(ax.HasTitle)),
                   "gridlines": tryget(lambda: bool(ax.HasMajorGridlines))}
            out.append(rec)
    return out


def series(ch) -> list[dict]:
    out = []
    sc = tryget(lambda: ch.FullSeriesCollection(), None) or ch.SeriesCollection()
    for i in range(1, sc.Count + 1):
        s = sc.Item(i)
        out.append({"name": tryget(lambda: s.Name), "formula": tryget(lambda: s.Formula), "type": tryget(lambda: s.ChartType),
                    "axisGroup": tryget(lambda: s.AxisGroup), "points": tryget(lambda: s.Points().Count)})
    return out


def group_info(ch) -> dict:
    g = tryget(lambda: ch.ChartGroups(1))
    if g is None:
        return {}
    keys = ["GapWidth", "Overlap", "SplitType", "SplitValue", "SecondPlotSize", "FirstSliceAngle", "DoughnutHoleSize", "BubbleScale",
            "HasHiLoLines", "HasUpDownBars", "HasDropLines", "HasRadarAxisLabels", "VaryByCategories", "ShowNegativeBubbles", "SizeRepresents"]
    out = {}
    for k in keys:
        v = tryget(lambda: getattr(g, k))
        if v is not None:
            out[k] = v if isinstance(v, (bool, str)) else r2(v)
    return out


def view3d(ch) -> dict:
    out = {}
    for k in ["Elevation", "Rotation", "Perspective", "RightAngleAxes", "DepthPercent", "HeightPercent", "AutoScaling"]:
        v = tryget(lambda: getattr(ch, k))
        if v is not None:
            out[k] = v if isinstance(v, bool) else r2(v)
    return out


def record(ch) -> dict:
    pa = ch.PlotArea
    return {
        "chartType": ch.ChartType,
        "series": series(ch),
        "hasTitle": tryget(lambda: bool(ch.HasTitle)),
        "title": tryget(lambda: ch.ChartTitle.Text) if tryget(lambda: ch.HasTitle) else None,
        "hasLegend": tryget(lambda: bool(ch.HasLegend)),
        "legendPos": tryget(lambda: ch.Legend.Position) if tryget(lambda: ch.HasLegend) else None,
        "axes": axes(ch),
        "group": group_info(ch),
        "view3d": view3d(ch),
        "plotInside": {k: r2(tryget(lambda: getattr(pa, "Inside" + k))) for k in ("Left", "Top", "Width", "Height")},
        "chartStyle": tryget(lambda: ch.ChartStyle),
    }


def scrub(path: Path) -> None:
    """The user name (document properties) and the save path (absPath) out of the saved file."""
    with zipfile.ZipFile(path) as z:
        items = [(info, z.read(info)) for info in z.infolist()]
    out = []
    for info, data in items:
        if info.filename == "docProps/core.xml":
            t = data.decode("utf-8")
            t = re.sub(r"<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>", f"<cp:lastModifiedBy>{NEUTRAL}</cp:lastModifiedBy>", t)
            t = re.sub(r"<dc:creator>[^<]*</dc:creator>", f"<dc:creator>{NEUTRAL}</dc:creator>", t)
            data = t.encode("utf-8")
        if info.filename == "xl/workbook.xml":
            t = data.decode("utf-8")
            t = re.sub(r'<mc:AlternateContent xmlns:mc="[^"]*"><mc:Choice Requires="x15"><x15ac:absPath [^>]*/></mc:Choice></mc:AlternateContent>', "", t)
            data = t.encode("utf-8")
        out.append((info, data))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for info, data in out:
            z.writestr(info, data)
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            d = z.read(n)
            for needle in ["C:\\Users", os.environ.get("USERNAME", "") or "\0"]:
                if needle.encode("utf-8") in d or needle.encode("utf-16-le") in d:
                    raise SystemExit(f"{path.name}: private text in {n}")


def main() -> None:
    only = sys.argv[1:]
    PICS.mkdir(exist_ok=True)
    x = ExcelTarget(visible=False)
    data: dict = {}
    try:
        wb = x.wb
        wb.ApplyTheme(THEME)
        first = wb.Worksheets(1)
        data["excel"] = {"version": x.app.Version, "theme": "Office 2013 - 2022", "size": [360, 216], "units": "points"}
        types = {k: v for k, v in TYPES.items() if not only or k in only}
        out: dict = {}
        for n, (tid, (xl, rows, tw)) in enumerate(types.items()):
            ws = first if n == 0 else wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            ws.Name = tid[:31]
            ws.Activate()
            addr = fill(ws, rows)
            ws.Range(addr).Select()
            try:
                shp = ws.Shapes.AddChart2(-1, xl, 230, 15, 360, 216)
            except Exception as e:  # noqa: BLE001
                out[tid] = {"error": str(e)}
                print(tid, "failed", e, flush=True)
                continue
            ch = shp.Chart
            tweak(ch, tw)
            ws.Range("A1").Select()
            rec = {"xlType": xl, **record(ch)}
            ch.Export(str(PICS / f"{tid}.png"))
            ch.Export(str(PICS / f"{tid}.svg"), "SVG")
            out[tid] = rec
            print(tid, rec["chartType"], len(rec["series"]), "series", flush=True)
        first.Activate()
        tmp = Path(tempfile.gettempdir()) / "excel-charttypes.xlsx"
        if tmp.exists():
            tmp.unlink()
        for prop in ("Author", "Last author"):
            tryget(lambda: setattr(wb.BuiltinDocumentProperties.Item(prop), "Value", NEUTRAL))
        wb.SaveAs(str(tmp), 51)
        data["types"] = out
    finally:
        x.close()
    scrub(tmp)
    # The Office user's name out of the file (CLAUDE.md); Excel's UserName option is never changed.
    import subprocess
    subprocess.run([sys.executable, str(HERE / "anonymize_office.py"), str(tmp)], check=True)
    if not only:
        shutil.copy(tmp, FIX / "excel-charttypes.xlsx")
        text = json.dumps(data, indent=1, ensure_ascii=False)
        OUT.write_text(text, encoding="utf-8", newline="\n")
        (FIX / "excel-charttypes.json").write_text(text, encoding="utf-8", newline="\n")
    print("saved", tmp)


if __name__ == "__main__":
    main()
