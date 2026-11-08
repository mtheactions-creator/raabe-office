"""Records Excel's sparklines (D-085) as the reference for packages/ooxml/test/xlsx-sparklines.test.ts:
for each option of the Sparkline tab (type, Anzeigen, colours, Stärke, Achse, empty and hidden cells,
date axis, Gruppieren, Gruppierung aufheben, Löschen) the x14:sparklineGroups XML Excel saves and what
COM reads back.

Starts an invisible Excel of its own (`ExcelTarget(visible=False)`), one sheet per case, saves one
workbook to TEMP and reads its sheet parts. No real input.

    .venv/Scripts/python.exe probe_excel_sparklines.py

Writes packages/ooxml/test/fixtures/excel-sparklines.json."""

from __future__ import annotations

import json
import re
import tempfile
import zipfile
from pathlib import Path

from feellab.excel import ExcelTarget

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "packages" / "ooxml" / "test" / "fixtures" / "excel-sparklines.json"
TMP = Path(tempfile.gettempdir())

# Three rows of data (a negative value, an empty cell, a tie for the highest value); sparklines in F1:F3.
DATA = [[3, 5, -2, 8, 4], [1, None, 4, -3, 6], [2, 7, 7, 1, 5]]
DATES = ["2026-01-01", "2026-01-02", "2026-01-05", "2026-01-06", "2026-01-12"]


def add(ws, typ=1, loc="F1:F3", src="A1:E3"):
    return ws.Range(loc).SparklineGroups.Add(typ, src)


def color(fc) -> dict:
    out = {"rgb": "#%06X" % _bgr(fc.Color)}
    try:
        out["theme"] = fc.ThemeColor
        out["tint"] = round(fc.TintAndShade, 6)
    except Exception:  # noqa: BLE001 (an RGB colour has no theme colour)
        pass
    return out


def _bgr(v: int) -> int:
    v = int(v)
    return ((v & 0xFF) << 16) | (v & 0xFF00) | ((v >> 16) & 0xFF)


def readback(g) -> dict:
    p = g.Points
    ax = g.Axes
    out = {"type": g.Type, "count": g.Count, "location": g.Location.Address.replace("$", ""), "source": g.SourceData,
           "lineWeight": g.LineWeight, "blanks": g.DisplayBlanksAs, "hidden": g.DisplayHidden, "series": color(g.SeriesColor)}
    for k in ("Highpoint", "Lowpoint", "Firstpoint", "Lastpoint", "Negative", "Markers"):
        sp = getattr(p, k)
        out[k] = {"visible": sp.Visible, "color": color(sp.Color)}
    out["axis"] = {"visible": ax.Horizontal.Axis.Visible, "color": color(ax.Horizontal.Axis.Color),
                   "date": ax.Horizontal.IsDateAxis, "rtl": ax.Horizontal.RightToLeftPlotOrder,
                   "minType": ax.Vertical.MinScaleType, "maxType": ax.Vertical.MaxScaleType}
    try:
        out["axis"]["min"] = ax.Vertical.CustomMinScaleValue
        out["axis"]["max"] = ax.Vertical.CustomMaxScaleValue
    except Exception:  # noqa: BLE001
        pass
    try:
        out["dateRange"] = g.DateRange
    except Exception:  # noqa: BLE001
        pass
    return out


def theme(fc, t: int, tint: float = 0.0) -> None:
    fc.ThemeColor = t
    if tint:
        fc.TintAndShade = tint


def cases():
    """(name, setup(ws)) for each case: setup adds the sparklines and changes one option."""
    def g1(ws):
        return ws.Range("F1").SparklineGroups(1)

    def two_groups(ws):
        add(ws, 1, "F1:F2", "A1:E2")
        b = add(ws, 2, "F3", "A3:E3")
        b.SeriesColor.Color = 0x0000FF  # BGR: red
        b.Points.Highpoint.Visible = True
        return b

    out = [
        ("line", lambda ws: add(ws, 1)),
        ("column", lambda ws: add(ws, 2)),
        ("winloss", lambda ws: add(ws, 3)),
        ("single", lambda ws: add(ws, 1, "F1", "A1:E1")),
        ("bycolumn", lambda ws: add(ws, 1, "A5:E5", "A1:E3")),
    ]
    for k in ("Highpoint", "Lowpoint", "Firstpoint", "Lastpoint", "Negative", "Markers"):
        out.append((f"show-{k}", lambda ws, k=k: setattr(getattr(add(ws).Points, k), "Visible", True)))
    def show_all(ws):
        g = add(ws)
        for k in ("Highpoint", "Lowpoint", "Firstpoint", "Lastpoint", "Negative", "Markers"):
            getattr(g.Points, k).Visible = True

    def everything(ws):
        g = add(ws)
        show_all_on(g)
        g.LineWeight = 2.25
        g.DisplayBlanksAs = 3
        g.DisplayHidden = True
        g.Axes.Horizontal.Axis.Visible = True
        g.Axes.Horizontal.RightToLeftPlotOrder = True
        g.Axes.Vertical.MinScaleType = 3
        g.Axes.Vertical.CustomMinScaleValue = -1.5
        g.Axes.Vertical.MaxScaleType = 1
        g.ModifyDateRange("A7:E7")
        theme(g.SeriesColor, 5, -0.499984740745262)

    def show_all_on(g):
        for k in ("Highpoint", "Lowpoint", "Firstpoint", "Lastpoint", "Negative", "Markers"):
            getattr(g.Points, k).Visible = True

    out.append(("show-all", show_all))
    out.append(("everything", everything))
    out.append(("column-negative", lambda ws: setattr(add(ws, 2).Points.Negative, "Visible", True)))
    out.append(("series-rgb", lambda ws: setattr(add(ws).SeriesColor, "Color", 0x00C0FF)))
    out.append(("series-theme", lambda ws: theme(add(ws).SeriesColor, 6)))
    out.append(("series-theme-tint", lambda ws: theme(add(ws).SeriesColor, 9, 0.399975585192419)))
    out.append(("high-rgb", lambda ws: (lambda g: (setattr(g.Points.Highpoint, "Visible", True), setattr(g.Points.Highpoint.Color, "Color", 0x00FF00)))(add(ws))))
    out.append(("markers-theme", lambda ws: (lambda g: (setattr(g.Points.Markers, "Visible", True), theme(g.Points.Markers.Color, 5, -0.249977111117893)))(add(ws))))
    for w in (0.25, 0.5, 1, 1.5, 2.25, 3, 4.5, 6):
        out.append((f"weight-{w}", lambda ws, w=w: setattr(add(ws), "LineWeight", w)))
    out.append(("blanks-zero", lambda ws: setattr(add(ws), "DisplayBlanksAs", 2)))
    out.append(("blanks-span", lambda ws: setattr(add(ws), "DisplayBlanksAs", 3)))
    out.append(("hidden", lambda ws: setattr(add(ws), "DisplayHidden", True)))
    out.append(("axis", lambda ws: setattr(add(ws).Axes.Horizontal.Axis, "Visible", True)))
    out.append(("axis-color", lambda ws: (lambda g: (setattr(g.Axes.Horizontal.Axis, "Visible", True), setattr(g.Axes.Horizontal.Axis.Color, "Color", 0xFF0000)))(add(ws))))
    out.append(("rtl", lambda ws: setattr(add(ws).Axes.Horizontal, "RightToLeftPlotOrder", True)))
    out.append(("min-group", lambda ws: setattr(add(ws).Axes.Vertical, "MinScaleType", 1)))
    out.append(("max-group", lambda ws: setattr(add(ws).Axes.Vertical, "MaxScaleType", 1)))
    out.append(("min-custom", lambda ws: (lambda g: (setattr(g.Axes.Vertical, "MinScaleType", 3), setattr(g.Axes.Vertical, "CustomMinScaleValue", -5)))(add(ws))))
    out.append(("max-custom", lambda ws: (lambda g: (setattr(g.Axes.Vertical, "MaxScaleType", 3), setattr(g.Axes.Vertical, "CustomMaxScaleValue", 10.5)))(add(ws))))
    out.append(("min-max-custom", lambda ws: (lambda g: (setattr(g.Axes.Vertical, "MinScaleType", 3), setattr(g.Axes.Vertical, "CustomMinScaleValue", 0),
                                                       setattr(g.Axes.Vertical, "MaxScaleType", 3), setattr(g.Axes.Vertical, "CustomMaxScaleValue", 10)))(add(ws))))
    out.append(("date-axis", lambda ws: add(ws).ModifyDateRange("A7:E7")))
    out.append(("type-change", lambda ws: setattr(add(ws), "Type", 2)))
    out.append(("modify-group", lambda ws: add(ws).Modify(ws.Range("G1:G3"), "B1:E3")))
    out.append(("modify-one", lambda ws: add(ws).Item(2).ModifySourceData("A2:C2")))
    out.append(("two-groups", two_groups))
    out.append(("group", lambda ws: (two_groups(ws), ws.Range("F1:F3").SparklineGroups.Group(ws.Range("F3")))))
    out.append(("group-first", lambda ws: (two_groups(ws), ws.Range("F1:F3").SparklineGroups.Group(ws.Range("F1")))))
    out.append(("ungroup", lambda ws: (add(ws).Points.Highpoint.__setattr__("Visible", True), ws.Range("F1:F3").SparklineGroups.Ungroup())))
    out.append(("ungroup-one", lambda ws: (add(ws), ws.Range("F2").SparklineGroups.Ungroup())))
    out.append(("clear-one", lambda ws: (add(ws), ws.Range("F2").SparklineGroups.Clear())))
    out.append(("clear-group", lambda ws: (two_groups(ws), ws.Range("F2").SparklineGroups.ClearGroups())))
    out.append(("other-sheet", lambda ws: ws.Range("F1").SparklineGroups.Add(1, "Daten!A1:E1")))
    return out


def ext_of(sheet_xml: str) -> str | None:
    m = re.search(r"<x14:sparklineGroups\b[\s\S]*?</x14:sparklineGroups>", sheet_xml)
    return m.group(0) if m else None


def main() -> None:
    x = ExcelTarget(visible=False)
    path = TMP / "raabe-sparklines-excel.xlsx"
    result: dict = {"data": DATA, "dates": DATES, "cases": {}}
    try:
        wb = x.wb
        data = wb.Worksheets(1)
        data.Name = "Daten"
        for r, row in enumerate(DATA, 1):
            for c, v in enumerate(row, 1):
                if v is not None:
                    data.Cells(r, c).Value2 = v
        names = []
        for name, setup in cases():
            ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            ws.Name = name[:31]
            for r, row in enumerate(DATA, 1):
                for c, v in enumerate(row, 1):
                    if v is not None:
                        ws.Cells(r, c).Value2 = v
            for c, d in enumerate(DATES, 1):
                ws.Cells(7, c).Formula = f"=DATEVALUE(\"{d}\")"
            try:
                setup(ws)
                groups = ws.Cells.SparklineGroups
                result["cases"][name] = {"groups": [readback(groups(i)) for i in range(1, groups.Count + 1)]}
            except Exception as e:  # noqa: BLE001
                result["cases"][name] = {"error": str(e)}
            names.append(name)
            print(name, "ok" if "error" not in result["cases"][name] else result["cases"][name]["error"], flush=True)
        wb.SaveAs(str(path), 51)
    finally:
        x.close()
    with zipfile.ZipFile(path) as z:
        wbx = z.read("xl/workbook.xml").decode("utf-8")
        rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        targets = dict(re.findall(r'<Relationship [^>]*Id="(\w+)"[^>]*Target="([^"]+)"', rels))
        targets.update({k: v for v, k in re.findall(r'<Relationship [^>]*Target="([^"]+)"[^>]*Id="(\w+)"', rels)})
        for sname, rid in re.findall(r'<sheet name="([^"]+)" sheetId="\d+" r:id="(\w+)"', wbx):
            if sname not in result["cases"]:
                continue
            t = targets[rid].lstrip("/")
            xml = z.read(t if t.startswith("xl/") else "xl/" + t).decode("utf-8")
            result["cases"][sname]["xml"] = ext_of(xml)
        result["theme"] = {k: v for k, v in re.findall(r"<a:(dk1|lt1|dk2|lt2|accent\d|hlink|folHlink)>\s*<a:(?:sysClr [^>]*lastClr|srgbClr val)=\"(\w+)\"",
                                                         z.read("xl/theme/theme1.xml").decode("utf-8"))}
    OUT.write_text(json.dumps(result, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print("written", OUT)


if __name__ == "__main__":
    main()
