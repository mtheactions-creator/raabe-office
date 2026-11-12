"""Probe (Step 12): Excel's automatic axis values and default gridlines. In an invisible Excel of our own (COM, no
input) it makes charts of several data sets and types at several sizes and reads, for each axis, the values Excel uses
while Grenzen and Einheiten are automatic (MinimumScale, MaximumScale, MajorUnit, MinorUnit: Excel reports the
computed ones) and which gridlines a new chart of that type shows (HasMajorGridlines / HasMinorGridlines of the value
and the category axis). Writes results/<date>-excel-axis-auto.json.

    .venv/Scripts/python.exe probe_excel_axis_auto.py
"""
from __future__ import annotations

import datetime
import json
from pathlib import Path

from feellab.excel import ExcelTarget

HERE = Path(__file__).parent

# Excel's chart type constants and the names our model uses.
TYPES = {"column": 51, "bar": 57, "line": 4, "lineMarkers": 65, "area": 1, "scatter": -4169, "scatterLines": 74,
         "columnStacked": 52, "columnStacked100": 53}
DATA = {
    "small": [1, 2, 3, 4],
    "school": [12, 47, 120, 83],
    "negative": [-5, 3, 8, -2],
    "allneg": [-40, -12, -75, -3],
    "decimals": [0.002, 0.009, 0.005, 0.0071],
    "large": [1000, 5000, 23000, 17250],
    "narrow": [101, 102, 104, 103],
    "zero": [0, 0, 0, 0],
    "grades": [1.7, 2.3, 3.0, 2.0],
    "percent": [0.12, 0.47, 0.3, 0.9],
}
X = [1, 2.5, 4, 7]
SIZES = [(360, 216), (480, 300), (240, 160)]


def axis(ch, which: int) -> dict:
    a = ch.Axes(which)
    out = {}
    for k in ("MinimumScale", "MaximumScale", "MajorUnit", "MinorUnit"):
        try:
            out[k] = round(float(getattr(a, k)), 10)
        except Exception:  # noqa: BLE001
            out[k] = None
    for k in ("HasMajorGridlines", "HasMinorGridlines"):
        try:
            out[k] = bool(getattr(a, k))
        except Exception:  # noqa: BLE001
            out[k] = None
    return out


def main() -> None:
    x = ExcelTarget(visible=False, place=False)
    out: dict = {"date": datetime.date.today().isoformat(), "x": X, "data": DATA, "cases": []}
    try:
        ws = x.sheet
        for name, vals in DATA.items():
            ws.Cells.Clear()
            ws.Range("A1").Value2 = "x"
            ws.Range("B1").Value2 = "Werte"
            for i, (xv, v) in enumerate(zip(X, vals), start=2):
                ws.Cells(i, 1).Value2 = xv
                ws.Cells(i, 2).Value2 = v
            for kind, xl in TYPES.items():
                for w, h in SIZES:
                    rng = ws.Range("A1:B5") if kind.startswith("scatter") else ws.Range("B1:B5")
                    shp = ws.Shapes.AddChart2(-1, xl, 200, 10, w, h)
                    ch = shp.Chart
                    ch.SetSourceData(rng)
                    if kind.startswith("scatter"):
                        # X from column A, Y from column B, one series.
                        while ch.SeriesCollection().Count > 1:
                            ch.SeriesCollection(ch.SeriesCollection().Count).Delete()
                        s = ch.SeriesCollection(1)
                        s.XValues = ws.Range("A2:A5")
                        s.Values = ws.Range("B2:B5")
                    case = {"data": name, "kind": kind, "size": [w, h], "val": axis(ch, 2)}
                    try:
                        case["cat"] = axis(ch, 1)
                    except Exception:  # noqa: BLE001
                        pass
                    case["legend"] = {-4107: "b", -4152: "r", -4160: "t", -4131: "l", 2: "tr"}.get(ch.Legend.Position, "?") if ch.HasLegend else "none"
                    case["title"] = bool(ch.HasTitle)
                    pa = ch.PlotArea
                    case["plot"] = [round(pa.InsideLeft, 2), round(pa.InsideTop, 2), round(pa.InsideWidth, 2), round(pa.InsideHeight, 2)]
                    out["cases"].append(case)
                    shp.Delete()
    finally:
        x.close()
    path = HERE / "results" / f"{out['date']}-excel-axis-auto.json"
    path.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(len(out["cases"]), "cases; wrote", path.name)


if __name__ == "__main__":
    main()
