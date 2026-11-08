"""Records where Excel puts the parts of its 3-D charts (D-125): our own invisible Excel, COM only, opens the chart
fixture (packages/ooxml/test/fixtures/excel-charttypes.xlsx, D-086) and reads, for every 3-D chart, the chart's size,
the 3-D view (Elevation, Rotation, Perspective, RightAngleAxes, DepthPercent, GapDepth, HeightPercent, AutoScaling),
the plot area (outer and inside box), each axis' box, the walls' and floor's format and every point's box
(Point.Left/Top/Width/Height: the projected outline). Written to results/2026-11-09-excel-3d-geometry.json.

    .venv/Scripts/python.exe probe_excel_3d_geometry.py
"""
from __future__ import annotations

import json
from pathlib import Path

from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
ROOT = HERE.parents[1]
FIXTURE = ROOT / "packages" / "ooxml" / "test" / "fixtures" / "excel-charttypes.xlsx"
OUT = HERE / "results" / "2026-11-09-excel-3d-geometry.json"
NAMES = ["column3D", "column3DClustered", "column3DStacked", "column3DStacked100", "bar3DClustered", "bar3DStacked", "bar3DStacked100",
         "line3D", "area3D", "area3DStacked", "area3DStacked100", "surface3D", "surfaceWireframe", "pie3D"]


def box(o) -> dict | None:
    try:
        return {k: round(getattr(o, k), 2) for k in ("Left", "Top", "Width", "Height")}
    except Exception:  # noqa: BLE001
        return None


def get(fn):
    try:
        return fn()
    except Exception:  # noqa: BLE001
        return None


def main() -> None:
    x = ExcelTarget(visible=False)
    out: dict = {}
    try:
        wb = x.app.Workbooks.Open(str(FIXTURE), ReadOnly=True)
        try:
            for name in NAMES:
                try:
                    ws = wb.Worksheets(name)
                except Exception:  # noqa: BLE001
                    continue
                if ws.ChartObjects().Count == 0:
                    continue
                co = ws.ChartObjects(1)
                ch = co.Chart
                rec = {"chartArea": {"w": co.Width, "h": co.Height}, "type": ch.ChartType,
                       "view": {k: get(lambda k=k: getattr(ch, k)) for k in ("Elevation", "Rotation", "Perspective", "RightAngleAxes", "DepthPercent", "GapDepth", "HeightPercent", "AutoScaling")},
                       "plotArea": box(ch.PlotArea), "plotInside": {k: round(getattr(ch.PlotArea, "Inside" + k), 2) for k in ("Left", "Top", "Width", "Height")},
                       "title": box(ch.ChartTitle) if ch.HasTitle else None, "legend": box(ch.Legend) if ch.HasLegend else None, "axes": {}, "points": []}
                for t, label in ((1, "cat"), (2, "val"), (3, "ser")):
                    a = get(lambda t=t: ch.Axes(t))
                    if a is not None:
                        rec["axes"][label] = {**(box(a) or {}), "min": get(lambda: a.MinimumScale), "max": get(lambda: a.MaximumScale), "major": get(lambda: a.MajorUnit)}
                rec["walls"] = {"fill": get(lambda: ch.Walls.Format.Fill.Visible), "line": get(lambda: ch.Walls.Format.Line.Visible)}
                rec["floor"] = {"fill": get(lambda: ch.Floor.Format.Fill.Visible), "line": get(lambda: ch.Floor.Format.Line.Visible)}
                for si in range(1, ch.SeriesCollection().Count + 1):
                    s = ch.SeriesCollection(si)
                    pts = []
                    for pi in range(1, (get(lambda: s.Points().Count) or 0) + 1):
                        pts.append(box(s.Points(pi)))
                    rec["points"].append(pts)
                out[name] = rec
                print(name, rec["view"], rec["plotInside"], flush=True)
        finally:
            wb.Close(False)
    finally:
        x.close()
    OUT.write_text(json.dumps(out, indent=1), encoding="utf-8", newline="\n")
    print("written", OUT)


if __name__ == "__main__":
    main()
