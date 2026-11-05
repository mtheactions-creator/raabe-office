"""Probe: what Excel's Schnelllayouts (Diagrammentwurf > Schnelllayout, ApplyLayout 1-11) and the colour
schemes of Farben ändern (ChartColor 10-26) do to a chart, per chart type (D-066). Own invisible
Excel, COM only, no input. Result: results/<date>-excel-chartlayouts.json.
"""
from __future__ import annotations

import datetime
import json
from pathlib import Path

from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
TYPES = {"column": 51, "bar": 57, "line": 4, "lineMarkers": 65, "pie": 5, "scatter": -4169, "area": 1}


def state(ch) -> dict:
    def safe(f, d=None):
        try:
            return f()
        except Exception:  # noqa: BLE001
            return d
    pie = ch.ChartType in (5, -4120)
    out = {
        "title": safe(lambda: bool(ch.HasTitle)), "legend": safe(lambda: bool(ch.HasLegend)),
        "legendPos": safe(lambda: ch.Legend.Position) if safe(lambda: ch.HasLegend) else None,
        "labels": safe(lambda: bool(ch.SeriesCollection(1).HasDataLabels)),
        "labelPos": safe(lambda: ch.SeriesCollection(1).DataLabels().Position) if safe(lambda: ch.SeriesCollection(1).HasDataLabels) else None,
        "showCat": safe(lambda: bool(ch.SeriesCollection(1).DataLabels().ShowCategoryName)) if safe(lambda: ch.SeriesCollection(1).HasDataLabels) else None,
        "showPct": safe(lambda: bool(ch.SeriesCollection(1).DataLabels().ShowPercentage)) if safe(lambda: ch.SeriesCollection(1).HasDataLabels) else None,
        "dataTable": safe(lambda: bool(ch.HasDataTable)),
    }
    if not pie:
        out.update({
            "catTitle": safe(lambda: bool(ch.Axes(1).HasTitle)), "valTitle": safe(lambda: bool(ch.Axes(2).HasTitle)),
            "valAxis": safe(lambda: bool(ch.HasAxis(2, 1))), "catAxis": safe(lambda: bool(ch.HasAxis(1, 1))),
            "majorGrid": safe(lambda: bool(ch.Axes(2).HasMajorGridlines)), "minorGrid": safe(lambda: bool(ch.Axes(2).HasMinorGridlines)),
            "catGrid": safe(lambda: bool(ch.Axes(1).HasMajorGridlines)), "gap": safe(lambda: ch.ChartGroups(1).GapWidth),
        })
    return out


def main() -> None:
    x = ExcelTarget(visible=False)
    out: dict = {"layouts": {}, "colors": {}}
    try:
        ws = x.sheet
        data = [("Monat", "A", "B"), ("Jan", 3, 4), ("Feb", 5, 2), ("Mrz", 2, 6), ("Apr", 4, 3)]
        for r, row in enumerate(data, start=1):
            for c, v in enumerate(row, start=1):
                ws.Cells(r, c).Value = v
        for name, t in TYPES.items():
            out["layouts"][name] = {}
            for n in range(1, 12):
                sh = ws.Shapes.AddChart2(-1, t, 300, 20, 360, 216)
                ch = sh.Chart
                ch.SetSourceData(ws.Range("A1:C5") if t not in (5,) else ws.Range("A1:B5"))
                try:
                    ch.ApplyLayout(n)
                    out["layouts"][name][n] = state(ch)
                except Exception as e:  # noqa: BLE001
                    out["layouts"][name][n] = {"error": str(e)[:80]}
                sh.Delete()
            print(name, "done", flush=True)
        # Farben ändern: the series colours of each scheme (RGB of series 1-3).
        sh = ws.Shapes.AddChart2(-1, 51, 300, 20, 360, 216)
        ch = sh.Chart
        ch.SetSourceData(ws.Range("A1:C5"))
        for k in range(10, 31):
            try:
                ch.ChartColor = k
                cols = []
                for i in range(1, 3):
                    v = ch.SeriesCollection(i).Format.Fill.ForeColor.RGB
                    cols.append("#%02X%02X%02X" % (v & 255, (v >> 8) & 255, (v >> 16) & 255))
                out["colors"][k] = cols
            except Exception as e:  # noqa: BLE001
                out["colors"][k] = {"error": str(e)[:60]}
        sh.Delete()
    finally:
        x.close()
    (HERE / "results" / f"{datetime.date.today().isoformat()}-excel-chartlayouts.json").write_text(json.dumps(out, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
