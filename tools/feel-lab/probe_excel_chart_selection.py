"""How Excel marks a selected chart part (Step 11 part 2, D-170): in an Excel of our own, a column, a line and a pie
chart; every selectable part is selected through COM (ChartObject.Activate, then e.g. Chart.PlotArea.Select()) and the
chart's region of the screen is captured (window verified in front, under the real-input lock). The crops stay in
TEMP (the title bar is never captured); the marks measured from them (selmarks.py: frame, handles, colours) go to
results/2026-11-11-excel-chart-selection.json.

    .venv/Scripts/python.exe probe_excel_chart_selection.py [--no-capture]   (--no-capture: re-measure the crops)
"""
from __future__ import annotations

import json
import sys
import tempfile
import time
from pathlib import Path

from feellab import selmarks, winapi
from feellab.inputlock import real_input

HERE = Path(__file__).parent
RESULT = HERE / "results" / "2026-11-11-excel-chart-selection.json"
CROPS = Path(tempfile.gettempdir()) / "raabe-chart-sel" / "excel"
DATA = [["Monat", "Nord", "Süd", "West"], ["Jan", 3, 4, 2], ["Feb", 5, 2, 3], ["Mrz", 2, 6, 4], ["Apr", 4, 3, 5]]
# Chart position and size in points (480 x 288 px at 100 %, Excel's new chart).
LEFT, TOP, W, H = 300.0, 30.0, 360.0, 216.0
MARGIN = 16  # device pixels around the chart's box in the crop
KINDS = {"column": 51, "line": 65, "pie": 5}


def parts(kind: str) -> list[tuple[str, str]]:
    """(name, COM expression on the chart `c`) for every part to select."""
    p = [("chart", "c.ChartArea.Select()"), ("plot", "c.PlotArea.Select()"), ("title", "c.ChartTitle.Select()"),
         ("legend", "c.Legend.Select()"), ("legendEntry", "c.Legend.LegendEntries(1).Select()")]
    if kind != "pie":
        p += [("axisTitle", "c.Axes(2).AxisTitle.Select()"), ("valAxis", "c.Axes(2).Select()"), ("catAxis", "c.Axes(1).Select()"),
              ("majorGrid", "c.Axes(2).MajorGridlines.Select()")]
    p += [("series", "c.SeriesCollection(1).Select()"), ("point", "c.SeriesCollection(1).Points(2).Select()"),
          ("labels", "c.SeriesCollection(1).DataLabels().Select()"), ("label", "c.SeriesCollection(1).Points(2).DataLabel.Select()")]
    if kind != "pie":
        p += [("trend", "c.SeriesCollection(1).Trendlines(1).Select()")]
    return p


def capture() -> dict:
    from feellab.excel import ExcelTarget
    CROPS.mkdir(parents=True, exist_ok=True)
    x = ExcelTarget(visible=True)
    out: dict = {"crops": {}}
    try:
        app = x.app
        ws = x.sheet
        for r, row in enumerate(DATA, 1):
            for c, v in enumerate(row, 1):
                ws.Cells(r, c).Value = v
        app.ActiveWindow.Zoom = 100
        app.ActiveWindow.ScrollRow = 1
        app.ActiveWindow.ScrollColumn = 1
        dpi = winapi.dpi_scale(x.hwnd)
        out["scale"] = dpi
        for kind, code in KINDS.items():
            co = ws.ChartObjects().Add(LEFT, TOP, W, H)
            c = co.Chart
            c.SetSourceData(ws.Range("A1:D5"))
            c.ChartType = code
            c.HasTitle = True
            c.ChartTitle.Text = "Diagrammtitel"
            c.HasLegend = True
            if kind != "pie":
                c.Axes(2).HasTitle = True
                c.Axes(2).AxisTitle.Text = "Achsentitel"
                c.SeriesCollection(1).Trendlines().Add()
            c.SeriesCollection(1).HasDataLabels = True
            # The chart's box on the screen: the grid's origin plus the points at Excel's 96 dpi * scale.
            ox = app.ActiveWindow.PointsToScreenPixelsX(0)
            oy = app.ActiveWindow.PointsToScreenPixelsY(0)
            px = lambda pt: round(pt / 72 * 96 * dpi)  # noqa: E731
            box = (ox + px(LEFT), oy + px(TOP), px(W), px(H))
            region = (box[0] - MARGIN, box[1] - MARGIN, box[2] + 2 * MARGIN, box[3] + 2 * MARGIN)
            out.setdefault("boxes", {})[kind] = {"chart": [MARGIN, MARGIN, box[2], box[3]], "region": list(region)}
            states = [("none", "ws.Range('F20').Select()")] + parts(kind)
            for name, expr in states:
                try:
                    if name != "none":
                        co.Activate()
                    eval(expr, {"c": c, "ws": ws})  # noqa: S307 - our own fixed expressions
                except Exception as e:  # noqa: BLE001
                    out["crops"][f"{kind}-{name}"] = {"error": str(e)[:200]}
                    continue
                time.sleep(0.6)
                path = CROPS / f"{kind}-{name}.png"
                winapi.activate(x.hwnd)
                time.sleep(0.3)
                fg = winapi.foreground()
                if winapi.window_pid(fg) != x.pid:
                    out["crops"][f"{kind}-{name}"] = {"error": "Excel not in front, not captured"}
                    continue
                winapi.move(region[0] + region[2] + 300, region[1] + region[3] + 300)
                time.sleep(0.2)
                winapi.save_region_png(str(path), *region)
                out["crops"][f"{kind}-{name}"] = {"file": path.name}
            ws.Range("F20").Select()
            co.Delete()
    finally:
        x.close()
    (CROPS / "capture.json").write_text(json.dumps(out, indent=1), encoding="utf-8")
    return out


def main() -> None:
    if "--no-capture" in sys.argv:
        cap = json.loads((CROPS / "capture.json").read_text(encoding="utf-8"))
    else:
        with real_input("chart selection marks (Excel)"):
            cap = capture()
    result = {"what": "Excel 365's marks of a selected chart part, measured in device pixels of the screen's scale "
                      "(selmarks.py: pixels that differ from the chart with nothing selected, or from the chart area "
                      "selected for inner parts)",
              "scale": cap["scale"], "parts": {}}
    for key, info in cap["crops"].items():
        if "error" in info or key.endswith("-none"):
            if "error" in info:
                result["parts"][key] = info
            continue
        kind, name = key.split("-", 1)
        base = CROPS / f"{kind}-{'none' if name == 'chart' else 'chart'}.png"
        result["parts"][key] = selmarks.measure(CROPS / info["file"], base, cap["boxes"][kind]["chart"])
    RESULT.write_text(json.dumps(result, indent=1, ensure_ascii=False) + "\n", encoding="utf-8", newline="\n")
    print(json.dumps({k: (v.get("summary") if isinstance(v, dict) else v) for k, v in result["parts"].items()}, indent=1, ensure_ascii=False))


if __name__ == "__main__":
    main()
