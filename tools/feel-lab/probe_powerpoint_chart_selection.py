"""How PowerPoint marks a selected chart and chart part (Step 11 part 2, D-171), against Excel's and Word's
(probe_excel_chart_selection.py, probe_word_chart_selection.py). Follows the PowerPoint rules (feellab/powerpoint.py):
only when no POWERPNT.EXE runs; our own presentation in TEMP; real input under the lock, window checked in front,
only the chart's region captured (crops in TEMP). A column chart (AddChart2, PowerPoint's sample data); the chart as a
shape (a click on its edge: Shape.Select), the chart area, plot area, title, legend, value axis, gridlines, a series
and a point.

    .venv/Scripts/python.exe probe_powerpoint_chart_selection.py [--no-capture]

Result: results/2026-11-11-powerpoint-chart-selection.json.
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
RESULT = HERE / "results" / "2026-11-11-powerpoint-chart-selection.json"
CROPS = Path(tempfile.gettempdir()) / "raabe-chart-sel" / "powerpoint"
MARGIN = 16
PARTS = [("shape", "sh.Select()"), ("chart", "c.ChartArea.Select()"), ("plot", "c.PlotArea.Select()"), ("title", "c.ChartTitle.Select()"),
         ("legend", "c.Legend.Select()"), ("valAxis", "c.Axes(2).Select()"), ("majorGrid", "c.Axes(2).MajorGridlines.Select()"),
         ("series", "c.SeriesCollection(1).Select()"), ("point", "c.SeriesCollection(1).Points(2).Select()")]


def capture() -> dict:
    from feellab.powerpoint import PowerPointTarget
    CROPS.mkdir(parents=True, exist_ok=True)
    t = PowerPointTarget(visible=True)
    out: dict = {"crops": {}}
    try:
        pres = t.new_presentation(with_window=True, name="chartsel")
        slide = pres.Slides.Add(1, 12)  # ppLayoutBlank
        sh = slide.Shapes.AddChart2(-1, 51, 120, 80, 480, 288)
        c = sh.Chart
        c.HasTitle = True
        c.HasLegend = True
        try:
            c.ChartData.Activate()
            c.ChartData.Workbook.Close()  # the chart's own data sheet that AddChart2 opened
        except Exception as e:  # noqa: BLE001
            out["dataSheet"] = str(e)[:200]
        time.sleep(1.0)
        win = t.app.ActiveWindow
        win.ViewType = 9  # ppViewNormal
        win.View.Zoom = 100
        out["scale"] = winapi.dpi_scale(t.hwnd)
        for name, expr in [("none", "win.Selection.Unselect()")] + PARTS:
            try:
                eval(expr, {"c": c, "sh": sh, "win": win})  # noqa: S307 - our own fixed expressions
            except Exception as e:  # noqa: BLE001
                out["crops"][name] = {"error": str(e)[:200]}
                continue
            time.sleep(0.6)
            x = win.PointsToScreenPixelsX(sh.Left)
            y = win.PointsToScreenPixelsY(sh.Top)
            x2 = win.PointsToScreenPixelsX(sh.Left + sh.Width)
            y2 = win.PointsToScreenPixelsY(sh.Top + sh.Height)
            region = (x - MARGIN, y - MARGIN, x2 - x + 2 * MARGIN, y2 - y + 2 * MARGIN)
            out.setdefault("box", [MARGIN, MARGIN, x2 - x, y2 - y])
            for _ in range(6):
                try:
                    winapi.activate(t.hwnd)
                except RuntimeError:
                    pass
                time.sleep(0.4)
                if winapi.window_pid(winapi.foreground()) == t.pid:
                    break
            if winapi.window_pid(winapi.foreground()) != t.pid:
                out["crops"][name] = {"error": "PowerPoint not in front, not captured"}
                continue
            winapi.move(region[0] + region[2] + 300, region[1] + region[3] + 150)
            time.sleep(0.2)
            winapi.save_region_png(str(CROPS / f"column-{name}.png"), *region)
            out["crops"][name] = {"file": f"column-{name}.png"}
    finally:
        import subprocess
        # The Excel that our PowerPoint started for the chart's data sheet: ended with it (only that one, by its parent).
        q = f"Get-CimInstance Win32_Process -Filter \"name='EXCEL.EXE' and ParentProcessId={t.pid}\" | ForEach-Object {{ $_.ProcessId }}"
        kids = subprocess.run(["powershell", "-NoProfile", "-Command", q], capture_output=True, text=True).stdout.split() if t.pid else []
        out["close"] = t.close(kill_if_unreachable=True)
        time.sleep(3)
        for k in kids:
            subprocess.run(["taskkill", "/PID", k, "/F"], capture_output=True)
        out["dataExcelEnded"] = len(kids)
    (CROPS / "capture.json").write_text(json.dumps(out, indent=1, default=str), encoding="utf-8")
    return out


def main() -> None:
    if "--no-capture" in sys.argv:
        cap = json.loads((CROPS / "capture.json").read_text(encoding="utf-8"))
    else:
        from feellab.powerpoint import owner_powerpoint_running
        if owner_powerpoint_running():
            print("POWERPNT.EXE is running: not ours, no PowerPoint probe now")
            sys.exit(2)
        with real_input("chart selection marks (PowerPoint)"):
            cap = capture()
    excel = json.loads((HERE / "results" / "2026-11-11-excel-chart-selection.json").read_text(encoding="utf-8"))
    result = {"what": "PowerPoint 365's marks of a selected chart and chart part (selmarks.py), and how they compare "
                      "with Excel's", "scale": cap.get("scale"), "parts": {}}
    for name, info in cap["crops"].items():
        if name == "none":
            continue
        if "error" in info:
            result["parts"][name] = info
            continue
        base = CROPS / ("column-none.png" if name in ("shape", "chart") else "column-chart.png")
        m = selmarks.measure(CROPS / info["file"], base, cap["box"])
        ref = excel["parts"].get(f"column-{'chart' if name == 'shape' else name}")
        if ref:
            m["vsExcel"] = selmarks.compare(m, ref)
        result["parts"][name] = m
    RESULT.write_text(json.dumps(result, indent=1, ensure_ascii=False) + "\n", encoding="utf-8", newline="\n")
    for k, v in result["parts"].items():
        print(k, v.get("summary", v), "| vs Excel:", v.get("vsExcel", {}).get("fails"))


if __name__ == "__main__":
    main()
