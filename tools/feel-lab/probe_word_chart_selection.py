"""How Word marks a selected chart part (Step 11 part 2, D-170), against Excel's (probe_excel_chart_selection.py): in a
Word of our own, an inline column chart (AddChart2, Word's own sample data); chart area, plot area, title, legend,
value axis, gridlines, series and a point are selected through COM and the chart's region captured (window verified in
front, under the real-input lock; crops in TEMP, never the title bar). Measured with selmarks.py.

    .venv/Scripts/python.exe probe_word_chart_selection.py [--no-capture]

Result: results/2026-11-11-word-chart-selection.json.
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
RESULT = HERE / "results" / "2026-11-11-word-chart-selection.json"
CROPS = Path(tempfile.gettempdir()) / "raabe-chart-sel" / "word"
MARGIN = 16
PARTS = [("chart", "c.ChartArea.Select()"), ("plot", "c.PlotArea.Select()"), ("title", "c.ChartTitle.Select()"),
         ("legend", "c.Legend.Select()"), ("valAxis", "c.Axes(2).Select()"), ("majorGrid", "c.Axes(2).MajorGridlines.Select()"),
         ("series", "c.SeriesCollection(1).Select()"), ("point", "c.SeriesCollection(1).Points(2).Select()")]


def capture() -> dict:
    from feellab.targets import WordTarget
    CROPS.mkdir(parents=True, exist_ok=True)
    w = WordTarget("short", visible=True)
    out: dict = {"crops": {}}
    try:
        doc = w.doc
        pid = winapi.window_pid(w.hwnd)
        doc.Content.Text = ""
        ish = doc.InlineShapes.AddChart2(-1, 51, doc.Range(0, 0))
        c = ish.Chart
        c.HasTitle = True
        c.HasLegend = True
        # AddChart2 opens the chart's own data sheet in an Excel window: close it, so Word comes to the front.
        try:
            c.ChartData.Activate()
            wb = c.ChartData.Workbook
            wb.Application.WindowState = -4140  # minimized
            wb.Close()
        except Exception as e:  # noqa: BLE001
            out["dataSheet"] = str(e)[:200]
        time.sleep(1.0)
        doc.Range(0, 0).Select()
        time.sleep(0.8)
        win = doc.Windows(1)
        out["scale"] = winapi.dpi_scale(w.hwnd)
        for name, expr in [("none", "doc.Range(doc.Content.End - 1, doc.Content.End - 1).Select()")] + PARTS:
            try:
                eval(expr, {"c": c, "doc": doc})  # noqa: S307 - our own fixed expressions
            except Exception as e:  # noqa: BLE001
                out["crops"][name] = {"error": str(e)[:200]}
                continue
            time.sleep(0.6)
            # The chart's box on the screen (GetPoint: screen pixels).
            x, y, cw, ch = win.GetPoint(0, 0, 0, 0, ish.Range)
            region = (x - MARGIN, y - MARGIN, cw + 2 * MARGIN, ch + 2 * MARGIN)
            out.setdefault("box", [MARGIN, MARGIN, cw, ch])
            for _ in range(6):
                try:
                    winapi.activate(w.hwnd)
                except RuntimeError:
                    pass
                time.sleep(0.4)
                if winapi.window_pid(winapi.foreground()) == pid:
                    break
            if winapi.window_pid(winapi.foreground()) != pid:
                out["crops"][name] = {"error": "Word not in front, not captured"}
                continue
            winapi.move(region[0] + region[2] + 300, region[1] + region[3] + 200)
            time.sleep(0.2)
            winapi.save_region_png(str(CROPS / f"column-{name}.png"), *region)
            out["crops"][name] = {"file": f"column-{name}.png"}
    finally:
        import subprocess
        word_pid = winapi.window_pid(w.hwnd)
        # The Excel that Word started for the chart's data sheet: ended with Word (only that one, by its parent).
        q = f"Get-CimInstance Win32_Process -Filter \"name='EXCEL.EXE' and ParentProcessId={word_pid}\" | ForEach-Object {{ $_.ProcessId }}"
        kids = subprocess.run(["powershell", "-NoProfile", "-Command", q], capture_output=True, text=True).stdout.split()
        try:
            w.doc.Close(0)
        finally:
            w.app.Quit(0)
        time.sleep(3)
        for k in kids:
            subprocess.run(["taskkill", "/PID", k, "/F"], capture_output=True)
    (CROPS / "capture.json").write_text(json.dumps(out, indent=1), encoding="utf-8")
    return out


def main() -> None:
    if "--no-capture" in sys.argv:
        cap = json.loads((CROPS / "capture.json").read_text(encoding="utf-8"))
    else:
        with real_input("chart selection marks (Word)"):
            cap = capture()
    excel = json.loads((HERE / "results" / "2026-11-11-excel-chart-selection.json").read_text(encoding="utf-8"))
    result = {"what": "Word 365's marks of a selected chart part (selmarks.py), and how they compare with Excel's",
              "scale": cap.get("scale"), "parts": {}}
    for name, info in cap["crops"].items():
        if name == "none":
            continue
        if "error" in info:
            result["parts"][name] = info
            continue
        base = CROPS / ("column-none.png" if name == "chart" else "column-chart.png")
        m = selmarks.measure(CROPS / info["file"], base, cap["box"])
        m["vsExcel"] = selmarks.compare(m, excel["parts"][f"column-{name}"])
        result["parts"][name] = m
    RESULT.write_text(json.dumps(result, indent=1, ensure_ascii=False) + "\n", encoding="utf-8", newline="\n")
    for k, v in result["parts"].items():
        print(k, v.get("summary", v), "| vs Excel:", v.get("vsExcel", {}).get("fails"))


if __name__ == "__main__":
    main()
