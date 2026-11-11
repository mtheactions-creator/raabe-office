"""Probe (Step 11 part 5, D-172): where Office's chart Format pane puts its parts. In our own Excel, Word and
PowerPoint (windows off screen, never activated, no input; PowerPoint only when no POWERPNT.EXE runs), a column chart's
Diagrammbereich is selected through COM and its pane opened with ExecuteMso("ChartFormatSelection"); the Füllung
section is expanded through UI Automation. The pane is read through UI Automation only: every element's type, name,
class and box in CSS px (screen px / the display's scale) relative to the pane's top left. Microsoft's labels only.

    .venv/Scripts/python.exe probe_office_pane_geometry.py [excel] [word] [powerpoint]

Result: results/2026-11-11-office-pane-geometry.json.
"""
from __future__ import annotations

import json
import sys
import tempfile
import time
from pathlib import Path

from feellab import uia, winapi
from probe_excel_chartpane import U, by_class, d_of, expander, kids, pattern

HERE = Path(__file__).parent
OUT = HERE / "results" / "2026-11-11-office-pane-geometry.json"
OFF = (-4200, 0, 1800, 2000)
KEEP = ("Button", "CheckBox", "RadioButton", "ComboBox", "Edit", "Spinner", "Slider", "TabItem", "ListItem", "Text",
        "Group", "SplitButton", "MenuItem", "Hyperlink", "Pane", "Custom", "Header", "HeaderItem")


def dump(root, scale: float) -> list[dict]:
    """The pane's elements, depth first, boxes relative to the pane."""
    r0 = d_of(root)["rect"]
    out = []
    stack = [(root, 0)]
    while stack:
        e, depth = stack.pop()
        d = d_of(e)
        if not d:
            continue
        x, y, w, h = d["rect"]
        if w > 0 and h > 0 and d["type"] in KEEP:
            out.append({"depth": depth, "type": d["type"], "name": d["name"], "class": d["class"],
                        "rect": [round((x - r0[0]) / scale, 1), round((y - r0[1]) / scale, 1), round(w / scale, 1), round(h / scale, 1)]})
        if depth < 14:
            stack.extend((c, depth + 1) for c in reversed(kids(e)))
    return out


CROPS = Path(tempfile.gettempdir()) / "raabe-pane-geo"


def title_ink(hwnd: int, frame, scale: float) -> dict | None:
    """The pane title's ink box (CSS px, relative to the pane), from a PrintWindow picture of our own window (off
    screen, no screen capture); the crop of the pane alone stays in TEMP."""
    from probe_excel_picturedrops import print_window
    from PIL import Image
    import numpy as np
    full = CROPS / "excel-window.png"
    if not print_window(hwnd, full):
        return None
    l, t, _, _ = winapi.window_rect(hwnd)
    x, y, w, h = d_of(frame)["rect"]
    im = Image.open(full).crop((x - l, y - t, x - l + w, y - t + h))
    im.save(CROPS / "excel-pane.png")
    full.unlink(missing_ok=True)
    return ink_box(np.asarray(im.convert("L")), int(5.3 * scale), int(4.7 * scale), int(200 * scale), int(28 * scale), scale)


def ink_box(g, x0: int, y0: int, w: int, h: int, scale: float) -> dict | None:
    """Dark pixels' box in a region of a grey picture, in CSS px relative to the picture."""
    import numpy as np
    reg = g[y0:y0 + h, x0:x0 + w] < 128
    ys, xs = np.nonzero(reg)
    if not len(xs):
        return None
    return {"x": round((x0 + xs.min()) / scale, 1), "y": round((y0 + ys.min()) / scale, 1),
            "w": round((xs.max() - xs.min() + 1) / scale, 1), "h": round((ys.max() - ys.min() + 1) / scale, 1)}


def open_pane(app, root, chart) -> object | None:
    chart.ChartArea.Select()
    app.CommandBars.ExecuteMso("ChartFormatSelection")
    for _ in range(40):
        time.sleep(0.25)
        body = by_class(root, "NetUInetpane")
        if body is not None and (d_of(body) or {}).get("name", "").endswith("formatieren"):
            break
    else:
        return None
    time.sleep(0.8)
    # Expand Füllung (the first section) so its rows can be measured.
    sv = by_class(body, "NetUIScrollViewer")
    for g in kids(sv) if sv is not None else []:
        ks = kids(g)
        h = d_of(ks[0]) if ks else None
        if h and h["type"] == "Button" and h["name"] == "Füllung":
            # The header button is a toggle (as probe_excel_chartpane.expand_section opens it).
            tp = pattern(ks[0], U.UIA_TogglePatternId, U.IUIAutomationTogglePattern)
            try:
                if tp is not None and tp.CurrentToggleState == 0:
                    tp.Toggle()
                elif tp is None and expander(g) is not None:
                    expander(g).Expand()
            except Exception:  # noqa: BLE001
                pass
            break
    time.sleep(0.8)
    # The pane's frame (title, close button) lies around the body: read from the body's parents up to the task pane.
    walker = uia._uia.ControlViewWalker
    frame = body
    for _ in range(4):
        par = walker.GetParentElement(frame)
        d = d_of(par) if par is not None else None
        if not d or d["rect"][2] > d_of(body)["rect"][2] + 60:
            break
        frame = par
    return frame


def record(app_name: str) -> dict:
    if app_name == "excel":
        from feellab.excel import ExcelTarget
        x = ExcelTarget(visible=False, place=False)
        try:
            winapi.place_quietly(x.hwnd, *OFF)
            x.app.Visible = True
            winapi.place_quietly(x.hwnd, *OFF)
            ws = x.wb.Worksheets(1)
            for r, row in enumerate([["", "A", "B"], ["x", 1, 2], ["y", 3, 4]], 1):
                for c, v in enumerate(row, 1):
                    ws.Cells(r, c).Value = v
            co = ws.ChartObjects().Add(300, 30, 360, 216)
            co.Chart.SetSourceData(ws.Range("A1:C3"))
            co.Activate()
            root = uia.element_for(x.hwnd)
            body = open_pane(x.app, root, co.Chart)
            scale = winapi.dpi_scale(x.hwnd)
            if body is None:
                return {"error": "pane not found"}
            out = {"scale": scale, "elements": dump(body, scale)}
            out["title"] = title_ink(x.hwnd, body, scale)
            return out
        finally:
            x.close()
    if app_name == "word":
        from feellab.targets import WordTarget
        import subprocess
        w = WordTarget("short", visible=False)
        word_pid = 0
        try:
            w.app.Visible = True
            winapi.place_quietly(w.hwnd, *OFF)
            word_pid = winapi.window_pid(w.hwnd)
            doc = w.doc
            doc.Content.Text = ""
            ish = doc.InlineShapes.AddChart2(-1, 51, doc.Range(0, 0))
            try:
                ish.Chart.ChartData.Activate()
                ish.Chart.ChartData.Workbook.Close()
            except Exception:  # noqa: BLE001
                pass
            winapi.place_quietly(w.hwnd, *OFF)
            root = uia.element_for(w.hwnd)
            body = open_pane(w.app, root, ish.Chart)
            scale = winapi.dpi_scale(w.hwnd)
            return {"scale": scale, "elements": dump(body, scale)} if body is not None else {"error": "pane not found"}
        finally:
            q = f"Get-CimInstance Win32_Process -Filter \"name='EXCEL.EXE' and ParentProcessId={word_pid}\" | ForEach-Object {{ $_.ProcessId }}"
            kids_ = subprocess.run(["powershell", "-NoProfile", "-Command", q], capture_output=True, text=True).stdout.split() if word_pid else []
            try:
                w.doc.Close(0)
            finally:
                w.app.Quit(0)
            time.sleep(3)
            for k in kids_:
                subprocess.run(["taskkill", "/PID", k, "/F"], capture_output=True)
    from feellab.powerpoint import PowerPointTarget, owner_powerpoint_running
    import subprocess
    if owner_powerpoint_running():
        return {"error": "POWERPNT.EXE running (not ours): skipped"}
    t = PowerPointTarget(visible=True, place=False)
    try:
        pres = t.new_presentation(with_window=True, name="panegeo")
        winapi.place_quietly(t.hwnd, *OFF)
        slide = pres.Slides.Add(1, 12)
        sh = slide.Shapes.AddChart2(-1, 51, 120, 80, 480, 288)
        try:
            sh.Chart.ChartData.Activate()
            sh.Chart.ChartData.Workbook.Close()
        except Exception:  # noqa: BLE001
            pass
        winapi.place_quietly(t.hwnd, *OFF)
        t.app.ActiveWindow.ViewType = 9  # ppViewNormal
        sh.Select()
        time.sleep(0.5)
        root = uia.element_for(t.hwnd)
        body = open_pane(t.app, root, sh.Chart)
        scale = winapi.dpi_scale(t.hwnd)
        return {"scale": scale, "elements": dump(body, scale)} if body is not None else {"error": "pane not found"}
    finally:
        q = f"Get-CimInstance Win32_Process -Filter \"name='EXCEL.EXE' and ParentProcessId={t.pid}\" | ForEach-Object {{ $_.ProcessId }}"
        kids_ = subprocess.run(["powershell", "-NoProfile", "-Command", q], capture_output=True, text=True).stdout.split() if t.pid else []
        t.close(kill_if_unreachable=True)
        time.sleep(3)
        for k in kids_:
            subprocess.run(["taskkill", "/PID", k, "/F"], capture_output=True)


def main() -> None:
    apps = [a for a in sys.argv[1:] if a in ("excel", "word", "powerpoint")] or ["excel", "word", "powerpoint"]
    data = json.loads(OUT.read_text(encoding="utf-8")) if OUT.exists() else {}
    for a in apps:
        data[a] = record(a)
        print(a, data[a].get("error") or f"{len(data[a]['elements'])} elements")
    OUT.write_text(json.dumps(data, indent=1, ensure_ascii=False) + "\n", encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
