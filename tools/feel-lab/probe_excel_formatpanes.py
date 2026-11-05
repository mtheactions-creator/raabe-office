"""Probe: Excel's Format panes (D-066). In an Excel of our own, each chart element (Diagrammbereich,
Zeichnungsfläche, Achsen, Datenreihe, Legende, Diagrammtitel, Gitternetzlinien, Datenbeschriftungen,
Trendlinie, Fehlerbalken) and a shape, picture and text box is selected through COM, Strg+1 opens
its pane, and every option set ("Diagrammoptionen", "Textoptionen"), icon tab and section (all
expanded through the automation pattern) is read: labels and control types only (D-058).

    .venv/Scripts/python.exe probe_excel_formatpanes.py

Result: results/<date>-excel-formatpanes.json.
"""
from __future__ import annotations

import datetime
import json
import tempfile
import time
from pathlib import Path

from feellab import uia, winapi
from feellab.calc_catalogue import Case, ExcelSide
from probe_excel_inventory import all_desc, has_expand, listed

HERE = Path(__file__).parent
OUT = HERE / "results" / f"{datetime.date.today().isoformat()}-excel-formatpanes.json"


def pane_root(side: ExcelSide):
    """The open task pane: the element whose name ends in "formatieren"."""
    for e, d in all_desc(side.root):
        if d["type"] in ("Pane", "Window", "Custom", "Group") and d["name"].endswith("formatieren") and d["rect"][2] > 50:
            return e, d["name"]
    return None, None


def expand_all(el) -> None:
    for _ in range(4):
        n = 0
        for e, d in all_desc(el):
            if not has_expand(e):
                continue
            try:
                p = e.GetCurrentPattern(uia.U.UIA_ExpandCollapsePatternId).QueryInterface(uia.U.IUIAutomationExpandCollapsePattern)
                if p.CurrentExpandCollapseState == 0 and d["type"] in ("Group", "Button", "TreeItem", "ListItem", "Custom"):
                    p.Expand()
                    n += 1
            except Exception:  # noqa: BLE001
                pass
        time.sleep(0.4)
        if not n:
            return


def read_tabs(side: ExcelSide, el) -> dict:
    """Every icon tab of the current option set, expanded and read."""
    out = {}
    tabs = [(e, d) for e, d in all_desc(el) if d["type"] in ("TabItem", "RadioButton") and d["name"]]
    names = [d["name"] for _, d in tabs]
    for name in names or [""]:
        if name:
            el2, _ = pane_root(side)
            hit = [(e, d) for e, d in all_desc(el2) if d["type"] in ("TabItem", "RadioButton") and d["name"] == name]
            if not hit:
                continue
            e, d = hit[0]
            try:
                uia.select(e)
            except Exception:  # noqa: BLE001
                x, y, w, h = d["rect"]
                winapi.click(x + w / 2, y + h / 2)
            time.sleep(0.6)
        el2, _ = pane_root(side)
        expand_all(el2)
        out[name or "(eine Seite)"] = listed(all_desc(el2))
    return out


def read_pane(side: ExcelSide) -> dict:
    el, title = pane_root(side)
    if el is None:
        return {"error": "no pane"}
    result = {"title": title, "sets": {}}
    # Option sets: "Diagrammoptionen" / "Textoptionen" (or "Formoptionen" …).
    sets = [d["name"] for _, d in all_desc(el) if d["type"] in ("TabItem", "Button", "SplitButton", "MenuItem") and d["name"].endswith("optionen")]
    seen = []
    for s in sets or [""]:
        if s in seen:
            continue
        seen.append(s)
        if s:
            el, _ = pane_root(side)
            hit = [(e, d) for e, d in all_desc(el) if d["name"] == s and d["type"] in ("TabItem", "Button", "SplitButton", "MenuItem")]
            if hit:
                x, y, w, h = hit[0][1]["rect"]
                # The text part, not the dropdown arrow of "Diagrammoptionen".
                winapi.click(x + min(w / 3, 40), y + h / 2)
                time.sleep(0.7)
                winapi.press("esc") if s.endswith("optionen") and side_menu_open(side) else None
        el, _ = pane_root(side)
        result["sets"][s or "(eine Seite)"] = read_tabs(side, el)
    return result


def side_menu_open(side: ExcelSide) -> bool:
    return bool(winapi.find_windows(lambda h: winapi.window_pid(h) == side.x.pid and winapi.window_class(h) == "Net UI Tool Window"))


def elements(side: ExcelSide) -> list[tuple[str, object]]:
    x = side.x
    ws = x.wb.Worksheets(1)
    ws.Cells.Clear()
    for s in list(ws.Shapes):
        s.Delete()
    data = [("Name", "Wert", "Wert2"), ("Anna", 3, 4), ("Ben", 5, 2), ("Cem", 2, 6), ("Dora", 4, 3)]
    for r, row in enumerate(data, start=1):
        for c, v in enumerate(row, start=1):
            ws.Cells(r, c).Value = v
    sh = ws.Shapes.AddChart2(-1, 51, 250, 20, 400, 250)
    ch = sh.Chart
    ch.SetSourceData(ws.Range("A1:C5"))
    ch.HasTitle = True
    ch.SetElement(101)       # legend right
    ch.SetElement(202)       # data labels
    ch.SeriesCollection(1).Trendlines().Add()
    ch.SeriesCollection(1).HasErrorBars = True
    out = [
        ("Diagrammbereich", lambda: ch.ChartArea.Select()),
        ("Zeichnungsfläche", lambda: ch.PlotArea.Select()),
        ("Diagrammtitel", lambda: ch.ChartTitle.Select()),
        ("Horizontale Achse", lambda: ch.Axes(1).Select()),
        ("Vertikale Achse", lambda: ch.Axes(2).Select()),
        ("Hauptgitternetz", lambda: ch.Axes(2).MajorGridlines.Select()),
        ("Datenreihe", lambda: ch.SeriesCollection(1).Select()),
        ("Datenpunkt", lambda: ch.SeriesCollection(1).Points(2).Select()),
        ("Datenbeschriftungen", lambda: ch.SeriesCollection(1).DataLabels().Select()),
        ("Legende", lambda: ch.Legend.Select()),
        ("Trendlinie", lambda: ch.SeriesCollection(1).Trendlines(1).Select()),
        ("Fehlerbalken", lambda: ch.SeriesCollection(1).ErrorBars.Select()),
    ]
    return [(n, (sh, f)) for n, f in out]


def main() -> None:
    side = ExcelSide()
    data: dict = {}
    try:
        side.reset(Case("panes", {}, [], "x"))
        winapi.activate(side.x.hwnd)
        for name, (shape, select) in elements(side):
            winapi.activate(side.x.hwnd)
            try:
                shape.Select()
                time.sleep(0.3)
                select()
            except Exception as e:  # noqa: BLE001
                data[name] = {"error": str(e)[:120]}
                continue
            time.sleep(0.5)
            winapi.press("ctrl+1")
            time.sleep(1.5)
            data[name] = read_pane(side)
            print(name, data[name].get("title"), {k: list(v) for k, v in data[name].get("sets", {}).items()}, flush=True)
            OUT.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        # A shape, a picture and a text box.
        ws = side.x.wb.Worksheets(1)
        from PIL import Image
        png = Path(tempfile.gettempdir()) / "raabe-panes.png"
        Image.new("RGB", (120, 80), (70, 130, 180)).save(png)
        for name, make in [("Form", lambda: ws.Shapes.AddShape(1, 20, 300, 120, 80)), ("Bild", lambda: ws.Shapes.AddPicture(str(png), False, True, 20, 400, 120, 80)),
                           ("Textfeld", lambda: ws.Shapes.AddTextbox(1, 20, 500, 120, 60))]:
            s = make()
            s.Select()
            time.sleep(0.5)
            winapi.activate(side.x.hwnd)
            winapi.press("ctrl+1")
            time.sleep(1.5)
            data[name] = read_pane(side)
            print(name, data[name].get("title"), {k: list(v) for k, v in data[name].get("sets", {}).items()}, flush=True)
            OUT.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    finally:
        OUT.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        side.close()


if __name__ == "__main__":
    main()
