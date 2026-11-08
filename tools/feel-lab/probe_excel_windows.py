"""Records German Excel 365's windows and views (D-093) in our own invisible Excel, COM only (no real input,
no window shown): Teilen (split panes, how Excel stores them in the sheetView's pane), Neues Fenster (captions,
the bookViews and the per-window sheetViews of the file), Ausblenden, Alle anordnen, Benutzerdefinierte Ansichten
(customWorkbookViews / customSheetViews) and the Seitenlayout view.

    .venv/Scripts/python.exe probe_excel_windows.py

Writes results/<date>-excel-windows.json and the fixture packages/ooxml/test/fixtures/excel-windows.json (the
XML parts Excel saved, the COM state beside them). Workbooks go to TEMP only."""

from __future__ import annotations

import datetime
import json
import re
import sys
import tempfile
import zipfile
from pathlib import Path

from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
ROOT = HERE.parents[1]
TMP = Path(tempfile.gettempdir())
FIXTURE = ROOT / "packages" / "ooxml" / "test" / "fixtures" / "excel-windows.json"


def to_fixture(src: Path, name: str) -> None:
    """Excel's saved workbook as a test fixture: the save folder (x15ac:absPath) left out, the user's name
    replaced (anonymize_office.py)."""
    import shutil
    import subprocess
    dst = ROOT / "packages" / "ooxml" / "test" / "fixtures" / name
    shutil.copyfile(src, dst)
    with zipfile.ZipFile(dst) as z:
        items = [(info, z.read(info)) for info in z.infolist()]
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for info, data in items:
            if info.filename == "xl/workbook.xml":
                data = re.sub(rb"<mc:AlternateContent\b(?:(?!</mc:AlternateContent>)[\s\S])*?x15ac:absPath[\s\S]*?</mc:AlternateContent>", b"", data)
            z.writestr(info, data)
    subprocess.run([sys.executable, str(HERE / "anonymize_office.py"), str(dst)], check=True, capture_output=True)


def parts(path: Path) -> dict:
    with zipfile.ZipFile(path) as z:
        wb = z.read("xl/workbook.xml").decode("utf-8")
        sheets = {}
        for n in sorted(z.namelist()):
            m = re.fullmatch(r"xl/worksheets/(sheet\d+)\.xml", n)
            if m:
                x = z.read(n).decode("utf-8")
                sheets[m.group(1)] = {
                    "sheetViews": (re.search(r"<sheetViews>[\s\S]*?</sheetViews>", x) or [None])[0],
                    "customSheetViews": (re.search(r"<customSheetViews>[\s\S]*?</customSheetViews>", x) or [None])[0],
                    "sheetFormatPr": (re.search(r"<sheetFormatPr\b[^>]*/>", x) or [None])[0],
                    "cols": (re.search(r"<cols>[\s\S]*?</cols>", x) or [None])[0],
                }
    return {"bookViews": (re.search(r"<bookViews>[\s\S]*?</bookViews>", wb) or [None])[0],
            "customWorkbookViews": (re.search(r"<customWorkbookViews>[\s\S]*?</customWorkbookViews>", wb) or [None])[0],
            "definedNames": (re.search(r"<definedNames>[\s\S]*?</definedNames>", wb) or [None])[0],
            "sheets": sheets}


def window_state(w) -> dict:
    out = {"caption": w.Caption, "zoom": w.Zoom, "split": bool(w.Split), "freeze": bool(w.FreezePanes),
           "splitRow": w.SplitRow, "splitColumn": w.SplitColumn, "splitHorizontal": w.SplitHorizontal, "splitVertical": w.SplitVertical,
           "scrollRow": w.ScrollRow, "scrollColumn": w.ScrollColumn, "headers": bool(w.DisplayHeadings), "view": w.View,
           "panes": [], "visible": bool(w.Visible), "activeSheet": w.ActiveSheet.Name, "activeCell": w.ActiveCell.Address.replace("$", "")}
    for i in range(1, w.Panes.Count + 1):
        p = w.Panes(i)
        out["panes"].append({"scrollRow": p.ScrollRow, "scrollColumn": p.ScrollColumn, "visible": p.VisibleRange.Address.replace("$", "")})
    out["activePane"] = w.ActivePane.Index
    return out


SPLITS = [
    # name, setup, split
    ("cols3", {}, {"col": 3}),
    ("rows5", {}, {"row": 5}),
    ("both3x5", {}, {"col": 3, "row": 5}),
    ("both-zoom150", {"zoom": 150}, {"col": 3, "row": 5}),
    ("both-zoom75", {"zoom": 75}, {"col": 3, "row": 5}),
    ("both-noheaders", {"headers": False}, {"col": 3, "row": 5}),
    ("both-widecols", {"widths": {"A": 20, "B": 5}}, {"col": 3, "row": 5}),
    ("both-scrolled", {"scroll": (10, 4)}, {"col": 3, "row": 5}),
    ("points100x60", {}, {"h": 100.0, "v": 60.0}),
    ("points-zoom150", {"zoom": 150}, {"h": 100.0, "v": 60.0}),
    ("both-bigrow", {"heights": {2: 40}}, {"col": 2, "row": 3}),
    ("cols3-panescroll", {}, {"col": 3, "row": 5, "pane_scroll": {4: (30, 12), 1: (2, 2)}, "active_pane": 4}),
]


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    res: dict = {"date": datetime.date.today().isoformat(), "splits": [], "windows": {}, "custom": {}, "pagelayout": {}}
    x = ExcelTarget(visible=False)
    app = x.app
    try:
        # ---- Teilen ------------------------------------------------------------------------
        wb = x.wb
        st = wb.Styles("Standard")
        st.Font.Name = "Calibri"
        st.Font.Size = 11
        while wb.Worksheets.Count < len(SPLITS):
            wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        for i, (name, setup, split) in enumerate(SPLITS, 1):
            ws = wb.Worksheets(i)
            ws.Name = name
            ws.Activate()
            w = app.ActiveWindow
            for col, width in setup.get("widths", {}).items():
                ws.Columns(col).ColumnWidth = width
            for row, height in setup.get("heights", {}).items():
                ws.Rows(row).RowHeight = height
            w.Zoom = setup.get("zoom", 100)
            w.DisplayHeadings = setup.get("headers", True)
            if "scroll" in setup:
                w.ScrollRow, w.ScrollColumn = setup["scroll"]
            if "col" in split:
                w.SplitColumn = split["col"]
            if "row" in split:
                w.SplitRow = split["row"]
            if "h" in split:
                w.SplitHorizontal = split["h"]
            if "v" in split:
                w.SplitVertical = split["v"]
            for pane, (r, c) in split.get("pane_scroll", {}).items():
                w.Panes(pane).ScrollRow = r
                w.Panes(pane).ScrollColumn = c
            if "active_pane" in split:
                w.Panes(split["active_pane"]).Activate()
            ws.Range("H20").Select() if "active_pane" in split else None
            res["splits"].append({"name": name, "setup": {k: (list(v) if isinstance(v, tuple) else v) for k, v in setup.items()},
                                  "split": {k: v for k, v in split.items() if k != "pane_scroll"},
                                  "pane_scroll": {str(k): list(v) for k, v in split.get("pane_scroll", {}).items()},
                                  "com": window_state(w), "standardWidth": ws.StandardWidth, "colWidthPt": ws.Columns("A").Width,
                                  "rowHeightPt": ws.Rows(1).Height})
            print(name, res["splits"][-1]["com"]["splitHorizontal"], res["splits"][-1]["com"]["splitVertical"], flush=True)
        # Teilen again removes the split (Window.Split = False); remembered for the last case.
        wb.Worksheets(1).Activate()
        path = TMP / "raabe-excel-split.xlsx"
        wb.SaveAs(str(path), 51)
        p = parts(path)
        for i, s in enumerate(res["splits"], 1):
            s["xml"] = p["sheets"][f"sheet{i}"]
        # Removing a split.
        w = app.ActiveWindow
        w.Split = False
        res["split_off"] = window_state(w)
        wb.Close(False)

        # ---- Neues Fenster -------------------------------------------------------------------
        wb = app.Workbooks.Add()
        wb.Worksheets.Add(After=wb.Worksheets(1))
        wb.Worksheets(1).Activate()
        res["windows"]["caption_one"] = wb.Windows(1).Caption
        w2 = wb.NewWindow()
        res["windows"]["captions_two"] = [wb.Windows(i).Caption for i in range(1, wb.Windows.Count + 1)]
        res["windows"]["new_window_state"] = window_state(w2)
        wb.NewWindow()
        res["windows"]["captions_three"] = [wb.Windows(i).Caption for i in range(1, wb.Windows.Count + 1)]
        # The newest window is window 1 of the collection (the active one): its caption is Mappe:3.
        res["windows"]["active_caption"] = app.ActiveWindow.Caption
        app.Windows(wb.Name + "  -  2").Close()
        res["windows"]["closed"] = "window 2 of 3"
        res["windows"]["captions_after_close3"] = [wb.Windows(i).Caption for i in range(1, wb.Windows.Count + 1)]
        # Window 2 shows another sheet, its own cell, zoom and scroll; window 1 stays.
        win2 = next(wb.Windows(i) for i in range(1, wb.Windows.Count + 1) if wb.Windows(i).Caption.endswith("-  2"))
        win1 = next(wb.Windows(i) for i in range(1, wb.Windows.Count + 1) if wb.Windows(i).Caption.endswith("-  1"))
        win2.Activate()
        wb.Worksheets(2).Activate()
        wb.Worksheets(2).Range("C3").Select()
        win2.Zoom = 150
        win2.ScrollRow = 10
        win1.Activate()
        wb.Worksheets(1).Range("B2").Select()
        wb.Worksheets(1).Range("A1").Value2 = 5
        res["windows"]["shown_in_2"] = None
        res["windows"]["state1"] = window_state(win1)
        res["windows"]["state2"] = window_state(win2)
        path = TMP / "raabe-excel-windows.xlsx"
        wb.SaveAs(str(path), 51)
        res["windows"]["xml_two"] = parts(path)
        res["windows"]["captions_saved"] = [wb.Windows(i).Caption for i in range(1, wb.Windows.Count + 1)]
        # Ausblenden: the active window hidden.
        win2.Activate()
        win2.Visible = False
        res["windows"]["after_hide"] = {"active": app.ActiveWindow.Caption, "visible": [bool(wb.Windows(i).Visible) for i in range(1, wb.Windows.Count + 1)]}
        path_h = TMP / "raabe-excel-windows-hidden.xlsx"
        wb.SaveCopyAs(str(path_h))
        res["windows"]["xml_hidden"] = parts(path_h)
        win2.Visible = True
        # Alle anordnen (Unterteilt, Horizontal, Vertikal, Überlappend).
        res["windows"]["arrange"] = {}
        for style, code in (("tiled", 1), ("horizontal", -4128), ("vertical", -4166), ("cascade", 7)):
            try:
                app.Windows.Arrange(code, True)
                res["windows"]["arrange"][style] = [{"caption": wb.Windows(i).Caption, "left": wb.Windows(i).Left, "top": wb.Windows(i).Top,
                                                     "width": wb.Windows(i).Width, "height": wb.Windows(i).Height, "state": wb.Windows(i).WindowState}
                                                    for i in range(1, wb.Windows.Count + 1)]
            except Exception as e:  # noqa: BLE001
                res["windows"]["arrange"][style] = {"error": str(e)[:200]}
        # Closing window 1 renames the other back to the workbook's name.
        win1.Close()
        res["windows"]["caption_after_close1"] = [wb.Windows(i).Caption for i in range(1, wb.Windows.Count + 1)]
        wb.Close(False)
        # Opening the saved file opens both windows again.
        wb = app.Workbooks.Open(str(path))
        res["windows"]["reopened"] = [{"caption": wb.Windows(i).Caption, "sheet": wb.Windows(i).ActiveSheet.Name, "cell": wb.Windows(i).ActiveCell.Address.replace("$", ""),
                                       "zoom": wb.Windows(i).Zoom} for i in range(1, wb.Windows.Count + 1)]
        wb.Close(False)

        # ---- Benutzerdefinierte Ansichten -----------------------------------------------------
        wb = app.Workbooks.Add()
        ws = wb.Worksheets(1)
        for r in range(1, 11):
            for c in range(1, 5):
                ws.Cells(r, c).Value2 = r * 10 + c
        ws.Range("B2").Select()
        app.ActiveWindow.Zoom = 100
        wb.CustomViews.Add("Alles", True, True)
        ws.Rows("3:4").Hidden = True
        ws.Columns("C").ColumnWidth = 20
        ws.PageSetup.Orientation = 2
        app.ActiveWindow.Zoom = 120
        ws.Range("D5").Select()
        wb.CustomViews.Add("Ohne Zeilen", True, True)
        ws.Range("A1").Select()
        wb.CustomViews.Add("Nur Anzeige", False, False)
        res["custom"]["names"] = [wb.CustomViews(i).Name for i in range(1, wb.CustomViews.Count + 1)]
        res["custom"]["flags"] = [[wb.CustomViews(i).PrintSettings, wb.CustomViews(i).RowColSettings] for i in range(1, wb.CustomViews.Count + 1)]
        wb.CustomViews("Alles").Show()
        res["custom"]["after_show_alles"] = {"hidden3": bool(ws.Rows(3).Hidden), "widthC": ws.Columns("C").ColumnWidth, "zoom": app.ActiveWindow.Zoom,
                                             "cell": app.ActiveCell.Address.replace("$", ""), "orientation": ws.PageSetup.Orientation}
        wb.CustomViews("Ohne Zeilen").Show()
        res["custom"]["after_show_ohne"] = {"hidden3": bool(ws.Rows(3).Hidden), "widthC": ws.Columns("C").ColumnWidth, "zoom": app.ActiveWindow.Zoom,
                                            "cell": app.ActiveCell.Address.replace("$", ""), "orientation": ws.PageSetup.Orientation}
        path = TMP / "raabe-excel-customviews.xlsx"
        wb.SaveAs(str(path), 51)
        res["custom"]["xml"] = parts(path)
        # A table in the workbook makes Benutzerdefinierte Ansichten unavailable.
        ws.ListObjects.Add(1, ws.Range("A1:D10"), None, 1)
        try:
            wb.CustomViews.Add("Mit Tabelle", True, True)
            res["custom"]["with_table"] = "added"
        except Exception as e:  # noqa: BLE001
            res["custom"]["with_table"] = "refused: " + str(e)[:200]
        wb.Close(False)

        # ---- Seitenlayout -------------------------------------------------------------------
        wb = app.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Range("A1").Value2 = 1
        w = app.ActiveWindow
        res["pagelayout"]["ruler_normal"] = bool(w.DisplayRuler)
        w.View = 3
        res["pagelayout"]["view"] = w.View
        res["pagelayout"]["zoom_in_layout"] = w.Zoom
        res["pagelayout"]["ruler"] = bool(w.DisplayRuler)
        res["pagelayout"]["whitespace"] = bool(w.DisplayWhitespace)
        w.Zoom = 85
        path = TMP / "raabe-excel-pagelayout.xlsx"
        wb.SaveAs(str(path), 51)
        res["pagelayout"]["xml"] = parts(path)
        w.DisplayRuler = False
        wb.SaveCopyAs(str(TMP / "raabe-excel-pagelayout-noruler.xlsx"))
        res["pagelayout"]["xml_noruler"] = parts(TMP / "raabe-excel-pagelayout-noruler.xlsx")
        w.DisplayRuler = True
        w.View = 2
        wb.SaveCopyAs(str(TMP / "raabe-excel-pagebreak.xlsx"))
        res["pagelayout"]["xml_pagebreak"] = parts(TMP / "raabe-excel-pagebreak.xlsx")
        w.View = 1
        res["pagelayout"]["zoom_back_normal"] = w.Zoom
        wb.Close(False)
    finally:
        try:
            x.wb = app.Workbooks.Add()
        except Exception:  # noqa: BLE001
            pass
        x.close()
    to_fixture(TMP / "raabe-excel-split.xlsx", "excel-split.xlsx")
    to_fixture(TMP / "raabe-excel-windows.xlsx", "excel-windows.xlsx")
    to_fixture(TMP / "raabe-excel-windows-hidden.xlsx", "excel-windows-hidden.xlsx")
    to_fixture(TMP / "raabe-excel-customviews.xlsx", "excel-customviews.xlsx")
    to_fixture(TMP / "raabe-excel-pagelayout.xlsx", "excel-pagelayout.xlsx")
    out = HERE / "results" / f"{res['date']}-excel-windows.json"
    out.write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    FIXTURE.write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print("written", out.name)


if __name__ == "__main__":
    main()
