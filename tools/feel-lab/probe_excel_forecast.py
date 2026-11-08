"""Records Daten > Prognose > Prognoseblatt in German Excel 365 (D-092), without real input: our own Excel,
made visible off screen; the button and the dialog's controls are used through UI Automation (Invoke,
Toggle, Value), the sheet Excel makes is read through COM and from the file Excel saves.

Runs: (1) the dialog's defaults; (2) with Optionen: Säulendiagramm, Konfidenzintervall 90 %, Saisonalität
manuell 4, Prognosestatistik einschließen. For each: the dialog's controls (labels, types, values), the new
sheet's name, cells (Formula2Local, Value2, NumberFormatLocal), table, column widths and the chart (type,
series formulas, title, legend). Labels and values only (D-058).

    .venv/Scripts/python.exe probe_excel_forecast.py

Result: results/2026-11-08-excel-forecast.json; Excel's workbook to TEMP/raabe-forecast-excel.xlsx and,
anonymized, packages/ooxml/test/fixtures/excel-forecast.xlsx.
"""
from __future__ import annotations

import json
import math
import subprocess
import sys
import tempfile
import threading
import time
from pathlib import Path

from feellab import uia, winapi
from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
ROOT = HERE.parents[1]
RESULT = HERE / "results" / "2026-11-08-excel-forecast.json"
TMP = Path(tempfile.gettempdir())
DIALOG_CLASSES = ("#32770", "bosa_sdm_XL9", "NUIDialog")
fg0 = winapi.foreground()


def give_back() -> None:
    if fg0 and winapi.foreground() != fg0:
        try:
            winapi.activate(fg0, 0)
        except Exception:  # noqa: BLE001
            pass


def descendants(root) -> list[tuple]:
    out = []
    try:
        f = root.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
    except Exception:  # noqa: BLE001
        return out
    for i in range(f.Length):
        e = f.GetElement(i)
        try:
            d = uia.describe(e)
        except Exception:  # noqa: BLE001
            continue
        try:
            vp = e.GetCurrentPattern(uia.U.UIA_ValuePatternId)
            if vp:
                d["value"] = vp.QueryInterface(uia.U.IUIAutomationValuePattern).CurrentValue
        except Exception:  # noqa: BLE001
            pass
        try:
            tp = e.GetCurrentPattern(uia.U.UIA_TogglePatternId)
            if tp:
                d["toggle"] = tp.QueryInterface(uia.U.IUIAutomationTogglePattern).CurrentToggleState
        except Exception:  # noqa: BLE001
            pass
        out.append((e, d))
    return out


def data() -> list[list]:
    rows = [["Monat", "Besucher"]]
    for i in range(24):
        rows.append([45292 + 31 * i - (i // 2), round(120 + 4 * i + 25 * math.sin(2 * math.pi * i / 6) + (i % 5) * 1.7, 1)])
    # Monthly-like first days: use real month starts.
    starts = [45292, 45323, 45352, 45383, 45413, 45444, 45474, 45505, 45536, 45566, 45597, 45627]
    starts += [s + 366 for s in starts]
    for i, s in enumerate(starts):
        rows[i + 1][0] = s
    return rows


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    out: dict = {"data": data(), "runs": {}}
    x = ExcelTarget(visible=False)
    pid = x.pid
    killed = {"v": False}
    deadline = {"t": None}

    def dog():
        while True:
            time.sleep(0.5)
            if deadline["t"] and time.time() > deadline["t"]:
                killed["v"] = True
                subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True)
                return
    threading.Thread(target=dog, daemon=True).start()

    def guarded(fn, seconds=25):
        deadline["t"] = time.time() + seconds
        try:
            return fn()
        finally:
            deadline["t"] = None
            give_back()

    def dialogs() -> list[int]:
        return winapi.find_windows(lambda h: winapi.window_pid(h) == pid and h != x.hwnd and winapi.window_class(h) in DIALOG_CLASSES)

    try:
        wb = x.wb
        ws = wb.Worksheets(1)
        ws.Name = "Werte"
        for r, row in enumerate(out["data"], 1):
            for c, v in enumerate(row, 1):
                ws.Cells(r, c).Value2 = v
        ws.Range("A2:A25").NumberFormatLocal = "TT.MM.JJJJ"
        winapi.place_quietly(x.hwnd, -4000, 0, 2400, 1300)
        x.app.Visible = True
        time.sleep(1.2)
        give_back()
        root = uia.element_for(x.hwnd)
        for run in ("default", "options"):
            ws.Activate()
            ws.Range("A1:B25").Select()
            time.sleep(0.4)
            tab = uia.find(root, lambda d: d["type"] == "TabItem" and d["name"] == "Daten", 25)
            guarded(lambda: uia.select(tab))
            time.sleep(1.0)
            btn = next(e for e, d in descendants(root) if d["name"] == "Prognoseblatt" and d["type"] == "Button")
            guarded(lambda: uia.invoke(btn))
            for _ in range(30):
                if dialogs():
                    break
                time.sleep(0.3)
            h = dialogs()[0]
            time.sleep(1.5)
            dlg = uia.element_for(h)
            rec: dict = {"title": winapi.window_title(h)}
            # Optionen expands the dialog.
            opt = next((e for e, d in descendants(dlg) if d["name"] == "Optionen" and d["type"] in ("Button", "Group", "TreeItem", "Hyperlink")), None)
            if opt is not None:
                for how in (lambda e: e.GetCurrentPattern(uia.U.UIA_TogglePatternId).QueryInterface(uia.U.IUIAutomationTogglePattern).Toggle(),
                            lambda e: e.GetCurrentPattern(uia.U.UIA_ExpandCollapsePatternId).QueryInterface(uia.U.IUIAutomationExpandCollapsePattern).Expand(),
                            uia.invoke):
                    try:
                        guarded(lambda: how(opt))
                        break
                    except Exception:  # noqa: BLE001
                        continue
                time.sleep(1.0)
            rec["controls"] = [{k: d[k] for k in ("type", "name", "key", "enabled", "value", "toggle") if k in d}
                               for _, d in descendants(dlg) if d["name"] or d.get("value")]
            if run == "options":
                els = descendants(dlg)

                def by(name, types):
                    return next((e for e, d in els if d["name"] == name and d["type"] in types), None)
                col = by("Säulendiagramm erstellen", ("Button", "RadioButton", "ListItem")) or next((e for e, d in els if "Säule" in d["name"] and d["type"] in ("Button", "RadioButton", "ListItem")), None)
                if col is not None:
                    for how in (uia.invoke, uia.select, lambda e: e.GetCurrentPattern(uia.U.UIA_TogglePatternId).QueryInterface(uia.U.IUIAutomationTogglePattern).Toggle(),
                                lambda e: e.GetCurrentPattern(uia.U.UIA_LegacyIAccessiblePatternId).QueryInterface(uia.U.IUIAutomationLegacyIAccessiblePattern).DoDefaultAction()):
                        try:
                            guarded(lambda: how(col))
                            rec["column"] = "set"
                            break
                        except Exception:  # noqa: BLE001
                            continue
                for e, d in els:
                    if d["type"] == "CheckBox" and d["name"].startswith("Prognosestatistik") and d.get("toggle") == 0:
                        guarded(lambda: e.GetCurrentPattern(uia.U.UIA_TogglePatternId).QueryInterface(uia.U.IUIAutomationTogglePattern).Toggle())
                    if d["type"] == "RadioButton" and d["name"].startswith("Manuell"):
                        guarded(lambda: e.GetCurrentPattern(uia.U.UIA_SelectionItemPatternId).QueryInterface(uia.U.IUIAutomationSelectionItemPattern).Select())
                time.sleep(0.6)
                els = descendants(dlg)
                edits = [(e, d) for e, d in els if d["type"] in ("Edit", "Spinner") and "value" in d]
                rec["edits_before"] = [{k: d[k] for k in ("type", "name", "value")} for _, d in edits]
                for e, d in edits:
                    v = None
                    if d["name"].startswith("Konfidenzintervall"):
                        v = "90%"
                    elif d["name"].startswith("Saisonalität") or d["name"] == "Manuell festlegen":
                        v = "4"
                    if v is not None:
                        try:
                            guarded(lambda: e.GetCurrentPattern(uia.U.UIA_ValuePatternId).QueryInterface(uia.U.IUIAutomationValuePattern).SetValue(v))
                        except Exception as ex:  # noqa: BLE001
                            rec.setdefault("errors", []).append(str(ex))
                time.sleep(0.6)
                rec["controls_set"] = [{k: d[k] for k in ("type", "name", "value", "toggle") if k in d} for _, d in descendants(dlg) if d["type"] in ("Edit", "Spinner", "CheckBox", "RadioButton", "ComboBox")]
            before = wb.Worksheets.Count
            make = next((e for e, d in descendants(dlg) if d["name"] == "Erstellen" and d["type"] == "Button"), None)
            if make is None:
                rec["error"] = "no Erstellen button"
                winapi.user32.PostMessageW(h, 0x0010, 0, 0)
                out["runs"][run] = rec
                continue
            guarded(lambda: uia.invoke(make))
            for _ in range(40):
                if wb.Worksheets.Count > before and not dialogs():
                    break
                time.sleep(0.3)
            time.sleep(1.0)
            ns = wb.ActiveSheet
            rec["sheet"] = ns.Name
            rec["index"] = ns.Index
            ur = ns.UsedRange
            cells = {}
            for r in range(1, ur.Row + ur.Rows.Count):
                for c in range(1, ur.Column + ur.Columns.Count):
                    cl = ns.Cells(r, c)
                    if cl.Formula2 in ("", None):
                        continue
                    cells[cl.Address.replace("$", "")] = {"f": cl.Formula2Local, "v": cl.Value2, "fmt": cl.NumberFormatLocal, "text": cl.Text}
            rec["cells"] = cells
            rec["widths"] = {ns.Cells(1, c).Address.replace("$", "")[:-1]: ns.Columns(c).ColumnWidth for c in range(1, ur.Column + ur.Columns.Count)}
            rec["tables"] = [{"name": lo.Name, "range": lo.Range.Address.replace("$", ""), "style": lo.TableStyle.Name if lo.TableStyle else None,
                              "columns": [lo.ListColumns(i).Name for i in range(1, lo.ListColumns.Count + 1)]} for lo in ns.ListObjects]
            charts = []
            for co in ns.ChartObjects():
                ch = co.Chart
                charts.append({"name": co.Name, "type": ch.ChartType, "left": co.Left, "top": co.Top, "width": co.Width, "height": co.Height,
                               "tl": co.TopLeftCell.Address.replace("$", ""), "br": co.BottomRightCell.Address.replace("$", ""),
                               "title": ch.ChartTitle.Text if ch.HasTitle else None, "legend": ch.Legend.Position if ch.HasLegend else None,
                               "series": [{"name": s.Name, "formula": s.Formula, "type": s.ChartType} for s in ch.SeriesCollection()]})
            rec["charts"] = charts
            out["runs"][run] = rec
            print(run, rec["sheet"], len(cells), json.dumps(rec["tables"], ensure_ascii=False), flush=True)
        x.app.Visible = False
        path = TMP / "raabe-forecast-excel.xlsx"
        if path.exists():
            path.unlink()
        for prop in ("Author", "Last author"):
            try:
                wb.BuiltinDocumentProperties.Item(prop).Value = "Raabe Office Test"
            except Exception:  # noqa: BLE001
                pass
        wb.SaveAs(str(path), 51)
    finally:
        give_back()
        try:
            x.close()
        except Exception:  # noqa: BLE001
            subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True)
        give_back()
    out["watchdogFired"] = killed["v"]
    RESULT.write_text(json.dumps(out, ensure_ascii=False, indent=1, default=str), encoding="utf-8", newline="\n")
    print("written", RESULT.name)


if __name__ == "__main__":
    main()
