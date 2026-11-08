"""Records the user interface around the Zeitachse in German Excel 365 (D-121) through UI Automation, without real
input: our own Excel, made visible off screen. The contextual tab Zeitachse (a timeline selected through COM) with every
control's label, keytip and state; the Zeitachsenformatvorlagen gallery's entries; the dialogs Zeitachsen einfügen
(Einfügen > Zeitachse and PivotTable-Analyse > Zeitachse einfügen) and Berichtsverbindungen (labels only, closed with
Abbrechen). Invoking a control makes our Excel the foreground window for a moment, so the probe holds the real-input
lock for its short run (another session's keys must not land in our Excel). Labels, types and keytips only (D-058).

    .venv/Scripts/python.exe probe_excel_timeline_ui.py

Result: merged into results/2026-11-09-excel-timeline.json ("ui").
"""
from __future__ import annotations

import datetime as dt
import json
import subprocess
import sys
import time
from pathlib import Path

from feellab import uia, winapi
from feellab.excel import ExcelTarget
from feellab.inputlock import real_input
from probe_excel_slicer_ui import DIALOG_CLASSES, Watchdog, descendants, listed, ribbon
from probe_excel_timeline import DATA

HERE = Path(__file__).parent
RESULT = HERE / "results" / "2026-11-09-excel-timeline.json"


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    out: dict = {}
    with real_input("timeline-ui probe (calc-misc fork)", wait=1800):
        fg0 = winapi.foreground()

        def give_back() -> None:
            if fg0 and winapi.foreground() != fg0:
                try:
                    winapi.activate(fg0, 0)
                except Exception:  # noqa: BLE001
                    pass

        x = ExcelTarget(visible=False)
        pid = x.pid
        dog = Watchdog(pid)

        def windows() -> list[int]:
            return winapi.find_windows(lambda h: winapi.window_pid(h) == pid and h != x.hwnd and winapi.window_class(h) in DIALOG_CLASSES)

        def popups() -> list[tuple]:
            res = []
            for h in winapi.find_windows(lambda h: winapi.window_pid(h) == pid and h != x.hwnd):
                res.extend(descendants(uia.element_for(h)))
            return res

        def invoke(el, wait: float = 0.8) -> None:
            dog.arm(25)
            try:
                uia.invoke(el)
            finally:
                dog.disarm()
            give_back()
            time.sleep(wait)

        def read_dialog(name: str) -> dict | None:
            for _ in range(20):
                hs = windows()
                if hs:
                    break
                time.sleep(0.3)
            else:
                out.setdefault("missing", []).append(name)
                return None
            h = hs[0]
            time.sleep(0.6)
            el = uia.element_for(h)
            rec = {"title": winapi.window_title(h), "controls": listed(descendants(el), with_lists=True)}
            btn = next((e for e, d in descendants(el) if d["type"] == "Button" and d["name"] in ("Abbrechen", "Schließen")), None)
            if btn is not None:
                invoke(btn, 0.6)
            for hh in windows():
                winapi.user32.PostMessageW(hh, 0x0010, 0, 0)
            time.sleep(0.5)
            print(name, rec["title"], len(rec["controls"]), flush=True)
            return rec

        def tab(root, name: str):
            t = uia.find(root, lambda d: d["type"] == "TabItem" and d["name"] == name, 25)
            if t is None:
                return None
            uia.select(t)
            give_back()
            time.sleep(1.0)
            return t

        try:
            wb = x.wb
            ws = wb.Worksheets(1)
            ws.Name = "Noten"
            for r, row in enumerate(DATA, 1):
                for c, v in enumerate(row, 1):
                    cell = ws.Cells(r, c)
                    if isinstance(v, dt.date):
                        cell.Formula = f"=DATE({v.year},{v.month},{v.day})"
                        cell.Value = cell.Value2
                        cell.NumberFormat = "dd.mm.yyyy"
                    else:
                        cell.Value = v
            src = ws.Range(ws.Cells(1, 1), ws.Cells(len(DATA), 3))
            ws2 = wb.Worksheets.Add(After=ws)
            ws2.Name = "Pivot"
            pc = wb.PivotCaches().Create(1, src)
            pt = pc.CreatePivotTable(ws2.Range("A3"), "PivotNoten")
            pt.PivotFields("Fach").Orientation = 1
            pt.AddDataField(pt.PivotFields("Note"), None, -4112)
            ws2.Range("A4").Select()
            winapi.place_quietly(x.hwnd, -4000, 0, 2400, 1300)
            x.app.Visible = True
            time.sleep(1.2)
            give_back()
            root = uia.element_for(x.hwnd)

            # PivotTable-Analyse > Zeitachse einfügen... (the dialog).
            for _ in range(3):
                t = tab(root, "PivotTable-Analyse")
                if t is not None:
                    break
            b = next((e for e, d in ribbon(root) if d["name"].startswith("Zeitachse einfügen") and d["type"] == "Button"), None)
            out["pivotInsert"] = uia.describe(b)["key"] if b is not None else None
            if b is not None:
                invoke(b)
                out["insertDialog"] = read_dialog("Zeitachsen einfügen")
            # Einfügen > Filter > Zeitachse.
            tab(root, "Einfügen")
            b = next((e for e, d in ribbon(root) if d["name"] == "Zeitachse" and d["type"] == "Button"), None)
            out["insertButton"] = uia.describe(b)["key"] if b is not None else None
            # A timeline, selected through COM: the Zeitachse tab.
            sc = wb.SlicerCaches.Add2(pt, "Datum", SlicerCacheType=2)
            tl = sc.Slicers.Add(ws2, Top=10, Left=300)
            tl.Shape.Select()
            time.sleep(1.2)
            t = None
            for _ in range(3):
                t = tab(root, "Zeitachse")
                if t is not None:
                    break
            out["timelineTab"] = {"tab": uia.describe(t)["key"] if t else None, "controls": listed(ribbon(root))}
            groups = [{"name": d["name"], "key": d["key"]} for _, d in ribbon(root) if d["type"] == "Group"]
            out["timelineGroups"] = groups
            more = next((e for e, d in ribbon(root) if "Zeitachsenformat" in d["name"] and d["type"] in ("Button", "MenuItem", "SplitButton")), None)
            if more is not None:
                try:
                    more.GetCurrentPattern(uia.U.UIA_ExpandCollapsePatternId).QueryInterface(uia.U.IUIAutomationExpandCollapsePattern).Expand()
                    give_back()
                    time.sleep(0.8)
                except Exception:  # noqa: BLE001
                    invoke(more)
                items = []
                for _, d in popups() + descendants(root):
                    if d["type"] == "ListItem" and "Zeitachse" in d["name"] and d["name"] not in items:
                        items.append(d["name"])
                out["timelineStyles"] = items
                for hh in winapi.find_windows(lambda h: winapi.window_pid(h) == pid and winapi.window_class(h) == "Net UI Tool Window"):
                    winapi.user32.PostMessageW(hh, 0x0010, 0, 0)
                time.sleep(0.4)
            tl.Shape.Select()
            time.sleep(0.6)
            tab(root, "Zeitachse")
            b = next((e for e, d in ribbon(root) if d["name"].startswith("Berichtsverbindungen") and d["type"] in ("Button", "SplitButton")), None)
            if b is not None:
                invoke(b)
                out["connections"] = read_dialog("Berichtsverbindungen")
        finally:
            give_back()
            try:
                x.app.Visible = False
                x.close()
            except Exception:  # noqa: BLE001
                subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True)
    rec = json.loads(RESULT.read_text(encoding="utf-8")) if RESULT.exists() else {}
    rec["ui"] = out
    RESULT.write_text(json.dumps(rec, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(json.dumps(out, ensure_ascii=False, indent=1)[:6000])


if __name__ == "__main__":
    main()
