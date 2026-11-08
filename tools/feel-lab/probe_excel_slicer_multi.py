"""Records how a Datenschnitt selects several items in Excel 365 (D-122), by hand in our own visible Excel (real input,
only under the real-input lock, keys and clicks only while our Excel is in front). A table with a column Fach (five
subjects) and a slicer on it; each case starts from "all items selected" (ClearManualFilter), then clicks, Strg+clicks,
Umschalt+clicks, drags over the buttons, Alt+S (Mehrfachauswahl) and the Mehrfachauswahl button. After each case the
items' Selected state (COM) and whether Mehrfachauswahl is on (UI Automation) are recorded.

    .venv/Scripts/python.exe probe_excel_slicer_multi.py

Result: results/2026-11-09-excel-slicer-multi.json.
"""
from __future__ import annotations

import json
import time
from pathlib import Path

from feellab import uia, winapi
from feellab.excel import ExcelTarget
from feellab.inputlock import real_input

HERE = Path(__file__).parent
RESULT = HERE / "results" / "2026-11-09-excel-slicer-multi.json"
DATA = [["Name", "Fach"], ["Anna", "Mathe"], ["Ben", "Deutsch"], ["Cem", "Englisch"], ["Dora", "Bio"], ["Emil", "Kunst"], ["Finn", "Mathe"]]
ITEMS = ["Bio", "Deutsch", "Englisch", "Kunst", "Mathe"]

# (name, steps) with steps ("click", i, mods) / ("drag", i, j, mods) / ("keys", combo) / ("multi",) the Mehrfachauswahl button.
CASES = [
    ("click", [("click", 1, "")]),
    ("click_twice_same", [("click", 1, ""), ("click", 1, "")]),
    ("ctrl_from_all", [("click", 2, "ctrl")]),
    ("click_then_ctrl", [("click", 0, ""), ("click", 3, "ctrl")]),
    ("ctrl_off_last", [("click", 0, ""), ("click", 0, "ctrl")]),
    ("shift_range", [("click", 1, ""), ("click", 3, "shift")]),
    ("shift_from_all", [("click", 3, "shift")]),
    ("drag_range", [("drag", 0, 2, "")]),
    ("drag_up", [("drag", 4, 2, "")]),
    ("ctrl_drag_add", [("click", 0, ""), ("drag", 2, 3, "ctrl")]),
    ("alt_s_click", [("click", 0, ""), ("keys", "alt+s"), ("click", 2, "")]),
    ("multi_from_all", [("multi",), ("click", 1, "")]),
    ("multi_two", [("multi",), ("click", 1, ""), ("click", 3, "")]),
    ("multi_all_off", [("multi",)] + [("click", i, "") for i in range(5)]),
    ("multi_shift", [("multi",), ("click", 0, ""), ("click", 3, "shift")]),
]


def main() -> None:
    out: dict = {"items": ITEMS, "cases": {}}
    with real_input("calc-misc probe_excel_slicer_multi"):
        x = ExcelTarget(visible=True)
        app = x.app
        pid = x.pid
        root = uia.element_for(x.hwnd)

        def front() -> None:
            for _ in range(10):
                try:
                    winapi.activate(x.hwnd, 0.2)
                except RuntimeError:
                    pass
                if winapi.window_pid(winapi.foreground()) == pid:
                    return
                time.sleep(0.3)
            raise RuntimeError("our Excel is not in front")

        try:
            ws = app.ActiveSheet
            for r, row in enumerate(DATA, 1):
                for c, v in enumerate(row, 1):
                    ws.Cells(r, c).Value = v
            lo = ws.ListObjects.Add(1, ws.Range(f"A1:B{len(DATA)}"), None, 1)
            sc = app.ActiveWorkbook.SlicerCaches.Add2(lo, "Fach")
            sl = sc.Slicers.Add(ws, Top=30, Left=250, Width=144, Height=220)
            win = app.ActiveWindow
            px = lambda pt: pt * 96 / 72 * 1.5  # noqa: E731
            x0, y0 = win.PointsToScreenPixelsX(0), win.PointsToScreenPixelsY(0)
            # The buttons: UI Automation names them by their text; else the measured geometry (D-092).
            time.sleep(1)
            btn = {}
            for name in ITEMS:
                el = uia.find(root, lambda d, n=name: d["name"] == n and d["type"] in ("ListItem", "Button", "CheckBox", "DataItem", "Custom"), 16)
                if el is not None:
                    r = uia.describe(el)["rect"]
                    btn[name] = (r[0] + r[2] / 2, r[1] + r[3] / 2)
            out["buttonsFrom"] = "uia" if len(btn) == 5 else "geometry"
            if len(btn) < 5:
                for i, name in enumerate(ITEMS):
                    btn[name] = (x0 + px(250 + 72), y0 + px(30 + 21 + 7 + 10 + i * (19.84 + 3.6)))
            out["buttons"] = btn
            multi_btn = uia.find(root, lambda d: d["name"].startswith("Mehrfachauswahl") and d["type"] in ("Button", "CheckBox"), 16)
            out["multiButton"] = uia.describe(multi_btn) if multi_btn is not None else None

            def multi_on():
                if multi_btn is None:
                    return None
                try:
                    return bool(multi_btn.GetCurrentPattern(uia.U.UIA_TogglePatternId).QueryInterface(uia.U.IUIAutomationTogglePattern).CurrentToggleState)
                except Exception:  # noqa: BLE001
                    return None

            def selected() -> list[str]:
                return [sc.SlicerItems(i).Name for i in range(1, sc.SlicerItems.Count + 1) if sc.SlicerItems(i).Selected]

            for name, steps in CASES:
                sc.ClearManualFilter()
                ws.Range("D2").Select()
                if multi_on():
                    front()
                    r = uia.describe(multi_btn)["rect"]
                    winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
                    time.sleep(0.3)
                front()
                trail = []
                for st in steps:
                    front()
                    if st[0] == "click":
                        bx, by = btn[ITEMS[st[1]]]
                        winapi.click(bx, by, 1, st[2])
                    elif st[0] == "drag":
                        (ax, ay), (bx, by) = btn[ITEMS[st[1]]], btn[ITEMS[st[2]]]
                        if st[3]:
                            winapi.key_down(winapi.vk_for(st[3]))
                        winapi.drag(ax, ay, bx, by, steps=15)
                        if st[3]:
                            winapi.key_up(winapi.vk_for(st[3]))
                    elif st[0] == "keys":
                        winapi.press(st[1])
                    elif st[0] == "multi":
                        r = uia.describe(multi_btn)["rect"]
                        winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
                    time.sleep(0.5)
                    trail.append({"selected": selected(), "multi": multi_on()})
                out["cases"][name] = {"steps": steps, "trail": trail}
                print(name, trail[-1], flush=True)
            # Mehrfachauswahl after the slicer was left and chosen again.
            sc.ClearManualFilter()
            front()
            r = uia.describe(multi_btn)["rect"] if multi_btn is not None else None
            if r:
                winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
                time.sleep(0.3)
                front()
                cx, cy = x0 + px(40), y0 + px(200)
                winapi.click(cx, cy)
                time.sleep(0.3)
                bx, by = btn[ITEMS[1]]
                winapi.click(bx, by)
                time.sleep(0.4)
                out["multiAfterLeaving"] = {"multi": multi_on(), "selected": selected()}
        finally:
            x.close()
    RESULT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
