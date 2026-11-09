"""Records Excel's Namenfeld / Bearbeitungsleiste splitter and the formula bar's expansion (D-120) by hand in our own visible
Excel (real input, only under the real-input lock, keys and clicks only while our Excel is in front):
- the handle between Namenfeld and the buttons: where it is, the mouse pointer over it, the Namenfeld's width after
  dragging it far left (smallest) and far right (largest), then dragged back to where it was;
- the expand button at the formula bar's right end: Application.FormulaBarHeight before and after a click;
- Strg+Umschalt+U in German Excel (formula bar or underline?);
- dragging the formula bar's lower edge: FormulaBarHeight after 1, 2.5 and 6 lines' worth of movement.
FormulaBarHeight and the Namenfeld width are Excel's settings for the user: both are put back.
Sizes in physical pixels at this PC's 150 % (CSS px = / 1.5).

    .venv/Scripts/python.exe probe_excel_formulabar.py

Result: results/2026-11-09-excel-formulabar.json.
"""
from __future__ import annotations

import ctypes
import json
import time
from pathlib import Path

from feellab import uia, winapi
from feellab.excel import ExcelTarget
from feellab.inputlock import real_input

HERE = Path(__file__).parent
RESULT = HERE / "results" / "2026-11-09-excel-formulabar.json"


class CURSORINFO(ctypes.Structure):
    _fields_ = [("cbSize", ctypes.c_uint32), ("flags", ctypes.c_uint32), ("hCursor", ctypes.c_void_p), ("x", ctypes.c_long), ("y", ctypes.c_long)]


CURSORS = {32512: "arrow", 32513: "ibeam", 32644: "sizewe", 32645: "sizens", 32649: "hand", 32646: "sizeall"}
LOADED = {}
for cid, name in CURSORS.items():
    LOADED[ctypes.windll.user32.LoadCursorW(None, ctypes.c_void_p(cid))] = name


def cursor() -> str:
    ci = CURSORINFO()
    ci.cbSize = ctypes.sizeof(CURSORINFO)
    ctypes.windll.user32.GetCursorInfo(ctypes.byref(ci))
    return LOADED.get(ci.hCursor, f"other:{ci.hCursor}")


def main() -> None:
    out: dict = {"scale": 1.5}
    with real_input("calc-misc probe_excel_formulabar"):
        x = ExcelTarget(visible=True)
        app = x.app
        pid = x.pid
        root = uia.element_for(x.hwnd)
        h0 = app.FormulaBarHeight

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

        def rect(name: str, typ: str) -> list[int]:
            return uia.describe(uia.find(root, lambda d: d["name"] == name and d["type"] == typ, 14))["rect"]

        try:
            app.ActiveSheet.Range("A1").Value = chr(10).join(f"Zeile {i}" for i in range(1, 9))
            app.ActiveSheet.Range("A1").Select()
            front()
            nb = rect("Namenfeld", "ComboBox")
            fx = rect("Bearbeitungsleiste", "Edit")
            out["namebox"], out["formulabar"] = nb, fx
            cy = nb[1] + nb[3] // 2
            # The handle: the pointer changes over it.
            scan = []
            for dx in range(-4, 30):
                front()
                winapi.move(nb[0] + nb[2] + dx, cy)
                time.sleep(0.08)
                scan.append((dx, cursor()))
            out["handleScan"] = scan
            hits = [dx for dx, c in scan if c == "sizewe"]
            out["handleCursor"] = "sizewe" if hits else sorted({c for _, c in scan})
            hx = nb[0] + nb[2] + (hits[len(hits) // 2] if hits else 8)
            out["handleX"] = hx - (nb[0] + nb[2])
            widths = {}
            for label, to in (("min", hx - 600), ("max", hx + 1600)):
                front()
                winapi.drag(hx, cy, to, cy, steps=25)
                time.sleep(0.5)
                r = rect("Namenfeld", "ComboBox")
                widths[label] = r[2]
                # The handle moves with the box.
                hx = r[0] + r[2] + (out["handleX"])
            out["nameboxWidth"] = {"start": nb[2], **widths}
            # Back to the width it had.
            front()
            winapi.drag(hx, cy, nb[0] + nb[2] + out["handleX"], cy, steps=25)
            time.sleep(0.4)
            out["nameboxBack"] = rect("Namenfeld", "ComboBox")[2]
            # A double-click on the handle.
            front()
            hx = nb[0] + nb[2] + out["handleX"]
            winapi.click(hx, cy, 2)
            time.sleep(0.4)
            out["nameboxAfterDoubleClick"] = rect("Namenfeld", "ComboBox")[2]
            if out["nameboxAfterDoubleClick"] != nb[2]:
                r = rect("Namenfeld", "ComboBox")
                front()
                winapi.drag(r[0] + r[2] + out["handleX"], cy, nb[0] + nb[2] + out["handleX"], cy, steps=25)
                time.sleep(0.3)
                out["nameboxBack2"] = rect("Namenfeld", "ComboBox")[2]

            # The expand button at the right end of the formula bar.
            fx = rect("Bearbeitungsleiste", "Edit")
            scan = []
            for dx in range(2, 40, 2):
                front()
                winapi.move(fx[0] + fx[2] - dx, fx[1] + fx[3] // 2)
                time.sleep(0.06)
                scan.append((dx, cursor()))
            out["expandScan"] = scan
            out["heightStart"] = app.FormulaBarHeight
            front()
            winapi.click(fx[0] + fx[2] - 28, nb[1] + nb[3] // 2)
            time.sleep(0.6)
            out["heightAfterExpandClick"] = app.FormulaBarHeight
            out["formulabarExpanded"] = rect("Bearbeitungsleiste", "Edit")
            front()
            winapi.click(fx[0] + fx[2] - 28, nb[1] + nb[3] // 2)
            time.sleep(0.6)
            out["heightAfterSecondClick"] = app.FormulaBarHeight
            # Strg+Umschalt+U with the grid active.
            front()
            winapi.press("esc")
            winapi.press("ctrl+shift+u")
            time.sleep(0.6)
            out["ctrlShiftU"] = {"height": app.FormulaBarHeight, "underline": app.ActiveSheet.Range("A1").Font.Underline}
            front()
            winapi.press("ctrl+shift+u")
            time.sleep(0.6)
            out["ctrlShiftU2"] = {"height": app.FormulaBarHeight, "underline": app.ActiveSheet.Range("A1").Font.Underline}
            # The lower edge, dragged.
            fx = rect("Bearbeitungsleiste", "Edit")
            bottom = fx[1] + fx[3]
            scan = []
            for dy in range(-6, 10):
                front()
                winapi.move(fx[0] + 400, bottom + dy)
                time.sleep(0.06)
                scan.append((dy, cursor()))
            out["edgeScan"] = scan
            edge = [dy for dy, c in scan if c == "sizens"]
            ey = bottom + (edge[len(edge) // 2] if edge else 2)
            drags = {}
            for label, dy in (("one", 28), ("twoAndHalf", 70), ("six", 168), ("back", 0)):
                front()
                fx2 = rect("Bearbeitungsleiste", "Edit")
                ey = fx2[1] + fx2[3] + (edge[len(edge) // 2] if edge else 2)
                target = (fx[1] + fx[3] + dy + (edge[len(edge) // 2] if edge else 2))
                winapi.drag(fx[0] + 400, ey, fx[0] + 400, target, steps=20)
                time.sleep(0.5)
                drags[label] = {"height": app.FormulaBarHeight, "rect": rect("Bearbeitungsleiste", "Edit")}
            out["edgeDrags"] = drags
        finally:
            try:
                app.FormulaBarHeight = h0
                app.ActiveSheet.Range("A1").Font.Underline = -4142
            except Exception:  # noqa: BLE001
                pass
            x.close()
    RESULT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(json.dumps({k: v for k, v in out.items() if "Scan" not in k}, ensure_ascii=False))


if __name__ == "__main__":
    main()
