"""Probe (Step 11, part 5): where Office's ribbon puts every group and control. Reads the Start tab of our own Excel,
Word and PowerPoint (the window at the harness's place and size, so the ribbon has the same width as our apps') through
UI Automation, no input: each group's and control's type, name and box in CSS pixels (screen pixels / the display's
scale) relative to the ribbon's top left. PowerPoint only when the owner's POWERPNT.EXE is not running.
Writes results/<date>-office-ribbon-geometry.json (Microsoft's labels only).

    .venv/Scripts/python.exe probe_office_ribbon_geometry.py [excel] [word] [powerpoint]
"""
from __future__ import annotations

import datetime
import json
import sys
import time
from pathlib import Path

from feellab import uia, winapi

HERE = Path(__file__).parent
KINDS = ("Button", "SplitButton", "MenuItem", "ComboBox", "CheckBox", "Edit", "Group", "List", "DataGrid")


def ribbon_of(root):
    """The lower ribbon panel: the "Start" tab's content (a Pane/Group holding the groups)."""
    tab = uia.find(root, lambda d: d["type"] == "TabItem" and d["name"] == "Start" and d["class"] == "NetUIRibbonTab", max_depth=12)
    uia.select(tab)
    time.sleep(0.8)
    # The groups: Group elements of class NetUIChunk; the panel is their common box.
    groups = []
    stack = [(root, 0)]
    while stack:
        e, d = stack.pop()
        try:
            dc = uia.describe(e)
            if dc["type"] == "Group" and dc["class"] == "NetUIChunk":
                groups.append((e, dc))
                continue
            if d < 12:
                stack.extend((c, d + 1) for c in uia.children(e))
        except Exception:  # noqa: BLE001
            continue
    return groups


def controls(group) -> list[dict]:
    out = []
    stack = [(group, 0)]
    while stack:
        e, d = stack.pop()
        for c in uia.children(e):
            try:
                dc = uia.describe(c)
            except Exception:  # noqa: BLE001
                continue
            if dc["type"] in KINDS and dc["name"] and dc["rect"][2] > 0:
                out.append({"type": dc["type"], "name": dc["name"], "rect": dc["rect"]})
            # Split buttons: their own box is the control; galleries and lists hold many entries: keep the outer box.
            if dc["type"] not in ("SplitButton", "List", "DataGrid", "ComboBox") and d < 6:
                stack.append((c, d + 1))
    return out


# Wide enough that every group shows its full layout (Office shrinks groups to make room for its add-in groups,
# which our apps do not have): the comparison is then between full layouts on both sides.
WIDE = (0, 0, 3000, 2050)


def record(root, hwnd: int) -> dict:
    winapi.place(hwnd, *WIDE)
    time.sleep(1.0)
    scale = winapi.dpi_scale(hwnd)
    groups = ribbon_of(root)
    if not groups:
        return {"error": "no ribbon groups"}
    left = min(g[1]["rect"][0] for g in groups)
    top = min(g[1]["rect"][1] for g in groups)
    css = lambda r: [round((r[0] - left) / scale, 1), round((r[1] - top) / scale, 1), round(r[2] / scale, 1), round(r[3] / scale, 1)]  # noqa: E731
    out = {"scale": scale, "window": list(WIDE), "groups": []}
    for g, dg in sorted(groups, key=lambda x: x[1]["rect"][0]):
        out["groups"].append({"name": dg["name"], "rect": css(dg["rect"]),
                              "controls": [{**c, "rect": css(c["rect"])} for c in controls(g)]})
    return out


def main() -> None:
    which = [a for a in sys.argv[1:]] or ["excel", "word", "powerpoint"]
    path = HERE / "results" / f"{datetime.date.today().isoformat()}-office-ribbon-geometry.json"
    out = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
    if "excel" in which:
        from feellab.excel import ExcelTarget
        x = ExcelTarget(visible=True)
        try:
            time.sleep(1.5)
            out["excel"] = record(uia.element_for(x.hwnd), x.hwnd)
        finally:
            x.close()
    if "word" in which:
        from feellab.targets import WordTarget
        w = WordTarget()
        try:
            top = winapi.user32.GetAncestor(int(w.app.ActiveWindow.Hwnd), 2)
            time.sleep(1.5)
            out["word"] = record(uia.element_for(top), top)
        finally:
            w.close()
    if "powerpoint" in which:
        from feellab.powerpoint import OwnerPowerPointRunning, PowerPointTarget
        try:
            p = PowerPointTarget(visible=True)
        except OwnerPowerPointRunning as e:
            print("PowerPoint skipped:", e)
        else:
            try:
                p.new_presentation(with_window=True)
                time.sleep(1.5)
                out["powerpoint"] = record(uia.element_for(p.hwnd), p.hwnd)
            finally:
                for pr in p.own:
                    pr.Close()
                if p.app.Presentations.Count == 0:
                    p.app.Quit()
    path.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    for k, v in out.items():
        print(k, len(v.get("groups", [])), "groups", sum(len(g["controls"]) for g in v.get("groups", [])), "controls")


if __name__ == "__main__":
    main()
