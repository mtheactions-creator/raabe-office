"""Probe (Step 11, part 1): how Excel's menus close. Opens a popup (context menu of a cell, a column header, a row
header, a sheet tab; the Füllfarbe palette; the Zellenformatvorlagen gallery), then does one thing (left click into
another cell, right click there, click a header, a sheet tab, a ribbon tab, Fett, the formula bar, Esc, a click into
another program's window) and records: is a popup still open, what is selected, which sheet and ribbon tab is active,
did Fett act. Own visible Excel, real input under the lock (tell the owner first).

    .venv/Scripts/python.exe probe_excel_menuclose.py
"""
from __future__ import annotations

import datetime
import os
import json
import time
from pathlib import Path

from feellab import uia, winapi
from feellab.calc_catalogue import Case, ExcelSide
from feellab.inputlock import real_input
from feellab.popups import ForeignWindow, find_named, own_popups, pick, rect_center, selected_ribbon_tab, split_arrow, wait_popup

HERE = Path(__file__).parent

OPENERS = ["ctx_cell", "ctx_colhead", "ctx_rowhead", "ctx_sheettab", "dd_fill", "gallery_styles"]
ACTIONS = ["click_cell", "rclick_cell", "click_colhead", "click_sheettab", "click_ribbontab", "click_bold",
           "click_fxbar", "click_zoomin", "esc", "click_other_app"]


class Probe:
    def __init__(self):
        self.side = ExcelSide()
        self.pid = winapi.window_pid(self.side.x.hwnd)
        self.root = self.side.root

    def ours(self) -> bool:
        return winapi.window_pid(winapi.foreground()) == self.pid

    def reset(self) -> None:
        for _ in range(3):
            if own_popups(self.pid):
                winapi.activate(self.side.x.hwnd)
                winapi.press("esc")
                time.sleep(0.3)
        self.side.reset(Case("menuclose", {"B2": "x", "D5": "y"}, [], "x", start="B2", sheets=4))
        tab = find_named(self.root, "TabItem", "Start")
        uia.select(tab)
        winapi.activate(self.side.x.hwnd)
        time.sleep(0.4)
        P, H, U = self.side.point, self.side.header_point, self.side.ui_point
        self.pts = {
            "B2": P("B2"), "colC": H("col", "C"), "row4": H("row", 4), "tab1": U("sheettab", "Tabelle1"),
            "fill": rect_center(split_arrow(self.root, "Füllfarbe")),
            "styles": rect_center(find_named(self.root, "MenuItem", "Zellenformatvorlagen")),
        }
        # Targets of the second action, in order of preference; the first one outside every open popup is used.
        self.targets = {
            "click_cell": {"D5": P("D5"), "J5": P("J5"), "J20": P("J20"), "D30": P("D30")},
            "rclick_cell": {"F7": P("F7"), "L8": P("L8"), "L22": P("L22")},
            "click_colhead": {"E": H("col", "E"), "M": H("col", "M"), "P": H("col", "P")},
            "click_sheettab": {"Tabelle3": U("sheettab", "Tabelle3"), "Tabelle5": U("sheettab", "Tabelle5")},
            "click_ribbontab": {n: rect_center(find_named(self.root, "TabItem", n)) for n in ("Einfügen", "Formeln", "Ansicht")},
            "click_bold": {"Fett": rect_center(find_named(self.root, "Button", "Fett"))},
            "click_fxbar": {"fx": U("fxbar")},
            "click_zoomin": {"Vergrößern": rect_center(find_named(self.root, "Button", "Vergrößern"))},
        }

    def open(self, opener: str) -> bool:
        p = self.pts
        if not self.ours():
            winapi.activate(self.side.x.hwnd)
        if opener == "ctx_cell":
            winapi.right_click(*p["B2"])
        elif opener == "ctx_colhead":
            winapi.right_click(*p["colC"])
        elif opener == "ctx_rowhead":
            winapi.right_click(*p["row4"])
        elif opener == "ctx_sheettab":
            winapi.right_click(*p["tab1"])
        elif opener == "dd_fill":
            winapi.click(*p["fill"])
        elif opener == "gallery_styles":
            winapi.click(*p["styles"])
        return wait_popup(self.pid)

    def act(self, action: str, other: ForeignWindow) -> str | None:
        """Does the action; returns the target's name (None when every target lies under a popup)."""
        if not self.ours():
            raise RuntimeError("our Excel is not in front")
        if action == "esc":
            winapi.press("esc")
            return "esc"
        if action == "click_other_app":
            winapi.click(*other.point)
            return "other"
        name, pt = pick(self.targets[action], own_popups(self.pid))
        if name is None:
            return None
        (winapi.right_click if action == "rclick_cell" else winapi.click)(*pt)
        return name

    def state(self) -> dict:
        time.sleep(0.7)
        pops = own_popups(self.pid)
        st = {"popup_open": bool(pops), "popups": pops, "foreground_ours": self.ours(),
              "ribbon_tab": selected_ribbon_tab(self.root), "mode": self.side.mode()}
        # A second context menu (right click) or a formula bar in edit mode: leave it before asking COM.
        for _ in range(3):
            if own_popups(self.pid) or st["mode"] not in ("Bereit", "?"):
                winapi.activate(self.side.x.hwnd)
                winapi.press("esc")
                time.sleep(0.3)
                st["mode_after_esc"] = self.side.mode()
        app = self.side.x.app
        st["sheet"] = app.ActiveSheet.Name
        st["zoom"] = app.ActiveWindow.Zoom
        sel, act = self.side.x.selection()
        st["selection"], st["active"] = sel, act
        st["b2_bold"] = bool(app.Worksheets("Tabelle1").Range("B2").Font.Bold)
        st["d5_bold"] = bool(app.Worksheets("Tabelle1").Range("D5").Font.Bold)
        return st


def main() -> None:
    out: dict = {"date": datetime.date.today().isoformat(), "app": "Excel", "cases": {}}
    with real_input("probe_excel_menuclose"):
        pr = Probe()
        try:
            other = ForeignWindow()
        except Exception:
            pr.side.close()
            raise
        try:
            only = set(filter(None, os.environ.get("ONLY", "").split(",")))
            for opener in OPENERS:
                for action in ACTIONS:
                    if only and f"{opener}+{action}" not in only:
                        continue
                    pr.reset()
                    winapi.activate(pr.side.x.hwnd)
                    time.sleep(0.3)
                    opened = pr.open(opener)
                    key = f"{opener}+{action}"
                    if not opened:
                        out["cases"][key] = {"error": "popup did not open"}
                        print(key, "NOT OPENED")
                        continue
                    target = pr.act(action, other)
                    if target is None:
                        out["cases"][key] = {"error": "every target under the popup"}
                        print(key, "NO TARGET")
                        continue
                    st = pr.state()
                    st["target"] = target
                    out["cases"][key] = st
                    print(key, {k: st[k] for k in ("target", "popup_open", "zoom", "selection", "sheet", "ribbon_tab", "b2_bold", "mode")})
                    winapi.move(1900, 1950)
        finally:
            other.close()
            pr.side.close()
    path = HERE / "results" / f"{out['date']}-excel-menuclose.json"
    if os.environ.get("ONLY") and path.exists():
        old = json.loads(path.read_text(encoding="utf-8"))
        old["cases"].update(out["cases"])
        out = old
    path.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print("wrote", path.name)


if __name__ == "__main__":
    main()
