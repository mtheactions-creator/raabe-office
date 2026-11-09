"""Probe (D-147): the names of Excel's Musterformat entries (Zellen formatieren > Ausfüllen). In an Excel of our own
(visible, real clicks under the real-input lock, each click only after checking that our Excel is in front) the
Musterformat drop-down below its label is opened and its entries are read through UI Automation, each hovered to
read its name if the list has none.

    .venv/Scripts/python.exe probe_excel_patterns_ui.py

Result: results/2026-11-09-excel-patterns-ui.json (labels only).
"""
from __future__ import annotations

import json
import time
from pathlib import Path

from feellab import winapi
from feellab.inputlock import real_input
from probe_excel_followup_ui import UI, visible

HERE = Path(__file__).parent


def run(p: UI) -> dict:
    out: dict = {}
    p.x.sheet.Range("A1").Select()
    winapi.activate(p.x.hwnd)
    before = set(p.dialogs())
    p.key("ctrl+1")
    h = p.wait_dialog(before)
    if not h:
        return {"error": "Zellen formatieren did not open"}
    try:
        p.click(p.in_dialog(h, ("TabItem",), "Ausfüllen"))
        time.sleep(0.5)
        ctl = p.controls(h)
        out["page"] = [[c["name"], c["type"]] for c in ctl]
        lab = next((c for c in ctl if c["name"] == "Musterformat:"), None)
        if not lab:
            return out
        x, y, w, hh = lab["rect"]
        # The drop-down sits right below its label.
        if winapi.window_pid(winapi.foreground()) != p.pid:
            raise RuntimeError("our Excel is not in front")
        winapi.click(x + 40, y + hh + 14)
        time.sleep(0.8)
        pops = winapi.find_windows(lambda w2: winapi.window_pid(w2) == p.pid and w2 not in (p.x.hwnd, h) and visible(w2))
        items = []
        for w2 in pops:
            items += [[c["name"], c["type"], c["rect"]] for c in p.controls(w2)]
        out["popup"] = [[n, t] for n, t, _ in items]
        p.esc()
        time.sleep(0.4)
    finally:
        for _ in range(3):
            if h in p.dialogs():
                p.esc()
                time.sleep(0.4)
    return out


def main() -> None:
    with real_input("calc-followup probe_excel_patterns_ui", wait=900):
        p = UI()
        try:
            data = run(p)
        finally:
            p.close()
    print(json.dumps(data, ensure_ascii=False)[:3000])
    out = HERE / "results" / "2026-11-09-excel-patterns-ui.json"
    out.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
