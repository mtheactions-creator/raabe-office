"""Probe: Excel's context menus (right click) on a cell, a column header, a row header and a
sheet tab, with their submenus' names. Own Excel instance; takes over mouse and screen briefly.
Writes the items (name, enabled) to results/<date>-excel-context.json (Microsoft's labels only,
as the ribbon's; no descriptions).

    .venv/Scripts/python.exe probe_excel_context.py
"""
from __future__ import annotations

import datetime
import json
import time
from pathlib import Path

from feellab import uia, winapi
from feellab.calc_catalogue import Case, ExcelSide
from probe_word_context import items_of, popups

HERE = Path(__file__).parent


def menu() -> list[dict]:
    end = time.time() + 3
    while time.time() < end:
        hs = popups()
        if hs:
            time.sleep(0.4)
            hs = popups()
            rects = {h: winapi.window_rect(h) for h in hs}
            m = max(hs, key=lambda h: rects[h][3] - rects[h][1])
            return [{"name": d["name"], "type": d["type"], "enabled": d["enabled"], "key": d["key"]} for _, d in items_of(m)]
        time.sleep(0.15)
    return []


def main() -> None:
    side = ExcelSide()
    out: dict[str, list] = {}
    try:
        side.reset(Case("ctx", {"A1": "1", "A2": "2", "B1": "Text"}, [], "x"))
        winapi.activate(side.x.hwnd)
        time.sleep(0.5)
        targets = {
            "cell": side.point("B2"),
            "column": side.header_point("col", "C"),
            "row": side.header_point("row", 4),
        }
        # The sheet tab: from the window's automation tree.
        tab = uia.find(side.root, lambda d: d["type"] == "TabItem" and d["name"] == "Tabelle1", max_depth=10)
        if tab:
            r = uia.describe(tab)["rect"]
            targets["sheettab"] = (r[0] + r[2] / 2, r[1] + r[3] / 2)
        for name, (x, y) in targets.items():
            if name != "sheettab":
                winapi.click(x, y)
                time.sleep(0.2)
            winapi.right_click(x, y)
            out[name] = menu()
            print(name, [i["name"] for i in out[name]])
            winapi.press("esc")
            time.sleep(0.3)
            winapi.press("esc")
            time.sleep(0.3)
    finally:
        side.close()
    path = HERE / "results" / f"{datetime.date.today().isoformat()}-excel-context.json"
    path.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
