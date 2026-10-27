"""Probe: Excel's AutoFilter dropdown (the panel under a filter button): its entries and sizes,
for RaabeCalc's filter panel (D-063). Own visible Excel, real clicks. Labels only.
Result: results/<date>-excel-filter.json.
"""
from __future__ import annotations

import datetime
import json
import time
from pathlib import Path

from feellab import uia, winapi
from feellab.calc_catalogue import Case, ExcelSide
from probe_word_context import popups

HERE = Path(__file__).parent


def main() -> None:
    side = ExcelSide()
    out: dict = {}
    try:
        side.reset(Case("f", {"A1": "Fach", "A2": "Deutsch", "A3": "Mathe", "A4": "Mathe", "A5": "Physik", "B1": "Note", "B2": "2", "B3": "3", "B4": "1", "B5": "2"}, [], "x"))
        side.x.sheet.Range("A1:B5").AutoFilter(1)
        winapi.activate(side.x.hwnd)
        time.sleep(0.5)
        for col, name in (("A", "text"), ("B", "number")):
            x, y = side._point(f"{col}1", "br")
            before = set(popups())
            winapi.click(x - 10, y - 10)
            time.sleep(1.2)
            hs = [h for h in popups() if h not in before]
            items = []
            rect = None
            for h in hs:
                l, t, r, b = winapi.window_rect(h)
                rect = [l, t, r - l, b - t]
                f = uia.element_for(h).FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
                for i in range(f.Length):
                    d = uia.describe(f.GetElement(i))
                    if d["name"] and len(d["name"]) < 60 and d["type"] in ("MenuItem", "Button", "CheckBox", "Edit", "TreeItem", "ListItem", "Text", "SplitButton"):
                        items.append({"name": d["name"], "type": d["type"], "rect": d["rect"]})
            out[name] = {"window": rect, "items": items}
            print(name, rect, [i["name"] for i in items])
            winapi.press("esc")
            time.sleep(0.5)
    finally:
        side.close()
    (HERE / "results" / f"{datetime.date.today().isoformat()}-excel-filter.json").write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
