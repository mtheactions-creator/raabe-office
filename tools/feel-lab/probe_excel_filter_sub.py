"""Probe: the Zahlenfilter / Textfilter submenus of Excel's AutoFilter panel and the dialog behind
Benutzerdefinierter Filter... (labels and the operator list), for RaabeCalc's custom filters.
Own visible Excel, real clicks; labels only (no help texts). Then, invisibly through COM, which
rows Excel shows for a set of criteria (numbers, texts, wildcards, Top 10, average), the rules
RaabeCalc must follow. Result: results/<date>-excel-filter-sub.json.
"""
from __future__ import annotations

import datetime
import sys
import json
import time
from pathlib import Path

import ctypes.wintypes as W

from feellab import uia, winapi
from feellab.calc_catalogue import Case, ExcelSide
from feellab.excel import _pid
from probe_word_context import popups

HERE = Path(__file__).parent
KINDS = ("MenuItem", "Button", "CheckBox", "Edit", "ListItem", "Text", "ComboBox", "RadioButton", "Group", "Window")


def items_of(hs: list) -> list[dict]:
    out = []
    for h in hs:
        f = uia.element_for(h).FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
        for i in range(f.Length):
            d = uia.describe(f.GetElement(i))
            if d["name"] and len(d["name"]) < 70 and d["type"] in KINDS:
                out.append({"name": d["name"], "type": d["type"], "rect": d["rect"]})
    return out


DATA = {"A1": "Fach", "A2": "Deutsch", "A3": "Mathe", "A4": "Mathematik", "A5": "Physik", "A6": "Chemie", "A7": "Mathe-AG",
        "B1": "Note", "B2": "2", "B3": "3", "B4": "1", "B5": "5", "B6": "4", "B7": "2,5"}

# (column, Criteria1, Operator, Criteria2) through COM: xlAnd 1, xlOr 2, xlTop10Items 3, xlBottom10Items 4,
# xlTop10Percent 5, xlFilterDynamic 11 (Criteria1 33 above average, 34 below).
CRITERIA = [
    (2, ">2", None, None), (2, ">=2,5", None, None), (2, "<>2", None, None), (2, ">1", 1, "<4"), (2, "<2", 2, ">4"),
    (2, "3", 3, None), (2, "2", 4, None), (2, "40", 5, None), (2, 33, 11, None), (2, 34, 11, None),
    (1, "Mathe*", None, None), (1, "*he*", None, None), (1, "*e", None, None), (1, "<>*the*", None, None), (1, "mathe", None, None),
    (1, ">M", None, None), (1, "?athe", None, None), (1, "=Mathe", None, None),
]


def main() -> None:
    side = ExcelSide()
    out: dict = {"menus": {}, "rules": []}
    try:
        side.reset(Case("f", DATA, [], "x"))
        ws = side.x.sheet
        ws.Range("A1:B7").AutoFilter(1)
        winapi.activate(side.x.hwnd)
        time.sleep(0.5)
        if "--colors" in sys.argv:
            # Coloured cells for the colour menus: fills and a font colour.
            ws.Range("B3").Interior.Color = 0x00FFFF
            ws.Range("B5").Interior.Color = 0x0000FF
            ws.Range("B6").Font.Color = 0x0000FF
        entries = (("B", "colorFilter", "Nach Farbe filtern"), ("B", "colorSort", "Nach Farbe sortieren")) if "--colors" in sys.argv else (("B", "number", "Zahlenfilter"), ("A", "text", "Textfilter"))
        for col, kind, entry in entries:
            x, y = side._point(f"{col}1", "br")
            before = set(popups())
            winapi.click(x - 10, y - 10)
            time.sleep(1.2)
            panel = [h for h in popups() if h not in before]
            it = next((i for i in items_of(panel) if i["name"] == entry), None)
            if not it:
                print("no", entry)
                winapi.press("esc")
                continue
            r = it["rect"]
            before2 = set(popups())
            winapi.click(r[0] + r[2] // 2, r[1] + r[3] // 2)
            time.sleep(1.2)
            # The submenu: read by point to the right of the entry (nested menus are not found as
            # windows of their own), only elements of our Excel.
            subitems, seen = [], set()
            for yy in range(r[1] - 40, r[1] + 520, 7):
                for xx in (r[0] + r[2] + 40, r[0] + r[2] + 120):
                    try:
                        hit = uia._uia.ElementFromPoint(W.POINT(xx, yy))
                        if not hit or hit.CurrentProcessId != side.x.pid:
                            continue
                        d = uia.describe(hit)
                    except Exception:  # noqa: BLE001
                        continue
                    if d["name"] and d["type"] in ("MenuItem", "Text", "Group", "ListItem", "Button") and d["rect"][0] > r[0] + r[2] - 5 and d["name"] not in seen:
                        seen.add(d["name"])
                        subitems.append({"name": d["name"], "type": d["type"], "rect": d["rect"]})
            subitems.sort(key=lambda d: d["rect"][1])
            out["menus"][kind] = [i["name"] for i in subitems]
            print(kind, out["menus"][kind])
            last = next((i for i in subitems if i["name"].startswith("Benutzerdefiniert")), None)
            if last and "--colors" not in sys.argv:
                r = last["rect"]
                winapi.click(r[0] + r[2] // 2, r[1] + r[3] // 2)
                time.sleep(1.5)
                # The dialog is a top-level window of our Excel.
                hs = winapi.find_windows(lambda h: _pid(h) == side.x.pid and "AutoFilter" in winapi.window_title(h))
                dlg = uia.element_for(hs[0]) if hs else None
                if dlg is not None:
                    f = dlg.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
                    names = []
                    for i in range(f.Length):
                        d = uia.describe(f.GetElement(i))
                        if d["name"] and len(d["name"]) < 70 and d["type"] in KINDS:
                            names.append([d["type"], d["name"]])
                    # The operator list: open the focused (first) operator box and read its entries.
                    winapi.press("alt+down")
                    time.sleep(0.8)
                    lbs = winapi.find_windows(lambda h: _pid(h) == side.x.pid and winapi.window_class(h) == "ComboLBox")
                    ops = [i["name"] for i in items_of(lbs) if i["type"] == "ListItem"]
                    winapi.press("esc")
                    time.sleep(0.3)
                    out["menus"][kind + "Dialog"] = {"title": uia.describe(dlg)["name"], "controls": names, "operators": ops}
                    print(kind, "dialog", uia.describe(dlg)["name"], names)
                winapi.press("esc")
                time.sleep(0.5)
            winapi.press("esc")
            time.sleep(0.3)
            winapi.press("esc")
            time.sleep(0.5)
        # The rules, invisibly: which rows stay visible for each criterion.
        for col, c1, op, c2 in ([] if "--menus" in sys.argv or "--colors" in sys.argv else CRITERIA):
            ws.ShowAllData() if ws.FilterMode else None
            kw = {"Field": col, "Criteria1": c1}
            if op is not None:
                kw["Operator"] = op
            if c2 is not None:
                kw["Criteria2"] = c2
            ws.Range("A1:B7").AutoFilter(**kw)
            shown = [r for r in range(2, 8) if not ws.Rows(r).Hidden]
            out["rules"].append({"col": col, "c1": c1, "op": op, "c2": c2, "shown": shown})
            print(col, c1, op, c2, shown)
        ws.ShowAllData() if ws.FilterMode else None
    finally:
        side.close()
    path = HERE / "results" / f"{datetime.date.today().isoformat()}-excel-filter-sub.json"
    if "--colors" in sys.argv:
        path = HERE / "results" / f"{datetime.date.today().isoformat()}-excel-filter-colors.json"
    elif "--menus" in sys.argv and path.exists():
        out["rules"] = json.loads(path.read_text(encoding="utf-8"))["rules"]
    path.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
