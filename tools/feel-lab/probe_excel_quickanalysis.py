"""Probe: Excel's Schnellanalyse (the button at the bottom right of a selection, Strg+Q): where the
button sits, its tabs, the entries of each tab, and what each entry does (clicked with real input,
the result read through COM, then undone): formulas and formats written, conditional formats,
charts, tables. Own visible Excel; labels and results only (no descriptions). Result:
results/<date>-excel-quickanalysis.json.
"""
from __future__ import annotations

import datetime
import json
import time
from pathlib import Path

from feellab import uia, winapi
from feellab.calc_catalogue import Case, ExcelSide
from feellab.excel import _pid

HERE = Path(__file__).parent
TABS = ("Formatierung", "Diagramme", "Ergebnisse", "Tabellen", "Sparklines")
SETUP = {"A1": "Name", "B1": "Mathe", "C1": "Deutsch", "A2": "Anna", "B2": "2", "C2": "3", "A3": "Ben", "B3": "3", "C3": "1",
         "A4": "Cem", "B4": "1", "C4": "2", "A5": "Dora", "B5": "4", "C5": "4"}


def tree(side: ExcelSide) -> list[dict]:
    """Everything in our Excel's other visible top-level windows (the Schnellanalyse gallery)."""
    out = []
    for h in winapi.find_windows(lambda h: _pid(h) == side.x.pid and h != side.x.hwnd):
        f = uia.element_for(h).FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
        for i in range(f.Length):
            d = uia.describe(f.GetElement(i))
            if d["name"] and len(d["name"]) < 60:
                out.append({"type": d["type"], "name": d["name"], "rect": d["rect"]})
    return out


def click(el: dict) -> None:
    r = el["rect"]
    winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)


def effect(side: ExcelSide) -> dict:
    """What the sheet holds now beyond the setup: new formulas and formats, rules, charts, tables."""
    ws = side.x.sheet
    cells = {}
    for c in ws.UsedRange.Cells:
        a = c.GetAddress(False, False)
        if c.HasFormula or (a not in SETUP and c.Formula != ""):
            cells[a] = {"formula": c.FormulaLocal, "text": c.Text, "format": c.NumberFormatLocal, "bold": bool(c.Font.Bold)}
    rules = []
    for i in range(1, ws.Cells.FormatConditions.Count + 1):
        fc = ws.Cells.FormatConditions(i)
        rules.append({"type": fc.Type, "range": fc.AppliesTo.GetAddress(False, False)})
    charts = [{"type": co.Chart.ChartType, "series": co.Chart.SeriesCollection().Count} for co in ws.ChartObjects()]
    tables = [{"range": lo.Range.GetAddress(False, False), "style": lo.TableStyle.Name if lo.TableStyle else None} for lo in ws.ListObjects]
    sparks = ws.Cells.SparklineGroups.Count
    return {"cells": cells, "rules": rules, "charts": charts, "tables": tables, "sparklines": sparks}


def open_gallery(side: ExcelSide, tab: str | None) -> list[dict]:
    side.x.sheet.Range("A1:C5").Select()
    time.sleep(0.4)
    winapi.press("ctrl+q")
    time.sleep(1.2)
    items = tree(side)
    if tab:
        t = next((i for i in items if i["name"] == tab), None)
        if t:
            click(t)
            time.sleep(0.8)
            items = tree(side)
    return items


def main() -> None:
    side = ExcelSide()
    out: dict = {"tabs": {}, "effects": {}}
    try:
        side.reset(Case("q", SETUP, [], "x"))
        winapi.activate(side.x.hwnd)
        time.sleep(0.5)
        x1, y1 = side._point("A1", "center")
        x2, y2 = side._point("C5", "center")
        winapi.drag(x1, y1, x2, y2)
        time.sleep(1.0)
        br = side._point("C5", "br")
        btn = uia.find(side.root, lambda d: d["type"] == "Button" and d["name"].startswith("Schnellanalyse"), max_depth=12)
        if btn is not None:
            d = uia.describe(btn)
            out["button"] = {"name": d["name"], "rect": d["rect"], "cellCorner": [round(br[0]), round(br[1])]}
        for tab in TABS:
            items = open_gallery(side, tab)
            names = [i for i in items if i["type"] in ("ListItem", "Button", "MenuItem") and i["name"] not in TABS and i["name"] != "Schnellanalyse"]
            out["tabs"][tab] = [[i["type"], i["name"]] for i in names]
            print(tab, [i["name"] for i in names])
            winapi.press("esc")
            time.sleep(0.5)
            for k, it in enumerate(names):
                if it["type"] != "ListItem":
                    continue
                items = open_gallery(side, tab)
                same = [i for i in items if i["type"] == "ListItem" and i["name"] not in TABS]
                if k >= len(same):
                    winapi.press("esc")
                    continue
                click(same[k])
                time.sleep(1.5)
                e = effect(side)
                out["effects"][f"{tab}/{k}/{it['name']}"] = e
                print(tab, it["name"], json.dumps(e, ensure_ascii=False)[:300])
                winapi.press("esc")
                time.sleep(0.3)
                winapi.press("ctrl+z")
                time.sleep(0.8)
                # Make sure the sheet is back as it was (undo can leave a chart behind).
                after = effect(side)
                if after["cells"] or after["rules"] or after["charts"] or after["tables"] or after["sparklines"]:
                    side.reset(Case("q", SETUP, [], "x"))
                    winapi.activate(side.x.hwnd)
                    time.sleep(0.5)
    finally:
        side.close()
    (HERE / "results" / f"{datetime.date.today().isoformat()}-excel-quickanalysis.json").write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
