"""Probe: Start > Bedingte Formatierung > Symbolsätze: the gallery's groups and entries (names
only), and for each entry the icon set Excel applies (IconSet index through COM, then undone).
Own visible Excel, real clicks. Result: results/<date>-excel-iconsets.json.
"""
from __future__ import annotations

import ctypes.wintypes as W
import datetime
import sys
import json
import time
from pathlib import Path

from feellab import uia, winapi
from feellab.calc_catalogue import Case, ExcelSide

HERE = Path(__file__).parent


def by_point(side: ExcelSide, l: int, t: int, w: int, h: int) -> list[dict]:
    out, seen = [], set()
    for y in range(t, t + h, 6):
        for x in range(l, l + w, 10):
            try:
                hit = uia._uia.ElementFromPoint(W.POINT(x, y))
                if not hit or hit.CurrentProcessId != side.x.pid:
                    continue
                d = uia.describe(hit)
            except Exception:  # noqa: BLE001
                continue
            key = (d["name"], tuple(d["rect"]))
            if d["name"] and d["type"] in ("ListItem", "Group", "Text", "MenuItem") and key not in seen:
                seen.add(key)
                out.append({"type": d["type"], "name": d["name"], "rect": d["rect"]})
    out.sort(key=lambda d: (d["rect"][1], d["rect"][0]))
    return out


def open_gallery(side: ExcelSide) -> tuple[list[dict], list[int]]:
    side.x.sheet.Range("A1:A5").Select()
    btn = uia.find(side.root, lambda d: d["name"] == "Bedingte Formatierung" and d["type"] in ("MenuItem", "Button", "SplitButton"), max_depth=14)
    r = uia.describe(btn)["rect"]
    winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
    time.sleep(1.0)
    item = uia.find(side.root, lambda d: d["name"] == "Symbolsätze" and d["type"] == "MenuItem", max_depth=16)
    r = uia.describe(item)["rect"]
    winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
    time.sleep(1.2)
    return by_point(side, r[0] + r[2] + 4, r[1] - 200, 420, 760), r


def main() -> None:
    side = ExcelSide()
    out: dict = {"entries": [], "applied": {}}
    try:
        side.reset(Case("i", {"A1": "1", "A2": "2", "A3": "3", "A4": "4", "A5": "5"}, [], "x"))
        winapi.activate(side.x.hwnd)
        time.sleep(0.5)
        items, _ = open_gallery(side)
        out["entries"] = [[i["type"], i["name"]] for i in items]
        # Only the gallery's own entries (the ribbon around it is read by point too).
        start = next((n for n, i in enumerate(items) if i["name"] == "Richtungen"), 0)
        items = items[start:]
        out["entries"] = [[i["type"], i["name"]] for i in items]
        print([i["name"] for i in items])
        winapi.press("esc")
        time.sleep(0.3)
        winapi.press("esc")
        time.sleep(0.3)
        tiles = [i for i in items if i["type"] == "ListItem"]
        for k, t in enumerate(tiles if "--apply" in sys.argv else []):
            again, _ = open_gallery(side)
            same = [i for i in again if i["type"] == "ListItem"]
            if k >= len(same):
                winapi.press("esc")
                continue
            r = same[k]["rect"]
            winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
            time.sleep(0.8)
            fcs = side.x.sheet.Range("A1:A5").FormatConditions
            if fcs.Count:
                fc = fcs(1)
                out["applied"][t["name"]] = {"iconSet": fc.IconSet.ID, "count": fc.IconCriteria.Count,
                                            "criteria": [[fc.IconCriteria(j).Type, fc.IconCriteria(j).Value, fc.IconCriteria(j).Operator] for j in range(2, fc.IconCriteria.Count + 1)]}
                print(t["name"], out["applied"][t["name"]])
                fcs.Delete()
            time.sleep(0.3)
    finally:
        side.close()
    (HERE / "results" / f"{datetime.date.today().isoformat()}-excel-iconsets.json").write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
