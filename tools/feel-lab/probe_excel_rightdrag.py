"""Probe: the menus Excel shows after dragging with the right mouse button: the fill handle
(Zellen kopieren, Datenreihe ausfüllen …) and the selection's border (Hierher verschieben …).
Labels and enabled state only, read by point from our own visible Excel; Esc closes the menu
without doing anything. Result: results/<date>-excel-rightdrag.json.
"""
from __future__ import annotations

import ctypes.wintypes as W
import datetime
import json
import time
from pathlib import Path

from feellab import uia, winapi
from feellab.calc_catalogue import Case, ExcelSide

HERE = Path(__file__).parent


def right_drag(x1: float, y1: float, x2: float, y2: float) -> None:
    winapi.move(x1, y1)
    time.sleep(0.1)
    winapi._send(winapi._mouse(x1, y1, 0x0008))  # right down
    time.sleep(0.1)
    for i in range(1, 21):
        winapi.move(x1 + (x2 - x1) * i / 20, y1 + (y2 - y1) * i / 20)
        time.sleep(0.015)
    time.sleep(0.1)
    winapi._send(winapi._mouse(x2, y2, 0x0010))  # right up


def menu_at(side: ExcelSide, x: float, y: float) -> list[dict]:
    """Menu entries near the drop point (the menu opens there), our Excel only."""
    out, seen = [], set()
    for yy in range(int(y) - 30, int(y) + 560, 6):
        for xx in (int(x) + 40, int(x) + 120, int(x) - 120, int(x) - 200):
            try:
                hit = uia._uia.ElementFromPoint(W.POINT(xx, yy))
                if not hit or hit.CurrentProcessId != side.x.pid:
                    continue
                d = uia.describe(hit)
            except Exception:  # noqa: BLE001
                continue
            if d["type"] == "MenuItem" and d["name"] and d["name"] not in seen:
                seen.add(d["name"])
                out.append({"name": d["name"], "enabled": d["enabled"], "y": d["rect"][1]})
    out.sort(key=lambda d: d["y"])
    return [{"name": d["name"], "enabled": d["enabled"]} for d in out]


def main() -> None:
    side = ExcelSide()
    out: dict = {}
    try:
        side.reset(Case("r", {"A1": "1", "A2": "2"}, [], "x"))
        winapi.activate(side.x.hwnd)
        time.sleep(0.5)
        side.x.sheet.Range("A1:A2").Select()
        time.sleep(0.3)
        x1, y1 = side._point("A2", "br")
        x2, y2 = side._point("A6", "br")
        right_drag(x1 - 1, y1 - 1, x2 - 1, y2 - 1)
        time.sleep(1.0)
        out["fillHandle"] = menu_at(side, x2, y2)
        print("fill", [d["name"] for d in out["fillHandle"]])
        winapi.press("esc")
        time.sleep(0.5)
        side.x.sheet.Range("A1:A2").Select()
        time.sleep(0.3)
        x1, y1 = side._point("A1", "right")
        x2, y2 = side._point("C4", "right")
        right_drag(x1 - 1, y1, x2 - 1, y2)
        time.sleep(1.0)
        out["border"] = menu_at(side, x2, y2)
        print("border", [d["name"] for d in out["border"]])
        winapi.press("esc")
        time.sleep(0.5)
    finally:
        side.close()
    (HERE / "results" / f"{datetime.date.today().isoformat()}-excel-rightdrag.json").write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
