"""Probe: the submenus of Word's equation context menu (matrix: Einfügen, Löschen,
Matrixausrichtung, Spaltenausrichtung; Formeloptionen: Ausrichtung) and the entries for
superscripts, accents and functions (right-click on a grid of points over the equation).
Own visible Word instance; takes over mouse and screen. Adds to
results/<date>-word-equation-context.json.
"""
from __future__ import annotations

import json
import sys
import time

from feellab import uia, winapi
from probe_word_equation_context import CAP, OUT, menu_items
from probe_word_equation_gallery import GalleryWord
from probe_word_picture import popups

BASE = {"Ausschneiden", "Kopieren", "Einfügen", "Copy Link to Location...", "Professionell", "Linear", "Als neue Formel speichern...", "Schriftart...", "Absatz..."}


def grid_menus(w: GalleryWord, typed: str, name: str) -> list[list[str]]:
    """Right-clicks on a grid over the equation; the distinct menus."""
    w.fresh()
    winapi.type_text(typed, gap=0.05)
    time.sleep(0.4)
    win = w.d.Windows(1)
    left, top, width, height = win.GetPoint(0, 0, 0, 0, w.d.OMaths(1).Range)
    menus: dict[str, list[str]] = {}
    for gy in range(4, max(5, height - 2), 6):
        for gx in range(2, max(3, width), 5):
            before = popups()
            winapi.right_click(left + gx, top + gy)
            time.sleep(0.8)
            names = [x["name"] for x in menu_items(before)]
            extra = [n for n in names if n not in BASE]
            key = " | ".join(extra)
            if extra and key not in menus:
                menus[key] = extra
                print(name, gx, gy, extra, flush=True)
            winapi.press("esc")
            time.sleep(0.2)
    return list(menus.values())


def submenu(w: GalleryWord, typed: str, entry: str, at: tuple[float, float]) -> list[str]:
    """Opens the context menu at a point (relative to the equation) and hovers an entry."""
    w.fresh()
    winapi.type_text(typed, gap=0.05)
    time.sleep(0.4)
    win = w.d.Windows(1)
    left, top, width, height = win.GetPoint(0, 0, 0, 0, w.d.OMaths(1).Range)
    before = popups()
    winapi.right_click(left + at[0] * width, top + at[1] * height)
    time.sleep(1.0)
    items = menu_items(before)
    # "Einfügen" is also the paste entry: the submenu is the last entry of that name.
    hit = next((x for x in reversed(items) if x["name"] == entry), None)
    menus = [h for h in popups() if h not in before]
    if not hit or not menus:
        winapi.press("esc")
        return []
    l, t, _, _ = winapi.window_rect(menus[0])
    known = popups()
    winapi.click(l + hit["rect"][0] + hit["rect"][2] / 2, t + hit["rect"][1] + hit["rect"][3] / 2)
    time.sleep(1.2)
    sub = [x["name"] for x in menu_items(known)]
    for h in [h for h in popups() if h not in known]:
        a, b, c, d = winapi.window_rect(h)
        winapi.save_region_png(str(CAP / f"submenu-{entry}.png"), a, b, c - a, d - b)
        break
    winapi.press("esc")
    time.sleep(0.2)
    winapi.press("esc")
    return sub


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    res = json.loads(OUT.read_text(encoding="utf-8")) if OUT.exists() else {}
    w = GalleryWord()
    try:
        for name, typed in (("superscript", "x^2 "), ("accent", "x\\hat "), ("function", "sin x "), ("subsup", "x_1^2 "), ("matrix", "\\matrix(a&b@c&d) ")):
            if "sub" in sys.argv:
                break
            res[f"grid {name}"] = grid_menus(w, typed, name)
        for entry in ("Einfügen",) if "insert" in sys.argv else ("Einfügen", "Löschen", "Matrixausrichtung", "Spaltenausrichtung"):
            res[f"matrix {entry}"] = submenu(w, "\\matrix(a&b@c&d) ", entry, (0.3, 0.3))
            print("matrix", entry, res[f"matrix {entry}"], flush=True)
        # Formeloptionen > Ausrichtung.
        w.fresh()
        winapi.type_text("x=1", gap=0.05)
        time.sleep(0.4)
        el = uia.find(w.root, lambda d: d["type"] == "MenuItem" and d["name"] == "Formeloptionen", max_depth=20)
        if el is not None:
            r = uia.describe(el)["rect"]
            before = popups()
            winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
            time.sleep(1.2)
            items = menu_items(before)
            hit = next((x for x in items if x["name"] == "Ausrichtung"), None)
            menus = [h for h in popups() if h not in before]
            if hit and menus:
                l, t, _, _ = winapi.window_rect(menus[0])
                known = popups()
                winapi.click(l + hit["rect"][0] + hit["rect"][2] / 2, t + hit["rect"][1] + hit["rect"][3] / 2)
                time.sleep(1.2)
                res["Formeloptionen Ausrichtung"] = [x["name"] for x in menu_items(known)]
                print("Ausrichtung", res["Formeloptionen Ausrichtung"], flush=True)
            winapi.press("esc")
            time.sleep(0.2)
            winapi.press("esc")
    finally:
        OUT.write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        w.close()


if __name__ == "__main__":
    main()
