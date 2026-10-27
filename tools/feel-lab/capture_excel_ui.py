"""Captures Excel's interface as the reference for RaabeCalc (D-061): UI Automation trees and
screenshots of every ribbon tab, the window (Namenfeld, Bearbeitungsleiste, sheet tabs, status
bar), the backstage pages and the ribbon at several window widths.

Output goes to captures/excel-ui/ (not committed: Microsoft screenshots and texts stay local).
Uses UI Automation only, in an Excel process of our own."""

from __future__ import annotations

import json
import sys
import time
from pathlib import Path

from capture_word_ui import summary
from feellab import uia, winapi
from feellab.excel import ExcelTarget

OUT = Path(__file__).parent / "captures" / "excel-ui"
WIDTHS = [2000, 1800, 1600, 1400, 1200, 1000, 800]


def save(name: str, node: dict) -> None:
    (OUT / f"{name}.json").write_text(json.dumps(node, ensure_ascii=False, indent=1), encoding="utf-8")
    (OUT / f"{name}.txt").write_text("\n".join(summary(node)), encoding="utf-8")


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    x = ExcelTarget()
    try:
        winapi.activate(x.hwnd)
        time.sleep(1.5)
        root = uia.element_for(x.hwnd)
        left, top, right, bottom = winapi.window_rect(x.hwnd)
        save("window", uia.tree(root, max_depth=7))
        winapi.save_region_png(str(OUT / "window.png"), left, top, right - left, bottom - top)
        tabs = []
        tablist = uia.find(root, lambda d: d["type"] == "Tab" and d["name"] == "Registerkarten des Menübands")
        for el in uia.children(tablist) if tablist else []:
            d = uia.describe(el)
            if d["type"] == "TabItem":
                tabs.append(d["name"])
        print("tabs:", tabs)
        for tab in tabs if "--widths-only" not in sys.argv else []:
            el = uia.find(root, lambda d, t=tab: d["type"] == "TabItem" and d["name"] == t)
            uia.select(el)
            time.sleep(0.8)
            ribbon = uia.find(root, lambda d: d["name"] == "Unteres Menüband")
            save(f"tab-{tab}", uia.tree(ribbon, max_depth=14))
            winapi.save_region_png(str(OUT / f"tab-{tab}.png"), left, top, right - left, 300)
            print("captured", tab)
        uia.select(uia.find(root, lambda d: d["type"] == "TabItem" and d["name"] == "Start"))
        for width in WIDTHS:
            for tab in ("Start", "Einfügen", "Seitenlayout", "Formeln", "Daten", "Überprüfen", "Ansicht"):
                el = uia.find(root, lambda d, t=tab: d["type"] == "TabItem" and d["name"] == t)
                if el is None:
                    continue
                uia.select(el)
                winapi.place(x.hwnd, left, top, width, bottom - top)
                time.sleep(0.8)
                ribbon = uia.find(root, lambda d: d["name"] == "Unteres Menüband")
                save(f"scale-{tab}-{width}", uia.tree(ribbon, max_depth=6))
        winapi.place(x.hwnd, left, top, right - left, bottom - top)
        time.sleep(1.0)
        status = uia.find(root, lambda d: d["name"] == "Statusleiste")
        if status:
            save("statusbar", uia.tree(status, max_depth=10))
        uia.invoke(uia.find(root, lambda d: d["name"] == 'Registerkarte "Datei"'))
        time.sleep(1.5)
        save("backstage", uia.tree(root, max_depth=12))
        winapi.save_region_png(str(OUT / "backstage.png"), left, top, right - left, bottom - top)
    finally:
        x.close()


if __name__ == "__main__":
    main()
