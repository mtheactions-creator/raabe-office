"""Captures Word's interface as a reference for the shell: UI Automation trees and
screenshots of every ribbon tab, the backstage view and the status bar, and the ribbon
at several window widths (to learn Word's group-shrinking order).

Output goes to captures/word-ui/ (not committed: Microsoft screenshots stay local).
Uses UI Automation only; keyboard and mouse are never touched."""

from __future__ import annotations

import json
import time
from pathlib import Path

from feellab import targets, uia, winapi

OUT = Path(__file__).parent / "captures" / "word-ui"
TABS = ["Start", "Einfügen", "Entwurf", "Layout", "Referenzen", "Überprüfen", "Ansicht"]
WIDTHS = [2000, 1800, 1600, 1400, 1200, 1000, 800]


def summary(node: dict, depth: int = 0) -> list[str]:
    """Compact text outline: type, name, size in CSS px (at 150 %), access key."""
    lines = []
    n = node
    if n["name"] or n["type"] not in ("Pane", "Custom"):
        w, h = round(n["rect"][2] / 1.5), round(n["rect"][3] / 1.5)
        x, y = round(n["rect"][0] / 1.5), round(n["rect"][1] / 1.5)
        lines.append(f"{'  ' * depth}{n['type']} | {n['name']} | {w}x{h} @{x},{y} | {n['key']}"
                     + ("" if n["enabled"] else " | disabled"))
        depth += 1
    for c in n.get("children", []):
        lines.extend(summary(c, depth))
    return lines


def save(name: str, node: dict) -> None:
    (OUT / f"{name}.json").write_text(json.dumps(node, ensure_ascii=False, indent=1), encoding="utf-8")
    (OUT / f"{name}.txt").write_text("\n".join(summary(node)), encoding="utf-8")


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    w = targets.WordTarget("short")
    try:
        winapi.activate(w.hwnd)
        root = uia.element_for(w.hwnd)
        left, top, right, bottom = winapi.window_rect(w.hwnd)
        is_tab = lambda name: (lambda d: d["type"] == "TabItem" and d["name"] == name)
        for tab in TABS:
            el = uia.find(root, is_tab(tab))
            if el is None:
                print("tab not found:", tab)
                continue
            uia.select(el)
            time.sleep(0.8)
            ribbon = uia.find(root, lambda d: d["name"] == "Unteres Menüband")
            save(f"tab-{tab}", uia.tree(ribbon, max_depth=14))
            winapi.save_region_png(str(OUT / f"tab-{tab}.png"), left, top, right - left, 300)
            print("captured", tab)

        # Ribbon scaling: Start tab at decreasing window widths.
        uia.select(uia.find(root, is_tab("Start")))
        for width in WIDTHS:
            winapi.place(w.hwnd, left, top, width, bottom - top)
            time.sleep(1.0)
            ribbon = uia.find(root, lambda d: d["name"] == "Unteres Menüband")
            save(f"scale-{width}", uia.tree(ribbon, max_depth=6))
            winapi.save_region_png(str(OUT / f"scale-{width}.png"), left, top, width, 300)
        winapi.place(w.hwnd, left, top, right - left, bottom - top)
        time.sleep(1.0)

        # Status bar and the whole window chrome.
        save("window", uia.tree(root, max_depth=5))
        status = uia.find(root, lambda d: d["name"] == "Statusleiste")
        save("statusbar", uia.tree(status, max_depth=10))
        winapi.save_region_png(str(OUT / "statusbar.png"), left, bottom - 60, right - left, 60)
        winapi.save_region_png(str(OUT / "window.png"), left, top, right - left, bottom - top)

        # Backstage (Datei) pages.
        uia.invoke(uia.find(root, lambda d: d["name"] == 'Registerkarte "Datei"'))
        time.sleep(1.5)
        save("backstage", uia.tree(root, max_depth=12))
        winapi.save_region_png(str(OUT / "backstage.png"), left, top, right - left, bottom - top)
        for page in ("Neu", "Öffnen", "Speichern unter", "Drucken", "Exportieren"):
            el = uia.find(root, lambda d, p=page: d["name"] == p and d["type"] in ("ListItem", "Button", "TabItem", "MenuItem"))
            if el is None:
                print("backstage page not found:", page)
                continue
            try:
                uia.select(el)
            except Exception:
                uia.invoke(el)
            time.sleep(1.5)
            save(f"backstage-{page}", uia.tree(root, max_depth=12))
            winapi.save_region_png(str(OUT / f"backstage-{page}.png"), left, top, right - left, bottom - top)
            print("captured backstage", page)
    finally:
        w.close()


if __name__ == "__main__":
    main()
