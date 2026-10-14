"""Word's ribbon layouts by window width (D-038, package A): for every tab and a range of
window widths, the UI Automation tree of the ribbon, reduced to groups and controls with
their size class (large / medium / small), row and column. RaabeWrite's ribbon definition
takes its layouts from this. Own Word instance; Word's window is resized, nothing is typed.

    .venv/Scripts/python.exe capture_ribbon_widths.py

Output: captures/word-ribbon/<tab>-<width>.json and a text summary per tab (local only).
"""

from __future__ import annotations

import json
import time
from pathlib import Path

from feellab import targets, uia, winapi

OUT = Path(__file__).parent / "captures" / "word-ribbon"
TABS = ["Start", "Einfügen", "Entwurf", "Layout", "Referenzen", "Überprüfen", "Ansicht"]
# CSS widths of the window (the display is at 150 %).
WIDTHS = [2560, 2300, 2100, 1900, 1700, 1500, 1333, 1200, 1050, 900, 800]
CONTROLS = {"Button", "SplitButton", "MenuItem", "ComboBox", "CheckBox", "Edit", "List", "Spinner"}


def size_class(w: float, h: float) -> str:
    if h >= 55:
        return "L"
    if w <= 30:
        return "S"
    return "M"


def reduce(node: dict, scale: float, tab: str = "") -> list[dict]:
    """Groups with their controls (CSS px, relative to the group)."""
    groups = []

    def controls_of(n: dict, out: list[dict]) -> None:
        if n["type"] == "Group" and n["name"]:
            x, y, w, h = (v / scale for v in n["rect"])
            out.append({"name": n["name"], "type": "Gallery", "x": x, "y": y, "w": round(w), "h": round(h),
                        "size": "G", "enabled": n["enabled"]})
            return
        if n["type"] in CONTROLS and n["name"] and n["rect"][2] > 0:
            x, y, w, h = (v / scale for v in n["rect"])
            out.append({"name": n["name"], "type": n["type"], "x": x, "y": y, "w": round(w), "h": round(h),
                        "size": size_class(w, h), "enabled": n["enabled"]})
            return
        for c in n.get("children", []):
            controls_of(c, out)

    def walk(n: dict) -> None:
        if n["type"] == "Group" and n["name"] == tab:
            for c in n.get("children", []):
                walk(c)
            return
        if n["type"] == "Group" and n["name"]:
            x, y, w, h = (v / scale for v in n["rect"])
            cs: list[dict] = []
            for c in n.get("children", []):
                controls_of(c, cs)
            for c in cs:
                c["x"] = round(c["x"] - x)
                c["y"] = round(c["y"] - y)
            groups.append({"name": n["name"], "x": round(x), "w": round(w), "controls": cs,
                           "collapsed": len(cs) == 1 and cs[0]["name"] == n["name"]})
            return
        for c in n.get("children", []):
            walk(c)

    walk(node)
    return groups


def outline(groups: list[dict]) -> str:
    lines = []
    for g in groups:
        lines.append(f"  [{g['name']}] w={g['w']}" + (" (collapsed)" if g["collapsed"] else ""))
        for c in sorted(g["controls"], key=lambda c: (c["x"] // 8, c["y"])):
            lines.append(f"      {c['size']} {c['type'][:6]:6} {c['name']}  @{c['x']},{c['y']} {c['w']}x{c['h']}")
    return "\n".join(lines)


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    w = targets.WordTarget("short")
    try:
        winapi.activate(w.hwnd)
        scale = winapi.dpi_scale(w.hwnd)
        root = uia.element_for(w.hwnd)
        for tab in TABS:
            el = uia.find(root, lambda d, t=tab: d["type"] == "TabItem" and d["name"] == t)
            uia.select(el)
            time.sleep(0.8)
            summary = []
            for width in WIDTHS:
                winapi.place(w.hwnd, 0, 0, round(width * scale), round(900 * scale))
                time.sleep(1.2)
                ribbon = uia.find(root, lambda d: d["name"] == "Unteres Menüband")
                groups = reduce(uia.tree(ribbon, max_depth=14), scale, tab)
                (OUT / f"{tab}-{width}.json").write_text(json.dumps(groups, ensure_ascii=False, indent=1), encoding="utf-8")
                summary.append(f"===== {tab} at {width} px =====\n{outline(groups)}")
                print(tab, width, len(groups), "groups", flush=True)
            (OUT / f"{tab}.txt").write_text("\n\n".join(summary), encoding="utf-8")
    finally:
        w.close()


if __name__ == "__main__":
    main()
