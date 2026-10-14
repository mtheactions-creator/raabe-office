"""Probe: Word's context menu inside an equation, per structure (the caret in a fraction, a
script, a root, a sum, brackets, a matrix …), and the Formeloptionen arrow's menu. Own visible
Word instance; takes over mouse and screen. Output: results/<date>-word-equation-context.json
and screenshots in captures/word-equation/ (local).

    .venv/Scripts/python.exe probe_word_equation_context.py
"""
from __future__ import annotations

import ctypes.wintypes as W
import datetime
import json
import sys
import time
from pathlib import Path

from feellab import uia, winapi
from probe_word_equation_gallery import GalleryWord
from probe_word_picture import popups

HERE = Path(__file__).parent
CAP = HERE / "captures" / "word-equation"
OUT = HERE / "results" / f"{datetime.date.today()}-word-equation-context.json"

# Typed in a new equation; then `left` presses put the caret inside the structure.
CASES = [
    ("plain", "x+1", 0),
    ("fraction", "a/b ", 1),
    ("superscript", "x^2 ", 1),
    ("subscript", "x_1 ", 1),
    ("root", "\\sqrt(x) ", 1),
    ("sum", "\\sum_(i=1)^n i ", 0),
    ("integral", "\\int_0^1 x ", 0),
    ("brackets", "(a+b) ", 1),
    ("matrix", "\\matrix(a&b@c&d) ", 1),
    ("function", "sin x ", 0),
    ("accent", "x\\hat ", 1),
    ("limit", "lim_(n->\\infty) a ", 0),
]


def menu_items(before: set[int]) -> list[dict]:
    """The entries of a context menu that opened (read by point; Office's menus are windows
    of their own)."""
    new = [h for h in popups() if h not in before]
    out: list[dict] = []
    for h in new:
        l, t, r, b = winapi.window_rect(h)
        if r - l < 60 or b - t < 40:
            continue
        seen = set()
        for y in range(t + 6, b - 3, 8):
            for x in (l + 60, l + (r - l) // 2):
                try:
                    hit = uia._uia.ElementFromPoint(W.POINT(x, y))
                    if not hit:
                        continue
                    d = uia.describe(hit)
                except Exception:  # noqa: BLE001
                    continue
                key = (d["type"], d["name"], tuple(d["rect"]))
                if d["name"] and key not in seen and d["type"] in ("MenuItem", "Button", "SplitButton", "ListItem", "CheckBox"):
                    seen.add(key)
                    out.append({"type": d["type"], "name": d["name"], "enabled": d["enabled"], "rect": [d["rect"][0] - l, d["rect"][1] - t, d["rect"][2], d["rect"][3]]})
        out.sort(key=lambda d: (d["rect"][1], d["rect"][0]))
    return out


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    w = GalleryWord()
    res: dict = {}
    try:
        for name, typed, lefts in CASES:
            w.fresh()
            winapi.type_text(typed, gap=0.05)
            time.sleep(0.4)
            win = w.d.Windows(1)
            om = w.d.OMaths(1).Range
            menus: dict[str, list[str]] = {}
            # Right-click on every character of the equation: each structure has its own entries.
            for k in range(om.Start, om.End):
                try:
                    left, top, width, height = win.GetPoint(0, 0, 0, 0, w.d.Range(k, k + 1))
                except Exception:  # noqa: BLE001
                    continue
                if width <= 0 or height <= 0:
                    continue
                before = popups()
                winapi.right_click(left + width / 2, top + height / 2)
                time.sleep(1.0)
                items = menu_items(before)
                names = [x["name"] for x in items]
                key = " | ".join(names)
                if key not in menus:
                    menus[key] = names
                    for h in [h for h in popups() if h not in before]:
                        l, t, r, bb = winapi.window_rect(h)
                        if r - l > 60:
                            winapi.save_region_png(str(CAP / f"context-{name}-{len(menus)}.png"), l, t, r - l, bb - t)
                            break
                winapi.press("esc")
                time.sleep(0.3)
            res[name] = list(menus.values())
            print(name, res[name], flush=True)
        # The arrow at the equation's box (Formeloptionen).
        w.fresh()
        winapi.type_text("x=1", gap=0.05)
        time.sleep(0.4)
        el = uia.find(w.root, lambda d: d["type"] == "MenuItem" and d["name"] == "Formeloptionen", max_depth=20)
        if el is not None:
            r = uia.describe(el)["rect"]
            before = popups()
            winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
            time.sleep(1.2)
            res["Formeloptionen"] = menu_items(before)
            for h in [h for h in popups() if h not in before]:
                l, t, rr, b = winapi.window_rect(h)
                if rr - l > 60:
                    winapi.save_region_png(str(CAP / "box-options.png"), l, t, rr - l, b - t)
                    break
            print("Formeloptionen", [x["name"] for x in res["Formeloptionen"]], flush=True)
            winapi.press("esc")
            time.sleep(0.3)
        # Word's equation box while editing, inline and display (screenshots).
        for kind in ("display", "inline"):
            w.d.Content.Delete()
            if kind == "inline":
                w.d.Content.Text = "Text "
                w.d.Range(5, 5).Select()
            winapi.press("alt+shift+0")
            time.sleep(0.4)
            winapi.type_text("a/b ", gap=0.05)
            if kind == "inline":
                winapi.press("right")
                winapi.press("right")
                winapi.type_text(" Text", gap=0.05)
                for _ in range(6):
                    winapi.press("left")
            time.sleep(0.6)
            win = w.d.Windows(1)
            r = w.d.OMaths(1).Range
            left, top, width, height = win.GetPoint(0, 0, 0, 0, r)
            winapi.save_region_png(str(CAP / f"box-{kind}.png"), left - 60, top - 30, width + 160, height + 60)
            res[f"box {kind}"] = {"equation": [left, top, width, height]}
            box = uia.find(w.root, lambda d: d["type"] == "MenuItem" and d["name"] == "Formeloptionen", max_depth=20)
            handle = uia.find(w.root, lambda d: d["type"] == "Button" and d["name"] == "Feld zum Verschieben", max_depth=20)
            res[f"box {kind}"]["options"] = uia.describe(box)["rect"] if box is not None else None
            res[f"box {kind}"]["handle"] = uia.describe(handle)["rect"] if handle is not None else None
            print(kind, res[f"box {kind}"], flush=True)
    finally:
        OUT.write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        w.close()


if __name__ == "__main__":
    main()
