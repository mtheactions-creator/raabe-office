"""Probe: Word's shape UI (Einfügen > Formen, the Formformat tab). Own visible Word instance;
takes over mouse and screen. Screenshots go to captures/word-shapes/ (local only), tooltips
and recorded XML to results/<date>-word-shapes-ui.json.

    .venv/Scripts/python.exe probe_word_shapes_ui.py shot <tab> <button>      # screenshot of a popup
    .venv/Scripts/python.exe probe_word_shapes_ui.py grid <tab> <button> <x0> <y0> <dx> <dy> <cols> <rows>
        # tooltips of a grid of gallery entries (pixel offsets inside the popup)
    .venv/Scripts/python.exe probe_word_shapes_ui.py apply <tab> <button> <x> <y> [<x> <y> ...]
        # click entries (pixel offsets) with the shape selected, record the shape's XML, undo
"""
from __future__ import annotations

import datetime
import json
import re
import sys
import time
from pathlib import Path

from feellab import uia, winapi
from feellab.targets import WINDOW_RECT
from probe_word_picture import popups, tooltip_text

HERE = Path(__file__).parent
CAP = HERE / "captures" / "word-shapes"
CAP.mkdir(parents=True, exist_ok=True)


def shape_xml(d) -> str:
    x = d.Content.WordOpenXML
    m = re.search(r"<wps:wsp>.*?</wps:wsp>", x, re.S)
    return m.group(0) if m else ""


class Word:
    def __init__(self):
        import win32com.client
        self.app = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx("Word.Application"))
        self.app.Visible = True
        d = self.app.Documents.Add()
        # The theme of new RaabeWrite documents (D-046), so colour names match ours.
        d.ApplyDocumentTheme(str(Path("C:/Program Files/Microsoft Office/root/Document Themes 16/Office 2013 - 2022 Theme.thmx")))
        d.Content.Text = "Formen:"
        d.Shapes.AddShape(1, 150, 150, 100, 70)
        self.d = d
        win = d.Windows(1)
        win.View.Type = 3
        winapi.place(win.Hwnd, *WINDOW_RECT)
        winapi.activate(win.Hwnd)
        time.sleep(0.8)
        self.root = uia.element_for(win.Hwnd)
        self.cache: dict = {}

    def tab(self, name: str) -> None:
        if name == "Formformat":
            self.d.Shapes(1).Select()
            time.sleep(0.4)
        key = ("tab", name)
        if key not in self.cache:
            el = uia.find(self.root, lambda x: x["type"] == "TabItem" and x["name"] == name, max_depth=14)
            self.cache[key] = uia.describe(el)["rect"] if el is not None else None
        r = self.cache[key]
        if r:
            winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
            time.sleep(0.7)

    def button(self, name: str):
        key = ("btn", name)
        if key not in self.cache:
            stack = [(self.root, 0)]
            found = None
            while stack and not found:
                e, depth = stack.pop()
                dd = uia.describe(e)
                if dd["name"] == name and dd["type"] in ("MenuItem", "SplitButton", "Button") and dd["rect"][2] > 0:
                    found = (dd["type"], dd["rect"])
                if depth < 18:
                    stack.extend((c, depth + 1) for c in uia.children(e))
            self.cache[key] = found
        return self.cache[key]

    def open(self, tab: str, name: str):
        self.tab(tab)
        b = self.button(name)
        if not b:
            print("no button", name)
            return None
        kind, (bx, by, bw, bh) = b
        before = popups()
        winapi.click(bx + bw / 2, by + bh - 8 if kind == "SplitButton" else by + bh / 2)
        time.sleep(1.3)
        new = [h for h in popups() if h not in before]
        if not new:
            return None
        return max(new, key=lambda h: (winapi.window_rect(h)[2] - winapi.window_rect(h)[0]) * (winapi.window_rect(h)[3] - winapi.window_rect(h)[1]))

    def close(self) -> None:
        self.d.Close(0)
        self.app.Quit(0)


def main() -> None:
    mode, tab, name = sys.argv[1], sys.argv[2], sys.argv[3]
    nums = [float(v) for v in sys.argv[4:]] if mode in ("grid", "apply") else []
    w = Word()
    out: dict = {}
    try:
        if mode == "nudge":
            # How far the arrow keys move a selected shape (with Strg, Umschalt), in points.
            sh = w.d.Shapes(1)
            for keys in ("right", "ctrl+right", "down", "ctrl+down", "shift+right"):
                sh.Select()
                time.sleep(0.3)
                winapi.activate(w.d.Windows(1).Hwnd)
                l0, t0 = sh.Left, sh.Top
                winapi.press(keys)
                time.sleep(0.4)
                out.setdefault("nudge", {})[keys] = [round(sh.Left - l0, 3), round(sh.Top - t0, 3)]
                print(keys, out["nudge"][keys], flush=True)
        elif mode == "dump":
            # The UIA subtree of a ribbon group (names, types, rects), to find gallery buttons.
            w.tab(tab)
            grp = uia.find(w.root, lambda x: x["type"] == "Group" and x["name"] == name, max_depth=16)
            def walk(e, depth):
                d = uia.describe(e)
                print("  " * depth + f'{d["type"]} {d["name"]!r} {d["rect"]}')
                if depth < 6:
                    for c in uia.children(e):
                        walk(c, depth + 1)
            if grp is not None:
                walk(grp, 0)
        elif mode == "applylist":
            # Every entry of an in-ribbon gallery (by its UIA name), applied to the shape; the
            # shape's XML recorded, then undone. `name` is the ribbon group, argv[4] the gallery.
            gallery = sys.argv[4]
            base = shape_xml(w.d)
            res = []
            w.tab(tab)

            def items():
                grp = uia.find(w.root, lambda x: x["type"] == "Group" and x["name"] == name, max_depth=16)
                gal = uia.find(grp, lambda x: x["type"] == "DataGrid" and x["name"] == gallery, max_depth=8) if grp is not None else None
                if gal is None:
                    return []
                return [c for c in uia.children(gal) if uia.describe(c)["type"] == "ListItem"]

            names = [uia.describe(e)["name"] for e in items()]
            print(len(names), "entries", flush=True)
            for i, label in enumerate(names):
                w.d.Shapes(1).Select()
                time.sleep(0.2)
                el = next((e for e in items() if uia.describe(e)["name"] == label), None)
                if el is None:
                    res.append({"label": label, "error": "not found"})
                    continue
                try:
                    uia.invoke(el)
                except Exception:  # noqa: BLE001
                    try:
                        uia.select(el)
                    except Exception as e:  # noqa: BLE001
                        res.append({"label": label, "error": str(e)})
                        continue
                time.sleep(0.6)
                xml = shape_xml(w.d)
                res.append({"label": label, "xml": xml, "changed": xml != base})
                print(i, label, "changed" if xml != base else "unchanged", flush=True)
                for _ in range(3):
                    if shape_xml(w.d) == base:
                        break
                    w.d.Undo(1)
                    time.sleep(0.3)
            out[f"applylist {tab} {name} {gallery}"] = res
        elif mode == "shot":
            pop = w.open(tab, name)
            if pop:
                l, t, r, b = winapi.window_rect(pop)
                dest = CAP / f"{tab}-{name}.png".replace(" ", "_")
                winapi.save_region_png(str(dest), l, t, r - l, b - t)
                print(dest, (l, t, r - l, b - t))
            winapi.press("esc")
        elif mode == "grids":
            # Several grids, as JSON: {"label": [x0, y0, dx, dy, cols, rows], ...}
            groups = json.loads(sys.argv[4])
            tips = {}
            for label, (x0, y0, dx, dy, cols, rows) in groups.items():
                pop = w.open(tab, name)
                if not pop:
                    continue
                l, t, _, _ = winapi.window_rect(pop)
                for r in range(int(rows)):
                    for c in range(int(cols)):
                        known = popups()
                        x, y = l + x0 + c * dx, t + y0 + r * dy
                        winapi.move(x, y + 2)
                        time.sleep(0.15)
                        winapi.move(x, y)
                        time.sleep(1.2)
                        tips.setdefault(label, []).append(tooltip_text(known))
                        print(label, r, c, tips[label][-1], flush=True)
                winapi.press("esc")
                time.sleep(0.4)
            out[f"tips {tab} {name}"] = tips
        elif mode == "grid":
            x0, y0, dx, dy, cols, rows = nums
            pop = w.open(tab, name)
            if pop:
                l, t, _, _ = winapi.window_rect(pop)
                tips = {}
                for r in range(int(rows)):
                    for c in range(int(cols)):
                        known = popups()
                        x, y = l + x0 + c * dx, t + y0 + r * dy
                        winapi.move(x, y + 2)
                        time.sleep(0.15)
                        winapi.move(x, y)
                        time.sleep(1.3)
                        tips[f"{r},{c}"] = tooltip_text(known)
                        print(r, c, tips[f"{r},{c}"], flush=True)
                out[f"tips {tab} {name} {sys.argv[4:]}"] = tips
            winapi.press("esc")
        elif mode == "apply":
            base = shape_xml(w.d)
            res = []
            for i in range(0, len(nums), 2):
                pop = w.open(tab, name)
                if not pop:
                    continue
                l, t, _, _ = winapi.window_rect(pop)
                winapi.click(l + nums[i], t + nums[i + 1])
                time.sleep(1.0)
                xml = shape_xml(w.d)
                res.append({"at": nums[i:i + 2], "xml": xml, "changed": xml != base})
                print(nums[i:i + 2], "changed" if xml != base else "unchanged", flush=True)
                winapi.press("esc")
                for _ in range(3):
                    if shape_xml(w.d) == base:
                        break
                    w.d.Undo(1)
                    time.sleep(0.3)
            out[f"apply {tab} {name}"] = res
    finally:
        w.close()
    if out:
        dest = HERE / "captures" / "word-shapes-ui.json"
        old = json.loads(dest.read_text(encoding="utf-8")) if dest.exists() else {}
        old.update(out)
        dest.write_text(json.dumps(old, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
