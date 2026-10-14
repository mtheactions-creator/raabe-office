"""Probe: Word's Formel tab (the contextual tab while the caret is in an equation): its groups
and controls, and what every dropdown offers (galleries with their headings, menus). Own
visible Word instance; takes over mouse and screen. Screenshots go to captures/word-equation/
(local only), the names to results/<date>-word-equation-ui.json.

    .venv/Scripts/python.exe probe_word_equation_ui.py            # tab, every dropdown
    .venv/Scripts/python.exe probe_word_equation_ui.py symbols    # the symbol categories
"""
from __future__ import annotations

import datetime
import json
import sys
import time
from pathlib import Path

from feellab import uia, winapi
from feellab.targets import WINDOW_RECT
from probe_word_picture import popups

HERE = Path(__file__).parent
CAP = HERE / "captures" / "word-equation"
CAP.mkdir(parents=True, exist_ok=True)
OUT = HERE / "results" / f"{datetime.date.today()}-word-equation-ui.json"


def walk(el, depth: int = 0, max_depth: int = 9) -> list[dict]:
    """Every element below el (type, name, access key, rect), flattened with depth."""
    out = []
    for c in uia.children(el):
        d = uia.describe(c)
        d["depth"] = depth
        out.append(d)
        if depth < max_depth:
            out.extend(walk(c, depth + 1, max_depth))
    return out


class Word:
    def __init__(self):
        import win32com.client
        self.app = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx("Word.Application"))
        self.app.Visible = True
        d = self.app.Documents.Add()
        d.Content.Text = "x"
        d.OMaths.Add(d.Paragraphs(1).Range)
        self.d = d
        win = d.Windows(1)
        win.View.Type = 3
        winapi.place(win.Hwnd, *WINDOW_RECT)
        winapi.activate(win.Hwnd)
        time.sleep(0.8)
        self.hwnd = win.Hwnd
        self.root = uia.element_for(win.Hwnd)
        self.into_equation()

    def into_equation(self) -> None:
        r = self.d.OMaths(1).Range
        self.d.Range(r.Start, r.Start + 1).Select()
        time.sleep(0.5)

    def tab(self, name: str = "Formel") -> None:
        el = uia.find(self.root, lambda x: x["type"] == "TabItem" and x["name"] == name, max_depth=14)
        if el is None:
            print("no tab", name)
            return
        r = uia.describe(el)["rect"]
        winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
        time.sleep(0.8)

    def lower_ribbon(self):
        return uia.find(self.root, lambda x: x["type"] == "Group" and x["name"] == "Formel", max_depth=16) or self.root

    def controls(self) -> list[dict]:
        groups = []
        for d in walk(self.lower_ribbon(), 0, 12):
            if d["rect"][2] <= 0:
                continue
            groups.append({k: d[k] for k in ("depth", "type", "name", "key", "enabled", "rect")})
        return groups

    def open_popup(self, rect, bottom: bool) -> int | None:
        before = popups()
        x, y, w, h = rect
        winapi.click(x + w / 2, y + h - 8 if bottom else y + h / 2)
        time.sleep(1.4)
        new = [p for p in popups() if p not in before]
        if not new:
            return None
        return max(new, key=lambda p: (winapi.window_rect(p)[2] - winapi.window_rect(p)[0]) * (winapi.window_rect(p)[3] - winapi.window_rect(p)[1]))

    def close(self) -> None:
        self.d.Close(0)
        self.app.Quit(0)


def popup_content(pop: int, root=None, name: str = "") -> list[dict]:
    """Every named element of a popup (headings, gallery entries, menu items), in tree order.
    Office shows a dropdown's content below the expanded control in the window's tree."""
    el = uia.element_for(pop)
    if root is not None and name:
        el = uia.find(root, lambda d: d["type"] in ("MenuItem", "Group") and d["name"] == name, max_depth=16) or el
    l, t, _, _ = winapi.window_rect(pop)
    items = []
    found = el.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
    for i in range(found.Length):
        d = uia.describe(found.GetElement(i))
        if not d["name"] and d["type"] != "Separator":
            continue
        x, y, w, h = d["rect"]
        items.append({"type": d["type"], "name": d["name"], "key": d["key"], "enabled": d["enabled"], "rect": [x - l, y - t, w, h]})
    return items


def shot(pop: int, name: str) -> None:
    l, t, r, b = winapi.window_rect(pop)
    winapi.save_region_png(str(CAP / f"{name}.png"), l, t, r - l, b - t)


def main() -> None:
    mode = sys.argv[1] if len(sys.argv) > 1 else "tab"
    w = Word()
    out: dict = json.loads(OUT.read_text(encoding="utf-8")) if OUT.exists() else {}
    try:
        w.tab()
        scale = winapi.dpi_scale(w.hwnd)
        l, t, _, _ = winapi.window_rect(w.hwnd)
        winapi.save_region_png(str(CAP / "tab.png"), l, t, int(2000 * scale) // 1, int(260 * scale) // 1)
        ctl = w.controls()
        if mode in ("tab", "popups"):
            if mode == "tab":
                out["tab"] = ctl
                for c in ctl:
                    print("  " * c["depth"] + f'{c["type"]} {c["name"]!r} key={c["key"]!r} {c["rect"]}', flush=True)
            # Every dropdown: open it, take its names and a screenshot.
            dropdowns = [c for c in ctl if (c["type"] in ("SplitButton", "MenuItem") and c["name"]) or (c["type"] == "Button" and c["name"] == "Formelsymbole")]
            out["popups"] = {}
            for c in dropdowns:
                w.into_equation()
                w.tab()
                pop = w.open_popup(c["rect"], c["type"] == "SplitButton")
                if not pop:
                    print("no popup", c["name"], flush=True)
                    continue
                shot(pop, c["name"].replace(" ", "_").replace("/", "_"))
                out["popups"][c["name"]] = popup_content(pop, w.root, c["name"])
                print("popup", c["name"], len(out["popups"][c["name"]]), flush=True)
                winapi.press("esc")
                time.sleep(0.3)
                winapi.press("esc")
                time.sleep(0.4)
        elif mode == "symbols":
            # The full symbol panel ("Mehr" in the Symbole gallery) and its categories.
            more = next((c for c in ctl if c["type"] == "Button" and c["name"] == "Formelsymbole"), None)
            if more is None:
                print("no Mehr button")
                return
            pop = w.open_popup(more["rect"], False)
            if not pop:
                print("no panel")
                return
            shot(pop, "symbols")
            first = popup_content(pop)
            out["symbols"] = {"panel": first}
            # The category menu is the first menu button in the panel.
            head = next((x for x in first if x["type"] in ("MenuItem", "SplitButton", "Button") and x["depth"] <= 4), None)
            print("head", head, flush=True)
            if head:
                pl, pt, _, _ = winapi.window_rect(pop)
                before = popups()
                x, y, ww, hh = head["rect"]
                winapi.click(pl + x + ww / 2, pt + y + hh / 2)
                time.sleep(1.2)
                menu = [p for p in popups() if p not in before]
                cats = []
                if menu:
                    shot(menu[0], "symbol-categories")
                    cats = [x for x in popup_content(menu[0]) if x["type"] == "MenuItem"]
                out["symbols"]["categories"] = cats
                winapi.press("esc")
                time.sleep(0.4)
                for i, cat in enumerate(cats):
                    before = popups()
                    winapi.click(pl + x + ww / 2, pt + y + hh / 2)
                    time.sleep(1.0)
                    menu = [p for p in popups() if p not in before]
                    if not menu:
                        continue
                    ml, mt, _, _ = winapi.window_rect(menu[0])
                    cx, cy, cw, ch = cat["rect"]
                    winapi.click(ml + cx + cw / 2, mt + cy + ch / 2)
                    time.sleep(1.0)
                    shot(pop, f"symbols-{i}")
                    out["symbols"][cat["name"]] = [x for x in popup_content(pop) if x["type"] in ("ListItem", "Button", "Group", "Header", "HeaderItem", "Text")]
                    print("category", cat["name"], len(out["symbols"][cat["name"]]), flush=True)
            winapi.press("esc")
    finally:
        OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        w.close()


if __name__ == "__main__":
    main()
