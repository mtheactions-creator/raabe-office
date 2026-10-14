"""Probe: what each entry of Word's equation galleries inserts. Every entry is applied (UI
Automation invoke) to a fresh empty equation and the equation's OMML recorded, with where the
caret is afterwards; the symbol gallery is walked category by category and the inserted
characters recorded. Own visible Word instance; takes over mouse and screen.
Output: captures/word-equation/gallery.json (Word's XML stays local).

    .venv/Scripts/python.exe probe_word_equation_gallery.py galleries [Bruch ...]
    .venv/Scripts/python.exe probe_word_equation_gallery.py symbols
"""
from __future__ import annotations

import json
import re
import sys
import time
from pathlib import Path

from feellab import uia, winapi
from probe_word_equation_ui import CAP, Word

OUT = CAP / "gallery.json"
GALLERIES = ["Formel", "Bruch", "Hoch/Tief", "Wurzel", "Integral", "Großer Operator", "Eckige Klammer", "Funktion",
             "Akzent", "Grenzwert und Logarithmus", "Operator", "Matrix"]


def math_xml(d) -> str:
    x = d.Content.WordOpenXML
    m = re.search(r"<m:oMathPara>.*?</m:oMathPara>", x, re.S) or re.search(r"<m:oMath>.*?</m:oMath>", x, re.S)
    return m.group(0) if m else ""


class GalleryWord(Word):
    def fresh(self) -> None:
        """An empty equation alone in the document, the caret in it."""
        d = self.d
        d.Content.Delete()
        d.Range(0, 0).Select()
        winapi.activate(self.hwnd)
        # Word's own command (Alt+Umschalt+0): an empty equation with the caret in it.
        winapi.press("alt+shift+0")
        time.sleep(0.5)

    def into_equation(self) -> None:
        if self.d.OMaths.Count == 0:
            self.fresh()
            return
        r = self.d.OMaths(1).Range
        self.d.Range(r.Start, r.Start).Select()
        time.sleep(0.3)

    def expanded(self, name: str):
        return uia.find(self.root, lambda d: d["type"] in ("MenuItem", "Group") and d["name"] == name, max_depth=16)

    def items(self, name: str) -> list:
        el = self.expanded(name)
        if el is None:
            return []
        f = el.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
        return [f.GetElement(i) for i in range(f.Length)]


def open_gallery(w: GalleryWord, ctl: list[dict], name: str) -> bool:
    c = next((x for x in ctl if x["name"] == name and x["type"] in ("MenuItem", "Button")), None)
    if c is None:
        return False
    x, y, ww, hh = c["rect"]
    winapi.click(x + ww / 2, y + hh / 2)
    time.sleep(1.2)
    return True


def galleries(w: GalleryWord, names: list[str], out: dict) -> None:
    w.fresh()
    w.tab()
    ctl = w.controls()
    for g in names:
        # The entries and their headings, in order.
        if not open_gallery(w, ctl, g):
            print("no gallery", g)
            continue
        entries, heading = [], ""
        for el in w.items(g):
            d = uia.describe(el)
            if d["type"] == "Group" and d["name"] not in ("Tools", ""):
                heading = d["name"]
            elif d["type"] == "ListItem":
                entries.append((heading, d["name"]))
        winapi.press("esc")
        time.sleep(0.4)
        res = []
        for i, (heading, label) in enumerate(entries):
            w.fresh()
            w.tab()
            if not open_gallery(w, ctl, g):
                break
            el = next((e for e in w.items(g) if uia.describe(e)["type"] == "ListItem" and uia.describe(e)["name"] == label), None)
            if el is None:
                res.append({"heading": heading, "label": label, "error": "not found"})
                winapi.press("esc")
                continue
            # Several entries share a name (Wurzel: two "Quadratwurzel"): take the i-th of that name.
            same = [e for e in w.items(g) if uia.describe(e)["type"] == "ListItem" and uia.describe(e)["name"] == label]
            k = sum(1 for _h, lab in entries[:i] if lab == label)
            if k < len(same):
                el = same[k]
            try:
                uia.invoke(el)
            except Exception:  # noqa: BLE001
                try:
                    uia.select(el)
                except Exception as e:  # noqa: BLE001
                    res.append({"heading": heading, "label": label, "error": str(e)})
                    winapi.press("esc")
                    continue
            time.sleep(0.7)
            sel = w.app.Selection
            m = w.d.OMaths(1).Range if w.d.OMaths.Count else w.d.Content
            res.append({"heading": heading, "label": label, "omml": math_xml(w.d), "text": m.Text,
                        "sel": [sel.Start - m.Start, sel.End - m.Start], "count": w.d.OMaths.Count})
            print(g, i, heading, "|", label, "|", m.Text, res[-1]["sel"], flush=True)
        out[g] = res
        OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


def symbols(w: GalleryWord, out: dict, only: list[str] | None = None) -> None:
    w.fresh()
    w.tab()
    ctl = w.controls()
    res: dict = dict(out.get("symbols", {})) if only else {}

    def open_panel() -> None:
        open_gallery(w, ctl, "Formelsymbole")

    def category_menu() -> list[tuple[str, list[int]]]:
        """Opens the Katalogfilter menu; its entries (name, rect), read by point (the menu is
        its own window that UI Automation does not show below the ribbon)."""
        panel = w.expanded("Formelsymbole")
        el = uia.find(panel, lambda d: d["type"] == "MenuItem" and d["name"] == "Katalogfilter", max_depth=10) if panel is not None else None
        if el is None:
            return []
        r = uia.describe(el)["rect"]
        before = set(winapi.find_windows(lambda h: True))
        winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
        time.sleep(1.0)
        menu = [h for h in winapi.find_windows(lambda h: True) if h not in before and winapi.window_class(h) == "Net UI Tool Window"]
        if not menu:
            return []
        import ctypes.wintypes as W
        l, t, rr, bb = winapi.window_rect(menu[0])
        out: list[tuple[str, list[int]]] = []
        for y in range(t + 8, bb - 4, 12):
            hit = uia._uia.ElementFromPoint(W.POINT((l + rr) // 2, y))
            if not hit:
                continue
            d = uia.describe(hit)
            if d["type"] == "ListItem" and d["name"] and d["name"] not in [n for n, _r in out]:
                out.append((d["name"], d["rect"]))
        return out

    open_panel()
    cats = [n for n, _r in category_menu()]
    print("categories", cats, flush=True)
    winapi.press("esc")
    time.sleep(0.3)
    winapi.press("esc")
    time.sleep(0.4)
    for cat in cats:
        if only and cat not in only:
            continue
        w.fresh()
        w.tab()
        open_panel()
        hit = next((r for n, r in category_menu() if n == cat), None)
        if hit is None:
            print("no category", cat)
            winapi.press("esc")
            continue
        winapi.click(hit[0] + hit[2] / 2, hit[1] + hit[3] / 2)
        time.sleep(1.0)
        entries, heading = [], ""
        for e in w.items("Formelsymbole"):
            d = uia.describe(e)
            if d["type"] == "Group" and d["name"] not in ("Katalogfilter", "Formelsymbole", ""):
                heading = d["name"]
            elif d["type"] == "ListItem":
                entries.append((heading, d["name"]))
        winapi.press("esc")
        time.sleep(0.5)
        # The ribbon's gallery now shows this category: insert each entry from there.
        got = []
        w.fresh()
        w.tab()
        grid_items = lambda: [e for e in w.items("Formelsymbole") if uia.describe(e)["type"] == "ListItem"]  # noqa: E731
        ribbon = grid_items()
        names = [uia.describe(e)["name"] for e in ribbon]
        for i, (heading, label) in enumerate(entries):
            k = sum(1 for _h, lab in entries[:i] if lab == label)
            idx = [j for j, n in enumerate(names) if n == label]
            if k >= len(idx):
                got.append({"heading": heading, "label": label, "error": "not in ribbon"})
                continue
            before = w.d.OMaths(1).Range.Text if w.d.OMaths.Count else ""
            try:
                uia.invoke(grid_items()[idx[k]])
            except Exception as e:  # noqa: BLE001
                got.append({"heading": heading, "label": label, "error": str(e)})
                continue
            time.sleep(0.35)
            after = w.d.OMaths(1).Range.Text if w.d.OMaths.Count else w.d.Content.Text
            got.append({"heading": heading, "label": label, "text": after[len(before):] if after.startswith(before) else after})
            print(cat, i, label, repr(got[-1]["text"]), flush=True)
            if i % 20 == 19:
                w.fresh()
                w.tab()
        res[cat] = got
        out["symbols"] = res
        OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    mode = sys.argv[1]
    out: dict = json.loads(OUT.read_text(encoding="utf-8")) if OUT.exists() else {}
    w = GalleryWord()
    try:
        if mode == "galleries":
            galleries(w, sys.argv[2:] or GALLERIES, out)
        elif mode == "symbols":
            symbols(w, out, sys.argv[2:])
    finally:
        OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        w.close()


if __name__ == "__main__":
    main()
