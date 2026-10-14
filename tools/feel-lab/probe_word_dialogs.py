"""Probe: Word's dialogs, menus and galleries for the remaining ribbon commands (Step 4b E).
Each entry opens a ribbon control with a real click in our own visible Word instance, then
records what UI Automation sees (names, types, access keys, positions) and a screenshot, and
closes it with Esc. Takes over mouse and screen. Screenshots: captures/word-dialogs/ (local);
names: results/<date>-word-dialogs.json.

    .venv/Scripts/python.exe probe_word_dialogs.py [name filter …]
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
CAP = HERE / "captures" / "word-dialogs"
CAP.mkdir(parents=True, exist_ok=True)
OUT = HERE / "results" / f"{datetime.date.today()}-word-dialogs.json"

# (name, tab, control name, how: "click" | "bottom" (split button arrow) | "launcher" (group name))
ITEMS = [
    ("symbol menu", "Einfügen", "Symbol", "click"),
    ("symbol dialog", "Einfügen", "Symbol", "menu:Weitere Symbole..."),
    ("link dialog", "Einfügen", "Link", "click"),
    ("bookmark dialog", "Einfügen", "Textmarke...", "click"),
    ("cross reference dialog", "Einfügen", "Querverweis...", "click"),
    ("caption dialog", "Referenzen", "Beschriftung einfügen...", "click"),
    ("citation menu", "Referenzen", "Zitat einfügen", "click"),
    ("sources dialog", "Referenzen", "Quellen verwalten...", "click"),
    ("citation style menu", "Referenzen", "Formatvorlage", "combo"),
    ("bibliography menu", "Referenzen", "Literaturverzeichnis", "click"),
    ("footnote dialog", "Referenzen", "Fußnoten", "launcher"),
    ("track options", "Überprüfen", "Optionen zum Nachverfolgen von Änderungen...", "click"),
    ("themes menu", "Entwurf", "Designs", "click"),
    ("colors menu", "Entwurf", "Farben", "click"),
    ("fonts menu", "Entwurf", "Schriftarten", "click"),
    ("paragraph spacing menu", "Entwurf", "Absatzabstand", "click"),
    ("effects menu", "Entwurf", "Effekte", "click"),
    ("watermark menu", "Entwurf", "Wasserzeichen", "click"),
    ("page color menu", "Entwurf", "Seitenfarbe", "click"),
    ("page borders dialog", "Entwurf", "Seitenränder...", "click"),
    ("hyphenation menu", "Layout", "Silbentrennung", "click"),
    ("hyphenation options", "Layout", "Silbentrennung", "menu:Silbentrennungsoptionen..."),
    ("text effects menu", "Start", "Texteffekte und Typografie", "click"),
    ("sort dialog", "Start", "Sortieren...", "click"),
    ("select menu", "Start", "Bearbeiten", "menu:Markieren"),
    ("clipboard pane", "Start", "Office-Zwischenablage...", "click"),
    ("view tab", "Ansicht", "", "tab"),
    ("tof dialog", "Referenzen", "Abbildungsverzeichnis einfügen...", "click"),
    ("toc menu", "Referenzen", "Inhaltsverzeichnis", "click"),
    ("tab Start", "Start", "", "tab"),
    ("tab Einfügen", "Einfügen", "", "tab"),
    ("tab Entwurf", "Entwurf", "", "tab"),
    ("tab Layout", "Layout", "", "tab"),
    ("tab Referenzen", "Referenzen", "", "tab"),
    ("tab Überprüfen", "Überprüfen", "", "tab"),
]


def snapshot(el, origin: tuple[int, int]) -> list[dict]:
    out = []
    try:
        f = el.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
    except Exception:  # noqa: BLE001
        return out
    for i in range(f.Length):
        d = uia.describe(f.GetElement(i))
        if not d["name"] and d["type"] not in ("Edit", "ComboBox", "List", "Separator"):
            continue
        x, y, w, h = d["rect"]
        out.append({"type": d["type"], "name": d["name"], "key": d["key"], "enabled": d["enabled"], "rect": [x - origin[0], y - origin[1], w, h]})
    return out


class Word:
    def __init__(self):
        import win32com.client
        self.app = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx("Word.Application"))
        self.app.Visible = True
        d = self.app.Documents.Add()
        d.Content.Text = "Überschrift\rDas ist ein Beispieltext mit einigen Wörtern.\rZweiter Absatz."
        d.Paragraphs(1).Style = d.Styles(-2)
        d.Bookmarks.Add("Beispiel", d.Paragraphs(2).Range)
        self.d = d
        win = d.Windows(1)
        win.View.Type = 3
        winapi.place(win.Hwnd, *WINDOW_RECT)
        winapi.activate(win.Hwnd)
        time.sleep(0.8)
        self.hwnd = win.Hwnd
        self.root = uia.element_for(win.Hwnd)
        d.Range(20, 20).Select()

    def tab(self, name: str) -> None:
        el = uia.find(self.root, lambda x: x["type"] == "TabItem" and x["name"] == name, max_depth=14)
        if el is not None:
            r = uia.describe(el)["rect"]
            winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
            time.sleep(0.8)

    def control(self, name: str):
        lower = uia.find(self.root, lambda x: x["type"] == "Pane" and x["name"].startswith("Unter"), max_depth=12) or self.root
        return uia.find(lower, lambda x: x["name"] == name and x["type"] in ("Button", "SplitButton", "MenuItem", "Group", "ComboBox") and x["rect"][2] > 0, max_depth=14)

    def group_launcher(self, group: str):
        g = uia.find(self.root, lambda x: x["type"] == "Group" and x["name"] == group, max_depth=14)
        if g is None:
            return None
        # The dialog launcher: the last small button of the group, at its bottom right.
        f = g.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
        kids = [f.GetElement(i) for i in range(f.Length)]
        gr = uia.describe(g)["rect"]
        best = None
        for e in kids:
            d = uia.describe(e)
            x, y, w, h = d["rect"]
            if d["type"] == "Button" and w < 40 and h < 40 and x + w >= gr[0] + gr[2] - 30 and y + h >= gr[1] + gr[3] - 30:
                best = e
        return best

    def close(self) -> None:
        self.d.Close(0)
        self.app.Quit(0)


def new_windows(before: set[int]) -> list[int]:
    return [h for h in winapi.find_windows(lambda h: True) if h not in before and winapi.window_rect(h)[2] - winapi.window_rect(h)[0] > 40]


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    only = sys.argv[1:]
    out: dict = json.loads(OUT.read_text(encoding="utf-8")) if OUT.exists() else {}
    w = Word()
    try:
        for name, tab, control, how in ITEMS:
            if only and not any(o in name for o in only):
                continue
            w.tab(tab)
            if how == "tab":
                l, t, r, b = winapi.window_rect(w.hwnd)
                winapi.save_region_png(str(CAP / f"{name}.png"), l, t, r - l, 280)
                lower = uia.find(w.root, lambda x: x["type"] == "Pane" and x["name"].startswith("Unter"), max_depth=12)
                out[name] = snapshot(lower, (l, t)) if lower is not None else []
                print(name, len(out[name]), flush=True)
                continue
            el = w.group_launcher(control) if how == "launcher" else w.control(control)
            if el is None:
                print("not found", name, flush=True)
                out[name] = {"error": "not found"}
                continue
            x, y, cw, ch = uia.describe(el)["rect"]
            before = set(winapi.find_windows(lambda h: True))
            kind = uia.describe(el)["type"]
            if how == "combo":
                winapi.click(x + cw - 8, y + ch / 2)
            else:
                winapi.click(x + cw / 2, y + ch - 8 if kind == "SplitButton" else y + ch / 2)
            time.sleep(1.5)
            if how.startswith("menu:"):
                item = how[5:]
                hit = uia.find(w.root, lambda d: d["name"] == item and d["type"] in ("MenuItem", "Button"), max_depth=20)
                if hit is None:
                    for h in new_windows(before):
                        hit = uia.find(uia.element_for(h), lambda d: d["name"] == item, max_depth=12) or hit
                if hit is None:
                    print("no menu item", item, flush=True)
                    winapi.press("esc")
                    continue
                r = uia.describe(hit)["rect"]
                before = set(winapi.find_windows(lambda h: True))
                winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
                time.sleep(1.5)
            wins = new_windows(before)
            # Dialogs are top-level windows; menus and galleries show below the expanded control.
            dialog = next((h for h in wins if winapi.window_class(h) in ("bosa_sdm_msword", "NUIDialog", "#32770")), None)
            target = dialog or next((h for h in wins if winapi.window_class(h) == "Net UI Tool Window"), None)
            if target is None:
                print("nothing opened", name, [winapi.window_class(h) for h in wins], flush=True)
                out[name] = {"error": "nothing opened"}
                winapi.press("esc")
                continue
            l, t, r, b = winapi.window_rect(target)
            winapi.save_region_png(str(CAP / f"{name}.png"), l, t, r - l, b - t)
            items = snapshot(uia.element_for(target), (l, t)) if dialog else []
            if not dialog:
                items = snapshot(el, (l, t)) or snapshot(uia.element_for(target), (l, t))
            out[name] = {"window": winapi.window_class(target), "title": winapi.window_title(target), "size": [r - l, b - t], "items": items}
            print(name, winapi.window_class(target), len(items), flush=True)
            winapi.press("esc")
            time.sleep(0.6)
            if dialog and target in winapi.find_windows(lambda h: True):
                winapi.press("esc")
                time.sleep(0.4)
    finally:
        OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        w.close()


if __name__ == "__main__":
    main()
