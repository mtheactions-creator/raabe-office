"""Probe: how Word shows the Ansicht commands (Lineal, Gitternetzlinien, Mehrere Seiten,
Lesemodus, Weblayout) and Layout > Zeilennummern, in our own visible Word instance with real
clicks; screenshots of the window. Gitternetzlinien is switched off again at the end.
Takes over mouse and screen. Screenshots: captures/word-views/ (local).

    .venv/Scripts/python.exe probe_word_views.py
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

from feellab import uia, winapi
from probe_word_dialogs import Word

CAP = Path(__file__).parent / "captures" / "word-views"
CAP.mkdir(parents=True, exist_ok=True)

TEXT = ("Die Photosynthese ist ein physiologischer Prozess zur Erzeugung von energiereichen Biomolekülen "
        "aus energieärmeren Stoffen mithilfe von Lichtenergie. Sie wird von Pflanzen, Algen und einigen "
        "Bakterien betrieben. ") * 6


def shot(w: Word, name: str) -> None:
    l, t, r, b = winapi.window_rect(w.hwnd)
    winapi.save_region_png(str(CAP / f"{name}.png"), l, t, r - l, b - t)
    print(name, flush=True)


def click(w: Word, name: str, types=("Button", "CheckBox", "MenuItem")) -> bool:
    el = uia.find(w.root, lambda d: d["name"] == name and d["type"] in types and d["rect"][2] > 0, max_depth=16)
    if el is None:
        print("not found", name, flush=True)
        return False
    r = uia.describe(el)["rect"]
    winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
    time.sleep(1.5)
    return True


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    w = Word()
    d = w.d
    d.Content.Text = "Überschrift\r" + "\r".join([TEXT] * 8)
    d.Paragraphs(1).Style = d.Styles(-2)
    d.Range(30, 30).Select()
    time.sleep(0.5)
    grid_was = w.app.Options.DisplayGridLines
    try:
        w.tab("Ansicht")
        shot(w, "plain")
        if click(w, "Lineal"):
            shot(w, "ruler")
            click(w, "Lineal")
        if click(w, "Gitternetzlinien"):
            shot(w, "gridlines")
            click(w, "Gitternetzlinien")
        if click(w, "Mehrere Seiten"):
            shot(w, "multiple pages")
            click(w, "100%")
        if click(w, "Weblayout"):
            shot(w, "web layout")
            click(w, "Drucklayout")
        if click(w, "Lesemodus"):
            time.sleep(1.5)
            shot(w, "read mode")
            winapi.press("esc")
            time.sleep(1.5)
        # Zeilennummern: fortlaufend (set through COM, a document property).
        ln = d.Sections(1).PageSetup.LineNumbering
        ln.Active = True
        ln.RestartMode = 2
        time.sleep(1.0)
        w.tab("Layout")
        shot(w, "line numbers")
    finally:
        w.app.Options.DisplayGridLines = grid_was
        w.close()


if __name__ == "__main__":
    main()
