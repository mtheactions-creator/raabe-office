"""Probe: Word's Tabellenentwurf tab (Step 4c D): the Tabellenformatvorlagen gallery (its
entries in order, screenshots of the expanded gallery scrolled through) and the Rahmen group's
menus (Rahmenarten, Stiftart, Stiftstärke, Stiftfarbe, Rahmen) with their entries. Own visible
Word instance; takes over mouse and screen. Screenshots: captures/word-table/ (local); names:
results/<date>-word-table-design.json.

    .venv/Scripts/python.exe probe_word_table_gallery.py
"""
from __future__ import annotations

import ctypes
import datetime
import json
import sys
import time
from pathlib import Path

from feellab import uia, winapi
from probe_word_dialogs import Word, snapshot

HERE = Path(__file__).parent
CAP = HERE / "captures" / "word-table"
CAP.mkdir(parents=True, exist_ok=True)
OUT = HERE / "results" / f"{datetime.date.today()}-word-table-design.json"


def popup_rect(before: set[int]) -> tuple[int, int, int, int] | None:
    new = [h for h in winapi.find_windows(lambda h: True) if h not in before and winapi.window_class(h) == "Net UI Tool Window"]
    if not new:
        return None
    rs = [winapi.window_rect(h) for h in new]
    return min(r[0] for r in rs), min(r[1] for r in rs), max(r[2] for r in rs), max(r[3] for r in rs)


def items_in(w: Word, pr: tuple[int, int, int, int]) -> list[str]:
    out = []
    for x in snapshot(w.root, (0, 0)):
        if x["name"] and pr[0] - 2 <= x["rect"][0] < pr[2] and pr[1] - 2 <= x["rect"][1] < pr[3] and x["type"] in ("ListItem", "MenuItem", "Button", "Group", "SplitButton", "CheckBox"):
            out.append(f"{x['type']}:{x['name']}")
    return out


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    w = Word()
    d = w.d
    out: dict = {}
    try:
        t = d.Tables.Add(d.Paragraphs(3).Range, 4, 4)
        t.Cell(2, 2).Range.Select()
        time.sleep(0.5)
        w.tab("Tabellenentwurf")
        grid = uia.find(w.root, lambda x: x["type"] == "DataGrid" and x["name"] == "Tabellenformatvorlagen", max_depth=16)
        if grid is not None:
            f = grid.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
            out["gallery"] = [uia.describe(f.GetElement(i))["name"] for i in range(f.Length) if uia.describe(f.GetElement(i))["type"] == "ListItem"]
            print("gallery", len(out["gallery"]), flush=True)
        # The expanded gallery (its groups), scrolled through.
        more = uia.find(w.root, lambda x: x["type"] == "Button" and x["name"] == "Tabellenformatvorlagen", max_depth=16)
        if more is not None:
            before = set(winapi.find_windows(lambda h: True))
            rr = uia.describe(more)["rect"]
            winapi.click(rr[0] + rr[2] / 2, rr[1] + rr[3] / 2)
            time.sleep(1.5)
            pr = popup_rect(before)
            if pr:
                out["gallery popup"] = {"rect": pr, "items": items_in(w, pr)}
                x0, y0, x1, y1 = pr
                for k in range(6):
                    winapi.save_region_png(str(CAP / f"gallery {k}.png"), x0, y0, x1 - x0, y1 - y0)
                    winapi.move((x0 + x1) / 2, (y0 + y1) / 2)
                    for _ in range(4):
                        ctypes.windll.user32.mouse_event(0x0800, 0, 0, -120, 0)
                        time.sleep(0.05)
                    time.sleep(0.6)
            winapi.press("esc")
            time.sleep(0.5)

        def open_menu(name: str, el, at: tuple[float, float]) -> None:
            before = set(winapi.find_windows(lambda h: True))
            winapi.click(*at)
            time.sleep(1.2)
            pr = popup_rect(before)
            if pr:
                winapi.save_region_png(str(CAP / f"menu {name}.png"), pr[0], pr[1], pr[2] - pr[0], pr[3] - pr[1])
                out[name] = items_in(w, pr)
            else:
                out[name] = None
            print(name, len(out[name] or []), flush=True)
            winapi.press("esc")
            time.sleep(0.6)

        def find(name: str, kind: str):
            return uia.find(w.root, lambda x: x["name"] == name and x["type"] == kind and x["rect"][2] > 0, max_depth=16)

        for name in ("Rahmenarten", "Stiftfarbe", "Rahmen"):
            el = find(name, "SplitButton")
            if el is None:
                out[name] = "not found"
                continue
            more = uia.find(el, lambda x: x["type"] == "MenuItem", max_depth=3)
            r = uia.describe(more if more is not None else el)["rect"]
            open_menu(name, el, (r[0] + r[2] / 2, r[1] + r[3] / 2))
        for name in ("Stiftart", "Stiftstärke"):
            el = find(name, "ComboBox")
            if el is None:
                out[name] = "not found"
                continue
            btn = uia.find(el, lambda x: x["type"] == "Button", max_depth=3)
            r = uia.describe(btn if btn is not None else el)["rect"]
            open_menu(name, el, (r[0] + r[2] / 2, r[1] + r[3] / 2))
    finally:
        OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        w.close()


if __name__ == "__main__":
    main()
