"""Probe: Word's task panes Grafik formatieren and Form formatieren (Step 4c B). Opens each
with the group's dialog launcher in our own visible Word instance, then for every tab of the
pane expands all sections and records what UI Automation sees (names, types, positions) and a
screenshot. Takes over mouse and screen. Screenshots: captures/word-dialogs/ (local); names:
results/<date>-word-format-pane.json.

    .venv/Scripts/python.exe probe_word_format_pane.py
"""
from __future__ import annotations

import datetime
import json
import sys
import time
from pathlib import Path

from feellab import uia, winapi
from probe_word_dialogs import CAP, Word, snapshot
from probe_word_panes import close_pane, pane_named

HERE = Path(__file__).parent
OUT = HERE / "results" / f"{datetime.date.today()}-word-format-pane.json"
PICTURE = HERE.parents[1] / "branding" / "original-logo-reference.png"


def expand_all(root, area: tuple[int, int, int]) -> int:
    """Opens every collapsed section of the pane (ExpandCollapse pattern; only elements inside
    the pane's columns of Word's window); returns how many."""
    n = 0
    ax, ay, aw = area
    for _ in range(40):
        f = root.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
        hit = None
        for i in range(f.Length):
            el = f.GetElement(i)
            try:
                p = el.GetCurrentPattern(uia.U.UIA_ExpandCollapsePatternId)
                if not p:
                    continue
                ec = p.QueryInterface(uia.U.IUIAutomationExpandCollapsePattern)
                d = uia.describe(el)
                if not (ax - 2 <= d["rect"][0] <= ax + aw and d["rect"][1] >= ay):
                    continue
                if ec.CurrentExpandCollapseState == 0 and d["type"] in ("Button", "TreeItem", "Group", "Custom", "ListItem") and d["name"]:
                    hit = (ec, d)
                    break
            except Exception:  # noqa: BLE001
                continue
        if hit is None:
            break
        try:
            hit[0].Expand()
            n += 1
        except Exception:  # noqa: BLE001
            break
        time.sleep(0.4)
    return n


def pane_items(w: Word, name: str) -> tuple[list[dict], list[int]] | None:
    """The task pane's elements: everything in Word's own window right of and below the pane's
    title bar (UI Automation names only the title). Never reads outside Word's window."""
    title = pane_named(w, name)
    if title is None:
        return None
    tx, ty, tw, _ = uia.describe(title)["rect"]
    l, t, r, b = winapi.window_rect(w.hwnd)
    items = [x for x in snapshot(w.root, (tx, ty)) if x["rect"][0] >= -2 and x["rect"][1] >= 0 and x["rect"][0] + x["rect"][2] <= tw + 4]
    return items, [tx, ty, tw, b - ty]


def record(w: Word, pane_name: str, key: str, out: dict) -> None:
    got = pane_items(w, pane_name)
    if got is None:
        print("no pane", pane_name, flush=True)
        out[key] = {"error": "no pane"}
        return
    first, rect = got
    winapi.save_region_png(str(CAP / f"{key}.png"), *rect)
    res: dict = {"rect": rect, "first": first, "tabs": {}}
    print(key, "items", [(x["type"], x["name"]) for x in first][:40], flush=True)
    def elements(kind: str) -> list:
        """(element, description) of the given type inside the pane's columns of Word's window."""
        title = pane_named(w, pane_name)
        if title is None:
            return []
        tx, ty, tw, _ = uia.describe(title)["rect"]
        f = w.root.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
        out = []
        for i in range(f.Length):
            el = f.GetElement(i)
            d = uia.describe(el)
            if d["type"] == kind and d["name"] and tx - 2 <= d["rect"][0] <= tx + tw and d["rect"][1] >= ty:
                out.append((el, d))
        return out

    def activate(el, d) -> None:
        try:
            uia.select(el)
        except Exception:  # noqa: BLE001
            try:
                uia.invoke(el)
            except Exception:  # noqa: BLE001
                r = d["rect"]
                winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
        time.sleep(1.0)

    def open_sections() -> int:
        """Opens every section with a click on its header, the lowest first (so the headers
        above stay where they are); a click that closed one is undone."""
        n = 0
        groups = sorted(((d["name"], d["rect"][1]) for _, d in elements("Group")), key=lambda g: -g[1])
        for g, _ in groups:
            before = len(pane_items(w, pane_name)[0])
            hit = next((d for _, d in elements("Button") if d["name"] == g), None)
            if hit is None:
                continue
            r = hit["rect"]
            winapi.click(r[0] + 40, r[1] + r[3] / 2)
            time.sleep(0.7)
            after = len(pane_items(w, pane_name)[0])
            if after < before:
                winapi.click(r[0] + 40, r[1] + r[3] / 2)
                time.sleep(0.7)
            n += 1
        return n

    def visit_tabs(prefix: str) -> None:
        for name in [d["name"] for _, d in elements("ListItem")]:
            hit = next(((el, d) for el, d in elements("ListItem") if d["name"] == name), None)
            if hit is None:
                continue
            activate(*hit)
            opened = open_sections()
            items, rect = pane_items(w, pane_name)
            winapi.save_region_png(str(CAP / f"{key} {prefix}{name}.png"), *rect)
            res["tabs"][prefix + name] = {"opened": opened, "items": items}
            print(key, prefix + name, opened, len(items), flush=True)

    outer = [d["name"] for _, d in elements("TabItem")]
    if outer:
        for o in outer:
            hit = next(((el, d) for el, d in elements("TabItem") if d["name"] == o), None)
            if hit:
                activate(*hit)
                visit_tabs(o + " - ")
    else:
        visit_tabs("")
    out[key] = res
    close_pane(w, pane_named(w, pane_name))


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    out: dict = {}
    w = Word()
    d = w.d
    try:
        d.Content.Text = "Ein Absatz mit einem Bild.\r"
        pic = d.InlineShapes.AddPicture(str(PICTURE), False, True, d.Paragraphs(2).Range)
        pic.Width = 160
        pic.Height = 120
        pic.Select()
        time.sleep(0.8)
        w.tab("Bildformat")
        el = w.group_launcher("Bildformatvorlagen")
        if el is None:
            print("no launcher Bildformatvorlagen", flush=True)
        else:
            r = uia.describe(el)["rect"]
            winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
            time.sleep(2.0)
            record(w, "Grafik formatieren", "format picture pane", out)
        shp = d.Shapes.AddShape(1, 200, 160, 140, 80, d.Paragraphs(1).Range)
        shp.TextFrame.TextRange.Text = "Text"
        shp.Select()
        time.sleep(0.8)
        w.tab("Formformat")
        el = w.group_launcher("Formenarten")
        if el is None:
            print("no launcher Formenarten", flush=True)
        else:
            r = uia.describe(el)["rect"]
            winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
            time.sleep(2.0)
            record(w, "Form formatieren", "format shape pane", out)
    finally:
        OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        w.close()


if __name__ == "__main__":
    main()
