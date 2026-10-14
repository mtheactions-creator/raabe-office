"""Probe: Word's context menus (right click) in text, a list, a table cell and on a picture,
with their submenus. Own Word instance; takes over mouse and screen briefly. Writes the
items (name, enabled, submenu items) to results/<date>-word-context.json and screenshots to
captures/word-context/ (local only).

    .venv/Scripts/python.exe probe_word_context.py
"""
from __future__ import annotations

import datetime
import json
import tempfile
import time
from pathlib import Path

from feellab import uia, winapi
from feellab.targets import WINDOW_RECT

HERE = Path(__file__).parent
CAP = HERE / "captures" / "word-context"
CAP.mkdir(parents=True, exist_ok=True)
PNG = Path(tempfile.gettempdir()) / "raabe-context-logo.png"


def make_png(path: Path) -> None:
    """A small plain PNG (no external file), written by hand."""
    import struct
    import zlib
    w, h = 120, 80
    raw = b"".join(b"\x00" + bytes((40, 110, 200)) * w for _ in range(h))
    def chunk(t, d):
        c = struct.pack(">I", len(d)) + t + d
        return c + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
    data = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)) + \
        chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")
    path.write_bytes(data)


def popups() -> list[int]:
    """Word's open popups (context menu, mini toolbar, submenus): "Net UI Tool Window"s."""
    return [h for h in winapi.find_windows(lambda h: winapi.window_class(h) == "Net UI Tool Window")]


def items_of(hwnd: int) -> list:
    """The controls inside a popup window (UIA may hand back more of Office's tree)."""
    l, t, r, b = winapi.window_rect(hwnd)
    out = []
    stack = [uia.element_for(hwnd)]
    while stack:
        e = stack.pop()
        for c in uia.children(e):
            dc = uia.describe(c)
            x, y, w, h = dc["rect"]
            inside = x >= l - 2 and y >= t - 2 and x + w <= r + 2 and y + h <= b + 2 and w > 0
            if dc["type"] in ("MenuItem", "Button", "SplitButton", "ComboBox", "Edit", "Separator", "CheckBox") and inside:
                out.append((c, dc))
            if dc["type"] in ("Menu", "Group", "ToolBar", "Pane", "Custom", "MenuBar", "List"):
                stack.append(c)
    out.sort(key=lambda x: (x[1]["rect"][1], x[1]["rect"][0]))
    return out


def menu_items(timeout: float = 2.0, exclude: set | None = None) -> tuple[list, list]:
    """The menu (the tallest popup) and the mini toolbar items."""
    end = time.time() + timeout
    while time.time() < end:
        hs = [h for h in popups() if h not in (exclude or set())]
        if hs:
            time.sleep(0.3)
            hs = [h for h in popups() if h not in (exclude or set())]
            rects = {h: winapi.window_rect(h) for h in hs}
            menu = max(hs, key=lambda h: rects[h][3] - rects[h][1])
            others = [h for h in hs if h != menu]
            return [menu] + others, items_of(menu)
        time.sleep(0.15)
    return [], []


def capture(kind: str, x: float, y: float, click_first: bool = True) -> dict:
    if click_first:
        winapi.click(x, y)
        time.sleep(0.3)
    before = set(popups())
    winapi.right_click(x, y)
    time.sleep(1.2)
    hs, items = menu_items(exclude=before)
    out = {"kind": kind, "items": [], "toolbar": []}
    if hs:
        r = winapi.window_rect(hs[0])
        x0 = min(winapi.window_rect(h)[0] for h in hs)
        y0 = min(winapi.window_rect(h)[1] for h in hs)
        winapi.save_region_png(str(CAP / f"{kind}.png"), x0 - 10, y0 - 10, max(winapi.window_rect(h)[2] for h in hs) - x0 + 20, r[3] - y0 + 20)
        for h in hs[1:]:
            out["toolbar"] += [{"type": d["type"], "name": d["name"], "enabled": d["enabled"]} for _, d in items_of(h)]
    for el, d in items:
        out["items"].append({"type": d["type"], "name": d["name"], "enabled": d["enabled"], "rect": d["rect"]})
    # Submenus: hovering an item opens one (a new popup window).
    known = set(hs) | before
    for i, (el, d) in enumerate(items):
        if d["type"] != "MenuItem" or not d["enabled"]:
            continue
        x0, y0, w0, h0 = d["rect"]
        winapi.move(x0 + w0 / 2, y0 + h0 / 2)
        time.sleep(0.9)
        new = [h for h in popups() if h not in known]
        if new:
            sub = []
            for h in new:
                sub += [{"type": dd["type"], "name": dd["name"], "enabled": dd["enabled"]} for _, dd in items_of(h)]
            out["items"][i]["submenu"] = sub
            r2 = winapi.window_rect(new[0])
            winapi.save_region_png(str(CAP / f"{kind}-sub{i}.png"), r2[0] - 10, r2[1] - 10, r2[2] - r2[0] + 20, r2[3] - r2[1] + 20)
            # Back onto the main menu so the submenu closes.
            winapi.move(x0 + 10, y0 + h0 / 2)
            time.sleep(0.5)
    winapi.press("esc")
    time.sleep(0.2)
    winapi.press("esc")
    time.sleep(0.4)
    return out


def main() -> None:
    import win32com.client
    make_png(PNG)
    app = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx("Word.Application"))
    results = []
    try:
        app.Visible = True
        d = app.Documents.Add()
        d.ShowSpellingErrors = False
        d.ShowGrammaticalErrors = False
        d.Content.Text = "\r".join(["Ein Absatz mit etwas Text zum Anklicken.", "Erster Punkt", "Zweiter Punkt", "Vor der Tabelle", "Nach der Tabelle", ""])
        d.Range(d.Paragraphs(2).Range.Start, d.Paragraphs(3).Range.End).ListFormat.ApplyListTemplate(app.ListGalleries(1).ListTemplates(1), False, 0, 2)
        r = d.Range(d.Paragraphs(4).Range.End, d.Paragraphs(4).Range.End)
        t = d.Tables.Add(r, 2, 2, 1, 0)
        for i in range(1, 3):
            for k in range(1, 3):
                t.Cell(i, k).Range.Text = f"Zelle{i}{k}"
        last = d.Paragraphs(d.Paragraphs.Count).Range
        last.InlineShapes.AddPicture(str(PNG), False, True)
        win = d.Windows(1)
        win.View.Type = 3
        win.View.Zoom.Percentage = 100
        hwnd = win.Hwnd
        winapi.place(hwnd, *WINDOW_RECT)
        winapi.activate(hwnd)
        time.sleep(0.8)

        def at(rg) -> tuple[float, float]:
            left, top, width, height = win.GetPoint(0, 0, 0, 0, rg)
            return float(left) + 4, float(top) + float(height) / 2

        def find(text: str):
            rg = d.Content.Duplicate
            rg.Find.Execute(FindText=text)
            return d.Range(rg.Start + 2, rg.Start + 2)

        results.append(capture("text", *at(find("etwas"))))
        results.append(capture("list", *at(find("Zweiter"))))
        results.append(capture("table", *at(find("Zelle12"))))
        # several cells selected
        app.Selection.SetRange(t.Cell(1, 1).Range.Start, t.Cell(2, 2).Range.End)
        sel_pt = at(find("Zelle21"))
        results.append(capture("table-cells", *sel_pt, click_first=False))
        shp = d.InlineShapes(1)
        left, top, width, height = win.GetPoint(0, 0, 0, 0, shp.Range)
        results.append(capture("picture", float(left) + float(width) / 2, float(top) + float(height) / 2))
        d.Close(0)
    finally:
        app.Quit(0)
    out = HERE / "results" / f"{datetime.date.today().isoformat()}-word-context.json"
    out.write_text(json.dumps(results, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    for r in results:
        print("==", r["kind"])
        print("   toolbar:", ", ".join(t["name"] for t in r.get("toolbar", [])))
        for it in r["items"]:
            print("  ", it["type"], it["name"], "" if it["enabled"] else "(disabled)")
            for s in it.get("submenu", []):
                print("      >", s["name"], "" if s["enabled"] else "(disabled)")


if __name__ == "__main__":
    main()
