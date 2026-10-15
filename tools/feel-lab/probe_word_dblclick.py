"""Probe: what a double-click does on each kind of object in Word (Step 4c F). In our own
visible Word instance a document with a footnote, a section break, a table and paragraphs is
double-clicked at: the footnote reference, the footnote's number in the note, a column border
of the table, the table's move handle, a cell's selection area, the left margin, a ribbon tab
and the section break (with Alle anzeigen on). Recorded after each: the selection (start, end,
type, story), a dialog that opened (title), the ribbon state and the table's column widths.
Takes over mouse and screen. Result: results/<date>-word-dblclick.json.

    .venv/Scripts/python.exe probe_word_dblclick.py [case words …]
"""
from __future__ import annotations

import datetime
import json
import sys
import time
from pathlib import Path

from feellab import uia, winapi
from feellab.targets import WINDOW_RECT

HERE = Path(__file__).parent
OUT = HERE / "results" / f"{datetime.date.today()}-word-dblclick.json"
TEXT = "Ein Absatz mit etwas Text, damit die Zeile nicht leer ist und man gut hineinklicken kann."


def dialogs(before: set[int]) -> list[int]:
    return [h for h in winapi.find_windows(lambda h: True) if h not in before and winapi.window_class(h) in ("bosa_sdm_msword", "NUIDialog", "#32770")]


def main() -> None:
    import win32com.client
    sys.stdout.reconfigure(encoding="utf-8")
    app = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx("Word.Application"))
    app.Visible = True
    out: dict = json.loads(OUT.read_text(encoding="utf-8")) if OUT.exists() and sys.argv[1:] else {}
    try:
        d = app.Documents.Add()
        win = d.Windows(1)
        win.View.Type = 3
        win.View.Zoom.Percentage = 100
        winapi.place(win.Hwnd, *WINDOW_RECT)
        winapi.activate(win.Hwnd)
        time.sleep(0.8)

        def build():
            d.Content.Delete()
            d.Content.Text = "\r".join([TEXT, TEXT, "", TEXT])
            d.Footnotes.Add(d.Range(d.Paragraphs(1).Range.End - 2, d.Paragraphs(1).Range.End - 2), Text="Eine Fußnote.")
            t = d.Tables.Add(d.Paragraphs(3).Range, 2, 3, 1, 0)
            t.Cell(1, 1).Range.Text = "kurz"
            t.Cell(1, 2).Range.Text = "ein längerer Text in der Zelle"
            t.Cell(2, 1).Range.Text = "Zelle"
            r = d.Range(d.Paragraphs(2).Range.End - 1, d.Paragraphs(2).Range.End - 1)
            r.InsertBreak(3)  # Abschnittswechsel (fortlaufend)
            win.View.ShowAll = False
            win.View.SeekView = 0
            d.Range(0, 0).Select()
            time.sleep(0.5)
            return d.Tables(1)

        def state() -> dict:
            s = app.Selection
            res = {"start": s.Start, "end": s.End, "type": s.Type, "story": s.StoryType, "text": s.Text[:40].replace("\r", "|")}
            # The ribbon: the height of its lower part (0 when reduced).
            lower = uia.find(uia.element_for(win.Hwnd), lambda x: x["type"] == "Pane" and x["name"].startswith("Unter"), max_depth=12)
            res["ribbonHeight"] = uia.describe(lower)["rect"][3] if lower is not None else 0
            if d.Tables.Count:
                t = d.Tables(1)
                res["widths"] = [round(t.Cell(1, k).Width, 1) for k in range(1, 4)]
            return res

        def point(rg) -> tuple[float, float, float, float]:
            left, top, width, height = win.GetPoint(0, 0, 0, 0, rg)
            return float(left), float(top), float(width), float(height)

        def run(name: str, where):
            t = build()
            x, y = where(t)
            before = set(winapi.find_windows(lambda h: True))
            winapi.activate(win.Hwnd)
            winapi.move(x, y)
            time.sleep(0.4)
            winapi.click(x, y, count=2)
            time.sleep(1.2)
            # A dialog first: while it is open Word rejects COM calls.
            found = dialogs(before)
            title = winapi.window_title(found[0]) if found else None
            if found:
                winapi.press("esc")
                time.sleep(0.6)
            res = state()
            if title:
                res["dialog"] = title
            if name == "ribbon tab":
                # Back as it was: the same double-click again.
                winapi.click(x, y, count=2)
                time.sleep(1.0)
                res["ribbonHeightAfterSecond"] = state()["ribbonHeight"]
            out[name] = res
            print(name, json.dumps(res, ensure_ascii=False), flush=True)
            win.View.SeekView = 0
            time.sleep(0.3)

        def ref(t):
            l, tp, w, h = point(d.Footnotes(1).Reference)
            return l + w / 2, tp + h / 2
        only = sys.argv[1:]
        if not only or "footnote" in only:
            run("footnote reference", ref)

        def note_number(t):
            win.View.SeekView = 0
            l, tp, w, h = point(d.Footnotes(1).Range.Paragraphs(1).Range.Characters(1))
            return l + 2, tp + h / 2
        if not only or "footnote" in only:
            run("footnote number in the note", note_number)

        def col_border(t):
            l, tp, w, h = point(t.Cell(1, 1).Range)
            l2, _, _, _ = point(t.Cell(1, 2).Range)
            # The border between the first two columns: the second cell's text start less its margin
            # (5.4 pt, in screen pixels at 150 % scaling).
            return l2 - 5.4 * 2, tp + h / 2
        if not only or any(o in "column border" for o in only):
            run("column border", col_border)

        def handle(t):
            l, tp, w, h = point(t.Range)
            winapi.move(l + 30, tp + 10)
            time.sleep(0.6)
            return l - 23, tp - 11
        if not only or any(o in "table move handle" for o in only):
            run("table move handle", handle)

        def cell_bar(t):
            l, tp, w, h = point(t.Cell(2, 1).Range)
            return l - 3, tp + h / 2
        if not only or any(o in "cell selection area" for o in only):
            run("cell selection area", cell_bar)

        def margin(t):
            l, tp, w, h = point(d.Paragraphs(1).Range.Characters(1))
            return l - 40, tp + h / 2
        if not only or any(o in "left margin" for o in only):
            run("left margin", margin)

        def tab(t):
            root = uia.element_for(win.Hwnd)
            el = uia.find(root, lambda x: x["type"] == "TabItem" and x["name"] == "Einfügen", max_depth=14)
            r = uia.describe(el)["rect"]
            return r[0] + r[2] / 2, r[1] + r[3] / 2
        if not only or any(o in "ribbon tab" for o in only):
            run("ribbon tab", tab)

        def section_break(t):
            win.View.ShowAll = True
            time.sleep(0.4)
            pos = d.Sections(1).Range.End - 1
            l, tp, w, h = point(d.Range(pos, pos + 1))
            return l + max(w, 30) / 2, tp + h / 2
        if not only or any(o in "section break" for o in only):
            run("section break", section_break)
        win.View.ShowAll = False
        d.Close(0)
    finally:
        OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        # Word keeps these for the user's next session: put them back as they were (Word's defaults).
        try:
            if app.CommandBars.GetPressedMso("MinimizeRibbon"):
                app.CommandBars.ExecuteMso("MinimizeRibbon")
            if app.Documents.Count:
                app.ActiveWindow.View.ShowAll = False
        except Exception:  # noqa: BLE001
            pass
        app.Quit(0)


if __name__ == "__main__":
    main()
