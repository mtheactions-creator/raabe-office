"""Probe: selecting table cells in Word (Step 4c D). (1) Screenshots of selected cells (an
empty cell, a multi-line cell, short text, a taller row) at 100 % zoom, and the same table saved
as packages/ooxml/test/fixtures/word-table-select.docx for comparing our drawing. (2) A mouse
drag from the text of one cell to the right and down, step by step, recording after each step
what Word selects (characters or cells). Own visible Word instance; takes over mouse and screen.
Result: results/<date>-word-table-select.json, screenshots in captures/word-table/ (local).

    .venv/Scripts/python.exe probe_word_table_select.py
"""
from __future__ import annotations

import datetime
import json
import sys
import time
from pathlib import Path

from feellab import winapi
from feellab.targets import WINDOW_RECT

HERE = Path(__file__).parent
CAP = HERE / "captures" / "word-table"
CAP.mkdir(parents=True, exist_ok=True)
OUT = HERE / "results" / f"{datetime.date.today()}-word-table-select.json"
FIXTURE = HERE.parents[1] / "packages" / "ooxml" / "test" / "fixtures" / "word-table-select.docx"


def main() -> None:
    import win32com.client
    sys.stdout.reconfigure(encoding="utf-8")
    app = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx("Word.Application"))
    app.Visible = True
    out: dict = {}
    try:
        d = app.Documents.Add()
        d.Content.Text = "Auswahl in Tabellen\r"
        t = d.Tables.Add(d.Range(d.Content.End - 1, d.Content.End - 1), 3, 3, 1, 0)
        texts = [["Kurz", "", "Zwei Wörter"],
                 ["Ein längerer Text, der in dieser Zelle über mehrere Zeilen umbrochen wird.", "Mitte", ""],
                 ["Unten", "Text", "Ende"]]
        for r in range(3):
            for c in range(3):
                t.Cell(r + 1, c + 1).Range.Text = texts[r][c]
        d.SaveAs2(str(FIXTURE), 16)
        win = d.Windows(1)
        win.View.Type = 3
        win.View.Zoom.Percentage = 100
        winapi.place(win.Hwnd, *WINDOW_RECT)
        winapi.activate(win.Hwnd)
        time.sleep(1.0)

        def shot(name: str) -> None:
            left, top, width, height = win.GetPoint(0, 0, 0, 0, t.Range)
            winapi.save_region_png(str(CAP / f"{name}.png"), int(left) - 20, int(top) - 20, int(width) + 40, int(height) + 40)

        selections = {
            "block 2x2 from empty": lambda: app.Selection.SetRange(t.Cell(1, 2).Range.Start, t.Cell(2, 3).Range.End),
            "whole table": lambda: t.Select(),
            "one cell multi-line": lambda: t.Cell(2, 1).Select(),
            "one empty cell": lambda: t.Cell(1, 2).Select(),
            "text in one cell": lambda: d.Range(t.Cell(2, 1).Range.Start + 4, t.Cell(2, 1).Range.Start + 30).Select(),
        }
        for name, sel in selections.items():
            sel()
            time.sleep(0.6)
            shot(name)
            out[f"shot: {name}"] = f"captures/word-table/{name}.png"
        # The drag: from inside "Kurz" to the right, across the border into the next cell.
        r11 = t.Cell(1, 1).Range
        x0, y0, _, h0 = win.GetPoint(0, 0, 0, 0, d.Range(r11.Start + 1, r11.Start + 1))
        c2 = t.Cell(1, 2).Range
        x2, _, _, _ = win.GetPoint(0, 0, 0, 0, d.Range(c2.Start, c2.Start))
        c3 = t.Cell(1, 3).Range
        x3, _, _, _ = win.GetPoint(0, 0, 0, 0, d.Range(c3.Start, c3.Start))
        y = y0 + h0 / 2
        # One drag per end point, the selection read after the button is let go (Word must
        # not be asked during a drag).
        steps = []
        x = x0
        while x < x3 + 40:
            x += 4
            winapi.drag(x0, y, x, y, steps=8)
            time.sleep(0.25)
            s = app.Selection
            steps.append({"dx": round(x - x0), "x_from_cell2_text": round(x - x2), "start": s.Start, "end": s.End, "cells": s.Cells.Count, "text": s.Text[:20]})
        out["drag right"] = {"cell2_text_x": round(x2 - x0), "cell3_text_x": round(x3 - x0), "steps": steps}
        # Downwards from "Kurz" into the multi-line cell below.
        r21 = t.Cell(2, 1).Range
        _, y2, _, _ = win.GetPoint(0, 0, 0, 0, d.Range(r21.Start, r21.Start))
        steps = []
        yy = y
        while yy < y2 + 30:
            yy += 3
            winapi.drag(x0, y, x0, yy, steps=8)
            time.sleep(0.25)
            s = app.Selection
            steps.append({"dy": round(yy - y), "start": s.Start, "end": s.End, "cells": s.Cells.Count, "text": s.Text[:20]})
        out["drag down"] = {"cell_below_text_y": round(y2 - y), "steps": steps}
        d.Close(0)
    finally:
        OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        try:
            app.Quit(0)
        except Exception:  # noqa: BLE001
            pass
    print(json.dumps({k: (v if isinstance(v, str) else {kk: vv for kk, vv in v.items() if kk != "steps"}) for k, v in out.items()}, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
