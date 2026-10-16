"""Probe: typing over selected table cells in Word (Step 4d). In our own visible Word instance a
3 x 3 table ("Z1S1" … "Z3S3") gets a selection (all nine cells, the whole table with its row
ends, two cells, one whole row) and "a" is typed with the real keyboard; the table afterwards
(texts per cell) and the caret are recorded. Takes over keyboard and screen.
Result: results/<date>-word-table-typing.json.

    .venv/Scripts/python.exe probe_word_table_typing.py
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
OUT = HERE / "results" / f"{datetime.date.today()}-word-table-typing.json"


def state(app, d) -> dict:
    out: dict = {"tables": d.Tables.Count, "text": d.Content.Text.replace("\r", "|").replace("\x07", "¤")}
    if d.Tables.Count:
        t = d.Tables(1)
        out["cells"] = [f"{c.RowIndex},{c.ColumnIndex}:{c.Range.Text.rstrip(chr(13) + chr(7))}" for c in t.Range.Cells]
    s = app.Selection
    out["caret"] = {"start": s.Start, "end": s.End, "inTable": bool(s.Information(12)),
                    "cell": [s.Information(13), s.Information(16)] if s.Information(12) else None}
    return out


def main() -> None:
    import win32com.client
    sys.stdout.reconfigure(encoding="utf-8")
    app = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx("Word.Application"))
    app.Visible = True
    out: dict = {}
    try:
        d = app.Documents.Add()
        win = d.Windows(1)
        winapi.place(win.Hwnd, *WINDOW_RECT)
        winapi.activate(win.Hwnd)
        time.sleep(0.8)

        def fresh():
            d.Content.Delete()
            d.Content.Text = "Vor\rNach"
            t = d.Tables.Add(d.Range(d.Paragraphs(1).Range.End, d.Paragraphs(1).Range.End), 3, 3, 1, 0)
            for r in range(1, 4):
                for c in range(1, 4):
                    t.Cell(r, c).Range.Text = f"Z{r}S{c}"
            return t

        cases = {
            "all nine cells": lambda t: app.Selection.SetRange(t.Cell(1, 1).Range.Start, t.Cell(3, 3).Range.End - 1),
            "whole table": lambda t: t.Select(),
            "two cells in a row": lambda t: app.Selection.SetRange(t.Cell(2, 1).Range.Start, t.Cell(2, 2).Range.End - 1),
            "whole row": lambda t: t.Rows(2).Select(),
            "block 2x2": lambda t: app.Selection.SetRange(t.Cell(1, 1).Range.Start, t.Cell(2, 2).Range.End - 1),
        }
        for key in ("a", "backspace then a"):
            for name, select in cases.items():
                t = fresh()
                select(t)
                time.sleep(0.3)
                sel = {"start": app.Selection.Start, "end": app.Selection.End, "cells": app.Selection.Cells.Count}
                winapi.activate(win.Hwnd)
                if key.startswith("backspace"):
                    winapi.press("backspace")
                    time.sleep(0.6)
                    # Word may ask (Zellen löschen): leave that dialog.
                    winapi.press("esc")
                    time.sleep(0.3)
                winapi.type_text("a")
                time.sleep(0.5)
                res = {"selection": sel, **state(app, d)}
                out[f"{key}: {name}"] = res
                print(key, name, json.dumps(res, ensure_ascii=False)[:400], flush=True)
        d.Close(0)
    finally:
        app.Quit(0)
        OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
