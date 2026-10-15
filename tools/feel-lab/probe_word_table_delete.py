"""Probe: Entf and Rücktaste on selected table cells in Word (Step 4c D). For each selection
(part of a cell, one whole cell, several cells, whole rows, whole columns, the whole table) the
key is pressed in our own visible Word instance; a dialog that opens is recorded (title, the
options UI Automation sees) and closed with Esc; otherwise the table afterwards (rows,
columns, texts). Takes over keyboard and screen. Result: results/<date>-word-table-delete.json.

    .venv/Scripts/python.exe probe_word_table_delete.py
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
OUT = HERE / "results" / f"{datetime.date.today()}-word-table-delete.json"


def table_state(d) -> dict:
    if d.Tables.Count == 0:
        return {"tables": 0, "text": d.Content.Text.replace("\r", "|")}
    t = d.Tables(1)
    cells = []
    for c in t.Range.Cells:
        cells.append(f"{c.RowIndex},{c.ColumnIndex}:{c.Range.Text.rstrip(chr(13) + chr(7))}")
    return {"tables": d.Tables.Count, "rows": t.Rows.Count, "cols": t.Columns.Count, "cells": cells}


def dialogs(before: set[int]) -> list[int]:
    return [h for h in winapi.find_windows(lambda h: True) if h not in before and winapi.window_class(h) in ("bosa_sdm_msword", "NUIDialog", "#32770")]


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
            "part of a cell": lambda t: d.Range(t.Cell(2, 2).Range.Start, t.Cell(2, 2).Range.Start + 2).Select(),
            "one whole cell": lambda t: t.Cell(2, 2).Select(),
            "two cells in a row": lambda t: d.Range(t.Cell(2, 1).Range.Start, t.Cell(2, 2).Range.End).Select(),
            "block 2x2": lambda t: app.Selection.SetRange(t.Cell(1, 1).Range.Start, t.Cell(2, 2).Range.End),
            "whole row": lambda t: t.Rows(2).Select(),
            "two whole rows": lambda t: d.Range(t.Rows(1).Range.Start, t.Rows(2).Range.End).Select(),
            "whole column": lambda t: t.Columns(2).Select(),
            "whole table": lambda t: t.Select(),
        }
        for key in ("delete", "backspace"):
            for name, select in cases.items():
                t = fresh()
                time.sleep(0.2)
                select(t)
                time.sleep(0.3)
                sel = {"start": app.Selection.Start, "end": app.Selection.End, "cells": app.Selection.Cells.Count if app.Selection.Information(12) else 0}
                before = set(winapi.find_windows(lambda h: True))
                winapi.activate(win.Hwnd)
                winapi.press(key)
                time.sleep(1.0)
                found = dialogs(before)
                res: dict = {"selection": sel}
                if found:
                    h = found[0]
                    el = uia.element_for(h)
                    f = el.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
                    items = []
                    for i in range(f.Length):
                        e = f.GetElement(i)
                        dd = uia.describe(e)
                        if dd["name"] and dd["type"] in ("RadioButton", "Button", "Text", "CheckBox"):
                            on = ""
                            if dd["type"] == "RadioButton":
                                try:
                                    pat = e.GetCurrentPattern(uia.U.UIA_SelectionItemPatternId).QueryInterface(uia.U.IUIAutomationSelectionItemPattern)
                                    on = "*" if pat.CurrentIsSelected else ""
                                except Exception:  # noqa: BLE001
                                    pass
                            items.append(f"{dd['type']}:{dd['name']}{on}")
                    res["dialog"] = {"title": winapi.window_title(h), "items": items}
                    winapi.press("esc")
                    time.sleep(0.5)
                res["after"] = table_state(d)
                out[f"{key}: {name}"] = res
                print(key, name, json.dumps(res, ensure_ascii=False)[:300], flush=True)
        # What the dialog's four choices do (the same operation through COM).
        for name in ("one whole cell", "two cells in a row", "block 2x2"):
            for k, label in enumerate(("Zellen nach links verschieben", "Zellen nach oben verschieben", "Ganze Zeile löschen", "Ganze Spalte löschen")):
                t = fresh()
                cases[name](t)
                app.Selection.Cells.Delete(k)
                out[f"dialog: {name}: {label}"] = {"after": table_state(d)}
                print("dialog", name, label, json.dumps(out[f"dialog: {name}: {label}"], ensure_ascii=False)[:300], flush=True)
        d.Close(0)
    finally:
        app.Quit(0)
        OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
