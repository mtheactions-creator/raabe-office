"""Probe: dragging a table by its move handle in Word (Step 4c D). In our own visible Word
instance a 3 × 3 table between two paragraphs is dragged by the handle at its top left (right,
down, and both); the saved file shows what Word makes of it (w:tblpPr: anchors, offsets, text
distances) and where the table lands. Takes over mouse and screen.
Result: results/<date>-word-table-move.json.

    .venv/Scripts/python.exe probe_word_table_move.py
"""
from __future__ import annotations

import datetime
import json
import re
import sys
import tempfile
import time
import zipfile
from pathlib import Path

from feellab import winapi
from feellab.targets import WINDOW_RECT

HERE = Path(__file__).parent
OUT = HERE / "results" / f"{datetime.date.today()}-word-table-move.json"
TMP = Path(tempfile.gettempdir()) / "raabe-table-move.docx"
TEXT = "Ein Absatz mit etwas Text, damit die Tabelle umflossen werden kann. " * 4


def main() -> None:
    import win32com.client
    sys.stdout.reconfigure(encoding="utf-8")
    app = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx("Word.Application"))
    app.Visible = True
    out: dict = {}
    try:
        for name, (dx, dy) in {"right 150": (150, 0), "down 120": (0, 120), "right 120 down 90": (120, 90)}.items():
            d = app.Documents.Add()
            d.Content.Text = "\r".join([TEXT, "", TEXT, TEXT])
            t = d.Tables.Add(d.Paragraphs(2).Range, 3, 3, 1, 0)
            t.PreferredWidthType = 3
            t.PreferredWidth = 200
            for i in range(1, 4):
                for j in range(1, 4):
                    t.Cell(i, j).Range.Text = f"Z{i}S{j}"
            win = d.Windows(1)
            win.View.Type = 3
            win.View.Zoom.Percentage = 100
            winapi.place(win.Hwnd, *WINDOW_RECT)
            winapi.activate(win.Hwnd)
            time.sleep(0.6)
            t.Cell(2, 2).Range.Select()
            app.Selection.Collapse(1)
            time.sleep(0.4)
            left, top, _, _ = win.GetPoint(0, 0, 0, 0, t.Range)
            # The move handle: 23 px left of and 11 px above the first cell's text (100 %, measured
            # on a screenshot); it shows while the pointer is over the table.
            hx, hy = left - 23, top - 11
            winapi.move(left + 30, top + 10)
            time.sleep(0.6)
            winapi.move(hx, hy)
            time.sleep(0.4)
            winapi.drag(hx, hy, hx + dx, hy + dy, steps=25)
            time.sleep(0.8)
            # Where things are now (points from the page's top; Word's own report).
            t2 = d.Tables(1)
            pos = {
                "table top": t2.Cell(1, 1).Range.Information(6),
                "para1 first line": d.Paragraphs(1).Range.Characters(1).Information(6),
                "para1 last line": d.Paragraphs(1).Range.Characters(d.Paragraphs(1).Range.Characters.Count).Information(6),
            }
            after = d.Range(t2.Range.End, t2.Range.End).Paragraphs(1).Range
            pos["para after table first line"] = after.Characters(1).Information(6)
            pos["para after table text"] = after.Text[:20]
            d.SaveAs2(str(TMP), 16)
            d.Close(0)
            with zipfile.ZipFile(TMP) as z:
                doc = z.read("word/document.xml").decode("utf-8")
            tblpr = re.search(r"<w:tblPr>(.*?)</w:tblPr>", doc)
            body = re.search(r"<w:body>(.*)</w:body>", doc, re.S).group(1)
            order = re.findall(r"<w:(tbl|p)\b", body)
            out[name] = {"tblPr": tblpr.group(1) if tblpr else None, "order": "".join("T" if x == "tbl" else "p" for x in order), "positions": pos}
            print(name, json.dumps(out[name], ensure_ascii=False)[:600], flush=True)
    finally:
        OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        try:
            app.Quit(0)
        except Exception:  # noqa: BLE001
            pass


if __name__ == "__main__":
    main()
