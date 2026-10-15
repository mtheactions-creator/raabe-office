"""Probe: how Word draws tables in some of its gallery styles (Step 4c D). Builds
packages/ooxml/test/fixtures/word-table-styles.docx (five 4 × 4 tables, each in another style,
Kopfzeile and Erste Spalte on, a total row on for the last) in our own visible Word instance and
takes a screenshot of each table at 100 % for comparing RaabeWrite's drawing. Takes over the
screen. Screenshots: captures/word-table/style *.png (local).

    .venv/Scripts/python.exe probe_word_table_looks.py
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

from feellab import winapi
from feellab.targets import WINDOW_RECT

HERE = Path(__file__).parent
CAP = HERE / "captures" / "word-table"
CAP.mkdir(parents=True, exist_ok=True)
FIXTURE = HERE.parents[1] / "packages" / "ooxml" / "test" / "fixtures" / "word-table-styles.docx"
STYLES = ["Gitternetztabelle 4 – Akzent 1", "Gitternetztabelle 1 hell  – Akzent 3", "Listentabelle 3 – Akzent 2", "Einfache Tabelle 3",
          "Gitternetztabelle 6 farbig – Akzent 5"]


def main() -> None:
    import win32com.client
    sys.stdout.reconfigure(encoding="utf-8")
    app = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx("Word.Application"))
    app.Visible = True
    try:
        d = app.Documents.Add()
        tables = []
        for k, name in enumerate(STYLES):
            r = d.Range(d.Content.End - 1, d.Content.End - 1)
            r.InsertAfter(f"Tabelle {k + 1}: {name}")
            r = d.Range(d.Content.End - 1, d.Content.End - 1)
            r.InsertParagraphAfter()
            r = d.Range(d.Content.End - 1, d.Content.End - 1)
            t = d.Tables.Add(r, 4, 4)
            t.Style = name
            t.ApplyStyleHeadingRows = True
            t.ApplyStyleFirstColumn = True
            t.ApplyStyleLastRow = k == len(STYLES) - 1
            for i in range(1, 5):
                for j in range(1, 5):
                    t.Cell(i, j).Range.Text = "Kopf" if i == 1 else f"Z{i}S{j}"
            tables.append(t)
            r = d.Range(d.Content.End - 1, d.Content.End - 1)
            r.InsertParagraphAfter()
        d.SaveAs2(str(FIXTURE), 16)
        win = d.Windows(1)
        win.View.Type = 3
        win.View.Zoom.Percentage = 100
        winapi.place(win.Hwnd, *WINDOW_RECT)
        winapi.activate(win.Hwnd)
        for k, t in enumerate(tables):
            t.Range.Cells(1).Range.Select()
            app.Selection.Collapse(1)
            win.ScrollIntoView(t.Range, True)
            time.sleep(0.8)
            left, top, width, height = win.GetPoint(0, 0, 0, 0, t.Range)
            winapi.save_region_png(str(CAP / f"style {k + 1}.png"), int(left) - 6, int(top) - 6, int(width) + 12, int(height) + 12)
        d.Close(0)
    finally:
        app.Quit(0)
    print("ok", FIXTURE)


if __name__ == "__main__":
    main()
