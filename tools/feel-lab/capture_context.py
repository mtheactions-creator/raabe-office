"""Screenshots of RaabeWrite's context menus (text, list, table cell, several cells, picture)
for comparison with Word's (probe_word_context.py). The document is made by Word; real mouse
input. Pictures go to captures/raabe-context/ (local only).

    .venv/Scripts/python.exe capture_context.py
"""
from __future__ import annotations

import tempfile
import time
from pathlib import Path

from feellab import winapi
from feellab.catalogue import Find, RaabeSide
from probe_word_context import PNG, make_png

HERE = Path(__file__).parent
CAP = HERE / "captures" / "raabe-context"
CAP.mkdir(parents=True, exist_ok=True)
DOC = Path(tempfile.gettempdir()) / "raabe-context.docx"


def build() -> None:
    import win32com.client
    make_png(PNG)
    app = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx("Word.Application"))
    app.Visible = False
    try:
        d = app.Documents.Add()
        d.Content.Text = "\r".join(["Ein Absatz mit etwas Text zum Anklicken.", "Erster Punkt", "Zweiter Punkt", "Vor der Tabelle", "Nach der Tabelle", ""])
        d.Range(d.Paragraphs(2).Range.Start, d.Paragraphs(3).Range.End).ListFormat.ApplyListTemplate(app.ListGalleries(1).ListTemplates(1), False, 0, 2)
        r = d.Range(d.Paragraphs(4).Range.End, d.Paragraphs(4).Range.End)
        t = d.Tables.Add(r, 2, 2, 1, 0)
        for i in range(1, 3):
            for k in range(1, 3):
                t.Cell(i, k).Range.Text = f"Zelle{i}{k}"
        d.Paragraphs(d.Paragraphs.Count).Range.InlineShapes.AddPicture(str(PNG), False, True)
        d.SaveAs2(str(DOC), 16)
        d.Close(0)
    finally:
        app.Quit(0)


def shot(name: str, x: float, y: float) -> None:
    winapi.save_region_png(str(CAP / f"{name}.png"), int(x) - 40, int(y) - 200, 760, 1150)


def main() -> None:
    build()
    side = RaabeSide()
    try:
        side.open(DOC)
        t = side.t

        def at(needle: str, delta: int = 2) -> tuple[float, float]:
            x, y, h = side.point_at(side.locate(Find(needle, delta)))
            return x + 2, y + h / 2

        for name, needle in (("text", "etwas"), ("list", "Zweiter"), ("table", "Zelle12")):
            x, y = at(needle)
            winapi.click(x, y)
            time.sleep(0.3)
            winapi.right_click(x, y)
            time.sleep(1.0)
            shot(name, x, y)
            winapi.press("esc")
            time.sleep(0.4)
        # Several cells.
        a = side.locate(Find("Zelle11"))
        b = side.locate(Find("Zelle22", 7))
        side.set_selection(a, b)
        x, y = at("Zelle21")
        winapi.right_click(x, y)
        time.sleep(1.0)
        shot("table-cells", x, y)
        winapi.press("esc")
        time.sleep(0.4)
        # The picture: the object character in the text.
        i = t.text().find("\ufffc")
        if i >= 0:
            x, y, h = side.point_at(i)
            x, y = x + 20, y + h * 0.6
            winapi.right_click(x, y)
            time.sleep(1.0)
            shot("picture", x, y)
            winapi.press("esc")
        print("saved to", CAP)
    finally:
        side.close()


if __name__ == "__main__":
    main()
