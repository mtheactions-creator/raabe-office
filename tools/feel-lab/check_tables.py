"""Table handling with the mouse in RaabeWrite (D-038, package B): the gridlines of a table with
and without borders, dragging a row's bottom border, the sizing handle and the move handle.
Real mouse input on a Word-made document; prints the table definition after each step and
saves screenshots to captures/tables/ (local only).

    .venv/Scripts/python.exe check_tables.py
"""
from __future__ import annotations

import json
import tempfile
import time
from pathlib import Path

from feellab import winapi
from feellab.catalogue import Find, RaabeSide

HERE = Path(__file__).parent
CAP = HERE / "captures" / "tables"
CAP.mkdir(parents=True, exist_ok=True)
DOC = Path(tempfile.gettempdir()) / "raabe-tables.docx"


def build() -> None:
    import win32com.client
    app = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx("Word.Application"))
    app.Visible = False
    try:
        d = app.Documents.Add()
        d.Content.Text = "Vor der Tabelle\rZwischen den Tabellen\rNach der Tabelle"
        for needle, grid in (("Vor der Tabelle", True), ("Zwischen den Tabellen", False)):
            para = next(p for p in d.Paragraphs if p.Range.Text.startswith(needle))
            r = d.Range(para.Range.End, para.Range.End)
            t = d.Tables.Add(r, 3, 3, 1 if grid else 0, 0)
            if not grid:
                t.Borders.Enable = False
            for i in range(1, 4):
                for k in range(1, 4):
                    t.Cell(i, k).Range.Text = f"{'G' if grid else 'O'}{i}{k}"
        d.SaveAs2(str(DOC), 16)
        d.Close(0)
    finally:
        app.Quit(0)


TABLE = """(() => { const e = __app.editor(); const p = e.doc.paras.find((q) => q.cell && q.text === 'G11');
  const t = p.cell.table; return { grid: t.grid.map((x) => Math.round(x * 10) / 10), indent: t.indent ?? 0,
  rows: t.rows.map((r) => [r.height ?? null, r.heightRule ?? null]) }; })()"""


def main() -> None:
    build()
    side = RaabeSide()
    out = {}
    try:
        side.open(DOC)
        t = side.t
        x0, y0, h0 = side.point_at(side.locate(Find("G11")))
        winapi.click(x0 + 5, y0 + h0 / 2)
        time.sleep(0.5)
        winapi.save_region_png(str(CAP / "gridlines.png"), int(x0) - 60, int(y0) - 60, 1500, 700)
        out["start"] = t.js(TABLE)
        # Row 1's bottom border: between G11 and G21.
        _, y1, _ = side.point_at(side.locate(Find("G21")))
        scale = winapi.dpi_scale(side.hwnd)
        border_y = y1 - 1 * scale  # the line is just above the next row's text top margin
        # Search the row-resize cursor position: move down from G11's line to G21 until the editor reports it.
        cx = x0 + 30
        found = None
        y = y0 + h0
        while y < y1 + 2:
            winapi.move(cx, y)
            time.sleep(0.03)
            if t.js("getComputedStyle(document.querySelector('.workspace')).cursor").endswith("row-resize"):
                found = y
                break
            y += 1
        out["row_border_found"] = found is not None
        if found:
            winapi.drag(cx, found, cx, found + 30 * scale, steps=15)
            time.sleep(0.4)
            out["after_row_drag"] = t.js(TABLE)
        # Sizing handle: just beyond the bottom-right corner of the table.
        x3, y3, h3 = side.point_at(side.locate(Find("G33", 3)))
        winapi.click(x3, y3 + h3 / 2)
        time.sleep(0.4)
        box = t.js("""(() => { const els = [...document.querySelectorAll('.workspace div')].filter((d) => d.style.width === '7px' && d.style.height === '7px');
          return els.map((d) => { const r = d.getBoundingClientRect(); return [r.left, r.top, r.width, r.height]; }); })()""")
        out["size_handles"] = box
        winapi.save_region_png(str(CAP / "handles.png"), int(x0) - 80, int(y0) - 80, 1600, 800)
        if box:
            # Client to screen: the window's client origin and device pixel ratio.
            ox, oy = winapi.client_origin(side.hwnd)
            l, tp, w, hh = box[0]
            sx, sy = ox + (l + w / 2) * scale, oy + (tp + hh / 2) * scale
            winapi.drag(sx, sy, sx + 60 * scale, sy + 20 * scale, steps=15)
            time.sleep(0.4)
            out["after_size_drag"] = t.js(TABLE)
        moves = t.js("""(() => { const els = [...document.querySelectorAll('.workspace div')].filter((d) => d.style.width === '13px' && d.style.height === '13px');
          return els.map((d) => { const r = d.getBoundingClientRect(); return [r.left, r.top, r.width, r.height]; }); })()""")
        out["move_handles"] = moves
        if moves:
            ox, oy = winapi.client_origin(side.hwnd)
            l, tp, w, hh = moves[0]
            sx, sy = ox + (l + w / 2) * scale, oy + (tp + hh / 2) * scale
            winapi.drag(sx, sy, sx + 40 * scale, sy, steps=15)
            time.sleep(0.4)
            out["after_move_drag"] = t.js(TABLE)
            out["selection_after_move"] = t.js("__app.editor().selectedText().slice(0, 40)")
    finally:
        side.close()
    print(json.dumps(out, indent=1, ensure_ascii=False))


if __name__ == "__main__":
    main()
