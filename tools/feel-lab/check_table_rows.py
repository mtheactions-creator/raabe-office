"""Row heights in RaabeWrite against Word (D-060).

(1) Zeilen verteilen: Word builds a table with rows of different heights (minimum, exact, and
    one tall by its text); Word and RaabeWrite each distribute the rows with the caret in one
    cell, with two rows selected and with the whole table selected; the rows' heights and
    rules are compared (the command runs by script, no input needed).
(2) The sizing handle of a new table (real mouse): dragged to make the table narrower with the
    mouse wandering a few pixels up and down, the table must keep its height (before the fix
    the first drag made it grow with every mouse move).

    .venv/Scripts/python.exe check_table_rows.py

Result: results/<date>-table-rows.json.
"""
from __future__ import annotations

import datetime
import json
import sys
import tempfile
import time
from pathlib import Path

from feellab import winapi
from feellab.catalogue import RaabeSide

HERE = Path(__file__).parent
OUT = HERE / "results" / f"{datetime.date.today()}-table-rows.json"
DOC = Path(tempfile.gettempdir()) / "raabe-table-rows.docx"
E = "__app.editor()"

# Word's HeightRule: 0 auto, 1 at least, 2 exact.
RULES = {0: None, 1: "atLeast", 2: "exact"}


def build(app):
    d = app.Documents.Add()
    d.Content.Text = "Vor\rNach"
    r = d.Range(d.Paragraphs(1).Range.End, d.Paragraphs(1).Range.End)
    t = d.Tables.Add(r, 4, 2, 1, 0)
    t.Rows(1).Height, t.Rows(1).HeightRule = 20, 1
    t.Rows(2).Height, t.Rows(2).HeightRule = 60, 1
    t.Rows(3).Height, t.Rows(3).HeightRule = 40, 2
    t.Cell(4, 1).Range.Text = "eins\vzwei\vdrei\vvier"
    return d, t


def word_side() -> dict:
    import win32com.client
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    out = {}
    try:
        d, _ = build(app)
        d.SaveAs2(str(DOC), 16)
        d.Close(0)
        for case in ("caret", "rows12", "table"):
            d = app.Documents.Open(str(DOC))
            t = d.Tables(1)
            if case == "caret":
                d.Range(t.Cell(2, 1).Range.Start, t.Cell(2, 1).Range.Start).Select()
            elif case == "rows12":
                d.Range(t.Rows(1).Range.Start, t.Rows(2).Range.End).Select()
            else:
                t.Select()
            app.Selection.Cells.DistributeHeight()
            out[case] = [[round(t.Rows(i).Height, 1), RULES[t.Rows(i).HeightRule]] for i in range(1, t.Rows.Count + 1)]
            d.Close(0)
    finally:
        app.Quit(0)
    return out


SELECT = {
    # Paragraphs: 0 "Vor", then the cells row by row (two per row).
    "caret": f"{E}.select({{p:3,o:0}},{{p:3,o:0}})",
    "rows12": f"{E}.select({{p:1,o:0}},{{p:4,o:{E}.doc.paras[4].text.length}})",
    "table": "__app.run('table.selectTable')",
}
ROWS = f"JSON.stringify({E}.doc.paras[1].cell.table.rows.map(r => [r.height === undefined ? null : Math.round(r.height * 10) / 10, r.heightRule ?? null]))"


def our_side(side: RaabeSide) -> dict:
    out = {}
    for case, sel in SELECT.items():
        # A fresh copy each time: opening a path that is open only switches to it.
        side.t.js("__app.newDoc()")
        side.t.js("__app.closeOthers && __app.closeOthers()")
        time.sleep(0.3)
        side.open(DOC)
        side.t.js(f"{E}.select({{p:3,o:0}},{{p:3,o:0}})")
        side.t.js(sel)
        time.sleep(0.2)
        side.t.js("__app.run('table.distributeRows')")
        time.sleep(0.3)
        out[case] = json.loads(side.t.js(ROWS))
    return out


HANDLE = """(() => { const d = [...document.querySelectorAll('div')].find((x) => x.style.width === '7px' && x.style.height === '7px');
  if (!d) return null; const r = d.getBoundingClientRect(); return [r.left + r.width / 2, r.top + r.height / 2]; })()"""
SIZE = f"""(() => {{ const e = {E}; const p = e.doc.paras.findIndex((q) => q.cell); const t = e.doc.paras[p].cell.table;
  return {{ width: Math.round(t.grid.reduce((a, b) => a + b, 0) * 10) / 10,
    height: Math.round(t.rows.reduce((a, r, i) => a + e.rowMeasurePt(t.id, i).inner, 0) * 10) / 10 }}; }})()"""


def size_handle(side: RaabeSide) -> dict:
    t = side.t
    out = {}
    t.js("__app.newDoc()")
    time.sleep(0.8)
    t.js("__app.closeOthers && __app.closeOthers()")
    t.js(f"{E}.insertTable(3, 3)")
    t.js(f"{E}.insertText('Wort')")
    time.sleep(0.5)
    dpr = t.js("devicePixelRatio")
    ox, oy = t._viewport_origin()
    for attempt in ("first drag", "second drag"):
        h = t.js(HANDLE)
        if not h:
            out[attempt] = "no sizing handle"
            break
        before = t.js(SIZE)
        sx, sy = ox + h[0] * dpr, oy + h[1] * dpr
        winapi.move(sx, sy)
        time.sleep(0.3)
        winapi.mouse_down(sx, sy)
        for i in range(1, 41):
            winapi.move(sx - 160 * dpr * i / 40, sy + (3 if i % 2 else -2) * dpr)
            time.sleep(0.02)
        winapi.mouse_up(sx - 160 * dpr, sy + 2 * dpr)
        time.sleep(0.5)
        after = t.js(SIZE)
        zoom = t.js(f"{E}.view.zoom")
        # The pointer ended 2 CSS px lower: the table may be that much taller, nothing more.
        expected = before["height"] + 2 / zoom * 0.75
        out[attempt] = {"before": before, "after": after, "height_ok": abs(after["height"] - expected) < 2,
                        "width_ok": after["width"] < before["width"] - 60}
    return out


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    word = word_side()
    side = RaabeSide()
    try:
        ours = our_side(side)
        handle = size_handle(side)
    finally:
        side.close()
    dist = {}
    for case in word:
        ok = len(word[case]) == len(ours[case]) and all(
            (w[1] or None) == (o[1] or None) and (w[1] is None or abs(w[0] - o[0]) <= 1.0) for w, o in zip(word[case], ours[case]))
        dist[case] = {"word": word[case], "ours": ours[case], "ok": ok}
    data = {"meta": {"date": str(datetime.date.today())}, "distribute_rows": dist, "size_handle": handle}
    OUT.write_text(json.dumps(data, indent=1, ensure_ascii=False), encoding="utf-8", newline="\n")
    for case, r in dist.items():
        print(f"Zeilen verteilen, {case}: {'ok' if r['ok'] else 'DIFFERENT'}  Word {r['word']}  ours {r['ours']}")
    for k, v in handle.items():
        print(f"sizing handle, {k}: {v}")
    print("saved to", OUT)


if __name__ == "__main__":
    main()
