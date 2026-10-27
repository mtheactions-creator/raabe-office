"""Cells through the clipboard between Excel and RaabeCalc, both ways (package F): numbers, text,
dates, percent, formulas (as values), bold, fill. Excel runs invisibly (own instance), RaabeCalc
off screen; the system clipboard is used (it is overwritten). Result: results/<date>-calc-clipboard.json.
"""
from __future__ import annotations

import datetime
import json
import time
from pathlib import Path

from feellab.calc_catalogue import CalcSide, Case
from feellab.excel import ExcelTarget

DATA = {"A1": "Name", "B1": "Wert", "A2": "Anna", "B2": "12,5", "A3": "Ben", "B3": "1.234", "A4": "Datum", "B4": "1.1.2026",
        "A5": "Anteil", "B5": "15%", "A6": "Summe", "B6": "=B2+B3"}


def main() -> None:
    out: dict = {}
    x = ExcelTarget(visible=False)
    c = CalcSide()
    try:
        ws = x.sheet
        for a, t in DATA.items():
            ws.Range(a).FormulaLocal = t
        ws.Range("A1:B1").Font.Bold = True
        ws.Range("B2").Interior.Color = 0x00FFFF  # yellow (BGR)
        # Excel -> RaabeCalc
        ws.Range("A1:B6").Copy()
        time.sleep(0.5)
        c.reset(Case("clip", {}, [], "x", start="D2"))
        c.js("__calc.app.paste()")
        time.sleep(0.8)
        st = c.js("__calc.state([])")
        got = {a: v["text"] for a, v in st["cells"].items()}
        bold = {a: v["bold"] for a, v in st["cells"].items()}
        expect = {chr(ord(a[0]) + 3) + str(int(a[1:]) + 1): ws.Range(a).Text for a in DATA}
        out["excel->raabecalc"] = {"expected": expect, "got": {k: got.get(k) for k in expect}, "boldD2E2": [bold.get("D2"), bold.get("E2")],
                                   "same": all(got.get(k) == v for k, v in expect.items())}
        x.app.CutCopyMode = False
        # RaabeCalc -> Excel
        c.reset(Case("clip2", DATA, [], "x"))
        c.js("__calc.app.grid.select({r0:0,c0:0,r1:5,c1:1}, {row:0,col:0}); __calc.app.commands.run('edit.copy')")
        time.sleep(0.8)
        ws.Range("D2").Select()
        ws.Paste()
        time.sleep(0.5)
        expect2 = {a: ws.Range(a).Text for a in DATA}
        got2 = {a: ws.Range(chr(ord(a[0]) + 3) + str(int(a[1:]) + 1)).Text for a in DATA}
        bold2 = [bool(ws.Range("D2").Font.Bold), bool(ws.Range("E2").Font.Bold)]
        out["raabecalc->excel"] = {"expected": expect2, "got": got2, "same": expect2 == got2, "boldD2E2": bold2,
                                   "values": {a: ws.Range(chr(ord(a[0]) + 3) + str(int(a[1:]) + 1)).Value2 for a in ("B2", "B3", "B5")}}
    finally:
        c.close()
        x.close()
    for k, v in out.items():
        print(k, "same" if v["same"] else "DIFFERENT", json.dumps({kk: vv for kk, vv in v.items() if kk != "same"}, ensure_ascii=False)[:600])
    Path(__file__).parent.joinpath("results", f"{datetime.date.today().isoformat()}-calc-clipboard.json").write_text(
        json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
