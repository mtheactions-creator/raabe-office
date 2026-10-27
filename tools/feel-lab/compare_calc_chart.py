"""Screenshots of the same column chart in Excel and RaabeCalc, for comparing by eye:
captures/calc-look/chart-excel.png and chart-raabecalc.png (local only). Takes over the screen briefly.
"""
from __future__ import annotations

import time
from pathlib import Path

from feellab import winapi
from feellab.calc_catalogue import Case, CalcSide, ExcelSide, settle

OUT = Path(__file__).parent / "captures" / "calc-look"
SETUP = {"A1": "Monat", "B1": "Klasse 8a", "C1": "Klasse 8b", "A2": "Jan", "B2": "12", "C2": "9", "A3": "Feb", "B3": "15", "C3": "11",
         "A4": "Mrz", "B4": "9", "C4": "14", "A5": "Apr", "B5": "17", "C5": "13"}


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    case = Case("chart", SETUP, [], "look", start="A1")
    x = ExcelSide()
    try:
        x.reset(case)
        ws = x.x.sheet
        ws.Range("A1:C5").Select()
        ws.Shapes.AddChart2(201, 51)
        x.x.app.ActiveWindow.ScrollRow = 1
        ws.Range("A1").Select()
        settle()
        winapi.activate(x.x.hwnd)
        time.sleep(1.0)
        l, t, r, b = winapi.window_rect(x.x.hwnd)
        winapi.save_region_png(str(OUT / "chart-excel.png"), l, t, 1500, 1100)
    finally:
        x.close()
    c = CalcSide()
    try:
        c.reset(case)
        c.js("__calc.app.grid.select({r0:0,c0:0,r1:4,c1:2}); __calc.app.charts.insert('column'); __calc.app.grid.selectCell(0,0)")
        settle()
        winapi.activate(c.hwnd)
        time.sleep(1.0)
        l, t, r, b = winapi.window_rect(c.hwnd)
        winapi.save_region_png(str(OUT / "chart-raabecalc.png"), l, t, 1500, 1100)
    finally:
        c.close()


if __name__ == "__main__":
    main()
