"""Screenshots of the same cells in Excel and in RaabeCalc (same window size and position), for
comparing the look by eye: captures/calc-look/excel.png and raabecalc.png (local only).
Takes over the screen briefly.

    .venv/Scripts/python.exe compare_calc_look.py
"""
from __future__ import annotations

import time
from pathlib import Path

from feellab import winapi
from feellab.calc_catalogue import Case, CalcSide, ExcelSide, run_step, settle

OUT = Path(__file__).parent / "captures" / "calc-look"
SETUP = {"A1": "Klasse 8b", "B1": "Note", "C1": "Punkte", "A2": "Anna", "B2": "2", "C2": "41,5", "A3": "Ben", "B3": "3", "C3": "33",
         "A4": "Ein langer Text, der überläuft", "C5": "=SUMME(C2:C3)", "D2": "1.1.2026", "D3": "15%", "E2": "=1/0", "B6": "WAHR"}
STEPS = [("click", "B2"), ("drag", "B2", "C4")]


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    case = Case("look", SETUP, STEPS, "look")
    for name, side_cls in (("excel", ExcelSide), ("raabecalc", CalcSide)):
        side = side_cls()
        try:
            side.reset(case)
            side.prepare(case)
            settle()
            hwnd = side.x.hwnd if name == "excel" else side.hwnd
            winapi.activate(hwnd)
            if name != "excel":
                side.js("__calc.grid.focus()")
            for st in STEPS:
                run_step(side, st)
            winapi.move(1900, 1990)
            time.sleep(0.8)
            l, t, r, b = winapi.window_rect(hwnd)
            winapi.save_region_png(str(OUT / f"{name}.png"), l, t, 900, 700)
        finally:
            side.close()


if __name__ == "__main__":
    main()
