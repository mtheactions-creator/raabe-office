"""Screenshots of Excel-built workbooks (packages/ooxml/test/fixtures/xlsx-*.xlsx) opened in Excel
and in RaabeCalc, same window size, for comparing by eye: captures/calc-files/<name>-excel.png and
<name>-raabecalc.png (local only). Takes over the screen briefly.

    .venv/Scripts/python.exe compare_calc_files.py [name …]
"""
from __future__ import annotations

import json
import sys
import time
from pathlib import Path

from feellab import winapi
from feellab.calc_catalogue import CalcSide, settle
from feellab.excel import ExcelTarget
from feellab.targets import WINDOW_RECT

ROOT = Path(__file__).resolve().parents[2]
FIX = ROOT / "packages" / "ooxml" / "test" / "fixtures"
OUT = Path(__file__).parent / "captures" / "calc-files"


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    names = sys.argv[1:] or ["styles", "numfmt", "layout", "objects", "values", "sheets"]
    x = ExcelTarget(visible=True)
    try:
        for n in names:
            wb = x.app.Workbooks.Open(str(FIX / f"xlsx-{n}.xlsx"), ReadOnly=True)
            # Every workbook has a window of its own in Excel.
            hwnd = int(wb.Windows(1).Hwnd)
            winapi.place(hwnd, *WINDOW_RECT)
            settle()
            winapi.activate(hwnd)
            time.sleep(1.2)
            l, t, r, b = winapi.window_rect(hwnd)
            winapi.save_region_png(str(OUT / f"{n}-excel.png"), l, t, 1600, 1300)
            wb.Close(False)
    finally:
        x.close()
    c = CalcSide()
    try:
        for n in names:
            c.js(f"__calc.open({json.dumps(str(FIX / f'xlsx-{n}.xlsx'))})")
            time.sleep(1.0)
            settle()
            winapi.activate(c.hwnd)
            time.sleep(0.8)
            l, t, r, b = winapi.window_rect(c.hwnd)
            winapi.save_region_png(str(OUT / f"{n}-raabecalc.png"), l, t, 1600, 1300)
    finally:
        c.close()


if __name__ == "__main__":
    main()
