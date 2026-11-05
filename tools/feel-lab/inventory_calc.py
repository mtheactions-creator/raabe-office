"""RaabeCalc's ribbon as data (`__calc.inventory()`) for inventory_report.py (D-066). Starts the app off
screen like the catalogue does; no input is sent.

    .venv/Scripts/python.exe inventory_calc.py

Result: results/<date>-raabe-inventory.json.
"""
from __future__ import annotations

import datetime
import json
from pathlib import Path

from feellab.calc_catalogue import CalcSide

HERE = Path(__file__).parent


def main() -> None:
    side = CalcSide()
    try:
        data = side.js("__calc.inventory()")
    finally:
        side.close()
    out = HERE / "results" / f"{datetime.date.today().isoformat()}-raabe-inventory.json"
    out.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(len(data), "tabs ->", out.name)


if __name__ == "__main__":
    main()
