"""Blatt schützen against Excel: RaabeCalc writes a sheet protected with the password "Schule"
(apps/raabe-calc/test/protection.test.ts, our SHA-512 hash); Excel (its own invisible instance)
opens the file, must see the sheet protected with Zellen formatieren allowed, must refuse a wrong
password and accept the right one. The other direction (Excel's hash checked by us) is in the
unit test. No real input. Result: results/<date>-calc-protection.json.

    .venv/Scripts/python.exe check_calc_protection.py
"""
from __future__ import annotations

import datetime
import json
import os
import subprocess
import tempfile
from pathlib import Path

from feellab.excel import ExcelTarget

ROOT = Path(__file__).resolve().parents[2]


def main() -> None:
    out = Path(tempfile.gettempdir()) / "raabe-protection-check.xlsx"
    out.unlink(missing_ok=True)
    env = {**os.environ, "PROTECTION_OUT": str(out)}
    subprocess.run("npx vitest run apps/raabe-calc/test/protection.test.ts", cwd=ROOT, env=env, shell=True, check=True, capture_output=True)
    res: dict = {}
    x = ExcelTarget(visible=False)
    try:
        wb = x.app.Workbooks.Open(str(out))
        ws = wb.Worksheets(2)
        res["protected"] = bool(ws.ProtectContents)
        res["allowFormattingCells"] = bool(ws.Protection.AllowFormattingCells)
        res["allowSorting"] = bool(ws.Protection.AllowSorting)
        res["firstSheetProtected"] = bool(wb.Worksheets(1).ProtectContents)
        try:
            ws.Unprotect("falsch")
            res["wrongPasswordRefused"] = False
        except Exception:  # noqa: BLE001
            res["wrongPasswordRefused"] = True
        ws.Unprotect("Schule")
        res["rightPasswordAccepted"] = not ws.ProtectContents
        wb.Close(False)
    finally:
        x.close()
        out.unlink(missing_ok=True)
    res["pass"] = res == {**res, "protected": True, "allowFormattingCells": True, "allowSorting": False, "firstSheetProtected": False,
                         "wrongPasswordRefused": True, "rightPasswordAccepted": True}
    print(json.dumps(res, indent=1))
    Path(__file__).parent.joinpath("results", f"{datetime.date.today().isoformat()}-calc-protection.json").write_text(json.dumps(res, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
