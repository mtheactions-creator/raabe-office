"""Records Excel's Spezialfilter (Daten > Sortieren und Filtern > Erweitert..., Range.AdvancedFilter; D-090) as the
reference for packages/sheet-engine/test/advfilter-excel.test.ts: for a list and twelve criteria ranges, which rows
Excel hides when filtering in place, and what it writes when copying to another place (with all or some headers,
with Keine Duplikate), plus the defined names it makes.

    .venv/Scripts/python.exe probe_excel_advfilter.py

Writes packages/sheet-engine/test/fixtures/excel-advfilter.json. Our own invisible Excel; no real input."""

from __future__ import annotations

import json
from pathlib import Path

from feellab.excel import ExcelTarget

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "packages" / "sheet-engine" / "test" / "fixtures" / "excel-advfilter.json"

LIST = [["Name", "Fach", "Note", "Datum"],
        ["Anna", "Mathe", 2, 46082], ["Ben", "Deutsch", 3, 46090], ["Cem", "Mathematik", 1, 46100], ["Dora", "Englisch", 4, 46050],
        ["Emil", "Mathe", 5, 46120], ["Fiona", "Deutsch", 2, 46075], ["Gül", "Bio", 1, 46110], ["Hannes", "Mathe", 2, 46082],
        ["Ida", "Englisch", 3, 46095], ["Jonas", "Deutsch", 4, 46130], ["Anna", "Mathe", 2, 46082], ["Kai", "", 3, 46100]]
# Criteria as typed (German, as FormulaLocal): rows of cells, the first row the headers.
CASES = [
    ("prefix", [["Fach"], ["Mathe"]]),
    ("exact", [["Fach"], ['="=Mathe"']]),
    ("and", [["Fach", "Note"], ["Deutsch", ">=3"]]),
    ("or", [["Fach", "Note"], ["Mathe", ""], ["", "1"]]),
    ("wildcard", [["Name"], ["?a*"]]),
    ("not", [["Note"], ["<>2"]]),
    ("formula", [["Gut"], ["=C2<3"]]),
    ("range", [["Note", "Note"], [">=2", "<=3"]]),
    ("date", [["Datum"], [">1.3.2026"]]),
    ("empty", [["Fach"], ['="="']]),
    ("blankrow", [["Fach", "Note"], ["", ""]]),
    ("text_gt", [["Name"], [">H"]]),
]


def put(ws, top: int, left: int, rows: list[list]) -> None:
    for i, r in enumerate(rows):
        for k, v in enumerate(r):
            c = ws.Cells(top + i, left + k)
            if isinstance(v, str) and v.startswith("="):
                c.Formula2Local = v
            elif v != "":
                c.Value = v


def record(name: str, crit: list[list], action: str) -> dict:
    """One case in a fresh Excel of our own (copying through COM can take Excel down: one case may not spoil the rest)."""
    case = {"name": name, "action": action, "criteria": crit}
    xl = ExcelTarget(visible=False)
    try:
        wb = xl.wb
        ws = wb.Worksheets(1)
        put(ws, 1, 1, LIST)
        put(ws, 1, 8, crit)  # criteria at H1
        lst = ws.Range(ws.Cells(1, 1), ws.Cells(len(LIST), 4))
        cr = ws.Range(ws.Cells(1, 8), ws.Cells(len(crit), 7 + len(crit[0])))
        if action == "inplace":
            lst.AdvancedFilter(1, cr)
            case["hidden"] = [r for r in range(2, len(LIST) + 1) if ws.Rows(r).Hidden]
        else:
            lst.AdvancedFilter(2, cr, ws.Range("M1"), action == "copy_unique")
            used = ws.Range("M1").CurrentRegion
            vals = used.Value2
            case["output"] = [list(r) for r in vals] if isinstance(vals, tuple) else [[vals]]
            case["output_address"] = used.Address.replace("$", "")
        case["names"] = {n.Name: n.RefersTo for n in wb.Names}
        case["filter_mode"] = bool(ws.FilterMode)
    except Exception as e:  # noqa: BLE001 - recorded
        case["error"] = str(e)[:120]
    finally:
        try:
            xl.close()
        except Exception:  # noqa: BLE001 - Excel went down; its process is ended below
            import subprocess
            subprocess.run(["taskkill", "/PID", str(xl.pid), "/F"], capture_output=True)
    return case


def main() -> None:
    out: dict = {"list": LIST, "cases": []}
    for name, crit in CASES:
        for action in ("inplace", "copy", "copy_unique"):
            if action != "inplace" and name not in ("prefix", "or", "formula", "and"):
                continue
            case = record(name, crit, action)
            out["cases"].append(case)
            print(name, action, case.get("hidden", case.get("output_address", case.get("error"))))
    OUT.write_text(json.dumps(out, indent=1, ensure_ascii=False) + chr(10), encoding="utf-8", newline=chr(10))


if __name__ == "__main__":
    main()
