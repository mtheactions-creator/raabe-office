"""Probe: Excel's Konsolidieren (Range.Consolidate) in an invisible Excel of our own, no real input:
by position and by labels (Oberster Zeile, Linker Spalte, both), all eleven functions, labels in another
case, text and empty cells, sources of different sizes, a destination range smaller than the result,
a source on the destination's own sheet, number formats, and Verknüpfungen mit Quelldaten (the outline
with formulas Excel builds). After each case the destination sheet is read cell by cell (formula, value,
text, format, bold), with its rows' and columns' outline levels, and the <dataConsolidate> element Excel
keeps in the sheet is read from a saved copy.

    .venv/Scripts/python.exe probe_excel_consolidate.py

Result: results/<date>-excel-consolidate.json and packages/sheet-engine/test/fixtures/excel-consolidate.json;
the workbook after the last case: packages/ooxml/test/fixtures/xlsx-consolidate.xlsx.
"""
from __future__ import annotations

import datetime
import json
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from feellab.excel import ExcelTarget
from feellab.sanitize import private_names
from make_xlsx_fixtures import NEUTRAL, anonymize, check_private
from probe_excel_scenarios import fill, neutralize, norm

HERE = Path(__file__).parent
ROOT = HERE.parents[1]
TMP = Path(tempfile.gettempdir())
NORD = {"B1": "Jan", "C1": "Feb", "A2": "Äpfel", "B2": 10, "C2": 20, "A3": "Birnen", "B3": 5, "C3": 7, "A4": "Kirschen", "B4": 3, "C4": "x"}
SUED = {"B1": "Feb", "C1": "Mär", "A2": "birnen", "B2": 1, "C2": 2, "A3": "Äpfel", "B3": 4, "A4": "Pflaumen", "B4": 6, "C4": 8, "A5": "Äpfel", "B5": 100, "C5": 200}
WEST = {"A1": 1, "B1": 2, "A2": 3, "B2": 4, "A3": 5, "B3": 6, "C3": 7, "D3": 9}
ORDER = [
    [["a", "b", "c"], ["d"]], [["a", "b", "c"], ["c", "d"]], [["a", "b", "c"], ["d", "a"]], [["a", "b", "c"], ["b", "a", "d"]],
    [["a", "b", "c"], ["b", "a", "d", "a"]], [["a", "b"], ["c"], ["d"]], [["a", "a", "b"]], [["a", "b", "c"], ["b", "d", "a"]],
    [["a", "b", "c"], ["d", "e", "a"]], [["a", "b", "c"], ["c", "d", "b"]], [["a", "b", "c"], ["a", "d", "c"]],
    [["a", "b", "c"], ["d"], ["e", "d"]], [["a", "b", "c"], ["d", "x", "c"]], [["A", "b"], ["a", "B"]],
]
FUNCS = {"Summe": -4157, "Anzahl": -4112, "Mittelwert": -4106, "Maximum": -4136, "Minimum": -4139, "Produkt": -4149,
         "Anzahl Zahlen": -4113, "Standardabweichung (Stichprobe)": -4155, "Standardabweichung (Grundgesamtheit)": -4156,
         "Varianz (Stichprobe)": -4164, "Varianz (Grundgesamtheit)": -4165}


def local(ref: str) -> str:
    """Excel takes the sources in R1C1 style of its own language: Z1S1 in German."""
    sheet, rc = ref.split("!")
    return f"{sheet}!" + re.sub(r"R(\d+)C(\d+)", lambda m: f"Z{m.group(1)}S{m.group(2)}", rc)


def dump(ws) -> dict:
    used = ws.UsedRange
    r1, c1 = used.Row + used.Rows.Count - 1, used.Column + used.Columns.Count - 1
    cells = {}
    for r in range(1, r1 + 1):
        for c in range(1, c1 + 1):
            cell = ws.Cells(r, c)
            if cell.Formula == "" and not cell.Font.Bold:
                continue
            cells[cell.Address.replace("$", "")] = {"formula": cell.Formula, "local": cell.FormulaLocal, "value": norm(cell.Value2),
                                                    "text": cell.Text, "format": cell.NumberFormat, "bold": bool(cell.Font.Bold)}
    rows = {r: [int(ws.Rows(r).OutlineLevel), bool(ws.Rows(r).Hidden)] for r in range(1, r1 + 1) if ws.Rows(r).OutlineLevel > 1 or ws.Rows(r).Hidden}
    cols = {c: [int(ws.Columns(c).OutlineLevel), bool(ws.Columns(c).Hidden), ws.Columns(c).ColumnWidth] for c in range(1, c1 + 1)}
    return {"used": used.Address.replace("$", ""), "cells": cells, "rows": rows, "cols": cols,
            "summaryRow": ws.Outline.SummaryRow, "summaryColumn": ws.Outline.SummaryColumn}


def consolidate_xml(wb, sheet_index: int) -> str | None:
    path = TMP / "raabe-cons.xlsx"
    if path.exists():
        path.unlink()
    wb.SaveCopyAs(str(path))
    with zipfile.ZipFile(path) as z:
        book = z.read("xl/workbook.xml").decode("utf-8")
        rid = re.findall(r'<sheet\b[^>]*r:id="([^"]+)"', book)[sheet_index - 1]
        rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        target = re.search(rf'<Relationship\b[^>]*Id="{rid}"[^>]*Target="([^"]+)"', rels) or re.search(rf'<Relationship\b[^>]*Target="([^"]+)"[^>]*Id="{rid}"', rels)
        x = z.read("xl/" + target.group(1).lstrip("/").removeprefix("xl/")).decode("utf-8")
    m = re.search(r"<dataConsolidate\b[^>]*?(/>|>.*?</dataConsolidate>)", x, re.S)
    return m.group(0) if m else None


def main() -> None:
    x = ExcelTarget(visible=False)
    out: dict = {"sources": {"Nord": NORD, "Süd": SUED, "West": WEST}, "cases": []}
    user = ""
    try:
        user = str(x.app.UserName)
        wb = x.wb
        ws = wb.Worksheets(1)
        ws.Name = "Gesamt"
        sheets = {}
        for name, data in (("Nord", NORD), ("Süd", SUED), ("West", WEST)):
            s = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            s.Name = name
            fill(s, data)
            sheets[name] = s
        sheets["Nord"].Range("B2").NumberFormat = "0.00"
        sheets["Süd"].Range("B3").NumberFormat = "0%"
        out["formats"] = {"Nord!B2": "0.00", "Süd!B3": "0%"}

        def case(name: str, sources: list[str], fn: int, top: bool, left: bool, links: bool = False, dest: str = "A1", pre=None) -> None:
            ws.Activate()
            ws.Cells.ClearOutline()
            ws.Rows.Hidden = False
            ws.Columns.Hidden = False
            ws.Cells.Delete()
            if pre:
                pre()
            try:
                ws.Range(dest).Consolidate([local(s) for s in sources], fn, top, left, links)
                res = dump(ws)
            except Exception as e:  # noqa: BLE001
                res = {"error": str(e)[:200]}
            res.update({"name": name, "sources": sources, "function": fn, "top": top, "left": left, "links": links, "dest": dest,
                        "dataConsolidate": consolidate_xml(wb, 1)})
            out["cases"].append(res)
            print(name, res.get("used"), res.get("error", ""), flush=True)

        nord, sued = "Nord!R1C1:R4C3", "Süd!R1C1:R5C3"
        case("Position", ["Nord!R2C2:R4C3", "Süd!R2C2:R5C3"], -4157, False, False)
        case("Position, verschieden groß", ["West!R1C1:R3C4", "Nord!R2C2:R4C3"], -4157, False, False)
        case("Beschriftung oben und links", [nord, sued], -4157, True, True)
        case("Beschriftung links", ["Nord!R2C1:R4C3", "Süd!R2C1:R5C3"], -4157, False, True)
        case("Beschriftung oben", ["Nord!R1C2:R4C3", "Süd!R1C2:R5C3"], -4157, True, False)
        case("Beschriftung oben und links, Bereich mit Ecke", ["Nord!R1C1:R4C3", "Süd!R1C1:R5C3"], -4157, True, True, dest="B2")
        for label, fn in FUNCS.items():
            case(f"Funktion {label}", [nord, sued], fn, True, True)
        case("Position, Funktion Anzahl", ["Nord!R2C2:R4C3", "Süd!R2C2:R5C3"], -4112, False, False)
        case("Zielbereich kleiner", [nord, sued], -4157, True, True, dest="A1:B2")
        case("Zielbereich eine Zeile", ["Nord!R2C2:R4C3", "Süd!R2C2:R5C3"], -4157, False, False, dest="A1:C1")
        case("Quelle auf dem Zielblatt", ["Gesamt!R10C1:R11C2", "Nord!R2C2:R3C3"], -4157, False, False,
             pre=lambda: fill(ws, {"A10": 1, "B10": 2, "A11": 3, "B11": 4}))
        case("Verknüpfungen", [nord, sued], -4157, True, True, links=True)
        case("Verknüpfungen, Mittelwert", [nord, sued], -4106, True, True, links=True)
        case("Verknüpfungen, nur links", ["Nord!R2C1:R4C3", "Süd!R2C1:R5C3"], -4157, False, True, links=True)
        case("Verknüpfungen, nur oben", ["Nord!R1C2:R4C3", "Süd!R1C2:R5C3"], -4157, True, False, links=True)
        case("Verknüpfungen, Position", ["Nord!R2C2:R4C3", "Süd!R2C2:R5C3"], -4157, False, False, links=True)
        case("Quellen umgekehrt", ["Süd!R1C1:R5C3", "Nord!R1C1:R4C3"], -4157, True, True)
        # Where a new label goes: before the next label of its own source that is already listed, else at the end.
        order = []
        srcs = [sheets["Nord"], sheets["Süd"], sheets["West"]]
        for labels_list in ORDER:
            for s in srcs:
                s.Cells.Clear()
            refs = []
            for i, labels in enumerate(labels_list):
                for r, lab in enumerate(labels):
                    srcs[i].Cells(r + 1, 1).Value2 = lab
                    srcs[i].Cells(r + 1, 2).Value2 = 1
                refs.append(f"{srcs[i].Name}!Z1S1:Z{len(labels)}S2")
            ws.Cells.Delete()
            ws.Range("A1").Consolidate(refs, -4157, False, True, False)
            got = []
            r = 1
            while ws.Cells(r, 1).Value2 is not None:
                got.append(ws.Cells(r, 1).Value2)
                r += 1
            order.append({"sources": labels_list, "rows": got})
        out["order"] = order
        for name, data in (("Nord", NORD), ("Süd", SUED), ("West", WEST)):
            sheets[name].Cells.Clear()
            fill(sheets[name], data)
        sheets["Nord"].Range("B2").NumberFormat = "0.00"
        sheets["Süd"].Range("B3").NumberFormat = "0%"
        # The last case stays in the fixture (labels, sum, no links).
        case("Fixture", [nord, sued], -4157, True, True)
        ws.Range("A1").Select()
        fixture = ROOT / "packages" / "ooxml" / "test" / "fixtures" / "xlsx-consolidate.xlsx"
        if fixture.exists():
            fixture.unlink()
        path = TMP / "raabe-cons-final.xlsx"
        if path.exists():
            path.unlink()
        wb.SaveCopyAs(str(path))
        shutil.copyfile(path, fixture)
    finally:
        x.close()
    anonymize(fixture)
    neutralize(fixture, user)
    check_private(fixture, user if user != NEUTRAL else "")
    text = json.dumps(out, ensure_ascii=False, indent=1, default=str)
    for n in sorted({user, *(p for p in user.split() if len(p) >= 4)} - {"", NEUTRAL}, key=len, reverse=True):
        text = text.replace(n, "<user>")
    assert not any(n in text.lower() for n in private_names()), "private name in the results"
    text = text.replace(str(TMP), "<temp>")
    (HERE / "results" / f"{datetime.date.today().isoformat()}-excel-consolidate.json").write_text(text, encoding="utf-8", newline="\n")
    (ROOT / "packages" / "sheet-engine" / "test" / "fixtures" / "excel-consolidate.json").write_text(text, encoding="utf-8", newline="\n")
    print("ok", len(out["cases"]))


if __name__ == "__main__":
    main()
