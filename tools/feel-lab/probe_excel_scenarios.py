"""Probe: Excel's Szenario-Manager (Worksheet.Scenarios) in an invisible Excel of our own, no real input:
- scenarios added through COM (contiguous and non-contiguous changing cells, a text value, protection flags,
  with and without a comment) and read back (name, cells, values, comment, flags);
- Anzeigen (Scenario.Show): the cells take the values, formulas depending on them recalculate, a formula
  among the changing cells is replaced by its value;
- Zusammenfassung (Scenarios.CreateSummary, Szenariobericht) with and without named cells: the new sheet's
  name, every cell (formula, value, text, number format, bold, fill, borders, alignment), row and column
  outline levels, widths and the sheet's own XML parts as Excel saves them;
- Zusammenführen (Scenarios.Merge) from another sheet, names that clash included;
- the file: the <scenarios> element (CT_Scenarios) as Excel writes it after each step.
Scenarios record Excel's user name: the fixture and the results get a neutral one instead.
The saved workbook becomes the reader's fixture packages/ooxml/test/fixtures/xlsx-scenarios.xlsx.

    .venv/Scripts/python.exe probe_excel_scenarios.py

Result: results/<date>-excel-scenarios.json and packages/sheet-engine/test/fixtures/excel-scenarios.json.
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

HERE = Path(__file__).parent
ROOT = HERE.parents[1]
TMP = Path(tempfile.gettempdir())
SETUP = {"A1": "Preis", "B1": 10, "A2": "Menge", "B2": 100, "A3": "Kosten", "B3": 500, "A4": "Gewinn", "B4": "=B1*B2-B3",
         "A5": "Steuer", "B5": "=B4*0.3", "D1": "Rabatt", "E1": 0.05, "E2": "=E1*2"}
ERRORS = {-2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?", -2146826288: "#NULL!",
          -2146826252: "#NUM!", -2146826265: "#REF!", -2146826273: "#VALUE!"}


def norm(v):
    if isinstance(v, int) and v in ERRORS:
        return {"error": ERRORS[v]}
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def fill(ws, setup: dict) -> None:
    for a, v in setup.items():
        if isinstance(v, str) and v.startswith("="):
            ws.Range(a).Formula = v
        else:
            ws.Range(a).Value2 = v


def scenarios_of(ws) -> list[dict]:
    out = []
    for i in range(1, ws.Scenarios().Count + 1):
        s = ws.Scenarios(i)
        vals = s.Values
        vals = [norm(v[0] if isinstance(v, tuple) else v) for v in (vals if isinstance(vals, tuple) else [vals])]
        out.append({"name": s.Name, "cells": s.ChangingCells.Address, "values": vals, "comment": s.Comment.replace(NEUTRAL, "<user>"),
                    "locked": bool(s.Locked), "hidden": bool(s.Hidden), "index": s.Index})
    return out


def rgb(c) -> str | None:
    c = int(c)
    return f"#{c & 255:02X}{(c >> 8) & 255:02X}{(c >> 16) & 255:02X}"


def dump_sheet(ws) -> dict:
    used = ws.UsedRange
    r0, c0 = used.Row, used.Column
    cells = {}
    for r in range(1, r0 + used.Rows.Count):
        for c in range(1, c0 + used.Columns.Count):
            cell = ws.Cells(r, c)
            f = cell.Formula
            fmt_fill = cell.Interior.Pattern != -4142
            borders = {}
            for name, idx in (("top", 8), ("bottom", 9), ("left", 7), ("right", 10)):
                b = cell.Borders(idx)
                if b.LineStyle != -4142:
                    borders[name] = {"style": b.LineStyle, "weight": b.Weight, "color": rgb(b.Color)}
            if f == "" and not fmt_fill and not borders and not cell.Font.Bold:
                continue
            cells[cell.Address.replace("$", "")] = {
                "formula": f, "local": cell.FormulaLocal, "value": norm(cell.Value2), "text": cell.Text, "format": cell.NumberFormat,
                "bold": bool(cell.Font.Bold), "fill": rgb(cell.Interior.Color) if fmt_fill else None,
                "fontColor": rgb(cell.Font.Color), "borders": borders, "h": cell.HorizontalAlignment, "indent": cell.IndentLevel,
                "wrap": bool(cell.WrapText), "v": cell.VerticalAlignment, "font": [cell.Font.Name, cell.Font.Size],
            }
    rows = {r: {"level": ws.Rows(r).OutlineLevel, "hidden": bool(ws.Rows(r).Hidden), "height": ws.Rows(r).RowHeight} for r in range(1, r0 + used.Rows.Count)}
    cols = {c: {"level": ws.Columns(c).OutlineLevel, "hidden": bool(ws.Columns(c).Hidden), "width": ws.Columns(c).ColumnWidth} for c in range(1, c0 + used.Columns.Count)}
    return {"name": ws.Name, "used": used.Address.replace("$", ""), "cells": cells, "rows": rows, "cols": cols,
            "summaryRow": ws.Outline.SummaryRow, "summaryColumn": ws.Outline.SummaryColumn, "gridlines": None}


def parts(path: Path) -> dict:
    """The sheets' XML from Excel's file: <scenarios> and the report sheet's parts."""
    out = {}
    with zipfile.ZipFile(path) as z:
        book = z.read("xl/workbook.xml").decode("utf-8")
        out["workbook_sheets"] = re.findall(r"<sheet\b[^>]*/>", book)
        out["definedNames"] = re.findall(r"<definedName\b.*?</definedName>", book)
        for n in sorted(z.namelist()):
            if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n):
                x = z.read(n).decode("utf-8").replace(NEUTRAL, "<user>")
                m = re.search(r"<scenarios\b[^>]*?(/>|>.*?</scenarios>)", x, re.S)
                pr = re.search(r"<sheetPr\b.*?(</sheetPr>|/>)", x)
                fmt = re.search(r"<sheetFormatPr\b[^>]*>", x)
                colsx = re.search(r"<cols>.*?</cols>", x)
                views = re.search(r"<sheetViews>.*?</sheetViews>", x)
                rows = re.findall(r"<row\b[^>]*>", x)
                out[n] = {"scenarios": m.group(0) if m else None, "sheetPr": pr.group(0) if pr else None, "sheetFormatPr": fmt.group(0) if fmt else None,
                          "cols": colsx.group(0) if colsx else None, "views": views.group(0) if views else None, "rows": rows,
                          "size": len(x)}
    return out


def neutralize(path: Path, user: str) -> None:
    """Scenarios record Excel's user name (user attribute, default comment): the neutral name instead.
    The Office user name is not changed for the run, since other probes may run at the same time."""
    with zipfile.ZipFile(path) as z:
        items = [(info, z.read(info)) for info in z.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for info, data in items:
            if info.filename.startswith("xl/worksheets/") or info.filename == "xl/sharedStrings.xml":
                text = data.decode("utf-8")
                text = re.sub(r'(<scenario\b[^>]*\buser=")[^"]*"', lambda m: m.group(1) + NEUTRAL + '"', text)
                if user:
                    text = text.replace(user, NEUTRAL)
                data = text.encode("utf-8")
            z.writestr(info, data)


def save_copy(wb, name: str) -> Path:
    path = TMP / name
    if path.exists():
        path.unlink()
    wb.SaveCopyAs(str(path))
    return path


def main() -> None:
    x = ExcelTarget(visible=False)
    out: dict = {"setup": SETUP}
    user = ""
    try:
        user = str(x.app.UserName)
        wb = x.wb
        ws = wb.Worksheets(1)
        ws.Name = "Tabelle1"
        fill(ws, SETUP)
        # 1. Adding: contiguous cells with a comment, protection flags, non-contiguous cells, a text value, no comment.
        sc = ws.Scenarios()
        sc.Add("Gut", ws.Range("B1:B2"), (12, 120), "Hohe Preise", True, False)
        sc.Add("Schlecht", ws.Range("B1:B2"), (8, 80), "", False, True)
        sc.Add("Mittel", x.app.Union(ws.Range("B1"), ws.Range("B3")), (10, 400))
        sc.Add("Text", ws.Range("E1"), ("abc",), "Text als Wert", True, False)
        sc.Add("Leer", ws.Range("B2"), (0.25,))
        out["added"] = scenarios_of(ws)
        out["file_added"] = parts(save_copy(wb, "raabe-scen-1.xlsx"))
        # 2. Anzeigen.
        shown = {}
        for name in ("Gut", "Schlecht", "Mittel", "Text"):
            ws.Scenarios(name).Show()
            shown[name] = {a: {"formula": ws.Range(a).Formula, "value": norm(ws.Range(a).Value2)} for a in ("B1", "B2", "B3", "B4", "B5", "E1", "E2")}
        out["shown"] = shown
        out["file_shown"] = parts(save_copy(wb, "raabe-scen-2.xlsx"))
        # A formula among the changing cells: Anzeigen writes the value over it.
        ws.Range("E1").Value2 = 0.05
        sc.Add("Formel", ws.Range("E2"), (7,))
        ws.Scenarios("Formel").Show()
        out["formula_cell"] = {"formula": ws.Range("E2").Formula, "value": norm(ws.Range("E2").Value2)}
        ws.Range("E2").Formula = "=E1*2"
        ws.Scenarios("Formel").Delete()
        # Back to the start values for the report.
        fill(ws, SETUP)
        # 3. Zusammenfassung (Szenariobericht) on B4:B5.
        sc.CreateSummary(1, ws.Range("B4:B5"))
        rep = wb.ActiveSheet
        out["report"] = dump_sheet(rep)
        out["report_index"] = rep.Index
        out["sheets_after_report"] = [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]
        out["file_report"] = parts(save_copy(wb, "raabe-scen-3.xlsx"))
        # A second report: the name it gets.
        ws.Activate()
        sc.CreateSummary(1, ws.Range("B4"))
        out["second_report_name"] = wb.ActiveSheet.Name
        out["report_single"] = dump_sheet(wb.ActiveSheet)
        x.app.DisplayAlerts = False
        wb.ActiveSheet.Delete()
        # Without result cells.
        ws.Activate()
        try:
            sc.CreateSummary(1)
            out["report_noresult"] = dump_sheet(wb.ActiveSheet)
            wb.ActiveSheet.Delete()
        except Exception as e:  # noqa: BLE001
            out["report_noresult"] = {"error": str(e)[:200]}
        # 4. Named cells: the report shows the names.
        ws.Activate()
        wb.Names.Add("Preis", "=Tabelle1!$B$1")
        wb.Names.Add("Gewinn", "=Tabelle1!$B$4")
        sc.CreateSummary(1, ws.Range("B4:B5"))
        out["report_named"] = dump_sheet(wb.ActiveSheet)
        wb.ActiveSheet.Delete()
        wb.Names("Preis").Delete()
        wb.Names("Gewinn").Delete()
        # 5. PivotTable report: whether Excel makes one here, and its sheet's name.
        ws.Activate()
        try:
            sc.CreateSummary(2, ws.Range("B4:B5"))
            out["pivot_report"] = {"name": wb.ActiveSheet.Name, "pivots": wb.ActiveSheet.PivotTables().Count}
            wb.ActiveSheet.Delete()
        except Exception as e:  # noqa: BLE001
            out["pivot_report"] = {"error": str(e)[:200]}
        # 6. Zusammenführen from another sheet, with a clashing name.
        ws.Activate()
        other = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        other.Name = "Quelle"
        fill(other, {"B1": 1, "B2": 2})
        other.Scenarios().Add("Gut", other.Range("B1:B2"), (20, 200), "Aus Quelle")
        other.Scenarios().Add("Neu", other.Range("B1"), (30,))
        ws.Activate()
        ws.Scenarios().Merge(other)
        out["merged"] = scenarios_of(ws)
        # 7. Löschen of the merged ones (back to the five of step 1, as the fixture keeps them).
        for s in [s["name"] for s in out["merged"][5:]]:
            ws.Scenarios(s).Delete()
        x.app.DisplayAlerts = False
        other.Delete()
        ws.Scenarios("Gut").Show()
        ws.Activate()
        ws.Range("A1").Select()
        # The fixture keeps no report (its note is Excel's prose, D-058); the report's file parts are in file_report.
        for i in range(wb.Worksheets.Count, 0, -1):
            if wb.Worksheets(i).Name != "Tabelle1":
                wb.Worksheets(i).Delete()
        out["final"] = scenarios_of(ws)
        fixture = ROOT / "packages" / "ooxml" / "test" / "fixtures" / "xlsx-scenarios.xlsx"
        if fixture.exists():
            fixture.unlink()
        path = save_copy(wb, "raabe-scen-final.xlsx")
        out["file_final"] = parts(path)
        shutil.copyfile(path, fixture)
    finally:
        x.close()
    anonymize(fixture)
    neutralize(fixture, user)
    check_private(fixture, user if user != NEUTRAL else "")
    # Excel's note below the report is its own prose (D-058): kept as its length and position only.
    for key in ("report", "report_named", "report_single", "report_noresult"):
        for a, c in out.get(key, {}).get("cells", {}).items():
            if a.startswith("B") and isinstance(c.get("value"), str) and len(c["value"]) > 40:
                c["value"] = c["local"] = c["text"] = c["formula"] = f"[Hinweis, {len(c['value'])} Zeichen]"
    text = json.dumps(out, ensure_ascii=False, indent=1, default=str)
    for n in sorted({user, *(p for p in user.split() if len(p) >= 4)} - {"", NEUTRAL}, key=len, reverse=True):
        text = text.replace(n, "<user>")
    assert not any(n in text.lower() for n in private_names()), "private name in the results"
    (HERE / "results" / f"{datetime.date.today().isoformat()}-excel-scenarios.json").write_text(text, encoding="utf-8", newline="\n")
    (ROOT / "packages" / "sheet-engine" / "test" / "fixtures" / "excel-scenarios.json").write_text(text, encoding="utf-8", newline="\n")
    print("ok", len(out["report"]["cells"]), "report cells")


if __name__ == "__main__":
    main()
