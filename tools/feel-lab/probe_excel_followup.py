"""Probe (D-145 to D-148): what Excel 365 does for the options the by-hand audit of RaabeCalc left open, read
through COM in an invisible Excel of our own (no real input, nothing of the owner's touched):

- sort: formulas whose references leave the sheet when their row moves (a sum row sorted with its own range);
- Zellen formatieren > Rahmen: every line style with its file name, colours, both diagonals;
- Ausfüllen: Musterformat (all 18 patterns) with Musterfarbe, and Fülleffekte (two-colour gradients, the six
  Schattierungsarten with their variants) as Excel writes them in the file;
- Ausrichtung: the degrees and vertical text as written in the file;
- Datenbalken: a rule with Einfarbige Füllung (BarFillType solid) as Excel writes it in the file.

    .venv/Scripts/python.exe probe_excel_followup.py

Result: results/2026-11-09-excel-followup.json.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from pathlib import Path

import pythoncom

from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
TMP = Path(tempfile.gettempdir()) / "raabe-excel-followup"


def styles_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        return z.read("xl/styles.xml").decode("utf-8")


def sheet_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        return z.read("xl/worksheets/sheet1.xml").decode("utf-8")


def part(xml: str, tag: str) -> list[str]:
    m = re.search(f"<{tag}s? [^>]*>(.*?)</{tag}s>", xml, re.S) or re.search(f"<{tag}s>(.*?)</{tag}s>", xml, re.S)
    return re.findall(f"<{tag}\\b.*?</{tag}>|<{tag}\\b[^>]*/>", m.group(1), re.S) if m else []


def sort_refs(app) -> list[dict]:
    """Rows with formulas sorted so that their relative references would leave the sheet."""
    out = []
    cases = [
        # (name, cells, sort range, key column, order)
        ("Summe nach oben", {"A1": "Name", "B1": "Punkte", "A2": "Anna", "B2": "12", "A3": "Ben", "B3": "7,5", "A4": "Cem", "B4": "15",
                             "A5": "Dora", "B5": "9", "A6": "Summe", "B6": "=SUMME(B2:B5)"}, "A1:B6", 1, 2),
        ("Einzelbezug nach oben", {"A1": "k", "B1": "x", "A2": "1", "B2": "5", "A3": "2", "B3": "6", "A4": "3", "B4": "=B2+1"}, "A1:B4", 1, 2),
        ("Bezug ganz ausserhalb", {"A1": "k", "B1": "x", "A2": "1", "B2": "5", "A3": "2", "B3": "6", "A4": "3", "B4": "=SUMME(B2:B3)"}, "A1:B4", 1, 2),
        ("Absoluter Bezug", {"A1": "k", "B1": "x", "A2": "1", "B2": "5", "A3": "2", "B3": "6", "A4": "3", "B4": "=SUMME($B$2:$B$3)"}, "A1:B4", 1, 2),
        ("Ohne Kopf", {"A1": "1", "B1": "5", "A2": "2", "B2": "6", "A3": "3", "B3": "=SUMME(B1:B2)"}, "A1:B3", 1, 2),
        ("Nach unten", {"A1": "k", "B1": "x", "A2": "3", "B2": "=B5*2", "A3": "2", "B3": "6", "A4": "1", "B4": "7", "A5": "0", "B5": "4"}, "A1:B5", 1, 1),
    ]
    for name, cells, rng, key, order in cases:
        wb = app.Workbooks.Add()
        try:
            ws = wb.Worksheets(1)
            for a, t in cells.items():
                ws.Range(a).FormulaLocal = t
            r = ws.Range(rng)
            header = 1 if name != "Ohne Kopf" else 2  # xlYes / xlNo
            r.Sort(Key1=r.Cells(1, key), Order1=order, Header=header, Orientation=1)
            got = {}
            for row in r.Rows:
                for c in row.Cells:
                    a = c.GetAddress(False, False)
                    got[a] = c.FormulaLocal if c.HasFormula else c.Text
            out.append({"name": name, "cells": cells, "range": rng, "after": got})
        finally:
            wb.Close(False)
    return out


LINE_STYLES = [
    # (LineStyle, Weight) as the dialog's Art list sets them (xlContinuous 1, xlDash -4115, xlDashDot 4, xlDashDotDot 5,
    # xlDot -4118, xlDouble -4119, xlSlantDashDot 13; xlHairline 1, xlThin 2, xlMedium -4138, xlThick 4)
    ("Haarstrich", 1, 1), ("Punktiert", -4118, 2), ("StrichPunktPunkt", 5, 2), ("StrichPunkt", 4, 2), ("Gestrichelt", -4115, 2),
    ("Dünn", 1, 2), ("StrichPunktPunkt Mittel", 5, -4138), ("SchrägstrichStrichPunkt", 13, -4138), ("StrichPunkt Mittel", 4, -4138),
    ("Gestrichelt Mittel", -4115, -4138), ("Mittel", 1, -4138), ("Dick", 1, 4), ("Doppelt", -4119, 4),
]

def formats(app) -> dict:
    """Borders, patterns, gradients, orientation as Excel writes them."""
    TMP.mkdir(parents=True, exist_ok=True)
    out: dict = {}
    wb = app.Workbooks.Add()
    try:
        ws = wb.Worksheets(1)
        for i, (name, ls, w) in enumerate(LINE_STYLES):
            b = ws.Range(f"A{i + 1}").Borders(9)  # xlEdgeBottom
            b.LineStyle = ls
            b.Weight = w
        ws.Range("C1").Borders(9).Color = 0x0000FF  # red (BGR)
        ws.Range("C1").Borders(9).Weight = 2
        ws.Range("C2").Borders(5).LineStyle = 1  # xlDiagonalDown
        ws.Range("C3").Borders(6).LineStyle = 1  # xlDiagonalUp
        ws.Range("C4").Borders(5).LineStyle = 1
        ws.Range("C4").Borders(6).LineStyle = 1
        ws.Range("C4").Borders(6).Color = 0xFF0000
        # Patterns: Excel's own constants for the 18 entries of Musterformat.
        pats = [1, -4126, -4125, -4124, 17, 18, -4128, -4166, -4121, -4162, 9, 10, 11, 12, 13, 14, 15, 16]
        for i, p in enumerate(pats):
            c = ws.Range(f"E{i + 1}")
            c.Interior.Pattern = p
            c.Interior.PatternColor = 0x0000FF
            c.Interior.Color = 0x00FFFF
        # Fülleffekte: linear gradients at the angles of the Schattierungsarten, and the path ones.
        for i, deg in enumerate((0, 90, 45, 135)):
            c = ws.Range(f"G{i + 1}")
            c.Interior.Pattern = 4000  # xlPatternLinearGradient
            c.Interior.Gradient.Degree = deg
            st = c.Interior.Gradient.ColorStops
            st.Clear()
            st.Add(0).Color = 0xFFFFFF
            st.Add(1).Color = 0xC07000
        c = ws.Range("G5")
        c.Interior.Pattern = 4001  # xlPatternRectangularGradient
        g = c.Interior.Gradient
        g.RectangleLeft, g.RectangleRight, g.RectangleTop, g.RectangleBottom = 0.5, 0.5, 0.5, 0.5
        g.ColorStops.Clear()
        g.ColorStops.Add(0).Color = 0xFFFFFF
        g.ColorStops.Add(1).Color = 0xC07000
        for i, deg in enumerate((45, -45, 90, -90)):
            ws.Range(f"I{i + 1}").Orientation = deg
        ws.Range("I5").Orientation = -4166  # xlVertical
        # Datenbalken with Einfarbige Füllung.
        for i in range(5):
            ws.Range(f"K{i + 1}").Value2 = i + 1
        db = ws.Range("K1:K5").FormatConditions.AddDatabar()
        db.BarFillType = 0  # xlDataBarFillSolid
        db.BarColor.Color = 0xC68E63
        out["solidBar"] = {"fill": db.BarFillType, "border": db.BarBorder.Type, "axis": db.AxisPosition, "direction": db.Direction}
        p = TMP / "formats.xlsx"
        wb.SaveAs(str(p), 51)
    finally:
        wb.Close(False)
    st = styles_xml(p)
    out["borders"] = part(st, "border")
    out["fills"] = part(st, "fill")
    out["xfs"] = re.findall(r"<xf [^>]*>(?:.*?</xf>)?", re.search(r"<cellXfs[^>]*>(.*?)</cellXfs>", st, re.S).group(1), re.S)
    sx = sheet_xml(p)
    out["sheetCf"] = re.findall(r"<conditionalFormatting.*?</conditionalFormatting>", sx, re.S)
    out["sheetExt"] = re.findall(r"<x14:conditionalFormattings>.*?</x14:conditionalFormattings>", sx, re.S)
    out["cells"] = re.findall(r'<c r="[A-K]\d+" s="\d+"', sx)
    return out


def main() -> None:
    pythoncom.CoInitialize()
    xl = ExcelTarget(visible=False, place=False)
    try:
        data = {"sort": sort_refs(xl.app), "formats": formats(xl.app)}
    finally:
        xl.close()
    out = HERE / "results" / "2026-11-09-excel-followup.json"
    out.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(json.dumps(data["sort"], ensure_ascii=False, indent=1))
    print("->", out.name)


if __name__ == "__main__":
    main()
