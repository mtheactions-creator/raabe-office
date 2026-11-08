"""Records Excel's ZELLE (CELL), INFO, BLATT (SHEET), BLÄTTER (SHEETS) and HYPERLINK (D-080) as the
reference for packages/sheet-engine/test/info-excel.test.ts.

Starts an invisible Excel of its own with three sheets; no file is saved (ZELLE("Dateiname") is
recorded for an unsaved workbook only, so no path of this PC lands in the results).

    .venv/Scripts/python.exe probe_excel_info.py

Writes packages/sheet-engine/test/fixtures/excel-info.json. No real input."""

from __future__ import annotations

import json
import time
from pathlib import Path

from feellab.excel import ExcelTarget

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "packages" / "sheet-engine" / "test" / "fixtures" / "excel-info.json"
ERRORS = {-2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?", -2146826288: "#NULL!",
          -2146826252: "#NUM!", -2146826265: "#REF!", -2146826273: "#VALUE!", -2146826215: "#SPILL!", -2146826238: "#CALC!"}

# Set up on Tabelle1 (value, number format).
SETUP = [("A1", 12.5, "Standard"), ("A2", "Text", "Standard"), ("A3", 0.25, "0%"), ("A4", 1234.5, "#.##0,00"),
         ("A5", 45000, "TT.MM.JJJJ"), ("A6", -3, "0;(0)"), ("A7", 3.14159, "0,00"), ("A8", 5, "0,00E+00"),
         ("A9", 7, "#.##0 €"), ("A10", 0.5, "hh:mm")]
KEYS = ["Adresse", "Spalte", "Farbe", "Inhalt", "Format", "Klammern", "Präfix", "Schutz", "Zeile", "Typ", "Breite",
        "address", "col", "contents", "format", "row", "type", "width", "Dateiname", "filename"]
CASES = [f'=ZELLE("{k}";A{i})' for k in KEYS for i in (1, 2, 11)]
CASES += [f'=ZELLE("Format";A{i})' for i in range(3, 11)] + [f'=ZELLE("Klammern";A{i})' for i in (6, 9)] + [f'=ZELLE("Farbe";A{i})' for i in (6,)]
CASES += ['=ZELLE("Adresse";Tabelle2!B3)', '=ZELLE("Adresse";A1:B2)', '=ZELLE("Zeile")', '=ZELLE("unbekannt";A1)', '=ZELLE("Breite";C1)',
          '=ZELLE("Präfix";A2)', '=ZELLE("Typ";A1:A2)']
CASES += ['=INFO("Rechenmodus")', '=INFO("recalc")', '=INFO("Dateienzahl")', '=INFO("numfile")', '=INFO("System")', '=INFO("system")',
          '=INFO("Ursprung")', '=INFO("origin")', '=INFO("Version")', '=INFO("release")', '=INFO("osversion")', '=INFO("Systemversion")',
          '=INFO("unbekannt")']  # INFO("Verzeichnis") is left out: it names a folder of this PC
CASES += ["=BLATT()", "=BLATT(Tabelle3!A1)", '=BLATT("Tabelle2")', '=BLATT("Fehlt")', "=BLATT(Tabelle1:Tabelle3!A1)", "=BLÄTTER()",
          "=BLÄTTER(Tabelle1:Tabelle3!A1)", "=BLÄTTER(A1)", "=BLATT(1)", "=BLÄTTER(Tabelle2!A1:B2)"]
CASES += ['=HYPERLINK("https://example.org")', '=HYPERLINK("https://example.org";"Beispiel")', '=HYPERLINK("#Tabelle2!A1";42)',
          '=HYPERLINK(A2)']


def val(v):
    if isinstance(v, int) and v in ERRORS:
        return {"error": ERRORS[v]}
    if isinstance(v, tuple):
        return [[val(x) for x in row] for row in v]
    return v


def main() -> None:
    xl = ExcelTarget(visible=False)
    out: dict = {"setup": SETUP, "cases": []}
    try:
        while xl.wb.Worksheets.Count < 3:
            xl.wb.Worksheets.Add(After=xl.wb.Worksheets(xl.wb.Worksheets.Count))
        sh = xl.wb.Worksheets(1)
        sh.Activate()
        out["sheets"] = [xl.wb.Worksheets(i).Name for i in range(1, xl.wb.Worksheets.Count + 1)]
        for a, v, fmt in SETUP:
            sh.Range(a).Value = v
            try:
                sh.Range(a).NumberFormatLocal = fmt
            except Exception:  # noqa: BLE001 - recorded as it failed
                out.setdefault("format_failed", []).append(fmt)
        sh.Range("A2").Value = "Text"
        out["col_width_A"] = sh.Range("A1").ColumnWidth
        for i, f in enumerate(CASES):
            r = sh.Range(f"E{i + 1}")
            try:
                r.Formula2Local = f
            except Exception:  # noqa: BLE001
                out["cases"].append({"typed": f, "rejected": True})
                continue
            time.sleep(0.01)
            spill = r.HasSpill
            area = r.SpillingToRange if spill else r
            out["cases"].append({"cell": f"E{i + 1}", "typed": f, "local": r.Formula2Local, "formula": r.Formula2, "value": val(area.Value2),
                                 "text": r.Text, "font_color": r.Font.Color, "underline": r.Font.Underline, "style": r.Style.NameLocal})
        # ZELLE("Format"/"Klammern"/"Farbe") for many number formats (English codes set, German results).
        # German codes through NumberFormatLocal (this COM binding reads NumberFormat in the local form too).
        fmts = ["Standard", "0", "0,00", "#.##0", "#.##0,00", "0%", "0,00%", "0,00E+00", "# ?/?", "#.##0 €", "#.##0,00 €",
                "#.##0;-#.##0", "#.##0;[Rot]-#.##0", "0;(0)", "0,00;(0,00)", "TT.MM.JJJJ", "T.M.JJJJ", "TT. MMM JJ", "T. MMM", "MMM JJ",
                "TT.MM.JJJJ hh:mm", "hh:mm", "hh:mm:ss", "h:mm AM/PM", "h:mm:ss AM/PM", "@", "0;[Rot]0", "TTTT, T. MMMM JJJJ", "[hh]:mm"]
        out["formats"] = []
        for i, fmt in enumerate(fmts):
            c = sh.Range(f"H{i + 1}")
            c.Value = 1.5
            try:
                c.NumberFormatLocal = fmt
            except Exception:  # noqa: BLE001
                out["formats"].append({"format": fmt, "failed": True})
                continue
            row = {"format": fmt, "english": c.NumberFormat}
            for k in ("Format", "Klammern", "Farbe"):
                r = sh.Range(f"I{i + 1}")
                r.Formula2Local = f'=ZELLE("{k}";H{i + 1})'
                row[k] = val(r.Value2)
            out["formats"].append(row)
    finally:
        xl.close()
    OUT.write_text(json.dumps(out, indent=1, ensure_ascii=False) + "\n", encoding="utf-8", newline="\n")
    for c in out["cases"]:
        print(c["typed"], "->", c.get("value"), "REJECTED" if c.get("rejected") else "", c.get("style", ""))


if __name__ == "__main__":
    main()
