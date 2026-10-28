"""Builds test workbooks with Microsoft Excel itself (an invisible instance of our own) into
packages/ooxml/test/fixtures/xlsx-*.xlsx, for RaabeCalc's .xlsx reader and writer (D-061).
The workbooks start from Excel's new workbook with the Office 2013-2022 design (Calibri, as
RaabeCalc's new workbooks, D-011); the contents are written for this project.

    .venv/Scripts/python.exe make_xlsx_fixtures.py            # all
    .venv/Scripts/python.exe make_xlsx_fixtures.py --values   # one
"""

from __future__ import annotations

import os
import re
import sys
import tempfile
import zipfile
from pathlib import Path

import pythoncom

from feellab.excel import ExcelTarget

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "packages" / "ooxml" / "test" / "fixtures"
THEME = r"C:\Program Files\Microsoft Office\root\Document Themes 16\Office 2013 - 2022 Theme.thmx"
NEUTRAL = "Raabe Office"


def english(obj, name: str, value) -> None:
    """Sets a property in Excel's English form (pywin32 passes the user's locale otherwise, and
    Excel then reads NumberFormat as NumberFormatLocal)."""
    o = obj._oleobj_
    o.Invoke(o.GetIDsOfNames(name), 1033, pythoncom.DISPATCH_PROPERTYPUT, 0, value)


def rgb(r: int, g: int, b: int) -> int:
    """Excel's colour value (BGR)."""
    return r + (g << 8) + (b << 16)


class Builder:
    def __init__(self, x: ExcelTarget, tmp: Path):
        self.x = x
        self.app = x.app
        # Template: Excel's new workbook with the Office 2013-2022 design applied, reopened so
        # row heights follow Calibri.
        wb = self.app.Workbooks.Add()
        wb.ApplyTheme(THEME)
        self.template = tmp / "vorlage.xlsx"
        wb.SaveAs(str(self.template), 51)
        wb.Close(False)

    def new(self):
        return self.app.Workbooks.Add(str(self.template))

    def save(self, wb, name: str) -> None:
        for prop in ("Author", "Last author"):
            wb.BuiltinDocumentProperties(prop).Value = NEUTRAL
        path = OUT / name
        if path.exists():
            path.unlink()
        wb.SaveAs(str(path), 51)
        wb.Close(False)
        print("saved", name, path.stat().st_size)


def values(b: Builder) -> None:
    wb = b.new()
    ws = wb.Worksheets(1)
    ws.Name = "Werte"
    for i, h in enumerate(["Zahlen", "Texte", "Wahrheitswerte", "Fehler", "Datum und Zeit", "Formeln", "Formatierter Text"]):
        ws.Cells(1, i + 1).Value2 = h
    nums = [42, -3.5, 0.1 + 0.2, 1e-7, 1.2345678901234567e19, 1 / 3, 3.14159, 0, -1e-300, 123456789012]
    for i, n in enumerate(nums):
        ws.Cells(i + 2, 1).Value2 = n
    texts = ["Grüße aus Hannover", "  führende und folgende Leerzeichen  ", "Zeile 1\nZeile 2", '<Tag> & "Zitat"',
             "_x000D_ wörtlich", "'0123", "Tab\tdazwischen", "Smiley \U0001F642", "'=1+1", "ß" * 3 + " " + "Lang " * 60]
    for i, t in enumerate(texts):
        ws.Cells(i + 2, 2).Value2 = t
    ws.Cells(4, 2).WrapText = True
    ws.Cells(2, 3).Value2 = True
    ws.Cells(3, 3).Value2 = False
    for i, e in enumerate(["#N/A", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#NULL!"]):
        ws.Cells(i + 2, 4).Formula = e
    ws.Cells(2, 5).Formula = "10/27/2026"
    ws.Cells(3, 5).Formula = "13:45"
    ws.Cells(4, 5).Formula = "10/27/2026 13:45"
    ws.Cells(5, 5).Formula = "5%"
    ws.Cells(6, 5).Formula = "1/1/1900"
    ws.Cells(7, 5).Formula = "2/28/1900"
    ws.Cells(2, 6).Formula = '=""'
    ws.Cells(3, 6).Formula = '="ab"&"cd"'
    ws.Cells(4, 6).Formula = "=1/0"
    ws.Cells(5, 6).Formula = "=A2>40"
    ws.Cells(6, 6).Formula = "=A2*2"
    r = ws.Cells(2, 7)
    r.Value2 = "fett kursiv rot normal"
    r.GetCharacters(1, 4).Font.Bold = True
    r.GetCharacters(6, 6).Font.Italic = True
    r.GetCharacters(13, 3).Font.Color = rgb(255, 0, 0)
    r = ws.Cells(3, 7)
    r.Value2 = "Fläche 25 m2, Wasser H2O"
    r.GetCharacters(13, 1).Font.Superscript = True
    r.GetCharacters(23, 1).Font.Subscript = True
    r = ws.Cells(4, 7)
    r.Value2 = "Arial 14 und Courier"
    r.GetCharacters(1, 8).Font.Name = "Arial"
    r.GetCharacters(1, 8).Font.Size = 14
    r.GetCharacters(14, 7).Font.Name = "Courier New"
    ws.Cells(8, 9).Interior.Color = rgb(255, 255, 0)  # format only, no value
    ws.Cells(1000, 1).Value2 = "Ende"
    ws.Cells(5, 27).Value2 = "Spalte AA"
    ws.Range("A1").Select()
    b.save(wb, "xlsx-values.xlsx")


def formulas(b: Builder) -> None:
    wb = b.new()
    f = wb.Worksheets(1)
    f.Name = "Formeln"
    d = wb.Worksheets.Add(After=f)
    d.Name = "Daten"
    m = wb.Worksheets.Add(After=d)
    m.Name = "Mein Blatt"
    m.Range("A1").Value2 = 7
    for i in range(10):
        d.Cells(i + 1, 1).Value2 = i + 1
    # Filled formulas: Excel stores them as shared formulas.
    d.Range("B1").Formula = "=A1*2"
    d.Range("B1:B10").FillDown()
    d.Range("C1").Formula = "=A1+$A$1+A$1+$A1"
    d.Range("C1:C10").FillDown()
    d.Range("D2").Formula = "=$A2*B$1+ROW()"
    d.Range("D2:F2").FillRight()
    d.Range("D2:F6").FillDown()
    wb.Names.Add("Zahlen", "=Daten!$A$1:$A$10")
    wb.Names.Add("Erste", "=Daten!$A$1")
    rows = [
        "=SUM(Daten!A1:A10)", "=AVERAGE(Daten!A:A)", "=Daten!$A$1+Daten!A$2+Daten!$A3", '=IF(A1>50,"groß","klein")',
        "=A1>10", "=1/0", "=ROUND(PI(),4)", '=CONCAT("a","b")', "=XLOOKUP(3,Daten!A1:A10,Daten!B1:B10)",
        "=SUM(Daten!1:1)", "=Erste*2", "=SUM(Zahlen)", "='Mein Blatt'!A1+1", '=TEXT(A1,"0.00")', "=-A1^2",
        "=A1&\" Punkte\"", "=COUNTIF(Daten!A1:A10,\">5\")", "=IFERROR(1/0,\"durch null\")", "=SUM(Daten!A1:B3 Daten!B2:C5)",
        "=SUM((Daten!A1:A2,Daten!A5:A6))", "=50%*A1", "=TRUE", "=#N/A",
    ]
    for i, t in enumerate(rows):
        f.Cells(i + 1, 1).Formula = t
    f.Range("C1:C3").FormulaArray = "=Daten!A1:A3*10"
    f.Range("D1").Formula2 = "=SEQUENCE(3)"
    for rng, formula in [("E1:E5", "=SUM(Daten!A1:A3)"), ("F1:F3", '="A1"&Daten!A1'),
                         ("G1:G4", "=SUM(Daten!A:A)+Daten!A1+SUM(Daten!$1:1)"), ("H2:I4", "=Daten!B1+'Mein Blatt'!$A$1+H$1"),
                         ("J1:J4", "=SUM(Daten!B$1:B1)"), ("K1:K3", "=Daten!A1:A2"), ("L1:L3", "=LOG10(Daten!A1)+ATAN2(1,Daten!A1)")]:
        first = rng.split(":")[0]
        f.Range(first).Formula = formula
        f.Range(rng).FillDown()
    f.Range("H2:I4").FillRight()
    f.Activate()
    f.Range("A1").Select()
    b.save(wb, "xlsx-formulas.xlsx")


def styles(b: Builder) -> None:
    wb = b.new()
    ws = wb.Worksheets(1)
    ws.Name = "Formate"
    row = 1

    def cell(label: str):
        nonlocal row
        ws.Cells(row, 1).Value2 = label
        c = ws.Cells(row, 2)
        c.Value2 = "Beispiel 123"
        row += 1
        return c

    cell("fett").Font.Bold = True
    cell("kursiv").Font.Italic = True
    cell("unterstrichen").Font.Underline = 2
    cell("doppelt").Font.Underline = -4119
    cell("buchhaltung einfach").Font.Underline = 4
    cell("buchhaltung doppelt").Font.Underline = 5
    cell("durchgestrichen").Font.Strikethrough = True
    cell("hochgestellt").Font.Superscript = True
    cell("tiefgestellt").Font.Subscript = True
    c = cell("Arial 10")
    c.Font.Name, c.Font.Size = "Arial", 10
    c = cell("Times 14 fett kursiv")
    c.Font.Name, c.Font.Size, c.Font.Bold, c.Font.Italic = "Times New Roman", 14, True, True
    c = cell("Überschriften-Schriftart")
    c.Font.Name = "Calibri Light"
    c.Font.Size = 16
    cell("Farbe RGB").Font.Color = rgb(192, 0, 0)
    c = cell("Designfarbe Akzent 1")
    c.Font.ThemeColor = 5
    c = cell("Designfarbe Akzent 2 heller 40")
    c.Font.ThemeColor = 6
    c.Font.TintAndShade = 0.3999
    c = cell("Text 1 heller 50")
    c.Font.ThemeColor = 2
    c.Font.TintAndShade = 0.4999
    cell("Farbindex 3").Font.ColorIndex = 3
    cell("Füllung RGB").Interior.Color = rgb(255, 230, 153)
    c = cell("Füllung Design dunkler 25")
    c.Interior.ThemeColor = 7
    c.Interior.TintAndShade = -0.2499
    c = cell("Füllung Hintergrund 1 dunkler 15")
    c.Interior.ThemeColor = 1
    c.Interior.TintAndShade = -0.1499
    c = cell("Muster")
    c.Interior.Pattern = 16  # xlPatternCrissCross
    c.Interior.PatternColor = rgb(0, 112, 192)
    c.Interior.Color = rgb(255, 255, 204)
    c = cell("Farbverlauf")
    c.Interior.Pattern = 4000  # xlPatternLinearGradient
    c.Interior.Gradient.Degree = 90
    c.Interior.Gradient.ColorStops.Clear()
    c.Interior.Gradient.ColorStops.Add(0).Color = rgb(255, 255, 255)
    c.Interior.Gradient.ColorStops.Add(1).Color = rgb(91, 155, 213)
    # Borders: (label, index, line style, weight)
    for label, style, weight in [("dünn", 1, 2), ("mittel", 1, -4138), ("dick", 1, 4), ("gestrichelt", -4115, 2), ("gepunktet", -4118, 2),
                                 ("doppelt", -4119, 4), ("haarfein", 1, 1), ("strichpunkt", 4, 2), ("strichpunktpunkt", 5, 2),
                                 ("mittel gestrichelt", -4115, -4138), ("schräg strichpunkt", 13, -4138)]:
        c = cell(f"Rahmen {label}")
        for idx in (7, 8, 9, 10):  # left, top, bottom, right
            c.Borders(idx).LineStyle = style
            c.Borders(idx).Weight = weight
    c = cell("Rahmen farbig unten")
    c.Borders(9).LineStyle = 1
    c.Borders(9).Weight = -4138
    c.Borders(9).Color = rgb(255, 0, 0)
    c = cell("Rahmen Designfarbe")
    c.Borders(8).LineStyle = 1
    c.Borders(8).ThemeColor = 5
    c = cell("Diagonal")
    c.Borders(5).LineStyle = 1  # xlDiagonalDown
    c.Borders(6).LineStyle = 1  # xlDiagonalUp
    for label, h in [("links", -4131), ("zentriert", -4108), ("rechts", -4152), ("füllen", 5), ("Blocksatz", -4130),
                     ("über Auswahl zentrieren", 7), ("verteilt", -4117)]:
        cell(f"horizontal {label}").HorizontalAlignment = h
    for label, v in [("oben", -4160), ("mittig", -4108), ("Blocksatz", -4130), ("verteilt", -4117)]:
        c = cell(f"vertikal {label}")
        c.VerticalAlignment = v
        c.RowHeight = 30
    c = cell("Zeilenumbruch mit einem längeren Text")
    c.Value2 = "Zeilenumbruch mit einem längeren Text, der umbricht"
    c.WrapText = True
    c = cell("Einzug 2")
    c.HorizontalAlignment = -4131
    c.IndentLevel = 2
    cell("gedreht 45").Orientation = 45
    cell("gedreht -90").Orientation = -90
    cell("vertikal gestapelt").Orientation = -4166  # xlVertical
    cell("verkleinern").ShrinkToFit = True
    c = cell("nicht gesperrt")
    c.Locked = False
    c = cell("ausgeblendet")
    c.FormulaHidden = True
    for st in ["Gut", "Schlecht", "Neutral", "Überschrift 1", "Überschrift 2", "Notiz", "Prozent", "Währung", "Ergebnis", "Akzent1", "40 % - Akzent3"]:
        cell(f"Formatvorlage {st}").Style = st
    ws.Columns("A:B").AutoFit()
    ws.Range("A1").Select()
    b.save(wb, "xlsx-styles.xlsx")


NUMFMTS = [
    "General", "0", "0.00", "#,##0", "#,##0.00", "0%", "0.00%", "0.00E+00", "# ?/?", "# ??/??", "m/d/yyyy", "d-mmm-yy", "d-mmm",
    "mmm-yy", "h:mm AM/PM", "h:mm:ss AM/PM", "h:mm", "h:mm:ss", "m/d/yyyy h:mm", "#,##0 ;(#,##0)", "#,##0 ;[Red](#,##0)",
    "mm:ss", "[h]:mm:ss", "mm:ss.0", "##0.0E+0", "@", "0.000", "#,##0.00 [$€-407]", '#,##0.00 "€"', "[Red]0.00",
    '[<0][Red]-0.0;[>100][Blue]0.0;0.0', "dddd, d. mmmm yyyy", "yyyy-mm-dd", '0" Stück"', "00000", "#,##0.00;-#,##0.00;\"-\"",
    "[$-de-DE]mmmm yyyy", "0.0%", "\\A\\b\\c 0",
]
NUMFMTS_LOCAL = ["TT.MM.JJJJ", "#.##0,00 €", "0,0", "TTTT", "MMM JJ", "hh:mm:ss", "_-* #.##0,00 €_-;-* #.##0,00 €_-;_-* \"-\"?? €_-;_-@_-"]


def numfmt(b: Builder) -> None:
    wb = b.new()
    ws = wb.Worksheets(1)
    ws.Name = "Zahlenformate"
    samples = [1234.5678, -1234.5678, 0.25, 46322.5729]
    row = 1
    for code in NUMFMTS:
        ws.Cells(row, 1).Value2 = code
        for i, v in enumerate(samples):
            c = ws.Cells(row, 2 + i)
            c.Value2 = v
            english(c, "NumberFormat", code)
        row += 1
    for code in NUMFMTS_LOCAL:
        ws.Cells(row, 1).Value2 = code
        for i, v in enumerate(samples):
            c = ws.Cells(row, 2 + i)
            c.Value2 = v
            c.NumberFormatLocal = code
        row += 1
    ws.Columns("A:E").AutoFit()
    ws.Range("A1").Select()
    b.save(wb, "xlsx-numfmt.xlsx")


def layout(b: Builder) -> None:
    wb = b.new()
    ws = wb.Worksheets(1)
    ws.Name = "Layout"
    ws.Range("A1").Value2 = "Zusammengeführter Titel"
    ws.Range("A1:C1").Merge()
    ws.Range("A1:C1").HorizontalAlignment = -4108
    ws.Range("B3").Value2 = "Block"
    ws.Range("B3:C5").Merge()
    for r in range(1, 31):
        ws.Cells(r, 8).Value2 = r
    ws.Columns("A").ColumnWidth = 20
    ws.Columns("D").Hidden = True
    ws.Columns("E").ColumnWidth = 2.5
    ws.Columns("F:G").ColumnWidth = 15.29
    ws.Rows(2).RowHeight = 30
    ws.Rows(6).Hidden = True
    ws.Rows(12).RowHeight = 4.5
    ws.Rows("8:10").Group()
    ws.Columns("I:J").Group()
    ws.Outline.ShowLevels(0, 1)  # columns I:J collapsed
    ws.Activate()
    win = wb.Windows(1)
    win.ScrollRow = 1
    win.ScrollColumn = 1
    # Excel freezes panes only in a visible window: show it for a moment, without drawing.
    b.app.ScreenUpdating = False
    b.app.Visible = True
    try:
        win.SplitColumn = 1
        win.SplitRow = 2
        win.FreezePanes = True
    finally:
        b.app.Visible = False
        b.app.ScreenUpdating = True
    win.Zoom = 120
    ws.Range("D12").Select()
    v = wb.Worksheets.Add(After=ws)
    v.Name = "Ansicht"
    v.Range("A1").Value2 = "ohne Gitternetz"
    v.Activate()
    win = wb.Windows(1)
    win.DisplayGridlines = False
    win.DisplayHeadings = False
    b.app.Union(v.Range("B2:C3"), v.Range("E5"), v.Range("G7:G9")).Select()
    w = wb.Worksheets.Add(After=v)
    w.Name = "Teilung"
    w.Range("A1").Value2 = "geteilt und gescrollt"
    w.StandardWidth = 12
    w.Activate()
    win = wb.Windows(1)
    win.SplitColumn = 2
    win.SplitRow = 3
    win.ScrollRow = 10
    w.Range("K20").Select()
    ws.Activate()
    b.save(wb, "xlsx-layout.xlsx")


def sheets(b: Builder) -> None:
    wb = b.new()
    s1 = wb.Worksheets(1)
    s1.Name = "Erstes"
    for i in range(4):
        s1.Cells(i + 1, 1).Value2 = i + 1
        s1.Cells(i + 1, 2).Value2 = f"Zeile {i + 1}"
    names = ["Versteckt", "Sehr versteckt", "Rot", "Akzent", "Letztes"]
    prev = s1
    made = {}
    for n in names:
        s = wb.Worksheets.Add(After=prev)
        s.Name = n
        s.Range("A1").Value2 = n
        made[n] = prev = s
    made["Versteckt"].Range("B2").Formula = "=Erstes!A1*3"
    made["Rot"].Tab.Color = rgb(255, 0, 0)
    made["Akzent"].Tab.ThemeColor = 5
    made["Akzent"].Tab.TintAndShade = 0.3999
    made["Letztes"].Range("A2").Formula = "=Versteckt!B2+Steuersatz"
    # Names.Add reads its formula in the user's locale (pywin32 passes it).
    wb.Names.Add("Steuersatz", "=0,19")
    wb.Names.Add("Bereich", "=Erstes!$A$1:$B$4")
    wb.Names.Add("Summe", "=SUMME(Erstes!$A$1:$A$4)")
    made["Rot"].Names.Add("Lokal", "=Rot!$C$3")
    wb.Names.Add("Unsichtbar", "=Erstes!$B$2", False)
    s1.PageSetup.PrintArea = "$A$1:$B$4"
    made["Rot"].Activate()
    made["Versteckt"].Visible = 0
    made["Sehr versteckt"].Visible = 2
    b.save(wb, "xlsx-sheets.xlsx")


def objects(b: Builder, tmp: Path) -> None:
    from PIL import Image, ImageDraw
    png = tmp / "bild.png"
    img = Image.new("RGB", (120, 80), (255, 255, 255))
    ImageDraw.Draw(img).ellipse((10, 10, 110, 70), fill=(0, 112, 192))
    img.save(png)
    wb = b.new()
    ws = wb.Worksheets(1)
    ws.Name = "Objekte"
    data = [("Fach", "Klasse 7", "Klasse 8"), ("Deutsch", 2.3, 2.1), ("Mathe", 2.8, 3.0), ("Englisch", 2.0, 2.4),
            ("Physik", 3.1, 2.7), ("Kunst", 1.7, 1.9)]
    for r, row in enumerate(data):
        for c, v in enumerate(row):
            ws.Cells(r + 1, c + 1).Value2 = v
    co = ws.Shapes.AddChart2(201, 51, 300, 10, 360, 220)
    co.Chart.SetSourceData(ws.Range("A1:C6"))
    co.Chart.HasTitle = True
    co.Chart.ChartTitle.Text = "Notenschnitt"
    ws.Range("B2").AddComment("Bitte prüfen")
    ws.Range("C4").AddComment("Zweite Notiz\nmit zwei Zeilen")
    fc = ws.Range("B2:C6").FormatConditions.Add(1, 5, "=2,5")  # xlCellValue, xlGreater
    fc.Interior.Color = rgb(255, 199, 206)
    fc.Font.Color = rgb(156, 0, 6)
    ws.Range("D2:D6").Formula = "=B2-C2"
    ws.Range("D2:D6").FormatConditions.AddColorScale(3)
    ws.Range("E2:E6").Value2 = 1
    for i in range(5):
        ws.Cells(i + 2, 5).Value2 = (i + 1) * 20
    ws.Range("E2:E6").FormatConditions.AddDatabar()
    v = ws.Range("F2:F6").Validation
    v.Add(3, 1, 1, "Ja;Nein")  # list
    v = ws.Range("G2:G6").Validation
    v.Add(1, 1, 1, "1", "6")  # whole number between
    v.InputTitle = "Note"
    v.InputMessage = "Eine Note von 1 bis 6"
    ws.Range("A1:C6").AutoFilter(2, ">2")
    ws.Hyperlinks.Add(Anchor=ws.Range("A9"), Address="https://www.raabeschule.de/", TextToDisplay="Raabeschule")
    ws.Hyperlinks.Add(Anchor=ws.Range("A10"), Address="", SubAddress="'Tabelle'!B2", TextToDisplay="Zur Tabelle",
                      ScreenTip="Springt zur Tabelle")
    ws.Hyperlinks.Add(Anchor=ws.Range("A11"), Address="mailto:sekretariat@example.org", TextToDisplay="E-Mail")
    pic = ws.Shapes.AddPicture(str(png), False, True, 300, 250, 90, 60)
    pic.Name = "Bild 1"
    ws.PageSetup.CenterHeader = "Raabeschule"
    ws.PageSetup.RightFooter = "Seite &P"
    t = wb.Worksheets.Add(After=ws)
    t.Name = "Tabelle"
    for r, row in enumerate([("Name", "Punkte", "Note"), ("Ada", 47, 1), ("Ben", 35, 2), ("Cem", 28, 3)]):
        for c, v in enumerate(row):
            t.Cells(r + 1, c + 1).Value2 = v
    lo = t.ListObjects.Add(1, t.Range("A1:C4"), None, 1)
    lo.Name = "Ergebnisse"
    lo.ShowTotals = True
    ws.Activate()
    ws.Range("A1").Select()
    b.save(wb, "xlsx-objects.xlsx")


def date1904(b: Builder) -> None:
    wb = b.new()
    wb.Date1904 = True
    ws = wb.Worksheets(1)
    ws.Name = "1904"
    ws.Range("A1").Formula = "10/27/2026"
    ws.Range("A2").Formula = "1/1/1904"
    ws.Range("A3").Formula = "=A1+1"
    ws.Range("A4").Formula = "=DATE(2000,1,1)"
    ws.Range("B1").Value2 = 0
    english(ws.Range("B1"), "NumberFormat", "dd.mm.yyyy")
    ws.Range("A1").Select()
    b.save(wb, "xlsx-1904.xlsx")


def blank(b: Builder) -> None:
    b.save(b.new(), "xlsx-blank.xlsx")


def validation(b: Builder) -> None:
    """Datenüberprüfung of every kind (probe_excel_validation.py records how Excel checks them):
    lists from text, cells, another sheet (x14 in the extLst), a name and entries with commas
    (x12ac:list); numbers, dates, times, text lengths, formulas; the three error styles."""
    wb = b.new()
    ws = wb.Worksheets(1)
    ws.Name = "Prüfung"
    lst = wb.Worksheets.Add(After=ws)
    lst.Name = "Listen"
    for i, t in enumerate(["Deutsch", "Mathe", "Physik"]):
        lst.Cells(i + 1, 1).Value2 = t
    lst.Range("B1").Value2 = 10
    wb.Names.Add("Faecher", "=Listen!$A$1:$A$3")
    ws.Activate()
    heads = ["Name", "Teilnahme", "Kurs", "Fach", "Note", "Punkte", "Datum", "Kurse", "Uhrzeit", "Kommentar", "Zahl", "Fach (Name)",
             "Stufe", "Hinweis", "Gruppe", "Höchstens B1"]
    for i, t in enumerate(heads):
        ws.Cells(1, i + 1).Value2 = t
    for i, t in enumerate(["Deutsch", "Mathe", "Englisch", "Physik"]):
        ws.Cells(i + 2, 8).Value2 = t

    def add(addr: str, kind: int, style: int, op: int, f1: str | None = None, f2: str | None = None):
        # Areas joined with Union: the range text's separator follows the locale through COM.
        parts = [ws.Range(x) for x in addr.split(",")]
        v = (b.app.Union(*parts) if len(parts) > 1 else parts[0]).Validation
        args = [kind, style, op] + [f for f in (f1, f2) if f is not None]
        v.Add(*args)
        return v

    v = add("B2:B6", 3, 1, 1, "Ja;Nein")
    v.InputTitle, v.InputMessage = "Teilnahme", "Ja oder Nein wählen"
    add("C2:C6", 3, 1, 1, "=$H$2:$H$5")
    add("D2:D6", 3, 1, 1, "=Listen!$A$1:$A$3")
    v = add("E2:E6", 1, 1, 1, "1", "6")
    v.ErrorTitle, v.ErrorMessage = "Ungültige Note", "Nur ganze Noten von 1 bis 6."
    add("F2:F6", 2, 2, 5, "0")
    add("G2:G6", 4, 3, 1, "01.08.2026", "31.07.2027")
    v = add("I2:I6", 5, 1, 6, "12:00")
    v.ShowError = False
    add("J2:J6", 6, 1, 8, "20")
    add("K2:K6", 7, 1, 1, "=ISTZAHL(K2)")
    v = add("L2:L6", 3, 1, 1, "=Faecher")
    v.InCellDropdown = False
    add("M2:M6", 3, 1, 1, "1;2;2,5")
    v = add("N2", 0, 1, 1)
    v.InputTitle, v.InputMessage = "Hinweis", "Hier steht nur ein Hinweis."
    add("O2:O4,O8:O9", 3, 1, 1, 'a,b;"c";d')
    add("P2:P6", 7, 2, 1, "=P2<=Listen!$B$1")
    ws.Range("A1").Select()
    b.save(wb, "xlsx-validation.xlsx")


def protection(b: Builder) -> None:
    """Blatt schützen with a password (Excel's SHA-512 hash) and two allowed actions; one sheet
    protected without a password; unlocked input cells (Zellen formatieren > Schutz)."""
    wb = b.new()
    ws = wb.Worksheets(1)
    ws.Name = "Geschützt"
    ws.Range("A1").Value = "Name"
    ws.Range("B1").Value = "Note"
    ws.Range("B2:B5").Locked = False
    ws.Protect(Password="Raabe", AllowFormattingCells=True, AllowSorting=True)
    two = wb.Worksheets.Add(After=ws)
    two.Name = "Ohne Kennwort"
    two.Range("A1").Value = 1
    two.Protect()
    ws.Activate()
    b.save(wb, "xlsx-protection.xlsx")


ALL = {"values": values, "formulas": formulas, "styles": styles, "numfmt": numfmt, "layout": layout, "sheets": sheets,
       "date1904": date1904, "blank": blank, "validation": validation, "protection": protection}


def anonymize(path: Path) -> None:
    """Excel writes the signed-in Office user as the last author whatever the document
    properties say: set it to the neutral name (the only change to Excel's file)."""
    with zipfile.ZipFile(path) as z:
        items = [(info, z.read(info)) for info in z.infolist()]
    out = []
    for info, data in items:
        if info.filename == "docProps/core.xml":
            text = data.decode("utf-8")
            text = re.sub(r"<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>", f"<cp:lastModifiedBy>{NEUTRAL}</cp:lastModifiedBy>", text)
            text = re.sub(r"<dc:creator>[^<]*</dc:creator>", f"<dc:creator>{NEUTRAL}</dc:creator>", text)
            data = text.encode("utf-8")
        if info.filename.startswith("xl/printerSettings/"):
            # DEVMODE: the owner's printer model (dmDeviceName, also repeated in the driver's
            # part) becomes a neutral name of the same length.
            name = data[:64].decode("utf-16-le").rstrip("\0").rstrip()
            if name:
                neutral = "Drucker".ljust(len(name))[: len(name)]
                data = data.replace(name.encode("utf-16-le"), neutral.encode("utf-16-le"))
        out.append((info, data))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for info, data in out:
            z.writestr(info, data)


def check_private(path: Path, user: str) -> None:
    """No user names or private paths in a fixture."""
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            data = z.read(n)
            texts = ["C:\\Users", user, os.environ.get("USERNAME", "")]
            # Printer settings (DEVMODE) are UTF-16.
            for needle in [t.encode(enc) for t in texts if t for enc in ("utf-8", "utf-16-le")]:
                if needle in data:
                    raise SystemExit(f"{path.name}: private text in {n}")


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    only = [a[2:] for a in sys.argv[1:] if a.startswith("--")]
    x = ExcelTarget(visible=False)
    user = ""
    try:
        user = str(x.app.UserName)
        # Notes record Excel's user name as their author.
        x.set_option("UserName", NEUTRAL)
        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            b = Builder(x, tmp)
            for name in only or [*ALL, "objects"]:
                if name == "objects":
                    objects(b, tmp)
                else:
                    ALL[name](b)
    finally:
        x.close()
    for f in sorted(OUT.glob("xlsx-*.xlsx")):
        anonymize(f)
        check_private(f, user)
    print("no private text in the fixtures")


if __name__ == "__main__":
    main()
