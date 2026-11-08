"""The look corpus, Excel part (D-080 D): school-like workbooks built by our own invisible Excel in
TEMP/raabe-corpus/excel (never committed, never the owner's files). Each one is print-ready, so Excel's
PDF and RaabeCalc's PDF of the same file can be compared page by page (compare_corpus.py).

    .venv/Scripts/python.exe make_corpus_excel.py [name ...]
"""
from __future__ import annotations

import random
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw

from feellab.excel import ExcelTarget

OUT = Path(tempfile.gettempdir()) / "raabe-corpus" / "excel"
NAMES = ["Anna", "Ben", "Cem", "Dora", "Emil", "Fiona", "Gül", "Hannes", "Ida", "Jonas", "Kira", "Leon", "Mia", "Noah", "Ole", "Paula",
         "Quentin", "Rosa", "Sami", "Tara", "Umut", "Vera", "Wim", "Xenia", "Yusuf", "Zoe"]
SUBJECTS = ["Mathe", "Deutsch", "Englisch", "Bio", "Physik", "Chemie", "Sport", "Kunst", "Musik", "Geschichte", "Erdkunde", "Religion"]
rnd = random.Random(8)

# Excel constants
XL_CENTER, XL_LEFT, XL_RIGHT = -4108, -4131, -4152
XL_CONT, XL_THIN, XL_MEDIUM, XL_THICK = 1, 2, -4138, 4
XL_LANDSCAPE = 2


def rgb(r: int, g: int, b: int) -> int:
    return r + (g << 8) + (b << 16)


def put(ws, top_left: str, rows: list[list]) -> None:
    r0 = ws.Range(top_left)
    for i, row in enumerate(rows):
        for k, v in enumerate(row):
            c = r0.Offset(i + 1, k + 1)
            if isinstance(v, str) and v.startswith("="):
                c.Formula2Local = v
            else:
                c.Value = v


def box(rng, weight=XL_THIN) -> None:
    for i in (7, 8, 9, 10, 11, 12):  # edges and inside lines
        try:
            b = rng.Borders(i)
            b.LineStyle = XL_CONT
            b.Weight = weight
        except Exception:  # noqa: BLE001 - inside lines of a single row/column
            pass


def header(rng, fill=rgb(31, 78, 121), color=rgb(255, 255, 255)) -> None:
    rng.Font.Bold = True
    rng.Interior.Color = fill
    rng.Font.Color = color
    rng.HorizontalAlignment = XL_CENTER


def notenliste(wb) -> None:
    ws = wb.Worksheets(1)
    ws.Name = "Klasse 8b"
    ws.Range("A1").Value = "Notenliste Mathematik, Klasse 8b"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    rows = [["Name", "KA 1", "KA 2", "KA 3", "Mündlich", "Schnitt", "Zeugnis"]]
    for i, n in enumerate(NAMES[:20]):
        r = i + 4
        rows.append([n] + [rnd.choice([1, 2, 2, 3, 3, 3, 4, 4, 5, 6]) for _ in range(4)] + [f"=MITTELWERT(B{r}:E{r})", f"=RUNDEN(F{r};0)"])
    put(ws, "A3", rows)
    header(ws.Range("A3:G3"))
    box(ws.Range("A3:G23"))
    ws.Range("F4:F23").NumberFormatLocal = "0,00"
    fc = ws.Range("G4:G23").FormatConditions.Add(1, 7, "=5")  # cell value >= 5
    fc.Font.Color = rgb(192, 0, 0)
    fc.Interior.Color = rgb(255, 199, 206)
    ws.Columns("A").ColumnWidth = 14
    ws.Range("A25").Value = "Klassenschnitt"
    ws.Range("F25").Formula2Local = "=MITTELWERT(F4:F23)"
    ws.Range("F25").NumberFormatLocal = "0,00"
    ws.Range("A25:G25").Font.Bold = True
    ws.PageSetup.PrintTitleRows = "$3:$3"
    ws.PageSetup.CenterHorizontally = True


def notenspiegel(wb) -> None:
    ws = wb.Worksheets(1)
    ws.Name = "Notenspiegel"
    ws.Range("A1").Value = "Notenspiegel Klassenarbeit"
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    put(ws, "A2", [["Note", "Anzahl"]] + [[n, c] for n, c in zip(range(1, 7), [2, 5, 8, 6, 3, 1])])
    header(ws.Range("A3:B3"), fill=rgb(226, 239, 218), color=rgb(0, 0, 0))
    box(ws.Range("A3:B9"))
    ws.Range("A11").Value = "Schnitt"
    ws.Range("B11").Formula2Local = "=SUMMENPRODUKT(A4:A9;B4:B9)/SUMME(B4:B9)"
    ws.Range("B11").NumberFormatLocal = "0,0"
    ch = ws.Shapes.AddChart2(201, 51, 180, 30, 360, 220).Chart
    ch.SetSourceData(ws.Range("A3:B9"))
    ch.HasTitle = True
    ch.ChartTitle.Text = "Verteilung der Noten"


def stundenplan(wb) -> None:
    ws = wb.Worksheets(1)
    ws.Name = "Stundenplan"
    days = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag"]
    ws.Range("A1:F1").Merge()
    ws.Range("A1").Value = "Stundenplan 8b, 2. Halbjahr"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").HorizontalAlignment = XL_CENTER
    put(ws, "A2", [["Stunde"] + days])
    colors = {"Mathe": rgb(189, 215, 238), "Deutsch": rgb(248, 203, 173), "Englisch": rgb(255, 230, 153), "Bio": rgb(198, 224, 180),
              "Physik": rgb(180, 198, 231), "Sport": rgb(237, 237, 237), "Kunst": rgb(255, 204, 255), "Musik": rgb(204, 236, 255),
              "Geschichte": rgb(244, 176, 132), "Erdkunde": rgb(169, 208, 142), "Religion": rgb(255, 242, 204), "Chemie": rgb(221, 235, 247)}
    for h in range(1, 7):
        ws.Cells(h + 3, 1).Value = f"{h}. Std."
        for d in range(5):
            s = rnd.choice(SUBJECTS)
            c = ws.Cells(h + 3, d + 2)
            c.Value = s
            c.Interior.Color = colors[s]
            c.HorizontalAlignment = XL_CENTER
            c.VerticalAlignment = -4108
    ws.Range("A6:F6").Insert(-4121)
    ws.Range("A6").Value = "Pause"
    ws.Range("A6:F6").Merge()
    ws.Range("A6").HorizontalAlignment = XL_CENTER
    ws.Range("A6").Font.Italic = True
    header(ws.Range("A3:F3"), fill=rgb(68, 84, 106))
    box(ws.Range("A3:F10"), XL_MEDIUM)
    ws.Rows("4:10").RowHeight = 30
    ws.Columns("B:F").ColumnWidth = 14
    ws.PageSetup.Orientation = XL_LANDSCAPE


def klassenkasse(wb) -> None:
    ws = wb.Worksheets(1)
    ws.Name = "Kasse"
    put(ws, "A1", [["Datum", "Zweck", "Einnahme", "Ausgabe", "Stand"]])
    stand = "=C{r}-D{r}"
    items = [("Klassenfahrt Anzahlung", 450, 0), ("Kopien", 0, 23.5), ("Sommerfest Kuchen", 87.2, 0), ("Ausflug Zoo", 0, 312), ("Spende", 50, 0),
             ("Bastelmaterial", 0, 41.9), ("Theaterkarten", 0, 120), ("Elternbeitrag", 380, 0)]
    for i, (z, e, a) in enumerate(items):
        r = i + 3
        ws.Cells(r, 1).Value = f"{(i * 9) % 28 + 1}.{i % 12 + 1}.2026"
        ws.Cells(r, 2).Value = z
        ws.Cells(r, 3).Value = e
        ws.Cells(r, 4).Value = a
        ws.Cells(r, 5).Formula2Local = stand.format(r=r) if i == 0 else f"=E{r - 1}+C{r}-D{r}"
    ws.Range("A3:A10").NumberFormatLocal = "TT.MM.JJJJ"
    ws.Range("C3:E10").NumberFormatLocal = "#.##0,00 €"
    header(ws.Range("A2:E2"), fill=rgb(112, 48, 160))
    box(ws.Range("A2:E10"))
    ws.Range("E3:E10").FormatConditions.AddDatabar()
    ws.Columns("B").ColumnWidth = 24
    ws.Columns("C:E").ColumnWidth = 13
    ws.Range("D12").Value = "Endstand"
    ws.Range("E12").Formula2Local = "=E10"
    ws.Range("E12").NumberFormatLocal = "#.##0,00 €"
    ws.Range("D12:E12").Font.Bold = True
    ws.Range("E12").Borders(9).LineStyle = -4119  # double bottom


def messreihe(wb) -> None:
    ws = wb.Worksheets(1)
    ws.Name = "Messreihe"
    ws.Range("A1").Value = "Fallversuch: Weg und Zeit"
    ws.Range("A1").Font.Bold = True
    put(ws, "A2", [["t in s", "s in m"]] + [[round(t * 0.1, 1), round(4.905 * (t * 0.1) ** 2 + rnd.uniform(-0.02, 0.02), 3)] for t in range(0, 11)])
    ws.Range("A3:B3").Font.Bold = True
    box(ws.Range("A3:B14"))
    ws.Range("B4:B14").NumberFormatLocal = "0,000"
    ch = ws.Shapes.AddChart2(240, -4169, 150, 20, 380, 240).Chart
    ch.SetSourceData(ws.Range("A3:B14"))
    ch.HasTitle = True
    ch.ChartTitle.Text = "s(t)"
    tl = ch.SeriesCollection(1).Trendlines().Add(3, Order=2)
    tl.DisplayEquation = True
    tl.DisplayRSquared = True


def funktion(wb) -> None:
    ws = wb.Worksheets(1)
    ws.Name = "Funktion"
    put(ws, "A1", [["x", "f(x)=x²-2x-3", "g(x)=2x+1"]])
    for i in range(21):
        r = i + 3
        ws.Cells(r, 1).Value = -5 + i * 0.5
        ws.Cells(r, 2).Formula2Local = f"=A{r}^2-2*A{r}-3"
        ws.Cells(r, 3).Formula2Local = f"=2*A{r}+1"
    ws.Range("A2:C2").Font.Bold = True
    ws.Range("A2:C2").Interior.Color = rgb(242, 242, 242)
    box(ws.Range("A2:C23"))
    ch = ws.Shapes.AddChart2(-1, 75, 230, 20, 380, 260).Chart  # Punkte mit Linien ohne Datenpunkte
    ch.SetSourceData(ws.Range("A2:C23"))


def vokabeln(wb) -> None:
    ws = wb.Worksheets(1)
    ws.Name = "Vokabeln"
    words = [("the attempt", "der Versuch", "We made an attempt to climb the hill."), ("to achieve", "erreichen", "She achieved her goal after a long and difficult year."),
             ("reliable", "zuverlässig", "A reliable friend is hard to find."), ("the evidence", "der Beweis", "There is no evidence that he was there."),
             ("to improve", "verbessern", "Practice will improve your skills."), ("the purpose", "der Zweck", "What is the purpose of this meeting?"),
             ("although", "obwohl", "Although it rained, we went outside."), ("to suggest", "vorschlagen", "I suggest that we leave early tomorrow morning.")]
    put(ws, "A1", [["Englisch", "Deutsch", "Beispielsatz"]] + [list(w) for w in words])
    lo = ws.ListObjects.Add(1, ws.Range("A2:C10"), None, 1)
    lo.TableStyle = "TableStyleMedium9"
    ws.Columns("A:B").ColumnWidth = 18
    ws.Columns("C").ColumnWidth = 40
    ws.Range("C3:C10").WrapText = True


def anwesenheit(wb) -> None:
    ws = wb.Worksheets(1)
    ws.Name = "Anwesenheit"
    ws.Range("A2").Value = "Name"
    for d in range(10):
        c = ws.Cells(2, d + 2)
        c.Value = f"{d + 3}.11.2026"
        c.NumberFormatLocal = "TT.MM."
        c.Orientation = 90
    for i, n in enumerate(NAMES[:12]):
        ws.Cells(i + 3, 1).Value = n
        for d in range(10):
            ws.Cells(i + 3, d + 2).Value = "✓" if rnd.random() > 0.12 else "–"
            ws.Cells(i + 3, d + 2).HorizontalAlignment = XL_CENTER
    ws.Range("L2").Value = "Fehltage"
    for i in range(12):
        ws.Cells(i + 3, 12).Formula2Local = f'=ZÄHLENWENN(B{i + 3}:K{i + 3};"–")'
    ws.Range("L3:L14").FormatConditions.AddColorScale(3)
    box(ws.Range("A2:L14"))
    ws.Columns("B:K").ColumnWidth = 4
    ws.Range("A2:L2").Font.Bold = True


def budget(wb) -> None:
    ws = wb.Worksheets(1)
    ws.Name = "AG-Budget"
    put(ws, "A1", [["Posten", "Betrag", "Anteil"]] + [[p, b, f"=B{i + 3}/SUMME($B$3:$B$7)"] for i, (p, b) in enumerate([("Material", 320), ("Ausflug", 450), ("Verpflegung", 180), ("Preise", 90), ("Sonstiges", 60)])])
    ws.Range("B3:B7").NumberFormatLocal = "#.##0 €"
    ws.Range("C3:C7").NumberFormatLocal = "0,0%"
    header(ws.Range("A2:C2"), fill=rgb(197, 90, 17))
    ch = ws.Shapes.AddChart2(251, 5, 220, 10, 300, 220).Chart
    ch.SetSourceData(ws.Range("A2:B7"))
    ch.SeriesCollection(1).HasDataLabels = True


def langeliste(wb) -> None:
    ws = wb.Worksheets(1)
    ws.Name = "Ausleihe"
    put(ws, "A1", [["Nr.", "Titel", "Schüler", "Ausgeliehen", "Rückgabe"]])
    for i in range(120):
        r = i + 3
        ws.Cells(r, 1).Value = i + 1
        ws.Cells(r, 2).Value = f"Lehrbuch {rnd.choice(SUBJECTS)} Band {rnd.randint(1, 4)}"
        ws.Cells(r, 3).Value = rnd.choice(NAMES)
        ws.Cells(r, 4).Value = f"{rnd.randint(1, 28)}.{rnd.randint(8, 9)}.2026"
        ws.Cells(r, 5).Value = f"{rnd.randint(1, 28)}.{rnd.randint(6, 7)}.2027"
    ws.Range("D3:E122").NumberFormatLocal = "TT.MM.JJJJ"
    header(ws.Range("A2:E2"))
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C:E").ColumnWidth = 12
    ws.PageSetup.PrintTitleRows = "$2:$2"
    ws.PageSetup.PrintGridlines = True
    ws.PageSetup.CenterFooter = "Seite &S von &A"
    ws.PageSetup.LeftHeader = "Schulbibliothek"


def schriften(wb) -> None:
    ws = wb.Worksheets(1)
    ws.Name = "Schriften"
    fonts = ["Calibri", "Arial", "Times New Roman", "Cambria", "Verdana", "Georgia", "Segoe UI", "Consolas"]
    for i, f in enumerate(fonts):
        for k, size in enumerate([8, 10, 11, 14, 18]):
            c = ws.Cells(i + 1, k + 1)
            c.Value = f"{f} {size}"
            c.Font.Name = f
            c.Font.Size = size
    ws.Range("A10").Value = "fett"
    ws.Range("A10").Font.Bold = True
    ws.Range("B10").Value = "kursiv"
    ws.Range("B10").Font.Italic = True
    ws.Range("C10").Value = "unterstrichen"
    ws.Range("C10").Font.Underline = 2
    ws.Range("D10").Value = "durchgestrichen"
    ws.Range("D10").Font.Strikethrough = True
    ws.Range("E10").Value = "rot"
    ws.Range("E10").Font.Color = rgb(255, 0, 0)
    ws.Range("A11").Value = "H"
    ws.Range("B11").Value = "x2"
    ws.Range("B11").GetCharacters(2, 1).Font.Superscript = True
    ws.Columns("A:E").AutoFit()


def zahlenformate(wb) -> None:
    ws = wb.Worksheets(1)
    ws.Name = "Zahlenformate"
    fmts = ["Standard", "0", "0,00", "#.##0", "#.##0,00", "0%", "0,00%", "0,00E+00", "# ?/?", "# ??/??", "#.##0,00 €", "#.##0 €;[Rot]-#.##0 €",
            "TT.MM.JJJJ", "T. MMMM JJJJ", "TTTT, T. MMMM JJJJ", "MMM JJ", "hh:mm", "hh:mm:ss", "[hh]:mm", "@", '0,0" kg"', "000"]
    for i, f in enumerate(fmts):
        ws.Cells(i + 1, 1).Value = f
        for k, v in enumerate([1234.5678, -0.4567, 46000.75, 7]):
            c = ws.Cells(i + 1, k + 2)
            c.Value = v
            try:
                c.NumberFormatLocal = f
            except Exception:  # noqa: BLE001
                pass
    ws.Columns("A:E").ColumnWidth = 22


def ausrichtung(wb) -> None:
    ws = wb.Worksheets(1)
    ws.Name = "Ausrichtung"
    labels = [("links", XL_LEFT), ("zentriert", XL_CENTER), ("rechts", XL_RIGHT)]
    for i, (t, a) in enumerate(labels):
        c = ws.Cells(1, i + 1)
        c.Value = t
        c.HorizontalAlignment = a
    for i, v in enumerate([-4160, -4108, -4107]):
        c = ws.Cells(2, i + 1)
        c.Value = ["oben", "Mitte", "unten"][i]
        c.VerticalAlignment = v
    ws.Rows(2).RowHeight = 40
    for i, deg in enumerate([45, 90, -45, -90]):
        c = ws.Cells(4, i + 1)
        c.Value = f"Drehung {deg}°"
        c.Orientation = deg
    ws.Rows(4).RowHeight = 80
    ws.Range("A6").Value = "Ein langer Text, der in der Zelle umbrochen wird, weil Zeilenumbruch eingeschaltet ist."
    ws.Range("A6").WrapText = True
    ws.Range("B6").Value = "Eingerückt"
    ws.Range("B6").IndentLevel = 2
    ws.Range("C6").Value = "An Zellgröße anpassen, ziemlich lang"
    ws.Range("C6").ShrinkToFit = True
    ws.Columns("A:D").ColumnWidth = 16
    box(ws.Range("A1:D6"))


def rahmen(wb) -> None:
    ws = wb.Worksheets(1)
    ws.Name = "Rahmen"
    styles = [(1, 2), (1, -4138), (1, 4), (-4115, 2), (4, 2), (-4119, 4), (5, 2), (13, 2), (-4142, 2)]
    for i, (ls, w) in enumerate(styles):
        c = ws.Cells(2 + i * 2, 2)
        c.Value = f"Linie {i + 1}"
        for e in (7, 8, 9, 10):
            try:
                c.Borders(e).LineStyle = ls
                c.Borders(e).Weight = w
            except Exception:  # noqa: BLE001
                pass
    c = ws.Range("D2:F6")
    box(c, XL_THIN)
    c.BorderAround(1, XL_THICK, Color=rgb(192, 0, 0))
    ws.Range("D8:F8").Borders(9).LineStyle = -4119


def tabelle_ergebnis(wb) -> None:
    ws = wb.Worksheets(1)
    ws.Name = "Sammelbestellung"
    put(ws, "A1", [["Artikel", "Menge", "Preis", "Summe"]] + [[a, m, p, f"=B{i + 3}*C{i + 3}"] for i, (a, m, p) in enumerate([("Heft A4", 30, 0.89), ("Füller", 12, 4.5), ("Lineal", 25, 1.2), ("Zirkel", 15, 6.99), ("Taschenrechner", 8, 17.95)])])
    lo = ws.ListObjects.Add(1, ws.Range("A2:D7"), None, 1)
    lo.TableStyle = "TableStyleMedium2"
    lo.ShowTotals = True
    lo.ListColumns(4).TotalsCalculation = 1
    ws.Range("C3:D8").NumberFormatLocal = "#.##0,00 €"
    ws.Columns("A").ColumnWidth = 18


def sparklines(wb) -> None:
    ws = wb.Worksheets(1)
    ws.Name = "Leistungen"
    put(ws, "A1", [["Name", "Sep", "Okt", "Nov", "Dez", "Jan", "Verlauf"]] + [[n] + [rnd.randint(4, 15) for _ in range(5)] + [""] for n in NAMES[:8]])
    ws.Range("G3:G10").SparklineGroups.Add(1, "B3:F10")
    header(ws.Range("A2:G2"))
    ws.Columns("G").ColumnWidth = 18


def formen(wb) -> None:
    ws = wb.Worksheets(1)
    ws.Name = "Ablauf"
    steps = ["Thema wählen", "Quellen suchen", "Gliederung", "Schreiben", "Abgabe"]
    for i, s in enumerate(steps):
        sh = ws.Shapes.AddShape(5, 20 + i * 110, 40, 90, 50)
        sh.TextFrame2.TextRange.Text = s
        if i < len(steps) - 1:
            ws.Shapes.AddConnector(1, 110 + i * 110, 65, 130 + i * 110, 65).Line.EndArrowheadStyle = 2
    tb = ws.Shapes.AddTextbox(1, 20, 120, 300, 40)
    tb.TextFrame2.TextRange.Text = "Hinweis: Die Facharbeit wird bis zum 15. März abgegeben."
    ws.Shapes.AddShape(9, 360, 120, 60, 60).Fill.ForeColor.RGB = rgb(255, 192, 0)


def bild(wb) -> None:
    png = OUT / "_logo.png"
    img = Image.new("RGB", (320, 200), (255, 255, 255))
    d = ImageDraw.Draw(img)
    d.ellipse((20, 20, 180, 180), fill=(237, 28, 36))
    d.rectangle((200, 60, 300, 140), fill=(31, 78, 121))
    img.save(png)
    ws = wb.Worksheets(1)
    ws.Name = "Bild"
    ws.Range("A1").Value = "Schullogo (Testbild)"
    ws.Shapes.AddPicture(str(png), False, True, 20, 30, 160, 100)
    ws.Range("A10").Value = "Unter dem Bild"


def fehler(wb) -> None:
    ws = wb.Worksheets(1)
    ws.Name = "Formeln"
    put(ws, "A1", [["Eingabe", "Formel", "Ergebnis"],
                   [5, "Kehrwert", "=1/A2"], [0, "Kehrwert", "=1/A3"], ["Text", "Doppelt", "=A4*2"], [7, "Note", '=WENN(A5>4;"nicht bestanden";"bestanden")'],
                   ["Ben", "Suche", "=SVERWEIS(A6;$E$2:$F$4;2;FALSCH)"], ["Xaver", "Suche", "=SVERWEIS(A7;$E$2:$F$4;2;FALSCH)"]])
    put(ws, "E1", [["Anna", 1], ["Ben", 2], ["Cem", 3]])
    ws.Columns("A:C").ColumnWidth = 16


def design_ion(wb) -> None:
    from pathlib import Path as P
    t = P(r"C:\Program Files\Microsoft Office\root\Document Themes 16\Ion.thmx")
    if t.exists():
        wb.ApplyTheme(str(t))
    ws = wb.Worksheets(1)
    ws.Name = "Design"
    for i in range(6):
        c = ws.Cells(2, i + 1)
        c.Value = f"Akzent {i + 1}"
        c.Interior.ThemeColor = 5 + i
        c.Font.ThemeColor = 1
        d = ws.Cells(3, i + 1)
        d.Value = "heller 60 %"
        d.Interior.ThemeColor = 5 + i
        d.Interior.TintAndShade = 0.6
    ws.Range("A1").Value = "Überschrift im Design"
    ws.Range("A1").Font.ThemeFont = 1  # major
    ws.Range("A1").Font.Size = 18


def mehrere_blaetter(wb) -> None:
    while wb.Worksheets.Count < 3:
        wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    for i, s in enumerate(["Halbjahr 1", "Halbjahr 2", "Gesamt"]):
        ws = wb.Worksheets(i + 1)
        ws.Name = s
        ws.Range("A1").Value = s
        ws.Range("A1").Font.Bold = True
        for r in range(8):
            ws.Cells(r + 2, 1).Value = NAMES[r]
            ws.Cells(r + 2, 2).Value = rnd.randint(1, 6) if i < 2 else None
            if i == 2:
                ws.Cells(r + 2, 2).Formula2Local = f"=MITTELWERT('Halbjahr 1'!B{r + 2};'Halbjahr 2'!B{r + 2})"
        ws.Tab.Color = [rgb(91, 155, 213), rgb(237, 125, 49), rgb(112, 173, 71)][i]


def skaliert(wb) -> None:
    ws = wb.Worksheets(1)
    ws.Name = "Jahresplan"
    months = ["Aug", "Sep", "Okt", "Nov", "Dez", "Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul"]
    put(ws, "A1", [["Fach"] + months] + [[s] + [rnd.choice(["KA", "Test", "", "", "Projekt"]) for _ in months] for s in SUBJECTS])
    header(ws.Range("A2:M2"))
    box(ws.Range("A2:M14"))
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = 1
    ws.PageSetup.Orientation = XL_LANDSCAPE


def balken(wb) -> None:
    ws = wb.Worksheets(1)
    ws.Name = "Umfrage"
    put(ws, "A1", [["Lieblingsfach", "Stimmen"]] + [[s, rnd.randint(2, 14)] for s in SUBJECTS[:8]])
    ch = ws.Shapes.AddChart2(216, 57, 160, 10, 360, 260).Chart
    ch.SetSourceData(ws.Range("A2:B10"))
    ch.SeriesCollection(1).HasDataLabels = True
    ch.HasTitle = True
    ch.ChartTitle.Text = "Umfrage 8b"


def gestapelt(wb) -> None:
    ws = wb.Worksheets(1)
    ws.Name = "Mensa"
    put(ws, "A1", [["Tag", "Nudeln", "Pizza", "Salat"]] + [[d, rnd.randint(20, 60), rnd.randint(20, 60), rnd.randint(5, 25)] for d in ["Mo", "Di", "Mi", "Do", "Fr"]])
    ch = ws.Shapes.AddChart2(297, 52, 230, 10, 380, 250).Chart
    ch.SetSourceData(ws.Range("A2:D7"))
    ch.SetElement(101)  # legend at the bottom
    ch.Axes(2).HasTitle = True
    ch.Axes(2).AxisTitle.Text = "Essen"


def ausgeblendet(wb) -> None:
    ws = wb.Worksheets(1)
    ws.Name = "Gliederung"
    put(ws, "A1", [["Bereich", "Wert"]] + [[f"Posten {i + 1}", rnd.randint(10, 99)] for i in range(12)])
    ws.Rows("5:7").Hidden = True
    ws.Columns("C").Hidden = True
    ws.Rows("9:12").Group()
    ws.Range("A16").Value = "Summe"
    ws.Range("B16").Formula2Local = "=TEILERGEBNIS(109;B3:B14)"


def raender(wb) -> None:
    ws = wb.Worksheets(1)
    ws.Name = "Raender"
    for r in range(60):
        ws.Cells(r + 1, 1).Value = f"Zeile {r + 1}"
        ws.Cells(r + 1, 2).Value = r * r
    ps = ws.PageSetup
    ps.LeftMargin = ps.RightMargin = 30
    ps.TopMargin = ps.BottomMargin = 40
    ws.HPageBreaks.Add(ws.Range("A31"))
    ps.CenterHorizontally = True
    ps.CenterHeader = "&D"


BUILDERS = {f.__name__: f for f in [notenliste, notenspiegel, stundenplan, klassenkasse, messreihe, funktion, vokabeln, anwesenheit, budget, langeliste,
                                    schriften, zahlenformate, ausrichtung, rahmen, tabelle_ergebnis, sparklines, formen, bild, fehler, design_ion,
                                    mehrere_blaetter, skaliert, balken, gestapelt, ausgeblendet, raender]}


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    only = sys.argv[1:] or list(BUILDERS)
    xl = ExcelTarget(visible=False)
    done, failed = [], {}
    try:
        for name in only:
            wb = xl.app.Workbooks.Add()
            try:
                BUILDERS[name](wb)
                path = OUT / f"{name}.xlsx"
                if path.exists():
                    path.unlink()
                wb.Worksheets(1).Activate()
                wb.SaveAs(str(path), 51)
                done.append(name)
            except Exception as e:  # noqa: BLE001 - recorded, the rest goes on
                failed[name] = str(e)[:200]
            finally:
                wb.Close(False)
    finally:
        xl.close()
    print(len(done), "built;", "failed:", failed)


if __name__ == "__main__":
    main()
