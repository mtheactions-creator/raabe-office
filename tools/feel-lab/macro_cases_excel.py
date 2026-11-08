"""Macro tests for RaabeCalc (D-080 B): our own versions of popular Excel macros and of macros teachers use (the forum
versions are not ours to copy), each with the cells it starts from. check_calc_macros.py runs every macro in Excel and in
RaabeCalc on the same workbook and compares what they leave: values, formulas, the formats named in `check`, sheet names,
hidden rows and columns, and the messages shown.

A case: name, setup ({"Tabelle1!A1": value or "=formula (German)"}), code (one standard module), macro (the Sub to run),
answers (what MsgBox/InputBox are answered, in order), check (extra things compared: "fill", "font", "hidden", "sheets").
"""
from __future__ import annotations

NAMES = ["Anna", "Ben", "Cem", "Dora", "Emil", "Fiona", "Gül", "Hannes", "Ida", "Jonas", "Kira", "Leon"]


def grades() -> dict:
    out = {"Tabelle1!A1": "Name", "Tabelle1!B1": "KA 1", "Tabelle1!C1": "KA 2", "Tabelle1!D1": "Mündlich"}
    vals = [(2, 3, 2), (4, 5, 3), (1, 1, 2), (3, 2, 3), (5, 6, 4), (2, 2, 1), (3, 4, 3), (6, 5, 5), (1, 2, 1), (4, 3, 4), (2, 1, 2), (3, 3, 3)]
    for i, (n, v) in enumerate(zip(NAMES, vals)):
        out[f"Tabelle1!A{i + 2}"] = n
        for k, x in enumerate(v):
            out[f"Tabelle1!{'BCD'[k]}{i + 2}"] = x
    return out


CASES: list[dict] = [
    {"name": "leere_zeilen_loeschen", "setup": {"Tabelle1!A1": "A", "Tabelle1!A3": "B", "Tabelle1!B3": 2, "Tabelle1!A6": "C", "Tabelle1!A7": "D", "Tabelle1!B9": 9},
     "macro": "LeereZeilenLoeschen", "code": """
Sub LeereZeilenLoeschen()
    Dim i As Long, letzte As Long
    letzte = Cells(Rows.Count, 1).End(xlUp).Row
    If Cells(Rows.Count, 2).End(xlUp).Row > letzte Then letzte = Cells(Rows.Count, 2).End(xlUp).Row
    For i = letzte To 1 Step -1
        If WorksheetFunction.CountA(Rows(i)) = 0 Then Rows(i).Delete
    Next i
End Sub"""},
    {"name": "duplikate_markieren", "setup": {**{f"Tabelle1!A{i + 1}": v for i, v in enumerate(["Mathe", "Deutsch", "Mathe", "Bio", "Kunst", "Bio", "Mathe", "Sport"])}},
     "macro": "DuplikateMarkieren", "check": ["fill"], "code": """
Sub DuplikateMarkieren()
    Dim c As Range
    For Each c In Range("A1:A8")
        If WorksheetFunction.CountIf(Range("A1:A8"), c.Value) > 1 Then
            c.Interior.Color = RGB(255, 199, 206)
            c.Font.Color = RGB(156, 0, 6)
        End If
    Next c
End Sub"""},
    {"name": "notenrechner", "setup": grades(), "macro": "Notenrechner", "answers": [1], "code": """
Sub Notenrechner()
    Dim i As Long, letzte As Long, summe As Double
    letzte = Range("A1").End(xlDown).Row
    Range("E1").Value = "Schnitt"
    Range("F1").Value = "Note"
    For i = 2 To letzte
        Cells(i, 5).Value = Round((Cells(i, 2).Value + Cells(i, 3).Value) / 2 * 0.6 + Cells(i, 4).Value * 0.4, 2)
        Cells(i, 6).Value = WorksheetFunction.Round(Cells(i, 5).Value, 0)
        summe = summe + Cells(i, 5).Value
    Next i
    Range("E" & letzte + 2).Value = Round(summe / (letzte - 1), 2)
    MsgBox "Klassenschnitt: " & Format(summe / (letzte - 1), "0.00"), vbInformation, "Notenrechner"
End Sub"""},
    {"name": "notenspiegel", "setup": grades(), "macro": "Notenspiegel", "code": """
Sub Notenspiegel()
    Dim n As Integer
    Range("H1").Value = "Note"
    Range("I1").Value = "Anzahl"
    For n = 1 To 6
        Cells(n + 1, 8).Value = n
        Cells(n + 1, 9).Formula = "=COUNTIF($C$2:$C$13," & n & ")"
    Next n
    Range("H8").Value = "Summe"
    Range("I8").FormulaLocal = "=SUMME(I2:I7)"
End Sub"""},
    {"name": "inhaltsverzeichnis", "setup": {"Tabelle1!A1": "Start"}, "macro": "Inhalt", "check": ["sheets"], "code": """
Sub Inhalt()
    Dim ws As Worksheet, i As Integer
    Worksheets.Add Before:=Worksheets(1)
    ActiveSheet.Name = "Inhalt"
    Worksheets.Add After:=Worksheets(Worksheets.Count)
    ActiveSheet.Name = "Daten"
    i = 1
    For Each ws In Worksheets
        Worksheets("Inhalt").Cells(i, 1).Value = ws.Index
        Worksheets("Inhalt").Cells(i, 2).Value = ws.Name
        i = i + 1
    Next ws
End Sub"""},
    {"name": "blaetter_einblenden", "setup": {"Tabelle1!A1": 1}, "macro": "Einblenden", "check": ["sheets"], "code": """
Sub Einblenden()
    Dim ws As Worksheet
    Worksheets.Add After:=Worksheets(1)
    ActiveSheet.Name = "Versteckt"
    Worksheets("Versteckt").Visible = xlSheetHidden
    Worksheets(1).Range("B1").Value = Worksheets("Versteckt").Visible
    For Each ws In Worksheets
        ws.Visible = xlSheetVisible
    Next ws
    Worksheets(1).Range("B2").Value = Worksheets("Versteckt").Visible
End Sub"""},
    {"name": "zeilen_faerben", "setup": grades(), "macro": "Faerben", "check": ["fill", "font"], "code": """
Sub Faerben()
    Dim i As Long
    For i = 2 To 13
        If Cells(i, 3).Value >= 5 Then
            Range(Cells(i, 1), Cells(i, 4)).Interior.Color = RGB(255, 235, 156)
            Cells(i, 1).Font.Bold = True
        ElseIf Cells(i, 3).Value <= 1 Then
            Range(Cells(i, 1), Cells(i, 4)).Interior.ColorIndex = 35
        End If
    Next i
End Sub"""},
    {"name": "summenzeile", "setup": {**{f"Tabelle1!A{i + 1}": n for i, n in enumerate(["Posten", "Heft", "Füller", "Lineal", "Zirkel"])}, **{f"Tabelle1!B{i + 2}": v for i, v in enumerate([2.5, 12.9, 1.2, 6.99])}},
     "macro": "Summe", "code": """
Sub Summe()
    Dim letzte As Long
    letzte = Cells(Rows.Count, 2).End(xlUp).Row
    Cells(letzte + 1, 1).Value = "Summe"
    Cells(letzte + 1, 2).Formula = "=SUM(B2:B" & letzte & ")"
    Cells(letzte + 1, 1).Font.Bold = True
    Range("B2:B" & letzte + 1).NumberFormat = "#,##0.00 ""€"""
End Sub"""},
    {"name": "einmaleins", "setup": {}, "macro": "Einmaleins", "code": """
Sub Einmaleins()
    Dim t(1 To 10, 1 To 10) As Variant, i As Integer, k As Integer
    For i = 1 To 10
        For k = 1 To 10
            t(i, k) = i * k
        Next k
    Next i
    Range("B2:K11").Value = t
    For i = 1 To 10
        Cells(1, i + 1).Value = i
        Cells(i + 1, 1).Value = i
    Next i
    Range("A1:K1").Font.Bold = True
    Range("A1:A11").Font.Bold = True
End Sub"""},
    {"name": "text_bereinigen", "setup": {f"Tabelle1!A{i + 1}": v for i, v in enumerate(["  anna MÜLLER ", "ben schmidt", " CEM yilmaz", "dora  meier "])},
     "macro": "Bereinigen", "code": """
Sub Bereinigen()
    Dim c As Range
    For Each c In Range("A1:A4")
        c.Offset(0, 1).Value = WorksheetFunction.Proper(Trim(c.Value))
        c.Offset(0, 2).Value = UCase(Left(Trim(c.Value), 1))
        c.Offset(0, 3).Value = Len(Trim(c.Value))
    Next c
End Sub"""},
    {"name": "select_case", "setup": grades(), "macro": "Bewertung", "code": """
Sub Bewertung()
    Dim i As Long, t As String
    For i = 2 To 13
        Select Case Cells(i, 2).Value
            Case 1: t = "sehr gut"
            Case 2: t = "gut"
            Case 3: t = "befriedigend"
            Case 4: t = "ausreichend"
            Case 5, 6: t = "nicht ausreichend"
            Case Else: t = "?"
        End Select
        Cells(i, 7).Value = t
    Next i
End Sub"""},
    {"name": "fehlerbehandlung", "setup": {"Tabelle1!A1": 10, "Tabelle1!A2": 0, "Tabelle1!A3": "x", "Tabelle1!A4": 4}, "macro": "Teilen", "code": """
Sub Teilen()
    Dim i As Integer
    On Error Resume Next
    For i = 1 To 4
        Err.Clear
        Cells(i, 2).Value = 100 / Cells(i, 1).Value
        If Err.Number <> 0 Then Cells(i, 2).Value = "Fehler " & Err.Number
    Next i
    On Error GoTo 0
End Sub"""},
    {"name": "do_while", "setup": {f"Tabelle1!A{i + 1}": n for i, n in enumerate(NAMES[:7])}, "macro": "Zaehlen", "code": """
Sub Zaehlen()
    Dim r As Long
    r = 1
    Do While Cells(r, 1).Value <> ""
        Cells(r, 2).Value = r
        r = r + 1
    Loop
    Range("D1").Value = r - 1
End Sub"""},
    {"name": "zeilen_ausblenden", "setup": grades(), "macro": "Ausblenden", "check": ["hidden"], "code": """
Sub Ausblenden()
    Dim i As Long
    For i = 2 To 13
        Rows(i).Hidden = (Cells(i, 3).Value > 3)
    Next i
    Columns("D").Hidden = True
End Sub"""},
    {"name": "farbige_zaehlen", "setup": {"Tabelle1!A1": 1, "Tabelle1!A2": 2, "Tabelle1!A3": 3, "Tabelle1!A4": 4}, "macro": "FarbigeZaehlen", "check": ["fill"], "code": """
Sub FarbigeZaehlen()
    Dim c As Range, n As Integer
    Range("A2").Interior.Color = RGB(255, 255, 0)
    Range("A4").Interior.Color = RGB(255, 255, 0)
    For Each c In Range("A1:A4")
        If c.Interior.Color = RGB(255, 255, 0) Then n = n + 1
    Next c
    Range("C1").Value = n
End Sub"""},
    {"name": "kopieren_neues_blatt", "setup": grades(), "macro": "Kopieren", "check": ["sheets"], "code": """
Sub Kopieren()
    Worksheets.Add After:=Worksheets(Worksheets.Count)
    ActiveSheet.Name = "Kopie"
    Worksheets(1).Range("A1:D13").Copy Destination:=Worksheets("Kopie").Range("B2")
    Worksheets("Kopie").Range("A1").Value = Worksheets(1).Range("A1").CurrentRegion.Rows.Count
End Sub"""},
    {"name": "sortieren", "setup": grades(), "macro": "Sortieren", "code": """
Sub Sortieren()
    Range("A1:D13").Sort Key1:=Range("C2"), Order1:=xlDescending, Key2:=Range("A2"), Order2:=xlAscending, Header:=xlYes
End Sub"""},
    {"name": "formatieren", "setup": grades(), "macro": "Formatieren", "check": ["fill", "font"], "code": """
Sub Formatieren()
    With Range("A1:D1")
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    Range("A1:D13").Borders.LineStyle = xlContinuous
    Columns("A:D").AutoFit
    Range("B2:D13").NumberFormat = "0.0"
End Sub"""},
    {"name": "adressen_offset", "setup": {"Tabelle1!B2": 5}, "macro": "Adressen", "code": """
Sub Adressen()
    Dim r As Range
    Set r = Range("B2")
    Range("D1").Value = r.Address
    Range("D2").Value = r.Offset(2, 1).Address(False, False)
    Range("D3").Value = r.Resize(3, 2).Address
    Range("D4").Value = r.Row + r.Column
    Range("D5").Value = Range("A1:C4").Cells.Count
    Range("D6").Value = Range("A1:C4").Rows.Count & "x" & Range("A1:C4").Columns.Count
End Sub"""},
    {"name": "worksheetfunction", "setup": grades(), "macro": "Statistik", "code": """
Sub Statistik()
    Range("G1").Value = WorksheetFunction.Average(Range("B2:B13"))
    Range("G2").Value = WorksheetFunction.Max(Range("C2:C13"))
    Range("G3").Value = WorksheetFunction.Min(Range("D2:D13"))
    Range("G4").Value = WorksheetFunction.Sum(Range("B2:D13"))
    Range("G5").Value = WorksheetFunction.VLookup("Dora", Range("A2:D13"), 3, False)
    Range("G6").Value = WorksheetFunction.Match("Ida", Range("A2:A13"), 0)
    Range("G7").Value = WorksheetFunction.CountIf(Range("B2:B13"), "<3")
    Range("G8").Value = WorksheetFunction.Median(Range("B2:B13"))
End Sub"""},
    {"name": "inputbox_msgbox", "setup": {}, "macro": "Fragen", "answers": ["Raabeschule", 6], "code": """
Sub Fragen()
    Dim s As String, a As Integer
    s = InputBox("Wie heißt die Schule?", "Frage", "")
    Range("A1").Value = s
    a = MsgBox("Weiter?", vbYesNo + vbQuestion, "Frage")
    Range("A2").Value = a
End Sub"""},
    {"name": "datum_zeit", "setup": {}, "macro": "Datum", "code": """
Sub Datum()
    Range("A1").Value = DateSerial(2026, 11, 8)
    Range("A2").Value = Format(DateSerial(2026, 11, 8), "dd.mm.yyyy")
    Range("A3").Value = Year(DateSerial(2026, 11, 8)) + Month(DateSerial(2026, 11, 8))
    Range("A4").Value = Weekday(DateSerial(2026, 11, 8), vbMonday)
    Range("A5").Value = DateAdd("d", 30, DateSerial(2026, 11, 8))
    Range("A5").NumberFormat = "dd.mm.yyyy"
End Sub"""},
    {"name": "schutz", "setup": {"Tabelle1!A1": 1}, "macro": "Schutz", "code": """
Sub Schutz()
    Range("B1").Locked = False
    ActiveSheet.Protect "raabe"
    Range("C1").Value = ActiveSheet.ProtectContents
    ActiveSheet.Unprotect "raabe"
    Range("C2").Value = ActiveSheet.ProtectContents
End Sub"""},
    {"name": "spezialzellen", "setup": {"Tabelle1!A1": 1, "Tabelle1!A3": 3, "Tabelle1!A5": "=A1+A3"}, "macro": "Leere", "check": ["fill"], "code": """
Sub Leere()
    Range("A1:A5").SpecialCells(xlCellTypeBlanks).Interior.Color = RGB(200, 200, 200)
    Range("A1:A5").SpecialCells(xlCellTypeFormulas).Font.Bold = True
    Range("C1").Value = Range("A1:A5").SpecialCells(xlCellTypeConstants).Count
End Sub"""},
]
