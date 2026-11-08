Attribute VB_Name = "Tests"
' RaabePresent's macro tests (D-080 B, D-105): our own versions of macros teachers often use in PowerPoint.
' Each test writes what it found into the shape "Ausgabe" on the last slide (no MsgBox: PowerPoint runs them
' through COM, where a dialog would wait for a click). check_present_macro_suite.py runs each in PowerPoint and in
' RaabePresent and compares the presentations afterwards. %NOTES% is replaced by the harness with a file path.
Option Explicit

Private Function Ausgabe() As Object
    Set Ausgabe = ActivePresentation.Slides(ActivePresentation.Slides.Count).Shapes("Ausgabe")
End Function

Private Sub Melde(ByVal text As String)
    Ausgabe.TextFrame.TextRange.Text = text
End Sub

' The names VBA sees for shapes PowerPoint named itself, and whether the name in the file finds the shape.
Public Sub FormNamen()
    Dim r As String
    r = ActivePresentation.Slides(1).Shapes(1).Name
    On Error Resume Next
    r = r & "|" & ActivePresentation.Slides(1).Shapes("Titel 1").Name
    If Err.Number <> 0 Then r = r & "|Fehler " & Err.Number
    Err.Clear
    r = r & "|" & ActivePresentation.Slides(1).Shapes("Title 1").Name
    If Err.Number <> 0 Then r = r & "|Fehler " & Err.Number
    On Error GoTo 0
    Melde r
End Sub

' Replaces one font by another in every text on every slide.
Public Sub SchriftErsetzen()
    Dim sld As Object, shp As Object, n As Long
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    If shp.TextFrame.TextRange.Font.Name = "Comic Sans MS" Then
                        shp.TextFrame.TextRange.Font.Name = "Verdana"
                        n = n + 1
                    End If
                End If
            End If
        Next shp
    Next sld
    Melde "Ersetzt: " & n
End Sub

' Gives every picture the same width, its height following its proportions.
Public Sub BilderGleichGross()
    Dim sld As Object, shp As Object, n As Long
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.Type = 13 Then ' msoPicture
                shp.LockAspectRatio = -1
                shp.Width = 200
                shp.Left = 40 + n * 220
                n = n + 1
            End If
        Next shp
    Next sld
    Melde "Bilder: " & n
End Sub

' Writes the notes of every slide into a text file.
Public Sub NotizenExportieren()
    Dim f As Integer, i As Long, t As String
    f = FreeFile
    Open "%NOTES%" For Output As #f
    For i = 1 To ActivePresentation.Slides.Count
        t = ActivePresentation.Slides(i).NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
        If Len(t) > 0 Then Print #f, "Folie " & i & ": " & Replace(t, vbCr, " / ")
    Next i
    Close #f
    Melde "Notizen exportiert"
End Sub

' Deletes placeholders that hold no text (from the last shape back, so the loop is not upset).
Public Sub LeerePlatzhalterLoeschen()
    Dim sld As Object, k As Long, n As Long
    For Each sld In ActivePresentation.Slides
        For k = sld.Shapes.Count To 1 Step -1
            With sld.Shapes(k)
                If .Type = 14 Then ' msoPlaceholder
                    If .HasTextFrame Then
                        If Not .TextFrame.HasText Then .Delete: n = n + 1
                    End If
                End If
            End With
        Next k
    Next sld
    Melde "Gelöscht: " & n
End Sub

' A countdown in the shape "Ausgabe" (short steps for the test).
Public Sub Countdown()
    Dim n As Integer, t As Single
    For n = 3 To 0 Step -1
        Melde CStr(n)
        t = Timer
        Do While Timer < t + 0.05
            DoEvents
        Loop
    Next n
    Melde "Countdown: " & Ausgabe.TextFrame.TextRange.Text
End Sub

' Quiz: buttons run Richtig or Falsch; the score is in the shape "Punkte".
Public Sub Richtig()
    With ActivePresentation.Slides(ActivePresentation.Slides.Count).Shapes("Punkte").TextFrame.TextRange
        .Text = CStr(Val(.Text) + 1)
    End With
End Sub

Public Sub Falsch()
    Dim s As Object
    Set s = ActivePresentation.Slides(ActivePresentation.Slides.Count).Shapes("Punkte")
    s.Fill.ForeColor.RGB = RGB(230, 60, 60)
End Sub

Public Sub Quiz()
    Richtig
    Richtig
    Falsch
    Richtig
    Melde "Punkte: " & ActivePresentation.Slides(ActivePresentation.Slides.Count).Shapes("Punkte").TextFrame.TextRange.Text
End Sub

' Random names from the list in "Namen", the same each time (a fixed seed).
Public Sub NamenZiehen()
    Dim namen() As String, i As Integer, r As String
    namen = Split(ActivePresentation.Slides(ActivePresentation.Slides.Count).Shapes("Namen").TextFrame.TextRange.Text, vbCr)
    Rnd -1
    Randomize 7
    For i = 1 To 5
        r = r & namen(Int(Rnd * (UBound(namen) + 1))) & " "
    Next i
    Melde Trim(r)
End Sub

' Runtime errors: their numbers as VBA and PowerPoint give them.
Public Sub Fehler()
    Dim r As String, x As Variant, a(2) As Integer
    On Error Resume Next
    x = 1 / 0: r = r & Err.Number & " ": Err.Clear
    x = a(5): r = r & Err.Number & " ": Err.Clear
    x = CInt("abc"): r = r & Err.Number & " ": Err.Clear
    x = ActivePresentation.Slides(99).Name: r = r & IIf(Err.Number <> 0, "S", "-") & " ": Err.Clear
    x = ActivePresentation.Slides(1).Shapes("Gibt es nicht").Name: r = r & IIf(Err.Number <> 0, "F", "-"): Err.Clear
    On Error GoTo 0
    Melde Trim(r)
End Sub

' Strings, numbers and dates (the built-in functions a quiz or a timer uses).
Public Sub Funktionen()
    Dim r As String
    r = Format(1234.5, "0.00") & "|" & Format(0.25, "0%") & "|" & UCase("äbc") & "|" & Left("Raabe", 3) & "|" & InStr("Hallo Welt", "Welt")
    r = r & "|" & Round(2.5) & "|" & Round(3.5) & "|" & Int(-2.5) & "|" & Fix(-2.5) & "|" & (7 Mod 3) & "|" & (7 \ 2)
    r = r & "|" & Year(DateSerial(2026, 11, 8)) & "|" & Weekday(DateSerial(2026, 11, 8)) & "|" & Format(DateSerial(2026, 11, 8), "dd.mm.yyyy")
    Melde r
End Sub
