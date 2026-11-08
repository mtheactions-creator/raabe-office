"""Cases for Makro aufzeichnen (D-119): the same actions done by hand in Excel (probe_excel_recorder.py, real input) and
in RaabeCalc (check_calc_recorder.py, trusted DevTools input). Steps as in feellab/calc_catalogue.py, plus
    ("keytips", "R H C")   Alt, then the keytip letters (a space between keytips)
    ("ribbon", "font.fill") the ribbon button itself (Excel: through UI Automation by its name; RaabeCalc: a click)
A case may record with Relative Verweise verwenden ("relative"), a shortcut letter and a description."""

CASES = [
    {"name": "Eingabe", "setup": {}, "start": "A1", "steps": [
        ("type", "Note"), ("keys", "enter"), ("type", "2"), ("keys", "enter"), ("type", "3,5"), ("keys", "enter"),
        ("type", "=SUMME(A2:A3)"), ("keys", "enter"), ("type", 'Text "zitiert"'), ("keys", "enter"),
        ("type", "01.02.2026"), ("keys", "enter"), ("type", "50%"), ("keys", "tab"), ("type", "=A2*2"), ("keys", "enter")]},
    {"name": "Format", "setup": {"A1": "Name", "B1": "Note", "C1": "Fach", "A2": "Anna", "B2": "0,25", "A3": "Ben", "B3": "0,5"}, "start": "A1",
     "shortcut": "q", "description": "Kopfzeile formatieren", "steps": [
        ("keys", "shift+right", 2), ("keys", "ctrl+shift+f"), ("keys", "ctrl+shift+k"), ("keytips", "R R Z"),
        ("keys", "down"), ("keys", "right"), ("keys", "shift+down"), ("keys", "ctrl+shift+5"),
        ("keys", "ctrl+home"), ("keytips", "R H C"), ("type", "14"), ("keys", "enter")]},
    {"name": "Fuellfarbe", "setup": {"A1": "x", "B2": "y"}, "start": "A1", "steps": [
        ("keys", "shift+down"), ("keys", "shift+right"), ("ribbon", "font.fill")]},
    {"name": "Ausfuellen", "setup": {"A1": "Mo", "B1": "1", "B2": "2"}, "start": "A1", "steps": [
        ("keys", "shift+down", 4), ("keys", "ctrl+u"), ("drag", "B1", "B2"), ("fill", "B1:B2", "B6")]},
    {"name": "Sortieren", "setup": {"A1": "Name", "B1": "Note", "A2": "Cem", "B2": "2", "A3": "Anna", "B3": "1", "A4": "Ben", "B4": "3"},
     "start": "A2", "steps": [("keytips", "V U"), ("keys", "right"), ("keytips", "V F1")]},
    {"name": "Zeilen", "setup": {"A1": "1", "A2": "2", "A3": "3", "A4": "4", "B1": "b"}, "start": "A2", "steps": [
        ("keys", "shift+space"), ("keys", "ctrl+plus"), ("keys", "down", 2), ("keys", "shift+space"), ("keys", "ctrl+minus"),
        ("keys", "right"), ("keys", "ctrl+space"), ("keys", "ctrl+plus")]},
    {"name": "Blaetter", "setup": {"A1": "eins"}, "start": "A1", "steps": [
        ("keys", "shift+f11"), ("type", "neu"), ("keys", "enter"), ("keys", "ctrl+pagedown"), ("tabdbl", "Tabelle1"),
        ("keys", "ctrl+a"), ("type", "Noten"), ("keys", "enter")]},
    {"name": "Kopieren", "setup": {"A1": "a", "A2": "b", "A3": "=A1&A2"}, "start": "A1", "steps": [
        ("keys", "shift+down", 2), ("keys", "ctrl+c"), ("keys", "right", 2), ("keys", "ctrl+v"), ("keys", "esc"),
        ("keys", "ctrl+home"), ("keys", "down"), ("keys", "delete")]},
    {"name": "Relativ", "setup": {}, "start": "B2", "relative": True, "steps": [
        ("type", "x"), ("keys", "enter"), ("type", "y"), ("keys", "tab"), ("keys", "shift+right"), ("keys", "ctrl+shift+f"),
        ("keys", "up"), ("keys", "shift+space"), ("keys", "ctrl+plus")]},
    {"name": "Spalte", "setup": {"A1": "ein langer Text", "B1": "kurz"}, "start": "B1", "steps": [
        ("colborder", "A"), ("keys", "down", 2), ("keys", "ctrl+9")]},
]
