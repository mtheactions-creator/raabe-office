"""RaabeCalc's options against Excel's on the same workbook (D-141): each case sets up the sweep's small
workbook in both, selects the same cells, triggers the option in RaabeCalc by hand (trusted DevTools clicks on
the ribbon, menus and dialog fields of the release exe, off screen) and in an invisible Excel of our own the
same ribbon command (CommandBars.ExecuteMso, or the dialog's effect through COM where a command has no id),
saves both, opens both files in Excel and compares what Excel sees: every cell's value, formula, number
format, font, fill, alignment, borders and merge, column widths, row heights, the window (frozen panes,
selection), AutoFilter, conditional formats, tables, page setup. Differences already there before the option
(the setup) are left out.

    .venv/Scripts/python.exe compare_calc_options.py [--only <name substring>]

Result: results/2026-11-09-calc-options-excel.json.
"""
from __future__ import annotations

import argparse
import json
import shutil
import sys
import tempfile
import time
from pathlib import Path

import pythoncom

import sweep_calc as SW
from feellab.excel import ExcelTarget
from xlsx_excel_check import diff, open_book, sheet_props

HERE = Path(__file__).parent
TMP = Path(tempfile.gettempdir()) / "raabe-calc-compare"
CELLS = [f"{c}{r}" for r in range(1, 9) for c in "ABCDEF"]

# A case: name, selection, RaabeCalc steps, Excel action. Steps: ("tab", label), ("ctrl", id), ("item", id),
# ("ctx", where), ("field", label, value) types into a dialog field, ("check", label) clicks a dialog checkbox/radio,
# ("ok",) clicks the dialog's default button, ("btn", label) a dialog button, ("key", key, mods), ("type", text).
# Excel: ("mso", idMso) or a function (app, ws) -> None.
T = ("tab", "Start")


def xl_numfmt(code: str):
    return lambda app, ws: setattr(app.Selection, "NumberFormatLocal", code)


def xl_set(prop: str, value):
    def run(app, ws):
        obj = app.Selection
        *path, last = prop.split(".")
        for p in path:
            obj = getattr(obj, p)
        setattr(obj, last, value)
    return run


def xl_cf(fc, dup: bool = False, top: int | None = None) -> None:
    """A rule as Excel's dialog makes it with its first preset (Hellrote Füllung 2 mit dunkelrotem Text)."""
    if dup:
        fc.DupeUnique = 1  # xlDuplicate
    if top is not None:
        fc.Rank = top
    fc.Interior.Color = 0xCEC7FF
    fc.Font.Color = 0x06009C


def xl_sort_region(order: int):
    """Excel's Von A bis Z / Z bis A on one cell: the current region, the header guessed, by the active cell's column."""
    def run(app, ws):
        app.CommandBars.ExecuteMso("SortAscendingExcel" if order == 1 else "SortDescendingExcel")
    return run


CASES: list[tuple] = [
    ("Fett", "A2:B4", [T, ("ctrl", "Schriftart|Fett")], ("mso", "Bold")),
    ("Fett auf gemischter Auswahl", "A1:B4", [T, ("ctrl", "Schriftart|Fett")], ("mso", "Bold")),
    ("Kursiv", "A2:B4", [T, ("ctrl", "Schriftart|Kursiv")], ("mso", "Italic")),
    ("Unterstreichen", "A2:B4", [T, ("ctrl", "Schriftart|Unterstreichen")], ("mso", "Underline")),
    ("Doppelt unterstreichen", "A2:B4", [T, ("ctrl", "Schriftart|Unterstreichen ▾"), ("item", "UDoppelt unterstreichen")], ("mso", "UnderlineDouble")),
    ("Schriftgrad vergrößern", "A2:B4", [T, ("ctrl", "Schriftart|Schriftgrad vergrößern")], ("mso", "FontSizeIncrease")),
    ("Schriftgrad verkleinern", "A2:B4", [T, ("ctrl", "Schriftart|Schriftgrad verkleinern")], ("mso", "FontSizeDecrease")),
    ("Schriftgrad 16", "A2:B4", [T, ("ctrl", "Schriftart|Schriftgrad"), ("item", "16")], xl_set("Font.Size", 16)),
    ("Schriftart Arial", "A2:B4", [T, ("ctrl", "Schriftart|Schriftart"), ("item", "Arial")], xl_set("Font.Name", "Arial")),
    ("Rahmenlinie unten", "A2:B4", [T, ("ctrl", "Schriftart|Rahmen ▾"), ("item", "Rahmenlinie unten")], ("mso", "BorderBottom")),
    ("Rahmenlinie oben", "A2:B4", [T, ("ctrl", "Schriftart|Rahmen ▾"), ("item", "Rahmenlinie oben")], ("mso", "BorderTop")),
    ("Rahmenlinie links", "A2:B4", [T, ("ctrl", "Schriftart|Rahmen ▾"), ("item", "Rahmenlinie links")], ("mso", "BorderLeft")),
    ("Rahmenlinie rechts", "A2:B4", [T, ("ctrl", "Schriftart|Rahmen ▾"), ("item", "Rahmenlinie rechts")], ("mso", "BorderRight")),
    ("Alle Rahmenlinien", "A2:B4", [T, ("ctrl", "Schriftart|Rahmen ▾"), ("item", "Alle Rahmenlinien")],
     lambda app, ws: [setattr(app.Selection.Borders(i), "LineStyle", 1) for i in (7, 8, 9, 10, 11, 12)]),
    ("Rahmenlinien außen", "A2:B4", [T, ("ctrl", "Schriftart|Rahmen ▾"), ("item", "Rahmenlinien außen")], ("mso", "BorderOutside")),
    ("Dicke Rahmenlinie außen", "A2:B4", [T, ("ctrl", "Schriftart|Rahmen ▾"), ("item", "Dicke Rahmenlinie außen")], ("mso", "BorderThickOutside")),
    ("Doppelte Rahmenlinien unten", "A2:B4", [T, ("ctrl", "Schriftart|Rahmen ▾"), ("item", "Doppelte Rahmenlinien unten")], ("mso", "BorderDoubleBottom")),
    ("Dicke Rahmenlinie unten", "A2:B4", [T, ("ctrl", "Schriftart|Rahmen ▾"), ("item", "Dicke Rahmenlinie unten")], ("mso", "BorderThickBottom")),
    ("Rahmenlinie oben und unten", "A2:B4", [T, ("ctrl", "Schriftart|Rahmen ▾"), ("item", "Rahmenlinie oben und unten")], ("mso", "BorderTopAndBottom")),
    ("Rahmenlinie oben und dicke unten", "A2:B4", [T, ("ctrl", "Schriftart|Rahmen ▾"), ("item", "Rahmenlinie oben und dicke unten")], ("mso", "BorderTopAndThickBottom")),
    ("Rahmenlinie oben und doppelte unten", "A2:B4", [T, ("ctrl", "Schriftart|Rahmen ▾"), ("item", "Rahmenlinie oben und doppelte unten")], ("mso", "BorderTopAndDoubleBottom")),
    ("Kein Rahmen nach Alle", "A2:B4", [T, ("ctrl", "Schriftart|Rahmen ▾"), ("item", "Alle Rahmenlinien"), ("ctrl", "Schriftart|Rahmen ▾"), ("item", "Kein Rahmen")],
     lambda app, ws: ([setattr(app.Selection.Borders(i), "LineStyle", 1) for i in (7, 8, 9, 10, 11, 12)], app.CommandBars.ExecuteMso("BorderNone"))),
    ("Oben ausrichten", "A2:B4", [T, ("ctrl", "Ausrichtung|Oben ausrichten")], ("mso", "AlignTopExcel")),
    ("Zentriert ausrichten", "A2:B4", [T, ("ctrl", "Ausrichtung|Zentriert ausrichten")], ("mso", "AlignMiddleExcel")),
    ("Linksbündig", "A2:B4", [T, ("ctrl", "Ausrichtung|Linksbündig")], ("mso", "AlignLeft")),
    ("Zentriert", "A2:B4", [T, ("ctrl", "Ausrichtung|Zentriert")], ("mso", "AlignCenter")),
    ("Rechtsbündig", "A2:B4", [T, ("ctrl", "Ausrichtung|Rechtsbündig")], ("mso", "AlignRight")),
    ("Einzug vergrößern", "A2:B4", [T, ("ctrl", "Ausrichtung|Einzug vergrößern")], ("mso", "IndentIncreaseExcel")),
    ("Einzug vergrößern und verkleinern", "A2:B4", [T, ("ctrl", "Ausrichtung|Einzug vergrößern"), ("ctrl", "Ausrichtung|Einzug vergrößern"), ("ctrl", "Ausrichtung|Einzug verkleinern")],
     lambda app, ws: [app.CommandBars.ExecuteMso(i) for i in ("IndentIncreaseExcel", "IndentIncreaseExcel", "IndentDecreaseExcel")]),
    ("Textumbruch", "A2:B4", [T, ("ctrl", "Ausrichtung|Textumbruch")], ("mso", "WrapText")),
    ("Verbinden und zentrieren", "A1:D1", [T, ("ctrl", "Ausrichtung|Verbinden und zentrieren"), ("ok",)], ("mso", "MergeCenter")),
    ("Verbinden über", "A2:C3", [T, ("ctrl", "Ausrichtung|Verbinden und zentrieren ▾"), ("item", "Verbinden über"), ("ok",)],
     lambda app, ws: app.Selection.Merge(True)),
    ("Zellen verbinden", "C1:D2", [T, ("ctrl", "Ausrichtung|Verbinden und zentrieren ▾"), ("item", "Zellen verbinden"), ("ok",)], ("mso", "MergeCells")),
    ("Gegen den Uhrzeigersinn drehen", "A1:D1", [T, ("ctrl", "Ausrichtung|Ausrichtung"), ("item", "Gegen den Uhrzeigersinn drehen")], xl_set("Orientation", 45)),
    ("Text nach oben drehen", "A1:D1", [T, ("ctrl", "Ausrichtung|Ausrichtung"), ("item", "Text nach oben drehen")], xl_set("Orientation", 90)),
    ("Prozentformat", "B2:B5", [T, ("ctrl", "Zahl|Prozentformat")], ("mso", "PercentStyle")),
    ("1.000er-Trennzeichen", "B2:B5", [T, ("ctrl", "Zahl|1.000er-Trennzeichen")], ("mso", "CommaStyle")),
    ("Dezimalstelle hinzufügen", "B2:B5", [T, ("ctrl", "Zahl|Dezimalstelle hinzufügen")], ("mso", "DecimalsIncrease")),
    ("Dezimalstelle entfernen", "B2:B5", [T, ("ctrl", "Zahl|Dezimalstelle entfernen")], ("mso", "DecimalsDecrease")),
    ("Dezimalstelle hinzufügen Standard", "D2:D5", [T, ("ctrl", "Zahl|Dezimalstelle hinzufügen")], ("mso", "DecimalsIncrease")),
    ("Buchhaltungszahlenformat", "B2:B5", [T, ("ctrl", "Zahl|Buchhaltungszahlenformat")], ("mso", "AccountingFormat")),
    ("Zahlenformat Währung", "D2:D5", [T, ("ctrl", "Zahl|Zahlenformat"), ("item", "Währung")], xl_numfmt("#.##0,00 €")),
    ("Zahlenformat Prozent", "D2:D5", [T, ("ctrl", "Zahl|Zahlenformat"), ("item", "Prozent")], xl_numfmt("0,00%")),
    ("Zahlenformat Datum lang", "C2:C5", [T, ("ctrl", "Zahl|Zahlenformat"), ("item", "Datum, lang")], xl_numfmt("[$-x-sysdate]TTTT, T. MMMM JJJJ")),
    ("Zahlenformat Text", "D2:D5", [T, ("ctrl", "Zahl|Zahlenformat"), ("item", "Text")], xl_numfmt("@")),
    ("Alle löschen", "A2:B4", [T, ("ctrl", "Bearbeiten|Löschen"), ("item", "Alle löschen")], ("mso", "ClearAll")),
    ("Formate löschen", "A1:B4", [T, ("ctrl", "Bearbeiten|Löschen"), ("item", "Formate löschen")], ("mso", "ClearFormats")),
    ("Inhalte löschen", "A2:B4", [T, ("ctrl", "Bearbeiten|Löschen"), ("item", "Inhalte löschen")], ("mso", "ClearContents")),
    ("Ausfüllen unten", "A2:B5", [T, ("ctrl", "Bearbeiten|Ausfüllen"), ("item", "Unten")], ("mso", "FillDown")),
    ("Ausfüllen rechts", "B2:D3", [T, ("ctrl", "Bearbeiten|Ausfüllen"), ("item", "Rechts")], ("mso", "FillRight")),
    ("Ausfüllen oben", "A2:B5", [T, ("ctrl", "Bearbeiten|Ausfüllen"), ("item", "Oben")], ("mso", "FillUp")),
    ("AutoSumme unter Spalte", "D6", [T, ("ctrl", "Bearbeiten|AutoSumme"), ("key", "Enter")], ("mso", "AutoSum")),
    ("AutoSumme Auswahl", "B2:B5", [T, ("ctrl", "Bearbeiten|AutoSumme")], ("mso", "AutoSum")),
    ("AutoSumme Mittelwert", "D2:D5", [T, ("ctrl", "Bearbeiten|AutoSumme ▾"), ("item", "Mittelwert")], lambda app, ws: setattr(ws.Range("D6"), "FormulaLocal", "=MITTELWERT(D2:D5)")),
    ("Nach Größe sortieren (aufsteigend)", "B3", [T, ("ctrl", "Bearbeiten|Sortieren und Filtern"), ("item", "Nach Größe sortieren (aufsteigend)")], ("mso", "SortAscendingExcel")),
    ("Von Z bis A sortieren", "A3", [T, ("ctrl", "Bearbeiten|Sortieren und Filtern"), ("item", "Von Z bis A sortieren")], ("mso", "SortDescendingExcel")),
    ("Filtern", "B3", [T, ("ctrl", "Bearbeiten|Sortieren und Filtern"), ("item", "Filtern")], ("mso", "Filter")),
    ("Blattzeilen einfügen", "A3:B4", [T, ("ctrl", "Zellen|Einfügen ▾"), ("item", "Blattzeilen einfügen")], ("mso", "SheetRowsInsert")),
    ("Blattspalten einfügen", "B3", [T, ("ctrl", "Zellen|Einfügen ▾"), ("item", "Blattspalten einfügen")], ("mso", "SheetColumnsInsert")),
    ("Blattzeilen löschen", "A3", [T, ("ctrl", "Zellen|Löschen ▾"), ("item", "Blattzeilen löschen")], ("mso", "SheetRowsDelete")),
    ("Blattspalten löschen", "C2", [T, ("ctrl", "Zellen|Löschen ▾"), ("item", "Blattspalten löschen")], ("mso", "SheetColumnsDelete")),
    ("Spaltenbreite automatisch anpassen", "A1:D6", [T, ("ctrl", "Zellen|Format"), ("item", "Spaltenbreite automatisch anpassen")], ("mso", "ColumnWidthAutoFit")),
    ("Spaltenbreite 20", "B2", [T, ("ctrl", "Zellen|Format"), ("item", "Spaltenbreite..."), ("field", "Spaltenbreite:", "20"), ("ok",)], lambda app, ws: setattr(app.Selection.EntireColumn, "ColumnWidth", 20)),
    ("Zeilenhöhe 30", "B2", [T, ("ctrl", "Zellen|Format"), ("item", "Zeilenhöhe..."), ("field", "Zeilenhöhe:", "30"), ("ok",)], lambda app, ws: setattr(app.Selection.EntireRow, "RowHeight", 30)),
    ("Zeilen ausblenden", "A3:A4", [T, ("ctrl", "Zellen|Format"), ("item", "Ausblenden und Einblenden"), ("item", "› Zeilen ausblenden")], lambda app, ws: setattr(app.Selection.EntireRow, "Hidden", True)),
    ("Spalten ausblenden", "C3", [T, ("ctrl", "Zellen|Format"), ("item", "Ausblenden und Einblenden"), ("item", "› Spalten ausblenden")], lambda app, ws: setattr(app.Selection.EntireColumn, "Hidden", True)),
    ("Zellenformatvorlage Gut", "A2:A4", [T, ("ctrl", "Formatvorlagen|Zellenformatvorlagen"), ("item", "Gut")], lambda app, ws: setattr(app.Selection, "Style", "Gut")),
    ("Zellenformatvorlage Überschrift 1", "A1:D1", [T, ("ctrl", "Formatvorlagen|Zellenformatvorlagen"), ("item", "Überschrift 1")], lambda app, ws: setattr(app.Selection, "Style", "Überschrift 1")),
    ("Bedingte Formatierung größer als", "B2:B5", [T, ("ctrl", "Formatvorlagen|Bedingte Formatierung"), ("item", "Regeln zum Hervorheben von Zellen"), ("item", "› Größer als..."), ("field", "", "10"), ("ok",)],
     lambda app, ws: xl_cf(app.Selection.FormatConditions.Add(1, 5, "=10"))),
    ("Bedingte Formatierung Textinhalt", "A2:A5", [T, ("ctrl", "Formatvorlagen|Bedingte Formatierung"), ("item", "Regeln zum Hervorheben von Zellen"), ("item", "› Textinhalt..."), ("field", "", "n"), ("ok",)],
     lambda app, ws: xl_cf(app.Selection.FormatConditions.Add(9, None, None, None, "n", 0))),
    ("Bedingte Formatierung Doppelte Werte", "D2:D5", [T, ("ctrl", "Formatvorlagen|Bedingte Formatierung"), ("item", "Regeln zum Hervorheben von Zellen"), ("item", "› Doppelte Werte..."), ("ok",)],
     lambda app, ws: xl_cf(app.Selection.FormatConditions.AddUniqueValues(), dup=True), {"D3": "2"}),
    ("Bedingte Formatierung Obere 2", "B2:B5", [T, ("ctrl", "Formatvorlagen|Bedingte Formatierung"), ("item", "Regeln für oberste/unterste Werte"), ("item", "› Obere 10 Elemente..."), ("field", "", "2"), ("ok",)],
     lambda app, ws: xl_cf(app.Selection.FormatConditions.AddTop10(), top=2)),
    ("Bedingte Formatierung Über dem Durchschnitt", "B2:B5", [T, ("ctrl", "Formatvorlagen|Bedingte Formatierung"), ("item", "Regeln für oberste/unterste Werte"), ("item", "› Über dem Durchschnitt..."), ("ok",)],
     lambda app, ws: xl_cf(app.Selection.FormatConditions.AddAboveAverage())),
    ("Neue Regel Formel", "A2:B5", [T, ("ctrl", "Formatvorlagen|Bedingte Formatierung"), ("item", "Neue Regel..."), ("click", ".cf-type:nth-child(6)"),
                                    ("field", "Werte formatieren", "=$B2>10"), ("ok",)],
     lambda app, ws: app.Selection.FormatConditions.Add(2, None, "=$B2>10")),
    ("Neue Regel Zellwert zwischen", "B2:B5", [T, ("ctrl", "Formatvorlagen|Bedingte Formatierung"), ("item", "Neue Regel..."),
                                               ("field", "Erster Wert", "8"), ("field", "Zweiter Wert", "13"), ("ok",)],
     lambda app, ws: app.Selection.FormatConditions.Add(1, 1, "=8", "=13")),
    ("Bedingte Formatierung Datum heute", "C2:C5", [T, ("ctrl", "Formatvorlagen|Bedingte Formatierung"), ("item", "Regeln zum Hervorheben von Zellen"), ("item", "› Datum..."), ("ok",)],
     lambda app, ws: xl_cf(app.Selection.FormatConditions.Add(11, None, None, None, None, None, 1))),
    ("Zellen formatieren Schriftfarbe", "A2:A4", [T, ("ctrl", "Schriftart|Zellenschriftart formatieren"), ("btn", "Schrift"), ("click", ".fc-color"), ("popswatch", 5), ("ok",)], None, None,
     lambda c: all(c[a]["Font"][8] not in (0, 0.0) for a in ("A2", "A3", "A4"))),
    ("Zellenformatvorlage Gut Schriftfarbe", "A2:A4", [T, ("ctrl", "Formatvorlagen|Zellenformatvorlagen"), ("item", "Gut")],
     lambda app, ws: (setattr(app.Selection.Interior, "Color", 0xCEEFC6), setattr(app.Selection.Font, "Color", 0x006100))),
    ("Formeln anzeigen", "A1", [("tab", "Formeln"), ("ctrl", "Formelüberwachung|Formeln anzeigen")], ("mso", "ShowFormulas")),
    ("Fenster fixieren", "B3", [("tab", "Ansicht"), ("ctrl", "Fenster|Fenster fixieren"), ("item", "Fenster fixieren")], ("mso", "FreezePanes")),
    ("Oberste Zeile fixieren", "B3", [("tab", "Ansicht"), ("ctrl", "Fenster|Fenster fixieren"), ("item", "Oberste Zeile fixieren")], ("mso", "FreezeTopRow")),
    ("Erste Spalte fixieren", "B3", [("tab", "Ansicht"), ("ctrl", "Fenster|Fenster fixieren"), ("item", "Erste Spalte fixieren")], ("mso", "FreezeFirstColumn")),
    ("Querformat", "A1", [("tab", "Seitenlayout"), ("ctrl", "Seite einrichten|Ausrichtung"), ("item", "Querformat")], lambda app, ws: setattr(ws.PageSetup, "Orientation", 2)),
    ("Druckbereich festlegen", "A1:B4", [("tab", "Seitenlayout"), ("ctrl", "Seite einrichten|Druckbereich"), ("item", "Druckbereich festlegen")], ("mso", "PrintAreaSetPrintArea")),
    ("Gruppieren Zeilen", "A2:A4", [("tab", "Daten"), ("ctrl", "Gliederung|Gruppieren"), ("ok",)], lambda app, ws: app.Selection.EntireRow.Group()),
    ("Zelle sperren aus", "A2", [T, ("ctrl", "Zellen|Format"), ("item", "Zelle sperren")], xl_set("Locked", False)),
    ("Zellen formatieren Ausfüllen gemischt", "A1:B3", [T, ("ctrl", "Zahl|Zellennummer formatieren"), ("btn", "Ausfüllen"), ("swatch", 12), ("ok",)], None, None,
     # Every cell gets the fill; the bold header stays bold, the others stay not bold (only the fill was changed).
     lambda c: all(c[a]["Interior"][0] == 1 for a in ("A1", "B1", "A2", "B3")) and c["A1"]["Font"][2] and not c["A2"]["Font"][2]),
    ("Zellen formatieren OK ohne Änderung", "A1:B3", [T, ("ctrl", "Zahl|Zellennummer formatieren"), ("ok",)], lambda app, ws: None),
    ("Zellen formatieren Zeilenumbruch", "A2:A4", [T, ("ctrl", "Ausrichtung|Zellenausrichtung formatieren"), ("check", "Zeilenumbruch"), ("ok",)], xl_set("WrapText", True)),
    ("Zellen formatieren Fett Kursiv", "A2:B4", [T, ("ctrl", "Schriftart|Zellenschriftart formatieren"), ("select", "Schriftschnitt:", "Fett Kursiv"), ("ok",)],
     lambda app, ws: (setattr(app.Selection.Font, "Bold", True), setattr(app.Selection.Font, "Italic", True))),
    ("Zellen formatieren Zahl 2 Dezimalstellen", "D2:D5", [T, ("ctrl", "Zahl|Zellennummer formatieren"), ("selectlist", "Kategorie:", "Zahl"), ("ok",)], xl_numfmt("0,00")),
    ("Sortieren mit Formeln", "B3", [T, ("ctrl", "Bearbeiten|Sortieren und Filtern"), ("item", "Nach Größe sortieren (aufsteigend)")], ("mso", "SortAscendingExcel"), {"D2": "=B2*2", "D3": "=B3*2", "D4": "=B4*2", "D5": "=B5*2"}),
]


class Raabe:
    def __init__(self):
        self.s = SW.Sweep(port=9442)
        self.p = self.s.p

    def dialog_field(self, label: str, value: str) -> bool:
        el = self.p.js(f"""(() => {{ const d = [...document.querySelectorAll('.dialog-layer .dialog')].pop(); if (!d) return null;
          const l = {json.dumps(label)};
          const f = l ? ([...d.querySelectorAll('label.field, label, .field')].find(e => e.textContent.trim().startsWith(l))?.querySelector('input, textarea') ?? [...d.querySelectorAll('input')].find(i => (i.getAttribute('aria-label') || '').startsWith(l))) : d.querySelector('.dlg-body input[type=text], .dlg-body input:not([type]), .dlg-body input[type=number]');
          if (!f) return null; const r = f.getBoundingClientRect(); return {{x: r.left + r.width / 2, y: r.top + r.height / 2}}; }})()""")
        if not el:
            return False
        self.p.click(el["x"], el["y"])
        self.p.key("a", "ctrl")
        self.p.type(value)
        return True

    def dialog_click(self, js_find: str) -> bool:
        el = self.p.js(f"""(() => {{ const d = [...document.querySelectorAll('.dialog-layer .dialog')].pop(); if (!d) return null;
          const e = {js_find}; if (!e) return null; e.scrollIntoView({{block: 'nearest'}}); const r = e.getBoundingClientRect(); return {{x: r.left + r.width / 2, y: r.top + r.height / 2}}; }})()""")
        if not el:
            return False
        self.p.click(el["x"], el["y"])
        self.p.pump(0.15)
        return True

    def select_option(self, label: str, value: str, listbox: bool = False) -> bool:
        """A <select>: a user clicks it and picks the entry; native drop-downs are not in the page, so the
        keyboard picks it as a user can (focus, then the entry's first letters)."""
        ok = self.p.js(f"""(() => {{ const d = [...document.querySelectorAll('.dialog-layer .dialog')].pop();
          const s = [...d.querySelectorAll('label.field, .field, div')].find(e => e.firstChild?.textContent?.trim() === {json.dumps(label)} || e.textContent.trim().startsWith({json.dumps(label)}))?.querySelector('select');
          if (!s) return false; const i = [...s.options].findIndex(o => o.textContent === {json.dumps(value)}); if (i < 0) return false;
          s.focus(); return {{i, cur: s.selectedIndex}}; }})()""")
        if not ok:
            return False
        steps = ok["i"] - ok["cur"]
        for _ in range(abs(steps)):
            self.p.key("ArrowDown" if steps > 0 else "ArrowUp")
        self.p.pump(0.1)
        return True

    def run(self, sel: str, steps: list[tuple], setup_extra: dict | None) -> dict:
        s = self.s
        s.reset(sel)
        if setup_extra:
            # Typed as a user would: the cell selected, the text typed, Enter.
            for a, t in setup_extra.items():
                self.p.js(f"(() => {{ const r = __calc.parse({json.dumps(a)}); __calc.app.grid.select(r, {{row: r.r0, col: r.c0}}); return true; }})()")
                self.p.type(t)
                self.p.key("Enter")
            self.p.js(f"(() => {{ const r = __calc.parse({json.dumps(sel)}); __calc.app.grid.select(r, {{row: r.r0, col: r.c0}}); __sweep.mark(); return true; }})()")
            self.p.errors.clear()
        out: dict = {"steps": []}
        for st in steps:
            k = st[0]
            if k in ("tab", "ctrl", "item"):
                r = s.step(k, st[1])
            elif k == "ctx":
                r = s.context_menu(st[1])
            elif k == "field":
                r = {"ok": self.dialog_field(st[1], st[2])}
            elif k == "check":
                r = {"ok": self.dialog_click(f"[...d.querySelectorAll('label')].find(l => l.textContent.trim() === {json.dumps(st[1])})?.querySelector('input')")}
            elif k == "btn":
                r = {"ok": self.dialog_click(f"[...d.querySelectorAll('button')].find(b => b.textContent.trim() === {json.dumps(st[1])})")}
            elif k == "swatch":
                r = {"ok": self.dialog_click(f"[...d.querySelectorAll('.fc-palette .swatch')].filter(e => e.offsetParent)[{st[1]}]")}
            elif k == "click":
                r = {"ok": self.dialog_click(f"d.querySelector({json.dumps(st[1])})")}
            elif k == "popswatch":
                el = self.p.js(f"""(() => {{ const b = [...document.querySelectorAll('.fc-pop .swatch')][{st[1]}]; if (!b) return null;
                  const r = b.getBoundingClientRect(); return {{x: r.left + r.width / 2, y: r.top + r.height / 2}}; }})()""")
                if el:
                    self.p.click(el["x"], el["y"])
                r = {"ok": bool(el)}
            elif k == "select":
                r = {"ok": self.select_option(st[1], st[2])}
            elif k == "selectlist":
                r = {"ok": self.select_option(st[1], st[2], True)}
            elif k == "ok":
                r = {"ok": self.dialog_click("d.querySelector('.dlg-buttons button.primary') || d.querySelector('.dlg-buttons button')")}
                self.p.pump(0.3)
            elif k == "key":
                self.p.key(st[1], st[2] if len(st) > 2 else "")
                r = {"ok": True}
            else:
                r = {"ok": False, "why": "unknown step"}
            out["steps"].append([str(st[1]) if len(st) > 1 else k, r.get("ok"), r.get("why")])
            if not r.get("ok"):
                out["failed"] = f"{st}: {r.get('why', 'not done')}"
                break
            self.p.pump(0.1)
        # A dialog still open (a message box) is shown in the result, then closed.
        out["dialog"] = self.p.js("__sweep.dialogTitle()")
        if out["dialog"]:
            self.p.key("Escape")
            self.p.pump(0.2)
        out["errors"] = list(self.p.errors)
        return out

    def save(self, path: Path) -> None:
        self.p.js(f"__calc.save({json.dumps(str(path))})")

    def close(self) -> None:
        self.s.close()


def excel_book(app, sel: str, action, setup_extra: dict | None, path: Path) -> str | None:
    wb = app.Workbooks.Add()
    try:
        st = wb.Styles("Standard")
        st.Font.Name = "Calibri"
        st.Font.Size = 11
        ws = wb.Worksheets(1)
        ws.Name = "Tabelle1"
        for a, t in {**SW.SETUP, **(setup_extra or {})}.items():
            ws.Range(a).FormulaLocal = t
        ws.Range("A1:D1").Select()
        app.CommandBars.ExecuteMso("Bold")
        ws.Range("B2:B5").Select()
        app.CommandBars.ExecuteMso("CommaStyle")
        ws.Range(sel).Select()
        ws.Range(sel.split(":")[0]).Activate()
        err = None
        if action is not None:
            try:
                if isinstance(action, tuple) and action[0] == "mso":
                    app.CommandBars.ExecuteMso(action[1])
                else:
                    action(app, ws)
            except Exception as e:  # noqa: BLE001
                err = str(e)[:300]
        wb.SaveAs(str(path), 51)
        return err
    finally:
        wb.Close(False)


def props(app, path: Path) -> dict:
    wb = open_book(app, path)
    try:
        ws = wb.Worksheets(1)
        p = sheet_props(wb, ws, CELLS)
        for k in ("charts", "shapes", "notes", "noteCount", "sparklines", "threads", "links", "validation", "protection", "UsedRange"):
            p.pop(k, None)
        p["cols"] = p["cols"][:7]
        p["rows"] = p["rows"][:9]
        for k in ("ScrollRow", "ScrollColumn", "Zoom"):
            p["window"].pop(k, None)
        p["cf"] = []
        def get(fn):
            try:
                return fn()
            except Exception:  # noqa: BLE001 - a property the rule's type does not have
                return None
        # Every rule as Excel reads it, and which of the cells it shows as matching (DisplayFormat).
        for fc in ws.Cells.FormatConditions:
            p["cf"].append([get(lambda: fc.Type), get(lambda: fc.AppliesTo.GetAddress(False, False)), get(lambda: fc.Operator), get(lambda: fc.Formula1),
                            get(lambda: fc.Formula2), get(lambda: fc.DateOperator), get(lambda: fc.TextOperator), get(lambda: fc.Text),
                            get(lambda: fc.Interior.Color if fc.Interior.ColorIndex != -4142 else None), get(lambda: fc.Font.Color)])
        if p["cf"]:
            p["cfShown"] = [a for a in CELLS if get(lambda a=a: ws.Range(a).DisplayFormat.Interior.Color != ws.Range(a).Interior.Color
                                                    or ws.Range(a).DisplayFormat.Font.Color != ws.Range(a).Font.Color)]
        return p
    finally:
        wb.Close(False)


def near(path: str, a, b) -> bool:
    """Sizes: row heights and column widths within what the pixel grid makes of them."""
    if isinstance(a, float) and isinstance(b, float) and ("/rows/" in path or "/cols/" in path):
        return abs(a - b) <= 0.8
    return False


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--only", default=None)
    ap.add_argument("--out", default="calc-options-excel")
    a = ap.parse_args()
    if TMP.exists():
        shutil.rmtree(TMP, ignore_errors=True)
    TMP.mkdir(parents=True, exist_ok=True)
    raabe = Raabe()
    xl = ExcelTarget(visible=False, place=False)
    app = xl.app
    app.DisplayAlerts = False
    results = []
    t0 = time.perf_counter()
    try:
        # The setup alone: what differs already is not the option's doing.
        raabe.s.reset("A1")
        raabe.save(TMP / "base-raabe.xlsx")
        excel_book(app, "A1", None, None, TMP / "base-excel.xlsx")
        base = {"raabe": props(app, TMP / "base-raabe.xlsx"), "excel": props(app, TMP / "base-excel.xlsx")}
        base_diff = set(d.split(":")[0] for d in diff(base["excel"], base["raabe"]))
        for case in CASES:
            name, sel, steps, action = case[:4]
            extra = case[4] if len(case) > 4 else None
            if a.only and a.only not in name:
                continue
            safe = "".join(ch if ch.isalnum() else "_" for ch in name)
            res: dict = {"name": name, "selection": sel}
            r = raabe.run(sel, steps, extra)
            res.update({k: v for k, v in r.items() if v})
            ours = TMP / f"{safe}-raabe.xlsx"
            raabe.save(ours)
            if action is None:
                # No Excel command for it (a palette swatch): the saved file, opened in Excel, against an expectation.
                res["excel"] = "no reference (checked by its own expectation)"
                cells = props(app, ours)["cells"]
                expect = case[5] if len(case) > 5 else None
                res["pass"] = bool(expect and expect(cells)) and not r.get("failed") and not r.get("errors")
                res["ours"] = {a: {k: cells[a][k] for k in ("Font", "Interior")} for a in ("A1", "A2", "B1", "B3")}
                results.append(res)
                print(name, "| ours only:", "ok" if res["pass"] else "FAILED", r.get("failed", ""), flush=True)
                continue
            theirs = TMP / f"{safe}-excel.xlsx"
            err = excel_book(app, sel, action, extra, theirs)
            if err:
                res["excelError"] = err
            try:
                pe, pr = props(app, theirs), props(app, ours)
            except Exception as e:  # noqa: BLE001
                res["openFailed"] = str(e)[:300]
                results.append(res)
                print(name, "| OPEN FAILED", flush=True)
                continue
            d = [x for x in diff(pe, pr) if x.split(":")[0] not in base_diff and not near(x.split(":")[0], *_vals(pe, pr, x.split(":")[0]))]
            res["diff"] = d
            res["pass"] = not d and not r.get("failed") and not r.get("errors")
            results.append(res)
            print(name, "|", "ok" if res["pass"] else "DIFF", r.get("failed", ""), d[:6], flush=True)
    finally:
        raabe.close()
        xl.close()
    out = HERE / "results" / f"2026-11-09-{a.out}.json"
    n = sum(1 for r in results if "pass" in r)
    summary = {"cases": len(results), "compared": n, "passed": sum(1 for r in results if r.get("pass")), "seconds": round(time.perf_counter() - t0),
               "baseDifferences": sorted(base_diff)}
    out.write_text(json.dumps({"summary": summary, "results": results}, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(json.dumps(summary, ensure_ascii=False, indent=1))


def _vals(a: dict, b: dict, path: str):
    def get(o, p):
        for k in p.strip("/").split("/"):
            if isinstance(o, dict):
                o = o.get(k)
            elif isinstance(o, (list, tuple)):
                try:
                    o = o[int(k)]
                except (ValueError, IndexError):
                    return None
            else:
                return None
        return o
    return get(a, path), get(b, path)


if __name__ == "__main__":
    pythoncom.CoInitialize()
    main()
