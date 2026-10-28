"""Behaviour catalogue for RaabeCalc against Excel (D-061). Real keyboard and mouse: run only
when nobody uses the PC.

    .venv/Scripts/python.exe catalogue_calc.py excel [area or name …]   record Excel's outcomes
    .venv/Scripts/python.exe catalogue_calc.py calc  [area or name …]   run RaabeCalc and compare

Excel's outcomes go to results/<date>-catalogue-calc-excel.json, the comparison to
results/<date>-catalogue-calc.json.
"""

from __future__ import annotations

import datetime as dt
import json
import sys
import time
from pathlib import Path

from feellab.calc_catalogue import Case, run_step, settle

C = Case
NUM3 = {"A1": "10", "A2": "20", "A3": "30"}
BLOCK = {"A1": "1", "A2": "2", "A3": "3", "B1": "4", "B2": "5", "B3": "6"}
GAPS = {"A1": "1", "A2": "2", "A3": "3", "A6": "6", "A7": "7", "A10": "10"}

CASES: list[Case] = [
    # ----- typing and committing -----------------------------------------------------------------
    C("type text Enter", {}, [("type", "abc"), ("keys", "enter")], "enter"),
    C("type number with comma", {}, [("type", "12,5"), ("keys", "enter")], "enter"),
    C("type thousands", {}, [("type", "1.234,5"), ("keys", "enter")], "enter"),
    C("type date", {}, [("type", "1.1.2026"), ("keys", "enter")], "enter"),
    C("type short date", {}, [("type", "3.4."), ("keys", "enter")], "enter"),
    C("type percent", {}, [("type", "15%"), ("keys", "enter")], "enter"),
    C("type euro", {}, [("type", "5 €"), ("keys", "enter")], "enter"),
    C("type time", {}, [("type", "12:30"), ("keys", "enter")], "enter"),
    C("type fraction", {}, [("type", "1 1/2"), ("keys", "enter")], "enter"),
    C("type scientific", {}, [("type", "1,5E3"), ("keys", "enter")], "enter"),
    C("type wahr", {}, [("type", "wahr"), ("keys", "enter")], "enter"),
    C("type apostrophe number", {}, [("type", "'0123"), ("keys", "enter")], "enter"),
    C("type leading zero", {}, [("type", "0123"), ("keys", "enter")], "enter"),
    C("type minus formula", {}, [("type", "-3+4"), ("keys", "enter")], "enter"),
    C("type plus formula", {}, [("type", "+2*3"), ("keys", "enter")], "enter"),
    C("type long text overflows", {}, [("type", "Ein langer Text, der überläuft"), ("keys", "enter")], "enter"),
    C("Tab then Enter returns to start column", {}, [("type", "a"), ("keys", "tab"), ("type", "b"), ("keys", "tab"), ("type", "c"), ("keys", "enter")], "enter"),
    C("Shift+Enter moves up", {}, [("type", "x"), ("keys", "shift+enter")], "enter", start="B3"),
    C("Shift+Tab moves left", {}, [("type", "x"), ("keys", "shift+tab")], "enter", start="C3"),
    C("arrow commits in Eingeben", {}, [("type", "x"), ("keys", "right")], "enter"),
    C("arrow down commits", {}, [("type", "x"), ("keys", "down")], "enter"),
    C("Esc cancels typing", {"A1": "alt"}, [("type", "neu"), ("keys", "esc")], "enter"),
    C("F2 edit and Left arrow moves caret", {"A1": "ab"}, [("keys", "f2"), ("keys", "left"), ("type", "X"), ("keys", "enter")], "edit"),
    C("double-click edits at the end", {"A1": "Hallo"}, [("dblclick", "A1")], "edit"),
    C("F2 shows Bearbeiten", {"A1": "ab"}, [("keys", "f2")], "edit"),
    C("F2 twice toggles to Eingeben", {"A1": "ab"}, [("keys", "f2"), ("keys", "f2"), ("keys", "right")], "edit"),
    C("Backspace clears and edits", {"A1": "abc"}, [("keys", "backspace"), ("keys", "enter")], "edit"),
    C("Delete clears contents", {"A1": "abc", "A2": "d"}, [("keys", "shift+down"), ("keys", "delete")], "edit"),
    C("Ctrl+Enter fills selection", {}, [("drag", "A1", "B3"), ("type", "7"), ("keys", "ctrl+enter")], "edit"),
    C("Alt+Enter line break", {}, [("type", "a"), ("keys", "alt+enter"), ("type", "b"), ("keys", "enter")], "edit"),
    C("Enter moves inside selection", {}, [("drag", "A1", "B2"), ("keys", "enter", 3)], "edit"),
    C("Tab moves inside selection", {}, [("drag", "A1", "B2"), ("keys", "tab", 3)], "edit"),
    C("Ctrl+Z undoes entry", {}, [("type", "x"), ("keys", "enter"), ("keys", "ctrl+z")], "edit"),
    C("Ctrl+Y redoes", {}, [("type", "x"), ("keys", "enter"), ("keys", "ctrl+z"), ("keys", "ctrl+y")], "edit"),
    C("AutoComplete from column", {"A1": "Apfel", "A2": "Birne"}, [("type", "Ap"), ("keys", "enter")], "edit", start="A3"),
    C("AutoComplete needs unique match", {"A1": "Apfel", "A2": "Aprikose"}, [("type", "Ap"), ("keys", "enter")], "edit", start="A3"),
    C("AutoComplete keeps typed case", {"A1": "Apfel"}, [("type", "ap"), ("keys", "enter")], "edit", start="A2"),
    # ----- moving and selecting -------------------------------------------------------------------
    C("Ctrl+Down to block end", GAPS, [("keys", "ctrl+down")], "move"),
    C("Ctrl+Down twice to next block", GAPS, [("keys", "ctrl+down", 2)], "move"),
    C("Ctrl+Down four times", GAPS, [("keys", "ctrl+down", 4)], "move"),
    C("Ctrl+Down from empty", {}, [("keys", "ctrl+down")], "move", start="B2"),
    C("Ctrl+Right from data", BLOCK, [("keys", "ctrl+right")], "move"),
    C("Ctrl+Up from below", GAPS, [("keys", "ctrl+up")], "move", start="A12"),
    C("Ctrl+Shift+Down extends", GAPS, [("keys", "ctrl+shift+down")], "move"),
    C("Shift+arrows extend", {}, [("keys", "shift+down", 2), ("keys", "shift+right")], "move"),
    C("Shift+arrow back shrinks", {}, [("keys", "shift+down", 3), ("keys", "shift+up")], "move"),
    C("Ctrl+Home", BLOCK, [("keys", "ctrl+home")], "move", start="D8"),
    C("Ctrl+End", BLOCK, [("keys", "ctrl+end")], "move"),
    C("Home to column A", {}, [("keys", "home")], "move", start="E5"),
    C("Ctrl+A selects region", BLOCK, [("keys", "ctrl+a")], "move"),
    C("Ctrl+A twice selects all", BLOCK, [("keys", "ctrl+a", 2)], "move"),
    C("Ctrl+A on empty selects all", {}, [("keys", "ctrl+a")], "move", start="E5"),
    C("Shift+Space selects row", {}, [("keys", "shift+space")], "move", start="C3"),
    C("Ctrl+Space selects column", {}, [("keys", "ctrl+space")], "move", start="C3"),
    C("PageDown", {}, [("keys", "pagedown")], "move"),
    C("click selects", {}, [("click", "C4")], "mouse"),
    C("Shift+click extends", {}, [("click", "B2"), ("click", "D5", "shift")], "mouse"),
    C("Ctrl+click adds", {}, [("click", "B2"), ("click", "D5", "ctrl")], "mouse"),
    C("drag selects", {}, [("drag", "B2", "D4")], "mouse"),
    C("drag upwards", {}, [("drag", "D4", "B2")], "mouse"),
    C("column header click", {}, [("colhead", "C")], "mouse"),
    C("column header shift-click", {}, [("colhead", "B"), ("colhead", "D", "shift")], "mouse"),
    C("row header click", {}, [("rowhead", 3)], "mouse"),
    C("row header then Enter", {}, [("rowhead", 3), ("keys", "enter")], "mouse"),
    # ----- formulas: pointing, F4, AutoSumme ----------------------------------------------------------
    C("= then Up points", NUM3, [("type", "="), ("keys", "up"), ("keys", "enter")], "formula", start="A4"),
    C("= then Up Up points further", NUM3, [("type", "="), ("keys", "up", 2), ("keys", "enter")], "formula", start="A4"),
    C("Shift+Up extends pointed range", NUM3, [("type", "=SUMME("), ("keys", "up"), ("keys", "shift+up", 2), ("keys", "enter")], "formula", start="A4"),
    C("operator then arrow points again", NUM3, [("type", "="), ("keys", "up"), ("type", "+"), ("keys", "up", 2), ("keys", "enter")], "formula", start="A4"),
    C("click points", NUM3, [("type", "="), ("click", "A2"), ("keys", "enter")], "formula", start="C1"),
    C("drag points range", NUM3, [("type", "=SUMME("), ("drag", "A1", "A3"), ("keys", "enter")], "formula", start="C1"),
    C("open bracket closed", NUM3, [("type", "=SUMME(A1:A3"), ("keys", "enter")], "formula", start="C1"),
    C("F4 makes absolute", NUM3, [("type", "=A1"), ("keys", "f4"), ("keys", "enter")], "formula", start="C1"),
    C("F4 twice", NUM3, [("type", "=A1"), ("keys", "f4", 2), ("keys", "enter")], "formula", start="C1"),
    C("F4 three times", NUM3, [("type", "=A1"), ("keys", "f4", 3), ("keys", "enter")], "formula", start="C1"),
    C("lower case function upper", NUM3, [("type", "=summe(a1:a3)"), ("keys", "enter")], "formula", start="C1"),
    C("Alt+= AutoSumme", NUM3, [("keys", "alt+shift+0"), ("keys", "enter")], "formula", start="A4"),
    C("AutoSumme mode", NUM3, [("keys", "alt+shift+0")], "formula", start="A4"),
    C("AutoSumme row", {"A1": "1", "B1": "2", "C1": "3"}, [("keys", "alt+shift+0"), ("keys", "enter")], "formula", start="D1"),
    C("pointing mode shown", NUM3, [("type", "="), ("keys", "up")], "formula", start="A4"),
    C("Enter mode after typing =", NUM3, [("type", "=")], "formula", start="A4"),
    C("division by zero", {}, [("type", "=1/0"), ("keys", "enter")], "formula"),
    C("unknown name", {}, [("type", "=foo+1"), ("keys", "enter")], "formula"),
    # ----- filling ---------------------------------------------------------------------------------
    C("fill one number copies", {"A1": "1"}, [("fill", "A1", "A5")], "fill"),
    C("fill one number with Ctrl counts", {"A1": "1"}, [("fill", "A1", "A5", "ctrl")], "fill"),
    C("fill two numbers continues", {"A1": "1", "A2": "3"}, [("drag", "A1", "A2"), ("fill", "A1:A2", "A6")], "fill"),
    C("fill weekday", {"A1": "Mo"}, [("fill", "A1", "A5")], "fill"),
    C("fill month name", {"A1": "Januar"}, [("fill", "A1", "A4")], "fill"),
    C("fill text with number", {"A1": "Woche 1"}, [("fill", "A1", "A4")], "fill"),
    C("fill date", {"A1": "1.1.2026"}, [("fill", "A1", "A4")], "fill"),
    C("fill formula shifts", {"A1": "1", "A2": "2", "A3": "3", "B1": "=A1*2"}, [("click", "B1"), ("fill", "B1", "B3")], "fill"),
    C("fill right", {"A1": "1", "A2": "2"}, [("drag", "A1", "A2"), ("fill", "A1:A2", "C2")], "fill"),
    C("fill handle double-click", {"A1": "1", "A2": "2", "A3": "3", "A4": "4", "B1": "=A1+1"}, [("click", "B1"), ("fillclick", "B1")], "fill"),
    C("Strg+U fills down", {"A1": "x"}, [("drag", "A1", "A4"), ("keys", "ctrl+u")], "fill"),
    C("Strg+R fills right", {"A1": "x"}, [("drag", "A1", "D1"), ("keys", "ctrl+r")], "fill"),
    # ----- moving and copying cells ------------------------------------------------------------------
    C("drag border moves", {"A1": "5", "B1": "=A1*2"}, [("border", "A1", "C3")], "move-cells"),
    C("drag border with Ctrl copies", {"A1": "5"}, [("border", "A1", "C3", "ctrl")], "move-cells"),
    C("copy and paste", {"A1": "1", "A2": "=A1+1"}, [("drag", "A1", "A2"), ("keys", "ctrl+c"), ("click", "C1"), ("keys", "ctrl+v"), ("keys", "esc")], "move-cells"),
    C("cut and paste updates references", {"A1": "1", "B1": "=A1"}, [("click", "A1"), ("keys", "ctrl+x"), ("click", "A3"), ("keys", "ctrl+v")], "move-cells"),
    C("paste into larger selection repeats", {"A1": "1"}, [("keys", "ctrl+c"), ("drag", "C1", "C4"), ("keys", "ctrl+v"), ("keys", "esc")], "move-cells"),
    C("Enter pastes and clears marquee", {"A1": "1"}, [("keys", "ctrl+c"), ("click", "C1"), ("keys", "enter")], "move-cells"),
    # ----- formatting shortcuts ------------------------------------------------------------------------
    C("Strg+Umschalt+F bold", {"A1": "x"}, [("keys", "ctrl+shift+f")], "format"),
    C("Strg+Umschalt+K italic", {"A1": "x"}, [("keys", "ctrl+shift+k")], "format"),
    C("Strg+Umschalt+U underline", {"A1": "x"}, [("keys", "ctrl+shift+u")], "format"),
    C("Strg+2 bold", {"A1": "x"}, [("keys", "ctrl+2")], "format"),
    C("Strg+5 strike", {"A1": "x"}, [("keys", "ctrl+5")], "format"),
    C("Strg+Umschalt+4 currency", {"A1": "1234,5"}, [("keys", "ctrl+shift+4")], "format"),
    C("Strg+Umschalt+5 percent", {"A1": "0,25"}, [("keys", "ctrl+shift+5")], "format"),
    C("Strg+Umschalt+1 number", {"A1": "1234,567"}, [("keys", "ctrl+shift+1")], "format"),
    C("Strg+Umschalt+3 date", {"A1": "46000"}, [("keys", "ctrl+shift+3")], "format"),
    C("Strg+Umschalt+6 scientific", {"A1": "1234"}, [("keys", "ctrl+shift+6")], "format"),
    C("Strg+9 hides row", {}, [("keys", "ctrl+9")], "format", start="A3"),
    C("Strg+0 hides column", {}, [("keys", "ctrl+0")], "format", start="C1"),
    C("column border double-click fits", {"B1": "Ein ziemlich langer Text"}, [("colborder", "B")], "format"),
    C("Strg+. inserts date", {}, [("keys", "ctrl+period"), ("keys", "enter")], "format", today=True),
    # ----- sheets ------------------------------------------------------------------------------------
    C("Umschalt+F11 new sheet", {}, [("keys", "shift+f11")], "sheets"),
    C("Strg+Bild ab next sheet", {}, [("keys", "ctrl+pagedown")], "sheets", sheets=2),
    # ----- more controls (second round) ------------------------------------------------------------
    C("F4 repeats bold", {"A1": "x", "B1": "y"}, [("keys", "ctrl+shift+f"), ("click", "B1"), ("keys", "f4")], "repeat"),
    C("Strg+Y repeats format", {"A1": "x", "B1": "y"}, [("keys", "ctrl+2"), ("click", "B1"), ("keys", "ctrl+y")], "repeat"),
    C("Strg+Minus deletes selected row", {"A1": "1", "A2": "2", "A3": "3"}, [("rowhead", 2), ("keys", "ctrl+minus")], "rows"),
    C("Strg+Plus inserts row", {"A1": "1", "A2": "2", "B3": "=A2*2"}, [("rowhead", 2), ("keys", "ctrl+plus")], "rows"),
    C("Strg+Minus deletes column", {"A1": "1", "B1": "2", "C1": "=B1+1"}, [("colhead", "A"), ("keys", "ctrl+minus")], "rows"),
    C("Ctrl+Shift+End", {"A1": "1", "C5": "2"}, [("keys", "ctrl+shift+end")], "move2"),
    C("Shift+Home", {}, [("keys", "shift+home")], "move2", start="D3"),
    C("Esc clears marquee then Enter moves", {"A1": "1"}, [("keys", "ctrl+c"), ("keys", "esc"), ("click", "C1"), ("keys", "enter")], "move2"),
    C("function AutoComplete Tab", NUM3, [("type", "=SUM"), ("keys", "tab"), ("type", "A1:A3)"), ("keys", "enter")], "formula2", start="B1"),
    C("fill handle back clears", {"A1": "1", "A2": "2", "A3": "3", "A4": "4"}, [("drag", "A1", "A4"), ("fill", "A1:A4", "A2")], "fill2"),
    C("fill Ctrl copies two numbers", {"A1": "1", "A2": "3"}, [("drag", "A1", "A2"), ("fill", "A1:A2", "A5", "ctrl")], "fill2"),
    C("fill weekday long", {"A1": "Montag"}, [("fill", "A1", "A3")], "fill2"),
    C("fill upwards", {"A5": "Woche 5"}, [("click", "A5"), ("fill", "A5", "A3")], "fill2", start="A5"),
    C("column resize drag", {}, [("colresize", "B", 60)], "sizes"),
    C("undo after delete", {"A1": "x"}, [("keys", "delete"), ("keys", "ctrl+z")], "repeat"),
    C("typing replaces selection range", {"A1": "1", "A2": "2"}, [("drag", "A1", "A2"), ("type", "9"), ("keys", "enter")], "edit2"),
    C("Backspace in edit removes char", {"A1": "abc"}, [("keys", "f2"), ("keys", "backspace"), ("keys", "enter")], "edit2"),
    C("Esc in Bearbeiten keeps old", {"A1": "abc"}, [("keys", "f2"), ("type", "zz"), ("keys", "esc")], "edit2"),
    C("arrow in formula after number commits", NUM3, [("type", "=A1+5"), ("keys", "down")], "formula2", start="B1"),
    # ----- the window's controls ------------------------------------------------------------------
    C("double-click tab renames", {}, [("tabdbl", "Tabelle1"), ("type", "Noten"), ("keys", "enter")], "window"),
    C("plus adds sheet after", {}, [("sheetadd",)], "window"),
    C("Namenfeld goes to cell", {}, [("namebox", "C5")], "window"),
    C("Namenfeld selects range", {}, [("namebox", "B2:D4")], "window"),
    C("status bar sums", {"A1": "1", "A2": "2,5", "A3": "x", "A4": "4"}, [("drag", "A1", "A4")], "window"),
    C("status bar count of text", {"A1": "a", "A2": "b"}, [("drag", "A1", "A2")], "window"),
    # ----- scrolling, zoom, display, more editing ---------------------------------------------------
    C("wheel down three notches", {}, [("wheel", "C5", -3)], "view"),
    C("wheel down then up", {}, [("wheel", "C5", -4), ("wheel", "C5", 1)], "view"),
    C("Strg+wheel zooms in", {}, [("ctrlwheel", "C5", 1)], "view"),
    C("Strg+wheel zooms out twice", {}, [("ctrlwheel", "C5", -2)], "view"),
    C("long number in General", {}, [("type", "123456789012"), ("keys", "enter")], "display"),
    C("many decimals in General", {}, [("type", "1234567,891"), ("keys", "enter")], "display"),
    C("small number in General", {}, [("type", "0,000012345"), ("keys", "enter")], "display"),
    C("two-digit year", {}, [("type", "1.1.26"), ("keys", "enter")], "display"),
    C("date and time", {}, [("type", "3.4.2026 14:30"), ("keys", "enter")], "display"),
    C("double-click border jumps down", GAPS, [("click", "A1"), ("borderdbl", "A1", "bottom")], "move3"),
    C("Pos1 in Bearbeiten", {"A1": "abc"}, [("keys", "f2"), ("keys", "home"), ("type", "X"), ("keys", "enter")], "edit3"),
    C("Shift+Pos1 selects text", {"A1": "abc"}, [("keys", "f2"), ("keys", "shift+home"), ("type", "Z"), ("keys", "enter")], "edit3"),
    C("undo column width", {}, [("colresize", "B", 60), ("keys", "ctrl+z")], "sizes2"),
    C("Strg+Pfeil in Bearbeiten moves by word", {"A1": "ein zwei drei"}, [("keys", "f2"), ("keys", "ctrl+left"), ("type", "X"), ("keys", "enter")], "edit3"),
    C("Alt+Unten pick list", {"A1": "Mathe", "A2": "Deutsch", "A3": "Physik", "A4": "Mathe"}, [("keys", "alt+down"), ("keys", "down"), ("keys", "enter")], "edit4", start="A5"),
    C("Alt+Unten pick list Esc", {"A1": "Mathe", "A2": "Deutsch"}, [("keys", "alt+down"), ("keys", "esc")], "edit4", start="A3"),
    C("Ende shows its mode", {"A1": "1", "A2": "2", "A3": "3"}, [("keys", "end")], "keys2"),
    C("Ende then Down jumps", {"A1": "1", "A2": "2", "A3": "3", "A6": "6"}, [("keys", "end"), ("keys", "down")], "keys2"),
    C("Ende then Shift+Right extends", {"A1": "1", "B1": "2", "C1": "3"}, [("keys", "end"), ("keys", "shift+right")], "keys2"),
    C("F8 extends with arrows", {"A1": "1"}, [("keys", "f8"), ("keys", "down"), ("keys", "right")], "keys2"),
    C("F8 twice ends extending", {"A1": "1"}, [("keys", "f8"), ("keys", "f8"), ("keys", "down")], "keys2"),
    C("Umschalt+F8 adds to selection", {"A1": "1"}, [("keys", "shift+f8")], "keys2"),
    C("Strg+Rücktaste shows active cell", {"A1": "1"}, [("wheel", "C5", -10), ("keys", "ctrl+backspace")], "keys2"),
    C("wheel away keeps active cell", {"A1": "1"}, [("wheel", "C5", -10)], "keys2"),
    C("Ctrl+Shift+Right extends to block end", {"A1": "1", "B1": "2", "C1": "3"}, [("keys", "ctrl+shift+right")], "keys3"),
    C("Ctrl+Shift+Home selects to A1", {}, [("keys", "ctrl+shift+home")], "keys3", start="C5"),
    C("Shift+Enter inside selection wraps", {}, [("keys", "shift+right"), ("keys", "shift+down"), ("keys", "shift+enter")], "keys3"),
    C("typing in selection with Tab wraps", {}, [("keys", "shift+right"), ("keys", "shift+down"), ("type", "a"), ("keys", "tab"), ("type", "b"), ("keys", "tab"), ("type", "c"), ("keys", "tab")], "keys3"),
    C("Strg+Z twice undoes two entries", {}, [("type", "1"), ("keys", "enter"), ("type", "2"), ("keys", "enter"), ("keys", "ctrl+z"), ("keys", "ctrl+z")], "keys3"),
    C("Delete clears a range", {"A1": "1", "A2": "2", "B1": "x"}, [("keys", "shift+down"), ("keys", "shift+right"), ("keys", "delete")], "keys3"),
    C("Ctrl+Up from below block", {"A1": "1", "A2": "2", "A3": "3"}, [("keys", "ctrl+up")], "keys3", start="A10"),
    C("Ctrl+Left in a gap row", {"A1": "1", "D1": "4"}, [("keys", "ctrl+left"), ("keys", "ctrl+left")], "keys3", start="F1"),
    C("Enter after Shift+Space keeps row", {}, [("keys", "shift+space"), ("keys", "enter"), ("keys", "enter")], "keys3"),
    C("Strg+Leertaste then type fills nothing", {}, [("keys", "ctrl+space"), ("type", "x"), ("keys", "enter")], "keys3"),
    C("undo after paste", {"A1": "5"}, [("keys", "ctrl+c"), ("keys", "down"), ("keys", "ctrl+v"), ("keys", "esc"), ("keys", "ctrl+z")], "keys3"),
    C("Alt+Enter twice", {}, [("type", "a"), ("keys", "alt+enter"), ("keys", "alt+enter"), ("type", "b"), ("keys", "enter")], "keys3"),
    C("Umschalt+drag inserts between rows", {"A1": "1", "A2": "2", "A3": "3", "A4": "4", "A5": "5"}, [("border", "A1", "A3@bottom", "shift")], "move4"),
    C("Umschalt+drag inserts between columns", {"A1": "1", "B1": "2", "C1": "3"}, [("border", "A1", "B1@right", "shift")], "move4"),
    C("Umschalt+drag upwards", {"A1": "1", "A2": "2", "A3": "3", "A4": "4"}, [("border", "A4", "A1@bottom", "shift")], "move4", start="A4"),
    C("Umschalt+drag to cell centre", {"A1": "1", "B1": "2", "C1": "3", "B2": "x"}, [("border", "A1", "B1@center", "shift")], "move5"),
    C("Umschalt+drag far right", {"A1": "1", "B1": "2", "C1": "3", "D1": "4"}, [("border", "A1", "C1@center", "shift")], "move5"),
    C("Umschalt+drag down a column", {"A1": "1", "B1": "2", "A2": "3", "B2": "4", "A3": "5"}, [("border", "A1", "A2@center", "shift")], "move5"),
    C("Umschalt+drag column block right", {"A1": "1", "A2": "2", "B1": "3", "B2": "4", "C1": "5", "C2": "6"}, [("keys", "shift+down"), ("border", "A1", "B1@br", "shift")], "move5"),
    C("Strg+E Blitzvorschau", {"A1": "Max Mustermann", "A2": "Erika Musterfrau", "A3": "Jonas Weber", "A4": "Ali Yilmaz", "B1": "Mustermann"}, [("keys", "ctrl+e")], "flash", start="B2"),
    C("Strg+E with header", {"A1": "Name", "B1": "Vorname", "A2": "Max Mustermann", "A3": "Erika Musterfrau", "A4": "Jonas Weber", "B2": "Max"}, [("keys", "ctrl+e")], "flash", start="B3"),
    C("Strg+Umschalt+L filter", {"A1": "Name", "A2": "x", "A3": "y"}, [("keys", "ctrl+shift+l")], "sheets"),
]


def main() -> None:
    which = sys.argv[1] if len(sys.argv) > 1 else "excel"
    pick = {a for a in sys.argv[2:] if a != "--from"}
    if "--from" in sys.argv:
        pick.discard(sys.argv[sys.argv.index("--from") + 1])
    cases = [c for c in CASES if not pick or c.area in pick or c.name in pick]
    if "--from" in sys.argv:
        start = sys.argv[sys.argv.index("--from") + 1]
        cases = cases[[c.name for c in cases].index(start):]
        pick.discard("--from"); pick.discard(start)
    out = Path(__file__).parent / "results"
    date = dt.date.today().isoformat()
    if which == "excel":
        from feellab import winapi
        from feellab.calc_catalogue import ExcelSide
        side = ExcelSide()
        results = {}
        try:
            for c in cases:
                side.reset(c)
                side.prepare(c)
                settle()
                winapi.activate(side.x.hwnd)
                time.sleep(0.25)
                for s in c.steps:
                    run_step(side, s)
                time.sleep(0.3)
                results[c.name] = side.state(c)
                print(f"{c.area:10} {c.name:45} {results[c.name]['selection']:12} {results[c.name]['mode']}")
        finally:
            side.close()
        path = out / f"{date}-catalogue-calc-excel.json"
        old = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
        old.update(results)
        path.write_text(json.dumps(old, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        print(f"{len(results)} cases recorded -> {path.name}")
    else:
        from feellab.calc_catalogue import compare_calc
        compare_calc(cases, out, date)


if __name__ == "__main__":
    main()
