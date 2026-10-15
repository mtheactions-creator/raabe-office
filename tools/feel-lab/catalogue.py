"""Behaviour catalogue against Word (D-038): runs every case with real keyboard and mouse in
Word and in RaabeWrite and compares the documents (read back by Word) and the caret. Takes
over keyboard, mouse and screen for the whole run (about 15 minutes).

    .venv/Scripts/python.exe catalogue.py [area or name …]

Results: results/<date>-catalogue.json and a summary on the console.
"""

from __future__ import annotations

import datetime as dt
import json
import sys
from pathlib import Path

from feellab.catalogue import Case, Find, run

F = Find
C = Case


def char(f): return ("char", f)
def caret(f): return ("caret", f)
def margin(f): return ("margin", f)


FILL_START = "Pflanzen nehmen"

CASES: list[Case] = [
    # ----- typing, AutoCorrect ----------------------------------------------------------------
    C("type plain text", "empty", [("type", "Hallo Welt")], "typing"),
    C("German quotation marks", "empty", [("type", 'Er sagte "Hallo" zu mir.')], "typing"),
    C("single quotes", "empty", [("type", "Er sagte 'ja' dazu.")], "typing"),
    C("capitalise sentence start", "empty", [("type", "der test. das ist gut")], "typing"),
    C("capitalise first word on Enter", "empty", [("type", "hallo welt"), ("keys", "enter"), ("type", "x")], "typing"),
    C("two initial capitals", "empty", [("type", "ZWei Wörter ")], "typing"),
    C("copyright sign", "empty", [("type", "(c) 2026 ")], "typing"),
    C("fraction one half", "empty", [("type", "1/2 Liter ")], "typing"),
    C("en dash between words", "empty", [("type", "Haus - Garten ")], "typing"),
    C("em dash from two hyphens", "empty", [("type", "Haus--Garten ")], "typing"),
    C("arrow from -->", "empty", [("type", "A --> B ")], "typing"),
    C("*bold* AutoFormat", "empty", [("type", "Das ist *wichtig* hier")], "typing"),
    C("_italic_ AutoFormat", "empty", [("type", "Das ist _wichtig_ hier")], "typing"),
    C("ordinal 1. not a list mid line", "empty", [("type", "Am 1. Mai")], "typing"),
    C("apostrophe in word", "empty", [("type", "Wie geht's dir")], "typing"),
    C("smiley", "empty", [("type", "Gut :) ")], "typing"),
    C("copyright then Enter", "empty", [("type", "(c)"), ("keys", "enter")], "typing"),
    C("ellipsis after word", "empty", [("type", "Und dann... ")], "typing"),
    C("weekday capitalised", "empty", [("type", "Am montag ")], "typing"),
    C("caps lock corrected", "empty", [("type", "Ein hALLO ")], "typing"),
    C("abbreviation keeps lower case", "empty", [("type", "Wir nehmen z.B. das ")], "typing"),
    C("ordinal keeps lower case", "empty", [("type", "Am 1. tag ")], "typing"),
    C("typo corrected", "empty", [("type", "Und dei Katze ")], "typing"),
    C("border line from ---", "empty", [("type", "---"), ("keys", "enter"), ("type", "Text")], "typing"),
    C("double border from ===", "empty", [("type", "==="), ("keys", "enter")], "typing"),
    C("Tab at paragraph start indents first line", "plain", [("caret", F("Das Wasser")), ("keys", "tab")], "typing"),
    C("Tab twice at paragraph start", "plain", [("caret", F("Das Wasser")), ("keys", "tab", 2)], "typing"),
    C("Tab in empty paragraph types a tab", "empty", [("keys", "tab"), ("type", "x")], "typing"),
    C("capital in table cell", "table", [("sel", F("Z2S2"), F("Z2S2", 4)), ("type", "neu ")], "typing"),
    C("typing over selection", "plain", [("sel", F("Pflanzen"), F("Pflanzen", 8)), ("type", "Bäume")], "typing"),
    C("typing after double click keeps space", "plain", [("click", char(F("Pflanzen", 3)), 2), ("type", "Bäume")], "typing"),
    C("(c) without space", "empty", [("type", "Text(c)")], "typing"),
    C("--> then text", "empty", [("type", "A -->B")], "typing"),
    C("... then letter", "empty", [("type", "Und...so")], "typing"),
    C(":) then text", "empty", [("type", "Gut :) ja")], "typing"),
    C("apostrophe at word end", "empty", [("type", "Peters' Haus")], "typing"),
    C("capital after ordinal date", "empty", [("type", "Am 1. januar ")], "typing"),
    C("**bold** Markdown", "empty", [("type", "Ein **fettes** Wort")], "typing"),
    C("~~strike~~ Markdown", "empty", [("type", "Ein ~~altes~~ Wort")], "typing"),
    C("# heading Markdown", "empty", [("type", "# Titel"), ("keys", "enter"), ("type", "Text")], "typing"),
    C("### heading Markdown", "empty", [("type", "### Titel")], "typing"),
    C("##### heading Markdown", "empty", [("type", "##### Titel")], "typing"),
    # ----- AutoFormat lists ---------------------------------------------------------------------
    C("autolist 1.", "empty", [("type", "1. Erster"), ("keys", "enter"), ("type", "Zweiter")], "autolist"),
    C("autolist dash", "empty", [("type", "- Erster"), ("keys", "enter"), ("type", "Zweiter")], "autolist"),
    C("autolist asterisk", "empty", [("type", "* Erster"), ("keys", "enter"), ("type", "Zweiter")], "autolist"),
    C("autolist a)", "empty", [("type", "a) Erster"), ("keys", "enter"), ("type", "Zweiter")], "autolist"),
    C("autolist 1)", "empty", [("type", "1) Erster"), ("keys", "enter"), ("type", "Zweiter")], "autolist"),
    C("autolist I.", "empty", [("type", "I. Erster"), ("keys", "enter"), ("type", "Zweiter")], "autolist"),
    C("autolist >", "empty", [("type", "> Erster"), ("keys", "enter"), ("type", "Zweiter")], "autolist"),
    C("autolist tab instead of space", "empty", [("type", "1."), ("keys", "tab"), ("type", "Erster")], "autolist"),
    C("autolist ends with two Enters", "empty", [("type", "1. Erster"), ("keys", "enter"), ("keys", "enter"), ("type", "Text")], "autolist"),
    C("autolist Tab makes level 2", "empty", [("type", "1. Erster"), ("keys", "enter"), ("keys", "tab"), ("type", "Unterpunkt")], "autolist"),
    C("autolist undo keeps text", "empty", [("type", "1. "), ("keys", "ctrl+z"), ("type", "x")], "autolist"),
    C("autolist (1)", "empty", [("type", "(1) Erster"), ("keys", "enter"), ("type", "Zweiter")], "autolist"),
    C("autolist 1-", "empty", [("type", "1- Erster")], "autolist"),
    C("autolist ->", "empty", [("type", "-> Erster")], "autolist"),
    C("autolist =>", "empty", [("type", "=> Erster")], "autolist"),
    C("autolist en dash", "empty", [("type", "\u2013 Erster")], "autolist"),
    C("autolist i)", "empty", [("type", "i) Erster")], "autolist"),
    C("autolist 2. stays text", "empty", [("type", "2. Erster")], "autolist"),
    C("autolist c) stays text", "empty", [("type", "c) Erster")], "autolist"),
    C("autolist * then Tab", "empty", [("type", "*"), ("keys", "tab"), ("type", "Erster")], "autolist"),
    C("autolist a) capital after", "empty", [("type", "a) erster ")], "autolist"),
    # ----- keyboard selection and movement -------------------------------------------------------
    C("shift+right x3", "plain", [("caret", F(FILL_START)), ("keys", "shift+right", 3)], "keys"),
    C("ctrl+shift+right x2", "plain", [("caret", F(FILL_START)), ("keys", "ctrl+shift+right", 2)], "keys"),
    C("ctrl+shift+left from mid word", "plain", [("caret", F("Kohlenstoffdioxid", 5)), ("keys", "ctrl+shift+left")], "keys"),
    C("shift+end", "plain", [("caret", F("Blätter")), ("keys", "shift+end")], "keys"),
    C("shift+home", "plain", [("caret", F("Blätter", 3)), ("keys", "shift+home")], "keys"),
    C("shift+down", "plain", [("caret", F("Blätter")), ("keys", "shift+down")], "keys"),
    C("shift+up", "plain", [("caret", F("Traubenzucker")), ("keys", "shift+up")], "keys"),
    C("ctrl+shift+end", "plain", [("caret", F("Wurzeln auf")), ("keys", "ctrl+shift+end")], "keys"),
    C("ctrl+shift+home", "plain", [("caret", F("Blätter")), ("keys", "ctrl+shift+home")], "keys"),
    C("ctrl+shift+down", "plain", [("caret", F("Blätter")), ("keys", "ctrl+shift+down")], "keys"),
    C("ctrl+right x5", "plain", [("caret", F(FILL_START)), ("keys", "ctrl+right", 5)], "keys"),
    C("ctrl+left x3", "plain", [("caret", F("Boden.")), ("keys", "ctrl+left", 3)], "keys"),
    C("ctrl+right over period", "plain", [("caret", F("Boden.")), ("keys", "ctrl+right", 2)], "keys"),
    C("home end", "plain", [("caret", F("Blätter", 2)), ("keys", "end"), ("type", "X")], "keys"),
    C("ctrl+home", "plain", [("caret", F("Blätter")), ("keys", "ctrl+home"), ("type", "X")], "keys"),
    C("ctrl+end", "plain", [("caret", F("Blätter")), ("keys", "ctrl+end"), ("type", "X")], "keys"),
    C("down x2 keeps column", "plain", [("caret", F("Kohlenstoffdioxid", 4)), ("keys", "down", 2), ("type", "X")], "keys"),
    C("up from first line of paragraph", "plain", [("caret", F("Das Wasser", 6)), ("keys", "up"), ("type", "X")], "keys"),
    C("ctrl+down", "plain", [("caret", F("Blätter")), ("keys", "ctrl+down"), ("type", "X")], "keys"),
    C("ctrl+up mid paragraph", "plain", [("caret", F("Blätter")), ("keys", "ctrl+up"), ("type", "X")], "keys"),
    C("right collapses selection", "plain", [("sel", F("Blätter"), F("Blätter", 4)), ("keys", "right"), ("type", "X")], "keys"),
    C("left collapses selection", "plain", [("sel", F("Blätter"), F("Blätter", 4)), ("keys", "left"), ("type", "X")], "keys"),
    C("ctrl+a", "plain", [("caret", F("Blätter")), ("keys", "ctrl+a")], "keys"),
    C("shift+f3 cycles case", "plain", [("sel", F("Blätter"), F("Blätter", 7)), ("keys", "shift+f3")], "keys"),
    C("shift+f3 twice", "plain", [("sel", F("Blätter"), F("Blätter", 7)), ("keys", "shift+f3", 2)], "keys"),
    # ----- mouse ----------------------------------------------------------------------------
    C("double click word", "plain", [("click", char(F("Kohlenstoffdioxid", 4)), 2)], "mouse"),
    C("double click word before period", "plain", [("click", char(F("Boden.", 2)), 2)], "mouse"),
    # Word needs a moment after opening a document it opened before ("Willkommen zurück").
    C("triple click paragraph", "plain", [("wait", 1.5), ("click", char(F("Blätter", 2)), 3)], "mouse"),
    C("ctrl+click sentence", "plain", [("click", char(F("Energie", 2)), 1, "ctrl")], "mouse"),
    C("click left margin selects line", "plain", [("click", margin(F(FILL_START)))], "mouse"),
    C("double click left margin selects paragraph", "plain", [("click", margin(F(FILL_START)), 2)], "mouse"),
    C("triple click left margin selects all", "plain", [("click", margin(F(FILL_START)), 3)], "mouse"),
    C("shift+click extends", "plain", [("click", char(F("Blätter"))), ("click", caret(F("Wasser")), 1, "shift")], "mouse"),
    C("drag across words snaps to words", "plain", [("drag", caret(F("Blätter", 3)), caret(F("Kohlenstoffdioxid", 5)))], "mouse"),
    C("drag inside one word", "plain", [("drag", caret(F("Kohlenstoffdioxid", 2)), caret(F("Kohlenstoffdioxid", 8)))], "mouse"),
    C("drag backwards", "plain", [("drag", caret(F("Luft", 2)), caret(F("Blätter", 3)))], "mouse"),
    C("click then type", "plain", [("click", caret(F("Blätter"))), ("type", "grünen ")], "mouse"),
    # ----- deleting ---------------------------------------------------------------------------
    C("backspace at start of paragraph", "plain", [("caret", F("Das Wasser")), ("keys", "backspace")], "delete"),
    C("backspace after heading joins", "plain", [("caret", F(FILL_START)), ("keys", "backspace")], "delete"),
    C("delete at end of heading", "plain", [("caret", F("Photosynthese", 13)), ("keys", "delete")], "delete"),
    C("ctrl+backspace mid word", "plain", [("caret", F("Kohlenstoffdioxid", 6)), ("keys", "ctrl+backspace")], "delete"),
    C("ctrl+backspace after word", "plain", [("caret", F("Blätter Kohlen")), ("keys", "ctrl+backspace")], "delete"),
    C("ctrl+delete mid word", "plain", [("caret", F("Kohlenstoffdioxid", 6)), ("keys", "ctrl+delete")], "delete"),
    C("ctrl+delete at word start", "plain", [("caret", F("Kohlenstoffdioxid")), ("keys", "ctrl+delete")], "delete"),
    C("delete double-clicked word", "plain", [("click", char(F("Kohlenstoffdioxid", 4)), 2), ("keys", "delete")], "delete"),
    C("backspace on selection", "plain", [("sel", F("aus der Luft"), F("aus der Luft", 12)), ("keys", "backspace")], "delete"),
    # ----- lists ---------------------------------------------------------------------------------
    C("list Enter adds item", "lists", [("caret", F("Becherglas", 10)), ("keys", "enter"), ("type", "Löffel")], "lists"),
    C("list Enter splits item", "lists", [("caret", F("Salz zugeben", 4)), ("keys", "enter")], "lists"),
    C("list Enter at item start", "lists", [("caret", F("Salz zugeben")), ("keys", "enter")], "lists"),
    C("list Enter on empty item ends list", "lists", [("caret", F("Umrühren", 8)), ("keys", "enter"), ("keys", "enter"), ("type", "Danach")], "lists"),
    C("list backspace removes bullet", "lists", [("caret", F("Becherglas")), ("keys", "backspace")], "lists"),
    C("list backspace twice", "lists", [("caret", F("Becherglas")), ("keys", "backspace", 2)], "lists"),
    C("list backspace three times", "lists", [("caret", F("Becherglas")), ("keys", "backspace", 3)], "lists"),
    C("list Tab demotes", "lists", [("caret", F("Salz")), ("keys", "tab")], "lists"),
    C("list shift+Tab promotes", "lists", [("caret", F("Leitungswasser")), ("keys", "shift+tab")], "lists"),
    C("list Tab at first item", "lists", [("caret", F("Becherglas")), ("keys", "tab")], "lists"),
    C("list Tab mid item types tab", "lists", [("caret", F("Becherglas", 4)), ("keys", "tab")], "lists"),
    C("list Tab at first item twice", "lists", [("caret", F("Becherglas")), ("keys", "tab", 2)], "lists"),
    C("list shift+Tab at first item", "lists", [("caret", F("Becherglas")), ("keys", "shift+tab")], "lists"),
    C("list backspace twice on second item", "lists", [("caret", F("Wasser")), ("keys", "backspace", 2)], "lists"),
    C("list backspace three times on second item", "lists", [("caret", F("Wasser")), ("keys", "backspace", 3)], "lists"),
    C("indent backspace", "plain", [("caret", F("Das Wasser")), ("keys", "ctrl+m", 2), ("keys", "backspace")], "lists"),
    C("first-line indent backspace", "plain", [("caret", F("Das Wasser")), ("keys", "tab", 2), ("keys", "backspace")], "lists"),
    C("hanging indent backspace", "plain", [("caret", F("Das Wasser")), ("keys", "ctrl+t"), ("keys", "backspace")], "lists"),
    C("numbered list Enter continues", "lists", [("caret", F("Umrühren", 8)), ("keys", "enter"), ("type", "Warten")], "lists"),
    C("ctrl+shift+n removes list", "lists", [("caret", F("Salz zugeben", 2)), ("keys", "ctrl+shift+n")], "lists"),
    C("delete at end of last item joins", "lists", [("caret", F("Umrühren", 8)), ("keys", "delete")], "lists"),
    # ----- tables ------------------------------------------------------------------------------------
    C("table Tab to next cell", "table", [("caret", F("Z1S1", 2)), ("keys", "tab"), ("type", "X")], "tables"),
    C("table shift+Tab", "table", [("caret", F("Z1S2", 2)), ("keys", "shift+tab"), ("type", "X")], "tables"),
    C("table Tab in last cell adds row", "table", [("caret", F("Z3S3", 4)), ("keys", "tab"), ("type", "Neu")], "tables"),
    C("table Enter in cell", "table", [("caret", F("Z2S2", 4)), ("keys", "enter"), ("type", "X")], "tables"),
    C("table backspace at cell start", "table", [("caret", F("Z2S1")), ("keys", "backspace"), ("type", "X")], "tables"),
    C("table down from last row", "table", [("caret", F("Z3S2", 2)), ("keys", "down"), ("type", "X")], "tables"),
    C("table up into last row", "table", [("caret", F("Nach der Tabelle", 3)), ("keys", "up"), ("type", "X")], "tables"),
    C("table right at cell end", "table", [("caret", F("Z1S1", 4)), ("keys", "right"), ("type", "X")], "tables"),
    C("table delete before table", "table", [("caret", F("Vor der Tabelle", 15)), ("keys", "delete")], "tables"),
    C("table delete cell text", "table", [("sel", F("Z2S2"), F("Z2S2", 4)), ("keys", "delete")], "tables"),
    C("table select row and delete", "table", [("sel", F("Z2S1"), F("Z2S3", 4)), ("keys", "delete")], "tables"),
    C("table shift+right across cells", "table", [("caret", F("Z1S1", 2)), ("keys", "shift+right", 4)], "tables"),
    C("table ctrl+a in cell", "table", [("caret", F("Z1S1", 2)), ("keys", "ctrl+a")], "tables"),
    C("table shift+right to cell end", "table", [("caret", F("Z1S1", 2)), ("keys", "shift+right", 3)], "tables"),
    C("table shift+right then back", "table", [("caret", F("Z1S1", 2)), ("keys", "shift+right", 4), ("keys", "shift+left")], "tables"),
    C("table shift+left across cells", "table", [("caret", F("Z1S2", 2)), ("keys", "shift+left", 4)], "tables"),
    C("table shift+down", "table", [("caret", F("Z1S2", 2)), ("keys", "shift+down")], "tables"),
    # ----- Step 4c: tables, selection by mouse (D-056, D-058) ----------------------------------------
    C("table ctrl+shift+enter splits the table", "table", [("caret", F("Z2S1")), ("keys", "ctrl+shift+enter"), ("type", "X")], "tables"),
    C("table ctrl+shift+enter in first row", "table", [("caret", F("Z1S2", 2)), ("keys", "ctrl+shift+enter"), ("type", "X")], "tables"),
    C("table backspace on cell content", "table", [("sel", F("Z2S2"), F("Z2S2", 4)), ("keys", "backspace"), ("type", "X")], "tables"),
    C("table delete on cell content then type", "table", [("sel", F("Z3S1"), F("Z3S1", 4)), ("keys", "delete"), ("type", "Y")], "tables"),
    C("table double-click word in cell", "table", [("click", char(F("Z2S3", 1)), 2), ("type", "W")], "tables"),
    C("margin double-click selects paragraph", "plain", [("click", margin(F("Das Wasser")), 2), ("type", "X")], "selection"),
    C("margin click selects line", "plain", [("click", margin(F("Letzter"))), ("type", "X")], "selection"),
    C("margin triple-click selects all", "plain", [("click", margin(F("Letzter")), 3), ("type", "X")], "selection"),
    C("double-click word then type", "plain", [("click", char(F("Wasser", 2)), 2), ("type", "Saft")], "selection"),
    C("triple-click paragraph then type", "plain", [("click", char(F("Wasser", 2)), 3), ("type", "Neu")], "selection"),
    # ----- headings and styles -----------------------------------------------------------------------
    C("Enter at end of heading", "plain", [("caret", F("Photosynthese", 13)), ("keys", "enter"), ("type", "Text")], "styles"),
    C("Enter at start of heading", "plain", [("caret", F("Photosynthese")), ("keys", "enter"), ("type", "X")], "styles"),
    C("Enter in middle of heading", "plain", [("caret", F("Photosynthese", 5)), ("keys", "enter")], "styles"),
    C("ctrl+alt+1 heading 1", "plain", [("caret", F("Letzter")), ("keys", "ctrl+alt+1")], "styles"),
    C("ctrl+alt+2 heading 2", "plain", [("caret", F("Letzter")), ("keys", "ctrl+alt+2")], "styles"),
    C("ctrl+shift+n on heading", "plain", [("caret", F("Photosynthese", 3)), ("keys", "ctrl+shift+n")], "styles"),
    C("ctrl+shift+l bullet style", "plain", [("caret", F("Letzter")), ("keys", "ctrl+shift+l")], "styles"),
    # ----- breaks -------------------------------------------------------------------------------------
    C("page break mid paragraph", "plain", [("caret", F("Das Wasser", 4)), ("keys", "ctrl+enter"), ("type", "X")], "breaks"),
    C("page break at paragraph end", "plain", [("caret", F("Letzter Absatz.", 15)), ("keys", "ctrl+enter"), ("type", "X")], "breaks"),
    C("page break at paragraph start", "plain", [("caret", F("Das Wasser")), ("keys", "ctrl+enter"), ("type", "X")], "breaks"),
    C("line break", "plain", [("caret", F("Das Wasser", 4)), ("keys", "shift+enter"), ("type", "X")], "breaks"),
    C("column break", "plain", [("caret", F("Das Wasser", 4)), ("keys", "ctrl+shift+enter"), ("type", "X")], "breaks"),
    C("backspace after page break", "plain", [("caret", F("Das Wasser", 4)), ("keys", "ctrl+enter"), ("keys", "backspace")], "breaks"),
    # ----- formatting shortcuts ---------------------------------------------------------------------
    C("bold inside word applies to word", "plain", [("caret", F("Kohlenstoffdioxid", 4)), ("keys", "ctrl+shift+f")], "format"),
    C("bold at word end then type", "plain", [("caret", F("Blätter", 7)), ("keys", "ctrl+shift+f"), ("type", "X")], "format"),
    C("italic inside word", "plain", [("caret", F("Kohlenstoffdioxid", 4)), ("keys", "ctrl+shift+k")], "format"),
    C("underline selection", "format", [("sel", F("kursives Wort"), F("kursives Wort", 13)), ("keys", "ctrl+shift+u")], "format"),
    C("toggle bold off", "format", [("click", char(F("fettes", 2)), 2), ("keys", "ctrl+shift+f")], "format"),
    C("ctrl+space resets word", "format", [("caret", F("fettes", 2)), ("keys", "ctrl+space")], "format"),
    C("ctrl+e centre", "plain", [("caret", F("Blätter")), ("keys", "ctrl+e")], "format"),
    C("ctrl+r right", "plain", [("caret", F("Blätter")), ("keys", "ctrl+r")], "format"),
    C("ctrl+b justify", "plain", [("caret", F("Blätter")), ("keys", "ctrl+b")], "format"),
    C("ctrl+m indent", "plain", [("caret", F("Blätter")), ("keys", "ctrl+m")], "format"),
    C("ctrl+m twice, ctrl+shift+m", "plain", [("caret", F("Blätter")), ("keys", "ctrl+m", 2), ("keys", "ctrl+shift+m")], "format"),
    C("ctrl+t hanging indent", "plain", [("caret", F("Blätter")), ("keys", "ctrl+t")], "format"),
    C("ctrl+q resets paragraph", "plain", [("caret", F("Blätter")), ("keys", "ctrl+e"), ("keys", "ctrl+m"), ("keys", "ctrl+q")], "format"),
    C("ctrl+2 double spacing", "plain", [("caret", F("Blätter")), ("keys", "ctrl+2")], "format"),
    C("ctrl+5 1.5 spacing", "plain", [("caret", F("Blätter")), ("keys", "ctrl+5")], "format"),
    C("grow font", "plain", [("sel", F("Blätter"), F("Blätter", 7)), ("keys", "ctrl+shift+lt")], "format"),
    C("shrink font", "plain", [("sel", F("Blätter"), F("Blätter", 7)), ("keys", "ctrl+lt")], "format"),
    C("subscript ctrl+#", "plain", [("sel", F("Blätter"), F("Blätter", 7)), ("keys", "ctrl+hash")], "format"),
    C("superscript ctrl+plus", "plain", [("sel", F("Blätter"), F("Blätter", 7)), ("keys", "ctrl+plus")], "format"),
    C("caps ctrl+shift+g", "plain", [("sel", F("Blätter"), F("Blätter", 7)), ("keys", "ctrl+shift+g")], "format"),
    # ----- undo --------------------------------------------------------------------------------------
    C("undo typing two words", "plain", [("caret", F("Blätter")), ("type", "rote Algen "), ("keys", "ctrl+z")], "undo"),
    C("undo typing three words twice", "plain", [("caret", F("Blätter")), ("type", "eins zwei drei "), ("keys", "ctrl+z", 2)], "undo"),
    C("undo AutoCorrect", "empty", [("type", "(c) "), ("keys", "ctrl+z")], "undo"),
    C("undo quotes", "empty", [("type", 'Er "'), ("keys", "ctrl+z")], "undo"),
    C("undo delete word", "plain", [("click", char(F("Kohlenstoffdioxid", 4)), 2), ("keys", "delete"), ("keys", "ctrl+z")], "undo"),
    C("redo after undo", "plain", [("caret", F("Blätter")), ("type", "rote "), ("keys", "ctrl+z"), ("keys", "ctrl+y")], "undo"),
    C("undo formatting", "plain", [("sel", F("Blätter"), F("Blätter", 7)), ("keys", "ctrl+shift+f"), ("keys", "ctrl+z")], "undo"),
    C("undo Enter", "plain", [("caret", F("Blätter")), ("keys", "enter"), ("keys", "ctrl+z")], "undo"),
    # ----- clipboard -----------------------------------------------------------------------------------
    C("copy and paste a word", "plain", [("click", char(F("Kohlenstoffdioxid", 4)), 2), ("keys", "ctrl+c"), ("caret", F("Wasser nimmt")), ("keys", "ctrl+v")], "clipboard"),
    C("cut and paste with smart spaces", "plain", [("click", char(F("Kohlenstoffdioxid", 4)), 2), ("keys", "ctrl+x"), ("caret", F("Traubenzucker")), ("keys", "ctrl+v")], "clipboard"),
    C("cut and paste keeps bold", "format", [("click", char(F("fettes", 2)), 2), ("keys", "ctrl+x"), ("caret", F("Wort.")), ("keys", "ctrl+v")], "clipboard"),
    C("paste over selection", "plain", [("click", char(F("Luft", 2)), 2), ("keys", "ctrl+c"), ("click", char(F("Boden", 2)), 2), ("keys", "ctrl+v")], "clipboard"),
    C("paste paragraph", "plain", [("click", margin(F("Letzter")), 2), ("keys", "ctrl+c"), ("caret", F("Photosynthese")), ("keys", "ctrl+v")], "clipboard"),
]


def main() -> None:
    out = Path(__file__).parent / "results" / f"{dt.date.today().isoformat()}-catalogue.json"
    if "--recompare" in sys.argv:
        from feellab.catalogue import recompare
        data = recompare(json.loads(out.read_text(encoding="utf-8"))["results"])
    else:
        only = [a for a in sys.argv[1:] if not a.startswith("--")] or None
        data = run(CASES, only)
        if only and out.exists():
            # A partial run updates its cases in the last full result.
            prev = json.loads(out.read_text(encoding="utf-8"))["results"]["cases"]
            prev.update(data["cases"])
            from feellab.catalogue import _by_area
            data = {"summary": f"{sum(1 for r in prev.values() if r.get('ok'))}/{len(prev)}", "by_area": _by_area(prev), "cases": prev}
    out.write_text(json.dumps({"meta": {"date": dt.datetime.now().isoformat(timespec="seconds")}, "results": data},
                              ensure_ascii=False, indent=1), encoding="utf-8")
    print("\n", data["summary"], data["by_area"])
    for name, r in data["cases"].items():
        if not r.get("ok"):
            what = r.get("error") or ", ".join(k for k in ("doc_ok", "caret_ok", "sel_ok") if not r.get(k))
            print(f"FAIL {name}: {what}")
    print("saved to", out)


if __name__ == "__main__":
    main()
