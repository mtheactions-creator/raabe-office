# Step 4c report: polish after trying Step 4b

Date: 2026-10-15 · Measured with `tools/feel-lab` (see `docs/FEEL.md`) · Raw data in
`tools/feel-lab/results/` (2026-10-15)

Step 4c (D-053) took the owner's list of differences found while trying Step 4b (pictures,
formulas, tables, cursors, selection, Formatvorlagen) and, in package F, looked for the same
kinds of difference across the whole app. As in 4b, Word's behaviour was recorded first in Word
instances of our own, then built and measured against it. Two owner decisions shaped it: a
Formatvorlage resets direct formatting cleanly, and tables can be moved freely with a second
handle of our own for wrapping and alignment; our look may differ from Word's, but everything
sits where a Word user looks for it.

## What was built

- **A. Bugs:** Zuschneiden drew the whole picture over the crop frame (it looked like a
  duplicate); the table handle selected only part of a table whose last cell was empty; only the
  first Tabellenformatvorlage worked; Bildformatvorlagen previews overlapped; Formatvorlagen
  names were cut at both sides.
- **B. Pictures and panes (D-054):** text on both sides of Quadrat and Eng pictures, Eng along
  Word's wrap polygon, snapping with Word's green alignment guides, a double-click opening the
  object's tab, Word's layout-options icons and tooltips, Transparente Farbe bestimmen, and a
  task pane in Word's layout (Grafik/Form formatieren, Texteffekte formatieren) instead of
  centred dialogs; turned selection frames and Word's context menu for shapes.
- **C. Formulas (D-055):** the formula's font in Start, "/" and space taking the whole operand
  (decimal commas included), quoted text becoming normal text, the caret around quotes, the
  handle selecting the whole formula, and equations through the clipboard with Word both ways.
- **D. Tables (D-056):** Word's drag thresholds and cell shading, the full Rahmen group
  (Rahmenarten, Stiftart/-stärke/-farbe, Rahmen übertragen, Rahmenpipette), all 105 of Word's
  table styles with previews and Livevorschau, Entf and Rücktaste on cells with Word's Zellen
  löschen dialog, and floating tables moved freely by the handle plus our layout handle.
- **E. Selection, pointers, Formatvorlagen (D-057):** a selection visible over highlighting,
  shading and coloured text, sharp pointers at every scaling, Livevorschau in the Formatvorlagen
  gallery, Formatvorlagen that reset direct formatting.
- **F. Audit of the whole app (D-058):**
  - *Tooltips* with Word's titles and shortcuts on every ribbon command, described in our own
    words (Microsoft's texts are not copied).
  - *Labels*: none cut off on any tab at widths from 1920 to 800 px (`audit_labels.py`).
  - *Menus and galleries* against Word's recordings, entry by entry (`audit_menus.py`, 19 menus
    and 7 context menus): no unexplained difference left; fixed on the way: Word's tint and
    shade rows, the colour names of the current Office theme, theme and font scheme names.
  - *Commands that did nothing* found and built: the whole Tabellenlayout tab with Word's labels
    and keytips (Tabelle teilen, Zellen einfügen, AutoAnpassen, Zellenbegrenzungen, In Text
    konvertieren), Als Grafik speichern, Neues Aufzählungszeichen/Zahlenformat definieren with
    Word's 13 number styles (German number words, Word's 001 formats, read by Word as we show
    them).
  - *Livevorschau* wherever Word has it: one mechanism for galleries, menus, the font lists and
    every colour palette.
  - *Double-clicks* as recorded in Word: notes, column borders, the move handle, section breaks,
    and the table selection zones (cell, row, column) with Word's black arrows.
  - *Panes*: Alternativtext as a task pane.
  - *Clipboard*: pictures and hyperlinks with Word both ways; Kopieren on a selected picture or
    shape; a hang after copying a picture in Word found and removed.
  - *Behaviour catalogue*: ten new cases, which found Word's intelligent paragraph selection
    (typing over a selected paragraph keeps its mark).

## Results

### Layout (F4b): Word-built test documents

| Documents | Line starts | Page breaks |
|---|---|---|
| the nineteen documents of Step 4b (fixtures re-made anonymized) | 963/968 | every one but one after the hyphenation differences |
| word-wrap (B: Quadrat, Eng, turned and cut-out pictures) | 164/164 | every line on Word's page (after the fix below) |
| word-table-float, word-table-select (D) | 76/76, 13/13 | every one |

**1216 of 1221 line starts identical to Word** in 22 documents; the five others are the known
hyphenation choices of D-050. word-table-styles (D) cannot be compared by text, because Word's
COM reports the header rows' "Kopf" in capitals; it was checked against Word's screenshots.
Floating tables are placed within 0.3 pt of Word.

### Behaviour with real input (F3)

The behaviour catalogue (`catalogue.py`, real keyboard and mouse in Word and RaabeWrite, 203
cases, 10 of them new): **202 of 203 as in Word**: typing and AutoCorrect 42/42, AutoFormat
lists 21/21, keys 25/25, mouse 11/12, delete 9/9, lists 21/21, tables 22/22, selection 5/5,
styles 7/7, breaks 6/6, formatting 20/20, undo 8/8, clipboard 5/5. The one difference is again
Word's: its own triple-click did not register in the harness (caret at the document start),
while the triple-clicks of the selection cases work in both.

The formula editor (`catalogue_math.py`): **146 of 146** recorded Word cases (114 from 4b and 32
new ones from C). The clipboard with Word: equations 0 differences both ways (C), lists and
tables 0 differences both ways (4b), pictures and links as described above
(`clipboard_objects.py`).

### Smoke tests in the real app (Word reads what RaabeWrite writes)

`smoke_write.py` 41/41, `smoke_step3.py` 28/28, `smoke_step4.py` 43/43, `smoke_languages.py`
25/25. Unit tests: 139 (TypeScript).

### Typing latency and memory (F1, F7)

| | RaabeWrite | Word | Step 4b (RaabeWrite / Word) |
|---|---|---|---|
| F1 typing, short document, median (p95) | 33.5 ms (64.7) | 36.3 ms (62.1) | 31.2 / 33.1 ms |
| F1 typing, long document, median (p95) | 31.8 ms (65.1) | 35.5 ms (63.3) | 32.2 / 36.0 ms |
| F7 memory with word-tables.docx (working set) | 489 MB | 290 MB | 462 / 291 MB |

In this run Word shows the same two-frame p95 tail as RaabeWrite (the Windows composition tail
of the Step 2 report); the time our app needs from key event to frame is unchanged (median 4.9
and 6.8 ms). The extra memory is Word's 105 table styles and the new panes and data.
F3b, the formatting commands against Word: **44/44**.

## Found and fixed during the audit and the wrap-up

- WebView2's clipboard API does not return while the clipboard holds only pictures: Einfügen
  hung after copying a picture in Word. The app now reads the clipboard natively.
- Typing over a selected paragraph merged it with the next one; Word keeps the paragraph mark.
- Kopieren and Ausschneiden were greyed out with only a picture or shape selected.
- Our Tabellenlayout keytips were guessed; they are now Word's.
- A probe left Word's ribbon reduced and Alle anzeigen on for the owner's next session (both
  put back at once; probes now restore them, and CLAUDE.md says so).
- A floating picture near the bottom of a page now takes its paragraph to the next page, as in
  Word (word-wrap.docx had 12 lines on the wrong page).
- The F3b test applied Überschrift 1 in Word through COM but sent our app Strg+Alt+1, which on a
  German keyboard is AltGr+1 and does nothing in Word either (the behaviour catalogue shows
  it); our app now gets the command, as Word does. run.py writes its reports in UTF-8 (an
  equation's object character had stopped a layout run before it saved).

## Known gaps

- Floating objects cross the clipboard as pictures in the line, both ways (Word makes floats
  only from its VML); shapes and text boxes copied from Word arrive as their pictures and text.
- Als Grafik speichern for shapes, Textrichtung in table cells, Tabelle zeichnen, Radierer,
  Formel in tables, Neue Liste mit mehreren Ebenen definieren and Neuen Listentyp definieren
  are deferred (greyed, "Zurückgestellt"); Abstand zwischen Zellen is not drawn.
- AutoAnpassen an Inhalt fits the columns once (Word keeps fitting while typing).
- The colours of themes other than Office's two are named "Text 2", "Akzent 1" … where Word
  names them by hue.
- Known from 4b and unchanged: deferred commands, equation heights of matrices and nested
  fractions, five hyphenation choices in a long justified text.
