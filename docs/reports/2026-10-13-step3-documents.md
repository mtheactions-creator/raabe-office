# Step 3 report: RaabeWrite documents

Date: 2026-10-13 · Measured with `tools/feel-lab` (see `docs/FEEL.md`) · Raw data in
`tools/feel-lab/results/2026-10-13-docx.json`, `-rendering.json`, `-memory.json`,
`-latency.json`

## What was built

- **Document structure (D-029):** stories (main text, headers, footers, notes, comments),
  sections, tables whose cells are ordinary paragraphs, Word's list definitions, floating
  pictures, markers for fields, hyperlinks and comments. Layout split into line breaking
  (`layout.ts`) and page flow (`flow.ts`). The .docx reader and writer cover all of it and
  write unchanged parts back byte for byte.
- **Lists:** Aufzählungszeichen, Nummerierung and Liste mit mehreren Ebenen with Word's
  library entries, heading numbering (1 / 1.1 / 1.1.1), Neu beginnen mit 1; AutoFormat while
  typing ("1. ", "1) ", "a) ", "I. ", "* ", "- ", "> "), Tab / Umschalt+Tab change the level,
  Enter on an empty item ends the list, Backspace at the start removes the number.
- **Tables:** Einfügen > Tabelle (grid and dialog); Tab and Umschalt+Tab between cells, a new
  row after the last cell; contextual tabs **Tabellenentwurf** (style options, six table
  styles, Schattierung, Rahmen) and **Layout** (Auswählen, Gitternetzlinien, Eigenschaften,
  Löschen, insert above/below/left/right, Zellen verbinden/teilen, Höhe/Breite, Zeilen/
  Spalten verteilen, nine cell alignments, Überschriften wiederholen); column borders can be
  dragged with the mouse; a selection across cells selects whole cells.
- **Pictures:** Einfügen > Bilder > Dieses Gerät; resizing with Word's eight handles (corners
  keep the proportions); contextual tab **Bildformat** with Textumbruch (all seven kinds) and
  Größe; floating pictures can be dragged; text flows around them.
- **Page setup:** Seitenränder (Normal, Schmal, Mittel, Breit, custom), Ausrichtung (margins
  rotate with the page, as measured in Word), Format (A3, A4, A5, B5, Letter, Legal), Spalten
  (1-3, Links, Rechts, custom with Zwischenlinie), Umbrüche (Seite, Spalte, Textumbruch;
  Abschnitt: Nächste Seite, Fortlaufend, Gerade/Ungerade Seite) and the dialog Seite
  einrichten.
- **Headers, footers, page numbers:** galleries Leer / Leer (drei Spalten), Seitenzahl at the
  top or bottom (left, centre, right), "Seite X von Y", Aktuelle Position, Seitenzahlen
  formatieren, entfernen; double click opens and closes them; the contextual tab **Kopf- und
  Fußzeile** (Erste Seite anders, Untersch. gerade ungerade Seiten, distances, Schließen).
- **Paragraph borders and shading** from the Start tab (also for table cells).
- **Suchen und Ersetzen:** the Navigationsbereich (Strg+F) with live results, all matches
  highlighted and the Überschriften outline; the dialog with Suchen, Ersetzen (Alle
  ersetzen as one undo step) and Gehe zu (Strg+G).
- **Rechtschreibung:** red wavy lines while typing, suggestions and Alle ignorieren /
  Zum Wörterbuch hinzufügen in the context menu, the F7 dialog, the Korrekturhilfesprache.
  German, US and British English dictionaries are compiled into the program (D-031).
- **Clipboard:** copying puts HTML with formatting on the clipboard (Word pastes it with
  fonts, sizes, bold …); HTML from Word or a browser is pasted with formatting, headings,
  lists and tables.
- **Drucken:** a preview with page navigation; all, current or chosen pages ("1-3, 5"). PDF
  export and printing keep each section's page size and orientation.
- **Text rendering** with grayscale smoothing like Word 365 (D-032).

## Results

### Layout (F4b): ten test documents built by Word

| Document | What it tests | Line starts | Page breaks | Baselines (Word's PDF) |
|---|---|---|---|---|
| word-layout.docx | Step 2 worksheet | 90/90 | 4/4 | within 0.18 pt |
| word-fonts.docx | 13 fonts, sizes 8-36 pt | 80/80 | 3/3 | within 0.18 pt |
| word-breaks.docx | space before after breaks | 81/81 | 6/6 | within 0.17 pt |
| word-top.docx | space before at the top | 5/5 | 2/2 | within 0.06 pt |
| word-lists.docx | bullets (2 levels), numbering with 10+ items, multilevel 1. 1.1. 1.1.1. | 33/33 | 1/1 | within 0.16 pt |
| word-tables.docx | grid table with wrapping cells, merged cells (across and down), centred table with minimum height and vertical centring, a 32-row table with header row over a page break | 143/143 | 2/2 | within 0.08 pt |
| word-sections.docx | first-page header, page numbers "Seite X von Y", a landscape section, two columns with a line between, a column break, continuous section breaks | 102/102 | 4/4 | within 0.14 pt |
| word-notes.docx | six footnotes, one long, near page ends | 43/43 | 2/2 | within 0.15 pt |
| word-borders.docx | boxed, shaded and grouped paragraphs, bottom rule, indented box | 23/23 | 1/1 | within 0.10 pt |
| word-images.docx | an inline picture in a line, a picture right with square wrapping, a picture with "oben und unten" | 32/32 | 1/1 | within 0.13 pt |

**690/690 line starts and every page break identical, all baselines within 0.18 pt.** The
measurements that got there are D-030: list label fonts, table border widths in rows and
cells, how rows break, line spacing on picture lines, continuous section breaks and column
breaks. The harness was extended for this: Word's lines are found by position (its line
navigation moves between table cells), table, section and picture characters are mapped
between Word and RaabeWrite, and baselines are read from the PDFs by position.

### Text rendering (F6)

Word and RaabeWrite showing the same lines of `word-layout.docx` at 100 %, captured from the
screen (`run.py rendering`):

| | Word | RaabeWrite |
|---|---|---|
| Edge sharpness (mean gradient on text) | 138.0 | 133.6 (97 %) |
| Ink pixels | 40 986 | 40 326 (98 %) |
| Colour fringes (ClearType) | 3.7 | 3.4 |
| Glyph positions (ink profile correlation per line, median) | | 0.97 |

The first measurement found ClearType colour fringes in RaabeWrite (74 against Word's 3.7):
Word 365 draws document text with grayscale smoothing. WebView2 now does the same (D-032).
The same run found two spelling false alarms (words with soft hyphens, long compounds),
both fixed.

### Memory

Same document (`word-tables.docx`) open, working set of the whole process tree
(`run.py memory`): RaabeWrite 345 MB (7 processes: app and WebView2), Word 300 MB.

### Typing latency (F1)

Real keystrokes, screen polled until the character appears (60 samples each, display at
30 Hz): in the long document RaabeWrite's median is **31.7 ms against Word's 35.3 ms** (p95
64.7 ms: the Windows composition tail described in the Step 2 report). Inside the app, a
key event reaches the screen update in 6.1 ms (median). The short-document run of Word
timed out in the harness (59 of 60 samples), so it gives no comparison; RaabeWrite's median
there was 33.2 ms.

### Smoke tests

- `smoke_step3.py` (new): **28/28** in the real app, files checked with Word: AutoFormat
  list, Tab levels, Listenabsatz style, table with Tab navigation and a new row, contextual
  tabs, inserting a column, page numbers "Seite X von Y" in the footer, section break with a
  landscape section, picture insert, selection and resizing, find, replace all, spelling
  marks, suggestions and dictionary, HTML copy and paste, Word reading table, list,
  sections, orientation, footer text, picture and page count, print preview.
- `smoke_write.py` (Steps 1 and 2): **41/41**, a Word table is now edited in place (typed
  text in a cell survives in Word).
- Unit tests: 44, including new ones for table operations, merges, lists and writing new
  sections, headers and footers.

### Performance

Typing in a 37-page document costs 4.3 ms per keystroke for layout and drawing (the full
page flow runs on every key; paragraphs are cached).

## Known gaps

- Pictures cannot be rotated or cropped yet; shapes and text boxes are shown as frames.
- Tables inside table cells are kept and shown read-only.
- The bullet button turned off in Word leaves a small left indent (measured with Word's own
  command); RaabeWrite returns the paragraph to no indent. Not yet checked with real input.
- Grammar checking is not part of Step 3 (SCOPE: later).
