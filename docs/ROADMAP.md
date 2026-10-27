# Roadmap

Each step ends with something that can be checked. No step starts without the project
owner's explicit go-ahead.

| Step | Content | Status |
|---|---|---|
| 0 | Dev environment, repository, documentation, logos, measurement harness against Word, two feel prototypes, report and engine decision (D-014) | done 2026-10-13: engine decided (D-017), published on GitHub ([report](reports/2026-10-13-step0-feel.md)) |
| 1 | Shared shell: window, document tabs, ribbon (incl. touch mode), backstage, file dialogs, AutoSave/recovery, theming in Raabe red | done 2026-10-13 (installer 0.1.0); real-input re-check: F1 32 ms vs Word 34-36 ms, F3 49/50 |
| 2 | RaabeWrite core: text model, layout, pagination, selection, character/paragraph formatting, styles, undo/redo, .docx read/write for that subset | done 2026-10-13 ([report](reports/2026-10-13-step2-formatting.md)): layout identical to Word in four test documents, formatting behaviour vs Word's own commands, lossless .docx; real-input re-check done (F1 tail: Windows composition, see report) |
| 3 | RaabeWrite documents: lists, tables, images, header/footer, page numbers, page setup, spell check, find/replace, print, PDF | done 2026-10-13 ([report](reports/2026-10-13-step3-documents.md)): layout identical to Word in ten test documents (632/632 lines), text rendering like Word (F6), smoke tests 28/28 and 41/41 |
| 4 | RaabeWrite school features: footnotes, table of contents, comments, track changes, formula editor, Raabeschule templates | done 2026-10-13 ([report](reports/2026-10-13-step4-school.md)): Word updates our tables of contents to the same page numbers, layout identical to Word in eleven documents (682/682 lines), smoke tests 43/43, 28/28, 41/41 |
| 4b | Word parity for RaabeWrite (D-038): feel basics and audit tools, tables/lists/clipboard, pictures and shapes, own formula editor, all remaining non-deferred ribbon commands, French and Spanish, own grammar checker | done 2026-10-15 ([report](reports/2026-10-15-step4b-word-parity.md)): every non-deferred ribbon command works; 960/965 line starts in 19 Word-built documents, behaviour catalogue 192/193 (the 193rd a Word glitch), formula editor 114/114, grammar in four languages finding more school errors than Word with fewer false alarms, smoke tests 41/41, 28/28, 43/43, 25/25 |
| 4c | Polish for RaabeWrite (D-053): the owner's list after trying Step 4b (pictures, formulas, tables, cursors, Formatvorlagen) and an audit of the whole app for the same kinds of difference | done 2026-10-15 ([report](reports/2026-10-15-step4c-polish.md)): the owner's list and an audit of the whole app (tooltips, labels, menus and context menus against Word, Livevorschau, double-clicks, panes, clipboard); 1216/1221 line starts in 22 Word-built documents, behaviour catalogue 202/203 (the 203rd a Word glitch), formula editor 146/146, smoke tests 41/41, 28/28, 43/43, 25/25 |
| 4d | Second polish list from the owner after trying Step 4c (D-059): pictures by drag and drop, text boxes and shapes while edited, the caret, ribbon menus, tables and the border pen, formulas | done 2026-10-16: the owner's list as recorded in Word (D-059); 1216/1221 line starts in 22 Word-built documents (unchanged), behaviour catalogue 206/207 (the 207th a Word glitch), formula editor 152/152, smoke tests 41/41, 28/28, 43/43, 25/25 |
| 5 | Pilot: one class uses RaabeWrite in lessons; fix what real use reveals | runs alongside Step 7 when a class fits the timetable (D-061) |
| 6 | iPad build of RaabeWrite (touch polish, Files app integration, distribution) | |
| 7 | RaabeCalc (D-061): handles like Excel in every way, the controls first; packages A-F below | in progress since 2026-10-27 (owner's go for the whole step, overnight run): A-E built (D-062, D-063), behaviour catalogue 142/142 |
| 7b | Diagramm in RaabeWrite: charts with their embedded workbook, drawn by RaabeCalc's chart engine (D-061) | right after Step 7; not to be forgotten |
| 8 | RaabePresent: proposal below | after Step 7 (D-061) |
| 9 | Optional features: read aloud, translator, dictation, ink | |

## Step 7: RaabeCalc (plan)

Agreed with the owner on 2026-10-27 (D-061). RaabeCalc comes before RaabePresent because
Present's charts come from RaabeCalc's chart engine and Excel, unlike PowerPoint, runs as a
process of our own for the harness. Goal: RaabeCalc handles like Excel in every way, the
controls above all. Same method as Step 4b, from the first package on: every control and
behaviour is recorded in an Excel instance of our own first, then built the same way and
compared automatically. The owner leaves the PC for the whole step, so real input is allowed
throughout. Default font Calibri 11 (Excel 2013-2022's defaults) instead of Excel 365's Aptos
Narrow (D-011). The owner tries the app after package D and at the end of the step.

| Package | Content |
|---|---|
| A. Foundation and Excel reference | The harness's Excel side (`ExcelTarget` through `DispatchEx`, workbooks built by Excel itself, cell positions and values read through COM, Excel's PDF export); the behaviour catalogue and screen comparison for Excel (real input, menus and galleries recorded entry by entry); the shared Rust code moved out of `apps/raabe-write/src-tauri` into a crate both apps use; `packages/ooxml` reading and writing .xlsx parts beside .docx; `apps/raabe-calc` with the shell and Excel's ribbon, labels, keytips and tooltip titles as recorded. |
| B. Grid and .xlsx | Reading and writing workbooks, everything not edited kept verbatim (D-024); column widths in Excel's character units and row heights as Excel computes them; fonts, fills, borders, alignment, wrapping, merged cells, hidden rows and columns, Fenster fixieren; a grid that draws only what is visible, so large sheets scroll at full frame rate (F2). Measured against Excel's own output. |
| C. Formula engine | German function names and separators in the UI, English in the file as Excel writes it; recalculation through a dependency graph that never blocks typing; the MVP function list (SCOPE); Excel's error values and its rounding, display and date rules. Every function compared with Excel's results on generated cases. |
| D. Working in the grid | Excel's controls and behaviour as recorded (see below); AutoAusfüllen, AutoSumme, copying with adjusted references, inserting and deleting cells, rows, columns and sheets, sorting and AutoFilter; behaviour catalogue cases for each. |
| E. Charts | The chart engine shared with RaabePresent and RaabeWrite: Säule, Balken, Linie, Kreis, Punkt (XY) with Trendlinie, Formel and Bestimmtheitsmaß, Fläche; Excel's default look and the Diagrammentwurf and Format tabs, recorded from Excel. |
| F. Formats, output, audit | Number formats in German conventions, bedingte Formatierung, printing (Druckbereich, an Seite anpassen, Gitternetzlinien and Überschriften), CSV; then the audit as in Step 4c F: tooltips, menus and context menus, double-clicks, Livevorschau, cut-off labels, the clipboard both ways with Excel, Word and RaabeWrite. |

**Excel's controls**, each recorded in Excel before it is built (pointers, keys, mouse, what
appears and where): the cell pointer and selections (Umschalt, Strg, whole rows and columns,
the select-all corner), Eingeben and Bearbeiten modes (F2) with what the arrow keys do in each,
Enter/Tab/Umschalt directions, Strg+arrows to the edge of the data, the fill handle (dragging,
double-click, Strg), moving and copying cells by their border, the Namenfeld and the
Bearbeitungsleiste (Funktion einfügen, its height), pointing at cells while a formula is typed
(coloured ranges, F4), AutoVervollständigen in cells and for functions with their screen tips,
column and row headers (sizing, double-click to fit, hiding), the sheet tabs (switching,
dragging, Strg-drag, double-click to rename, context menu, the scroll arrows, Neues Blatt), the
scroll bars and how far the sheet scrolls, Strg+mouse wheel zoom, the status bar (Summe,
Mittelwert, Anzahl, its context menu, views, zoom), the AutoFilter buttons and their menus, the
Einfügeoptionen and AutoAusfülloptionen buttons, error indicators and their menu, the
Schnellanalyse button, context menus of cells, headers and sheet tabs, and every keyboard
shortcut Excel reports on the German keyboard.

Status 2026-10-27 night: A (Excel reference side, raabe-host crate, .xlsx package, the app with
Excel's ribbon), B (grid, formats, .xlsx with 0 differences in Excel), C (formula engine,
4229/4234 Excel results), D (controls: behaviour catalogue 142/142 cases as in Excel, context
menus and tooltips recorded) and E (charts) are built (D-062, D-063). F (printing, conditional
formatting, AutoFilter, CSV are in; the audit is open).

## Step 8: RaabePresent (proposal, not yet agreed)

Proposed on 2026-10-27 and set back behind Step 7 (D-061); to be discussed again before it
starts. Packages as in Step 4b: A. PowerPoint reference side for the harness (fixtures built by
PowerPoint, line positions through COM, slides exported as PDF/PNG) and `apps/raabe-present`;
B. reading, drawing and saving .pptx (masters, layouts, themes, placeholders, text bodies with
autofit and bullet levels, shapes, pictures, tables, equations; everything else verbatim);
C. editing in Normalansicht (slide panel, Neue Folie with the layout gallery, text editing,
PowerPoint's two selection levels, moving and sizing, smart guides, Anordnen, the notes pane);
D. Foliensortierung, Abschnitte, PDF, printing, designs incl. a Raabeschule design; E.
Bildschirmpräsentation (navigation, pen and laser pointer, Referentenansicht, a curated set of
transitions and animations); F. the audit. Open questions for then: PowerPoint runs as a single
process, so the harness cannot start one of its own and must refuse to run while the owner's
PowerPoint is open; new presentations 16:9 with the Office 2013-2022 design (Calibri Light,
Calibri); PowerPoint's own designs are Microsoft artwork and are not copied (school material
for a Raabeschule design welcome).

## Step 4d: Second polish list

Agreed with the owner on 2026-10-16 after trying Step 4c (D-059), in one run with real input.
The owner's list: pictures dropped from Explorer; text boxes and shapes while their text is
edited (pointers, yellow handles, moving by the frame); a caret "afterglow" on fast Rücktaste;
ribbon tabs jumping and chevron menus that do not close on a second press; tables (right-click
on selected cells, whole table vs. cells, typing over cells, the layout handle, at most 14
Tabellenformatvorlagen, the border pen's lines and exit, gridlines only while the pen is out);
formulas ("/" + space alone and after one operand, script sizes).

**Done 2026-10-16**: everything on the list as recorded in Word or decided with the owner
(D-059); the ribbon jump could not be reproduced (tabs switch only on insertion, double-click
and header editing). Found on the way: after a right-click and Esc the keys went nowhere, and a
drag from a cell into the text after the table ended mid-word (both fixed). Behaviour catalogue
206/207, F3 49/50, F3b 44/44, formula editor 152/152, F4b 1216/1221 (unchanged), F1 32 ms vs
Word 36 ms, smoke tests all passed.

Fixes after Step 4d, 2026-10-26 (D-060): the sizing handle of a new table no longer makes it
grow on the first drag; Zeilen verteilen gives the rows Word's heights (3/3 recorded cases).

## Step 4c: Polish (plan)

Agreed with the owner on 2026-10-15 after trying Step 4b. Same working mode as 4b: Word's
behaviour is recorded first in Word instances of our own, then built and measured against it;
the owner leaves the PC, so real input is allowed throughout. Our look may differ from Word's,
but everything sits where a Word user expects it (D-053).

| Package | Content |
|---|---|
| A. Bugs | Zuschneiden creates duplicates; the table's top-left handle selects only part of the table; only the first Tabellenformatvorlage works; Bildformatvorlagen previews overlap; Formatvorlagen labels cut off. |
| B. Pictures and panes | Text flow on both sides for Quadrat, Eng along the picture's outline (as measured in Word); snapping to margins, page centre and paragraphs with guide lines; double-click opens the contextual tab; icons and tooltips in the layout-options menu; Transparente Farbe bestimmen; a right-side task pane in our style for everything Word puts in a pane (Grafik formatieren, Bildfarboptionen …) instead of centred dialogs. |
| C. Formulas | The Start tab shows the formula's font; copy and paste with Word both ways; "/" + space builds a fraction (with the operand before it) and "…" + space makes plain text, as recorded from Word; the caret around quotation marks; the top-left handle selects the whole formula. |
| D. Tables | Drag-selection thresholds and the shading of selected cells as in Word (text area, multi-line cells, gridlines visible); the full Rahmen group (Rahmenarten, Linienart, Stärke, Stiftfarbe, Rahmenpinsel, Rahmenkopierer); all of Word's table styles with the in-ribbon preview; Entf and Rücktaste on selected cells with Word's Zellen löschen dialog; floating tables moved freely by the handle, plus a layout handle of our own for wrapping and alignment (D-053). |
| E. Selection, cursors, Formatvorlagen | A selection visible over every background (highlighting, shading, coloured text, page colour); sharp cursors at every scaling (Windows' own where Word uses them); live preview on hover in the galleries; a Formatvorlage applied to a selection resets its direct formatting cleanly (D-053). |
| F. Audit of the whole app | The same kinds of difference everywhere: live preview in every gallery, panes instead of dialogs, double-click on every object, tooltips and icons, cut-off labels at several widths (by script), selection visibility, clipboard with Word for every object, snapping for shapes and text boxes, menus and galleries entry by entry against recordings of Word, new behaviour-catalogue cases; what it finds is added to the list. |

Status: A done 2026-10-15. B done 2026-10-15 (D-054): text on both sides and Eng along
outlines (all 162 line starts of the Word-built word-wrap.docx as in Word), snapping with guide
lines, double-click, wrap icons, Transparente Farbe bestimmen, the Grafik/Form formatieren pane
in Word's layout, turned selection frames, Word's shape context menu.
C done 2026-10-15 (D-055): "/" takes the whole operand and quoted text becomes normal text as
in Word (146/146 recorded cases in the app), equations through the clipboard with Word both ways
(0 differences), the formula's font in Start, the handle selects the whole formula. D done
2026-10-15 (D-056): Entf/Rücktaste and Zellen löschen as recorded, cell drag and shading as Word's,
the full Rahmen group with Rahmen übertragen and Rahmenpipette, Word's 105 table styles with
previews and Livevorschau, floating tables (word-table-float.docx 100 %) with dragging and our
layout handle. E done 2026-10-15 (D-057): selection visible over every background, sharp
pointers, Livevorschau in the Formatvorlagen gallery, Formatvorlagen reset cleanly.
F done 2026-10-15 (D-058): tooltips, labels, menus and context menus audited against Word, the
Tabellenlayout tab and the list dialogs completed; Livevorschau in every gallery and palette
Word has it in; Word's double-clicks and the table selection zones; Alternativtext as a pane;
pictures and links through the clipboard with Word both ways; ten new behaviour cases.
**Step 4c done 2026-10-15** ([report](reports/2026-10-15-step4c-polish.md)): F4b 1216/1221 line
starts in 22 Word-built documents, behaviour catalogue 202/203, formula editor 146/146, smoke
tests all passed.

## Step 4b: Word parity (plan)

Agreed with the owner on 2026-10-14 after trying Step 4. Goal: RaabeWrite behaves and looks
like Word wherever a student or teacher would notice, before the pilot. The owner leaves the
PC for the whole step, so real-input measurements (mouse, hover, clipboard, typing in Word)
are allowed throughout. Every package ends with measurements against Word and an updated
report.

| Package | Content |
|---|---|
| A. Basics and audit tools | Mouse cursors as in Word (I-beam over text, mirrored arrow in the selection bar, move/resize cursors on pictures, table borders and handles); caret hidden while the window is inactive; caret to the next page after a page break; deferred commands marked in their tooltip (D-041). Two audit tools used by every later package: a **behaviour catalogue** (hundreds of scripted interactions run in Word and RaabeWrite, comparing document and caret) and a **screen comparison** of every ribbon tab, menu and dialog, normal and on hover (icons, chevrons, split buttons). |
| B. Tables, lists, clipboard | Gridlines only where a cell edge has no border, default as in Word; Word's table context menu; row heights by dragging; the move handle (top left) and resize handle (bottom right); list gallery entries and AutoFormat prefixes ("- ", "* ", "1. ", "a) " …) 1:1 with Word (symbol, font, size, indent, tab); copy and paste of lists and tables to and from Word keeping indents and levels. |
| C. Pictures and drawing | Fix inserting the reported PNG, paste pictures (Strg+V), pictures in table cells; the full Bildformat tab except Künstlerische Effekte (Korrekturen, Farbe, Transparenz, Bildformatvorlagen with frames and shadows, Zuschneiden, Drehen, Position, Ausrichten, Vorder-/Hintergrund, Alternativtext, the layout-options button); Formen and Textfeld with the Formformat tab; Anordnen in the Layout tab. Künstlerische Effekte last, if worth it. |
| D. Formula editor | Our own editor replaces MathLive (D-039): editing in the text line without an overlay, Word's slot navigation and selection with the mouse, the two handles, Word's Backspace and size rules (inline/display), Word's Math AutoCorrect (\cdot + space …), the Formel tab exactly as in Word except Freihandgleichung. Word's behaviour is recorded first as a specification (real keystrokes in our own invisible Word instance). |
| E. Remaining commands | Every ribbon command that is not deferred: Symbol, Link, Textmarke, Querverweis, Beschriftung, Zitat einfügen, Quellen verwalten, Literaturverzeichnis, Fuß- und Endnote dialog, Optionen zum Nachverfolgen, Entwurf tab (Designs, Farben, Schriftarten, Absatzabstand, Wasserzeichen, Seitenfarbe, Seitenränder), Silbentrennung (measured against Word's line breaks), Lineal, Lesemodus, Mehrere Seiten, Zwischenablage pane, Texteffekte, Sortieren, Markieren menu. |
| F. Languages and grammar | French and Spanish spelling, hyphenation and Thesaurus (also German and English); our own grammar checker in Rust (D-040): engine, then German, English, French, Spanish, each switched on once its measurements are good. |

Status: A done 2026-10-14 (cursors, caret, page break, deferred tooltips, catalogue and
screen comparison; the catalogue found 27 differences, all fixed, D-042). B done 2026-10-14
(gridlines, context menus and mini toolbar D-044, row heights, handles, Word's list libraries
and AutoFormat prefixes, clipboard D-043). C done 2026-10-14 apart from what moved to E: the
Bildformat tab (D-045), Formen and Textfeld with the Formformat tab and Layout > Anordnen
(D-046: all 152 preset shapes checked against Word's drawing, Word's 77 Formenarten). D done
2026-10-14: our own formula editor replaces MathLive (D-047): 114 recorded Word cases replayed
like Word, the Formel tab with Word's galleries (231 entries) and symbols (8 categories) as
Word inserts them, Word's context menus, Konvertierungen and Formeloptionen; Word's 24
equations of `word-math.docx` written back with Word's structure. E in progress (D-048): Symbol, Link,
Textmarke, Querverweis, Beschriftung, Abbildungsverzeichnis, Zitate, Literaturverzeichnis,
Fuß- und Endnote, Optionen zum Nachverfolgen, Überarbeitungsbereich, Sortieren, Markieren,
Auswahlbereich, Formatvorlagen pane with Formatvorlage ändern, Zwischenablage, Lineal,
Gitternetzlinien, Mehrere Seiten, Weblayout, Lesemodus, Zeilennummern, Notizen anzeigen, the
Entwurf tab and Rahmen und Schattierung done; E done 2026-10-14 with Texteffekte, Gruppieren,
verknüpfte Textfelder and Deckblatt (every non-deferred ribbon command now works; Silbentrennung
and Thesaurus come with the languages in F). F in progress (D-049): French and Spanish spelling, Sprache
automatisch erkennen, Word's dialog Sprache, Thesaurus pane and Synonyme, the Editor pane (F7)
and the options Dokumentprüfung and Sprache done. Silbentrennung done (D-050: Word's breaks on 120
of 125 measured lines). Our own grammar checker in Rust done for German
(D-051: 63 of 96 school errors found where Word finds 46, false alarms on clean texts as low as
Word's), then English, French and Spanish (D-052: every school error Word finds and more, fewer
false alarms on Wikipedia than Word), all four switched on. F done 2026-10-15; Step 4b done
2026-10-15 with the wrap-up measurements in the
[report](reports/2026-10-15-step4b-word-parity.md).

Order: A, B, C, D, E, F. A comes first because its tools find the differences nobody has
reported yet; the behaviour catalogue grows with each package. F is last and the largest
open-ended part; German grammar first.
