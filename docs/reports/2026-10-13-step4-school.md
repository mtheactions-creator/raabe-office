# Step 4 report: RaabeWrite school features

Date: 2026-10-13 · Measured with `tools/feel-lab` (see `docs/FEEL.md`) · Raw data in
`tools/feel-lab/results/2026-10-13-docx.json`, `-latency.json`, `-memory.json`

## What was built

- **Fuß- und Endnoten (Referenzen):** Fußnote einfügen (Strg+Alt+F) and Endnote einfügen
  (Strg+Alt+D) put the reference at the caret and the caret into the new note; Nächste
  Fußnote with Vorherige Fußnote, Nächste/Vorherige Endnote. Footnotes stand at the page
  bottom (Step 3), endnotes follow the text under a separator line and are numbered i, ii,
  iii as in Word. A note whose reference is deleted is left out of the file.
- **Inhaltsverzeichnis:** the gallery Automatische Tabelle 1 ("Inhalt") and 2
  ("Inhaltsverzeichnis"), Text hinzufügen (Ebene 1-3, Nicht im Inhaltsverzeichnis anzeigen),
  Tabelle aktualisieren with Word's dialog (Nur Seitenzahlen / Gesamtes Verzeichnis),
  Inhaltsverzeichnis entfernen. The table is a real Word field (TOC \o "1-3" \h \z \u) with
  linked entries, _Toc bookmarks and PAGEREF page numbers; numbered headings appear with
  their numbers ("2.1 Unterkapitel"). Page numbers come from the layout and are laid out
  again until they are stable (a longer table can push headings to later pages).
- **Kommentare (Überprüfen):** Neuer Kommentar (Strg+Alt+K, also Einfügen > Kommentar) on the
  selection or the word at the caret; cards right of the page as in Word 365 (avatar,
  name, time, text, Antworten box; Strg+Eingabe posts, Esc cancels) with the commented text
  highlighted, darker for the active card; replies, Kommentar bearbeiten, Thread löschen,
  Thread auflösen / erneut öffnen; Löschen (one / all), Zurück, Weiter; the Kommentare
  button opens a pane with all threads. Comments, replies and "erledigt" are written in
  Word's format (comments.xml, commentsExtended.xml).
- **Änderungen nachverfolgen:** Strg+Umschalt+E; typed and pasted text is marked as
  inserted, deleted text stays struck through, paragraph marks are tracked too, and
  deleting one's own insertion removes it (D-034). Annehmen/Ablehnen with Word's menus
  (… und weiter, Diese Änderung, Alle Änderungen, … und Nachverfolgung beenden), Zurück and
  Weiter, Markup für Überarbeitung anzeigen (Einfaches Markup, Alle Markups, Kein Markup,
  Original) and change bars in the margin. Author: the user name from Optionen.
- **Formeln (D-036):** Einfügen > Formel (Alt+Umschalt+0) with Word's built-in equations
  (Binomischer Lehrsatz, Fläche eines Kreises, Pythagoreischer Lehrsatz, Quadratische
  Formel, Taylor-Entwicklung …); editing in place with MathLive (offline, MIT); the
  contextual tab **Formel** with Normaler Text, Symbole (basic mathematics and Greek
  letters) and all eleven Strukturen galleries (Bruch, Skript, Radikal, Integral, Großer
  Operator, Klammer, Funktion, Akzent, Grenzwert und Logarithmus, Operator, Matrix). The
  arrow keys move into and out of an equation, a click opens it. Equations are drawn as
  MathML in Cambria Math, alone in a paragraph as a centred display equation, and saved as
  Office Math (OMML); Word's equations are read and become editable.
- **Raabeschule-Vorlagen (D-035):** Datei > Neu shows Arbeitsblatt, Facharbeit,
  Versuchsprotokoll and Handout with page sketches. The Facharbeit has a title page without
  page number, an Inhaltsverzeichnis, numbered chapters, Literaturverzeichnis, Anhang and
  Erklärung, Times New Roman 12 pt, 1,5-zeilig, Blocksatz and a 4 cm correction margin. A
  new document from a template counts as unchanged until it is edited.

## Results

### Layout (F4b): eleven test documents built by Word

The new `word-step4.docx` (Word's own table of contents, two footnotes, an endnote, a
comment, two inline and three display equations built from Word's linear format) joins the
ten documents of Step 3:

| Document | Line starts | Page breaks | Baselines (Word's PDF) |
|---|---|---|---|
| ten Step 3 documents | 632/632 | 26/26 | within 0.18 pt (unchanged) |
| word-step4.docx | 50/50 | 2/2 | before the equations within 0.11 pt; after five equations within 0.8 pt |

**682/682 line starts and every page break identical.** (The Step 3 report gave 690 for its ten documents; its own table sums to 632, corrected there.) Equation lines were measured
against Word and tuned (sum limits above and below in display equations, matrix row
spacing, the extra space Word gives display sums and integrals). The harness learnt to map
Word's hidden field codes and the characters inside equations, which Word counts as
document positions but RaabeWrite holds as one object.

### Word reads what RaabeWrite writes (`smoke_step4.py`, new): 43/43

In the real app, files checked with Word: footnote and endnote with their texts, the
comment on the right words with its reply, a tracked insertion and deletion, two equations
(one display) with the right content, a table of contents for which **Word's own update
gives exactly the entries and page numbers RaabeWrite wrote**; the same for the Facharbeit
template (numbered headings, Times New Roman, 1,5-zeilig, no page number on the title
page). All four templates open unchanged and Word counts the same pages. Reopened, nothing
is left read-only. Also checked: the shortcuts Strg+Alt+F, Strg+Alt+K, Strg+Umschalt+E and
Alt+Umschalt+0, the Formel tab, the comment card beside the page, Markup Original.

### Unchanged

- `smoke_step3.py` 28/28, `smoke_write.py` 41/41, unit tests 59 (new: round trips of notes,
  comments, tables of contents and tracked changes; LaTeX ↔ OMML round trips of 18 formulas; the
  Word-made Step 4 document).
- F1 typing latency with real keystrokes: median **31.6 ms (short) and 32.0 ms (long)
  against Word's 36.8 and 34.3 ms**; the p95 tail (65 ms) is the Windows composition tail
  described in the Step 2 report.
- F7 memory with the same document: RaabeWrite 346 MB, Word 302 MB (Step 3: 345/300).
  MathLive is only loaded when an equation is edited.

## Found and fixed on the way

- Alt+Umschalt combinations went to the keytips (Alt+letter jumps to a tab); Word treats
  them as shortcuts, and so does RaabeWrite now.
- Word ends the field of a table of contents at the start of the first heading; updating
  the table would have replaced that heading. The heading is now kept.
- A reply box could take the focus a moment after posting when the window draws late; it
  now only does so if the focus has not moved on.
- MathLive's own MathML writes sum limits beside the ∑ even in display equations, and the
  browser's default math font falls back to Times for digits. RaabeWrite now builds the
  MathML from the same parse as the saved OMML and names Cambria Math explicitly.

## Known gaps

- Deleting another author's insertion makes it a plain deletion: rejecting it brings the
  text back, where Word would also undo the insertion (D-034).
- The equation editor (MathLive) draws with its own fonts while editing; the document shows
  Cambria Math.
- Equation alignment in multi-line equations (& marks) is kept but not shown.
- Comments on headers, footers or notes, and tracked formatting changes, are kept from Word
  files but not created.
- Captions, cross-references, citations and the bibliography remain "later" (SCOPE).
