# Step 4b report: Word parity for RaabeWrite

Date: 2026-10-15 · Measured with `tools/feel-lab` and `tools/lang-data` (see `docs/FEEL.md`) ·
Raw data in `tools/feel-lab/results/` (2026-10-14 and 2026-10-15)

Step 4b (D-038) took the differences the owner listed after trying Step 4 (cursor, icons,
tables, lists, formulas, pictures, behaviour, "all ribbon commands should work") and worked
through them in six packages. Everything was first recorded from Word on this PC, in Word
instances of our own (real input where needed; Word's options never changed), then built and
measured against it.

## What was built

- **A. Basics and audit tools (D-041, D-042):** Word's mouse cursors (I-beam, the mirrored
  arrow in the selection bar, move and resize cursors on pictures, table borders and handles),
  the caret hidden in inactive windows, the caret on the next page after a page break,
  deferred commands greyed out with "Zurückgestellt" in their tooltip. Two audit tools that
  every later package used: the **behaviour catalogue** (scripted real-input interactions run
  in Word and RaabeWrite, comparing text, formatting and caret) and the **screen comparison**
  of every ribbon tab, menu and dialog, normal and on hover. AutoCorrect and AutoFormat as
  Word 365 does them, including its Markdown shortcuts and emoji.
- **B. Tables, lists, clipboard (D-043, D-044):** gridlines only where a cell edge has no
  border, Word's context menus and mini toolbar, row heights by dragging, the move and
  resize handles, Word's list libraries and AutoFormat prefixes, and the clipboard with Word
  in both directions through the whole HTML document (lists, levels, indents, column widths).
- **C. Pictures and drawing (D-045, D-046):** the Bildformat tab (Korrekturen, Farbe,
  Transparenz, 28 Bildformatvorlagen, Bildeffekte, Zuschneiden, Drehen, Position, Ausrichten,
  Alternativtext, the layout-options button), pasted pictures and pictures in table cells;
  Formen and Textfeld with the Formformat tab: all 152 preset shapes drawn from ECMA-376's
  formulas and checked against Word's own drawing (within 0.5 pt), Word's 77 Formenarten.
- **D. Formula editor (D-039, D-047):** our own editor replaces MathLive: editing in the text
  line, Word's slot navigation, mouse selection, the two handles, Backspace and size rules,
  Math AutoCorrect, and the Formel tab with Word's galleries (231 entries) and symbols.
- **E. Remaining commands (D-048, D-050):** Symbol, Link, Textmarke, Querverweis,
  Beschriftung, Abbildungsverzeichnis, Zitate and Literaturverzeichnis, the Fuß- und Endnote
  dialog, Optionen zum Nachverfolgen, the Überarbeitungsbereich, Sortieren, Markieren, the
  Entwurf tab (Designs, Farben, Schriftarten, Absatzabstand, Wasserzeichen, Seitenfarbe,
  Seitenränder), Lineal, Gitternetzlinien, Mehrere Seiten, Weblayout, Lesemodus,
  Zeilennummern, Zwischenablage, Texteffekte, Gruppieren, verknüpfte Textfelder, Deckblatt.
  Every ribbon command that is not deferred by agreement now works.
- **F. Languages and grammar (D-049 to D-052):** French and Spanish spelling next to German
  and English, Sprache automatisch erkennen, Word's dialog Sprache, the Thesaurus pane and
  Synonyme menu in four languages, the Editor pane (F7), Dokumentprüfung and Sprache in the
  options; Silbentrennung as Word divides words; and our own grammar checker in Rust for
  German, English, French and Spanish: LanguageTool's open rules run by our engine on open
  lexicons, plus our own rules for pupils' errors, all offline and compiled into the program.

## Results

### Layout (F4b): nineteen test documents built by Word

| Documents | Line starts | Page breaks |
|---|---|---|
| the eleven documents of Step 4 | 682/682 | every one; baselines in Word's PDF unchanged (word-step4: within 1.3 pt instead of 1.6 pt) |
| word-formatting, word-rich, word-short (Step 2) and word-shapes (C) | 85/85 | every one |
| word-math.docx (D: 24 equations made by Word) | 73/73 | 3/3 |
| word-hyphen, -limit, -long (F: automatic hyphenation) | 120/125 | one shifted by the five lines below |

**960 of 965 line starts identical to Word.** The five differences are the known ones of
D-050: in a long justified text Word sometimes keeps a word whole that RaabeWrite divides
(or the other way round), when both choices fit.

### Behaviour with real input (F3)

The behaviour catalogue (`catalogue.py`, real keyboard and mouse in Word and RaabeWrite):
**192 of 193 cases as in Word**: typing and AutoCorrect 42/42, AutoFormat lists 21/21, keys
25/25, mouse 11/12, delete 9/9, lists 21/21, tables 17/17, styles 7/7, breaks 6/6, formatting
20/20, undo 8/8, clipboard 5/5. The one difference is on Word's side: in both of today's runs
Word's triple-click left its caret at the start of the document, while RaabeWrite selected the
paragraph exactly as Word did in the run of 2026-10-14.

The formula editor (`catalogue_math.py`): **114 of 114** recorded Word cases replayed as in
Word. From the packages: the clipboard with Word without differences both ways (B), all 152
preset shapes within 0.5 pt of Word's drawing (C).

### Grammar against Word's offline grammar check (D-051, D-052)

Sentence level; Word in its own invisible instance, its options untouched. Word has no Spanish
proofing tools on this PC.

| Corpus | de ours / Word | en ours / Word | fr ours / Word | es ours |
|---|---|---|---|---|
| typical school errors (found) | 63 / 46 of 96 | 47 / 39 of 50 | 35 / 29 of 44 | 24 of 28 |
| 400 of LanguageTool's incorrect examples (found) | 392 / 72 | 393 / 92 | 387 / 103 | 383 |
| 500 Wikipedia paragraphs on school topics (flagged: false alarms) | 30 / 50 | 17 / 130 | 76 / 112 | 58 |
| 400 of LanguageTool's correct examples (flagged) | 36 / 28 | 13 / 59 | 79 / 81 | 31 |

In every language RaabeWrite finds more of the school errors than Word and flags fewer clean
Wikipedia paragraphs; all four languages are switched on. Checking takes about 3 ms per
sentence in the background; a language loads on first use in at most 0.36 s.
Silbentrennung (D-050): Word's breaks on 120 of 125 lines.

### Smoke tests in the real app (Word reads what RaabeWrite writes)

`smoke_write.py` 41/41, `smoke_step3.py` 28/28, `smoke_step4.py` 43/43 (now driving our own
formula editor), `smoke_languages.py` 25/25 (spelling in three languages found by itself,
Thesaurus, Editor pane, grammar marks and suggestions in German, English, French and
Spanish, Sprache and nicht prüfen as Word reads them). Unit tests: 117 (TypeScript) and 6
(Rust).

### Typing latency and memory (F1, F7)

| | RaabeWrite | Word | Step 4 (RaabeWrite / Word) |
|---|---|---|---|
| F1 typing, short document, median (p95) | 31.2 ms (66.7) | 33.1 ms (38.0) | 31.6 / 36.8 ms |
| F1 typing, long document, median (p95) | 32.2 ms (36.1) | 36.0 ms (46.2) | 32.0 / 34.3 ms |
| F7 memory with word-tables.docx (working set) | 462 MB | 291 MB | 346 / 302 MB |

The p95 tail of the short document is the Windows composition tail described in the Step 2
report. The extra memory since Step 4 is the language data and the grammar checker (the Rust
process now 115 MB instead of about 25 MB); the wrap-up measurement first showed 636 MB, which
led to the fix described below.

## Found and fixed during the wrap-up

- **Memory of the grammar checker** (D-052): compiled with the regex crate, the rules' tens of
  thousands of token expressions took 14-20 KB each: 174 MB for German, 336 MB for French. Word
  lists now become word sets and the rest runs on a small matcher of our own, with the regex
  crate only for the few hundred expressions used most; the four checkers now take 43, 22, 59
  and 19 MB, load in at most 0.36 s and find exactly what they found before.
- Grammar findings on a misspelled word now give way to the spelling mark, as in Word (the
  French rule for "marchè" had added a second card in the Editor pane).

## Known gaps

- Deferred by agreement (SCOPE, greyed out with "Zurückgestellt"): among others Künstlerische
  Effekte (Word's 23 artistic filters cannot be drawn close enough), Freihandgleichung, the
  online and Copilot commands.
- Freistellen removes the background around the edges without Word's interactive marking;
  Bildlayout (SmartArt) is not built; 3D rotation of pictures is approximated.
- Equation heights: matrices in brackets are 6-7 pt taller than Word's and nested display
  fractions 2.3-2.6 pt (D-047); Argumentgröße and Matrixabstand are not offered.
- Silbentrennung: 5 of 125 measured lines break differently from Word where both choices fit.
- Grammar: LanguageTool rules that need its Java chunker or filters are not used (English
  agreement through chunks, many French filter rules); our own rules cover the most common
  of those errors. Spanish could not be compared with Word on this PC.
- Deleting another author's tracked insertion still becomes a plain deletion (D-034).
