# Step 2 report: RaabeWrite formatting

Date: 2026-10-13 · Measured with `tools/feel-lab` (see `docs/FEEL.md`) · Raw data in
`tools/feel-lab/results/2026-10-13-docx.json` and `-formatting.json`

## What was built

- **Formatting model:** character formatting (font, size, bold, italic, underline: single,
  double, thick, dotted, dashed, wavy, words only; strikethrough and double strikethrough, super-/subscript, colour incl. theme
  colours, highlight, caps, small caps, hidden, spacing, kerning, character styles) and
  paragraph formatting (alignment, indents incl. first-line/hanging, spacing before/after,
  line spacing in all four kinds, tab stops incl. centre/right/decimal, keep with next,
  keep lines together, widow control, page break before, contextual spacing), resolved
  through Word's style hierarchy. Word 2016-2019's built-in styles with German names.
- **Start tab:** font and size boxes with font list and live values, Fett/Kursiv/
  Unterstreichen (with menu), Durchgestrichen, Tief-/Hochgestellt, Alle Formatierungen
  löschen, Texthervorhebungs- and Schriftfarbe with Word's palettes, Groß-/Kleinschreibung
  menu, Schriftart vergrößern/verkleinern; Einzug, Ausrichtung, Zeilen- und Absatzabstand
  menu, Alle anzeigen (formatting marks); style gallery with 16 styles; Format übertragen.
  Dialogs Schriftart and Absatz. **Layout tab:** indent and spacing spinners.
- **Keyboard:** Word's shortcuts as Word itself reports them on this PC (German keyboard),
  e.g. Strg+Umschalt+F/K/U, Strg+< / Strg+Umschalt+<, Strg+Leertaste, Strg+L/E/R/B,
  Strg+M, Strg+1/2/5, Umschalt+F3, Strg+Alt+1-3, Strg+Alt+C/V, Umschalt+Eingabe, Strg+Eingabe.
- **.docx:** formatting is read and written; everything RaabeWrite does not edit is kept
  exactly as Word wrote it (D-024). Saving into the original file is therefore allowed again.
- **Harness:** layout comparison on real .docx files with Word's PDF output as the exact
  reference, a formatting-behaviour catalogue with Word's own commands as reference, and a
  background mode: apart from latency and real-input behaviour, every test now runs without
  taking over the PC (D-028).

## Results

### Layout (F4b): four test documents built by Word

| Document | What it tests | Line starts | Page breaks | Baselines (Word's PDF) | Word positions in lines |
|---|---|---|---|---|---|
| word-layout.docx | worksheet: title, headings, mixed sizes, justified text with soft hyphens, lists with hanging indent, exact/at-least/1.5 spacing, tabs incl. decimal, mixed fonts, super/subscript, keep together, page break before | 90/90 | 4/4 | 90/90 within 0.18 pt | 87/88 within 0.25 pt ² |
| word-fonts.docx | 13 fonts × spacing rules, sizes 8-36 pt, mixed fonts and sizes in one line, super/subscript | 80/80 | 3/3 | 71/71 within 0.18 pt ¹ | 78/80 within 0.25 pt ² |
| word-breaks.docx | space before after natural, manual and "page break before" breaks, a break inside a paragraph | 81/81 | 6/6 | 81/81 within 0.17 pt | 76/76 within 0.07 pt |
| word-top.docx | space before on the first page and after a page-break character | 5/5 | 2/2 | 5/5 within 0.06 pt | 5/5 within 0.06 pt |

Word's PDF coordinates are quantised to 0.12 pt, so ≤ 0.2 pt is "identical". In total:
**256/256 line starts and 15/15 page breaks identical.**

¹ The other 9 lines are on a page whose lines could not be paired with Word's PDF.
² The exceptions are kerning at 28 and 36 pt (0.6-0.7 pt over a line).

### Formatting behaviour (F3b): 44/44 cases identical to Word

Word's own commands (the ones its shortcuts are bound to) against RaabeWrite's shortcuts,
comparing every character's formatting, paragraph values, text, selection and the
formatting for typing: bold/italic/underline at the caret, at word ends and on mixed
selections, all underline kinds, hidden, grow/shrink (also by one point), Strg+Leertaste,
caps, small caps, subscript, Umschalt+F3 cycles, Enter after formatted text, joining
paragraphs with Backspace/Entf, headings, alignment, Strg+Q, indents, hanging indents and
line spacing.

### Smoke test: 41/41

Typing, undo, styles, formatting shortcuts, Format kopieren/einfügen, saving and Word reading
the result (bold word, red 12 pt text), a Word file with a table edited and saved over the
original with the table intact, keytips, search, backstage, PDF, ribbon collapsing,
AutoRecover after a crash, close prompt. All off screen.

### Files

Word opens all seven test documents after an edit-and-save by RaabeWrite and sees them
unchanged (tables, image, shape, footnote, comment, hyperlink, revisions, equation, fields,
bookmark, lists, header; `docx_word_check.py`). 33 unit tests.

## What the measurements taught us

The rules are recorded in D-025 and D-026. The ones that were not obvious:

1. **Line height** is the larger of two font tables (hhea and OS/2 Windows), and the space
   below the baseline is always the Windows descent. Using only hhea was right for Calibri,
   Times and Arial but wrong for Century Gothic, Book Antiqua and Trebuchet MS.
2. **"Genau" spacing** puts the baseline at 80 % of the line for every font.
3. **"Abstand vor"** is dropped at the top of every page except the first, also after a
   manual page break or "Seitenumbruch oberhalb".
4. **Character widths:** the browser's canvas reports hinted whole-pixel widths for some
   fonts (Cambria lines came out 0.7 % too wide and broke differently); measuring at a large
   size gives the design widths, which both Word and the browser's own text drawing use.
5. **Super-/subscript** is 65 % of the size rounded to half points, not two thirds.
6. **Blocksatz** lets a line overflow by up to 25 % of the width of its spaces and narrows
   them; without this one justified line in the worksheet broke one word earlier.
7. **Paragraph marks:** formatting a paragraph's text does not format its mark; typing at
   the end does. Enter gives the first paragraph a new mark with the formatting at the caret
   and passes the old mark on, so "Fett, Eingabe, tippen" does not continue in bold unless
   bold text was typed first.
8. **Strg+Umschalt+G** is "Großbuchstaben" formatting in Word, not a change of the letters;
   Schriftgrad verkleinern skips 10.5.
9. A bug outside the engine: the font database matched family names case-sensitively, so
   every font apart from Calibri, Times and Arial silently used rough browser estimates.

## Known limitations

- **Aptos** (default of documents from current Microsoft 365 templates) is an Office cloud
  font, not installed in Windows, and has no metric-compatible free substitute. Such
  documents use a fallback font and break lines differently (D-027).
- In our own PDF/print output, text baselines can be up to about 1 pt off the exact layout
  because the browser snaps text to its pixel grid; on screen this is below one pixel.
- Kerning at large sizes differs slightly from Word (up to 0.7 pt over a 36 pt line).
- Borders and shading of paragraphs, lists, tables and images are Step 3.

## Real-input session (same day, PC handed over by the owner)

| Check | Word | RaabeWrite |
|---|---|---|
| F1 keystroke → screen, median (end of short doc / page 12 of 50) | 66.1 / 66.2 ms | 66.3 / 66.0 ms |
| F1 keystrokes ≥ 4 frames | 0 / 120 | 8 / 120 |
| F3 behaviours (50 cases, real mouse and keys) | reference | 49 / 50 ¹ |
| Alignment shortcuts pressed once, twice, mixed (`keys_check.py`) | reference | 9 / 9 |
| Mouse on palettes, font and size lists, underline menu, gallery, Absatz and Schriftart dialogs, format painter (`ui_check.py`) | | 13 / 13 |

¹ The miss is the known Word glitch from Step 1: Word did not react to that double-click.

- Both apps needed two display frames today where they needed one in Step 1, so the
  screen pipeline itself was slower; the medians are equal.
- RaabeWrite has a tail Word does not have. A CPU profile while typing in the 50-page
  document shows about 3.5 ms of our code per keystroke and the frame request answered
  within one refresh, so the delay arises after our code, in composition or presentation.
  **Open:** trace WebView2's GPU process to find it.
- Real keystrokes showed one more Word rule: Strg+L on a left-aligned paragraph switches
  to Blocksatz (fixed; the COM command does not do this).
