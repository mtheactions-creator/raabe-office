# Measuring "feel"

Decision D-013: we do not rely on gut feeling. Subtle differences in feel are noticed
subconsciously during normal use and vanish as soon as someone tests consciously. So we
turn them into numbers and measure Microsoft Word and our apps under identical conditions.
The harness lives in `tools/feel-lab/`.

## Metrics

| ID | Metric | How | Target |
|---|---|---|---|
| F1 | **Keystroke-to-screen latency** | Real key events via `SendInput`; the screen region where the character appears is polled until it differs from every pre-press state (so a blinking caret is not mistaken for the character). 100+ samples, median and p95. | Not slower than Word; median < 1 frame above Word |
| F2 | **Frame stability** while typing, scrolling, dragging, resizing | Inside our app: `requestAnimationFrame` deltas plus the Long Animation Frames API during scripted input. | No frame longer than one refresh interval in a 50-page document |
| F3 | **Behaviour conformance** | A catalogue of micro-behaviours (double-click, triple-click, drag-select word snapping, Strg+arrow movement, Shift+click, list handling, …). Each case sends *real* mouse/keyboard input to Word, reads the resulting text and selection through Word's COM interface, then replays the same input in our app and compares. | 100 % of the catalogue |
| F3b | **Formatting behaviour** | What formatting commands do at the caret and on selections (word under the caret, toggles, paragraph marks, Enter, typing afterwards). Word runs the commands its shortcuts are bound to (WordBasic, found with `FindKey`); our app gets the shortcut as key events. Every character's formatting, the paragraph values, the text, the selection and the formatting for typing are compared. | 100 % of the catalogue |
| F4 | **Layout fidelity** | Same document in Word and in our app; compare where every line and page breaks. Word's line breaks come from its COM interface. | ≥ 99 % identical line breaks for documents in our supported subset |
| F4b | **Layout fidelity on .docx files** | The same Word-made file opened in both. Line and page starts (COM), left edge of every line and the x of every word start (COM at 400 % zoom), and the baseline of every line from Word's PDF export (exact glyph positions, read by `feellab/pdftext.py`). | 100 % identical line and page starts; baselines within 0.2 pt |
| F5 | **Start and load time** | Process launch → window shows a ready document; open a 30-page .docx. | Start < 1 s, open < 1 s |
| F6 | **Text rendering** | The same lines in Word and our app at 100 %, captured from the screen (`run.py rendering`): edge sharpness, ink, colour fringes (ClearType) and the horizontal ink profile of each line. | Sharpness ≥ 95 % of Word's, same smoothing, profile correlation ≥ 0.95 |
| F7 | **Memory** | Working set of the whole process tree with the same document open (`run.py memory`). | Not far above Word |

## Conditions

- Development PC: Intel i5-8600K, 32 GB RAM, GTX 1070, Windows 10 Pro, 3840×2160 at
  150 % scaling. **The display currently runs at 30 Hz**, so screen-based timings (F1)
  are quantised to ~33 ms and need many samples; comparisons stay valid because both
  apps are measured on the same screen.
- Word: Microsoft 365, version 16.0.20430, 64-bit.
- F1 and F3 take over keyboard, mouse and screen for a few minutes; they are only run when
  nobody is using the PC. Everything else runs invisibly (Word without a window, RaabeWrite
  off screen and never focused, D-028).

## Reference values

Measured results are recorded in `tools/feel-lab/results/` with date, versions and raw
samples, so we can track regressions over time. Reports that interpret them live in
`docs/reports/`.

| Date | Word | Prototype B | Report |
|---|---|---|---|
| 2026-10-13 | F1 36 ms · F5 1185 ms | F1 32 ms · F3 50/50 · F4 2000/2000 lines, 50/50 pages · F5 526 ms | [Step 0](reports/2026-10-13-step0-feel.md) |
| 2026-10-13 | F1 34-36 ms · F5 1215 ms (opening a .docx) | RaabeWrite 0.1.0 with full shell: F1 32 ms (9/120 late) · F3 49/50 (the 50th a Word glitch) · F4 2000/2000 lines, 50/50 pages · F5 603 ms · smoke test 32/32 | Step 1 |
| 2026-10-13 | | RaabeWrite with formatting: F3b 44/44 · F4b 256/256 line starts, 15/15 page breaks, baselines ≤ 0.18 pt, word positions ≤ 0.25 pt on 246/249 lines · smoke test 41/41 | [Step 2](reports/2026-10-13-step2-formatting.md) |
| 2026-10-13 | F1 35.3 ms · F7 300 MB | RaabeWrite with lists, tables, sections, notes, pictures: F4b 632/632 line starts in ten documents (first reported as 690: a summing error), every page break, baselines ≤ 0.18 pt · F6 sharpness 97 %, fringes 3.4 vs 3.7, correlation 0.97 · F1 31.7 ms · F7 345 MB · smoke tests 28/28 and 41/41 | [Step 3](reports/2026-10-13-step3-documents.md) |
| 2026-10-13 | F1 36.8/34.3 ms · F7 302 MB | RaabeWrite with notes, tables of contents, comments, tracked changes, equations, templates: F4b 682/682 line starts in eleven documents, every page break (text after equations within 0.8 pt) · F1 31.6/32.0 ms · F7 346 MB · smoke tests 43/43, 28/28 and 41/41 | [Step 4](reports/2026-10-13-step4-school.md) |
| 2026-10-14 | | RaabeWrite Step 4b A/B, real input: F3 behaviour catalogue 193/193 (typing and AutoCorrect, AutoFormat lists, keys, mouse, delete, lists, tables, styles, breaks, formatting, undo, clipboard; first run 134/161) · clipboard with Word 0 differences both ways (lists, levels, indents, column widths) | Step 4b |
| 2026-10-14 | | RaabeWrite Step 4b D, real input: formula editor 114/114 recorded Word cases (build-up, Math AutoCorrect, arrows, Up/Down, Home/End, Strg+arrows, Backspace/Entf, selection, Enter) · formula UI check (mouse, galleries, symbols, context menu, Konvertieren, box menu, Word reads the file) · word-math.docx 73/73 line starts, 3/3 pages (equation heights: D-047) · Word's 24 equations written back with Word's OMML structure | Step 4b |
| 2026-10-15 | F1 33.1/36.0 ms · F7 291 MB | RaabeWrite Step 4b done: F4b 960/965 line starts in 19 documents (the 5 others: hyphenation choices, D-050), every page break but one after them · F3 behaviour catalogue 192/193 (the 193rd: Word's triple-click did not register) · formula editor 114/114 · F1 31.2/32.2 ms · F7 462 MB (grammar checker memory cut from 636 MB, D-052) · grammar de/en/fr/es against Word (D-051, D-052) · smoke tests 41/41, 28/28, 43/43, 25/25 | [Step 4b](reports/2026-10-15-step4b-word-parity.md) |
| 2026-10-15 | F1 36.3/35.5 ms · F7 290 MB | RaabeWrite Step 4c done: F4b 1216/1221 line starts in 22 documents (the 5 others: hyphenation choices, D-050), every page break but one after them, floating tables within 0.3 pt · F3 behaviour catalogue 202/203 (10 new cases; the 203rd: Word's triple-click did not register) · F3b 44/44 · formula editor 146/146 · clipboard with Word: equations 0 differences both ways, pictures and links both ways · F1 33.5/31.8 ms · F7 489 MB · smoke tests 41/41, 28/28, 43/43, 25/25 | [Step 4c](reports/2026-10-15-step4c-polish.md) |
| 2026-10-16 | F1 35.8/35.7 ms | RaabeWrite Step 4d done: F4b 1216/1221 line starts in 22 documents (unchanged) · F3 behaviour catalogue 206/207 (4 new cases; the 207th: Word's triple-click did not register) · F3 49/50 · F3b 44/44 · formula editor 152/152 (6 new cases) · script sizes 10.99/8.02/6.6 pt vs Word 11.04/8.04/6.48 · F1 32.2/32.1 ms · smoke tests 41/41, 28/28, 43/43, 25/25 | Step 4d (D-059) |
| 2026-10-26 | | RaabeWrite fixes after Step 4d (D-060): Zeilen verteilen 3/3 recorded cases as in Word, the sizing handle keeps a new table's height (`check_table_rows.py`) · F4b 1216/1221 (unchanged) · smoke tests 41/41, 28/28, 43/43, 25/25 | D-060 |
| 2026-10-27 | Excel 365 (German) | RaabeCalc Step 7, real input: F3 behaviour catalogue 142/142 (`catalogue_calc.py`: typing, Eingeben/Bearbeiten/Zeigen, moving, selecting, pointing at cells, F4, AutoSumme, filling, moving and copying cells, shortcuts, sheets) · formulas 4229/4234 Excel results · number formats and input ≈ 66 000 cases · .xlsx 0 differences in Excel (17 files) · chart axes 145/145 · speed with 20 000 rows (`perf_calc.py`): one typed value 11 ms (66 ms when 20 000 chained formulas depend on it), redraw while scrolling 6 ms median, 14 ms worst, full recalculation 0,26 s · printing (`check_calc_print.py`, PDF from both): pages, lines per page and line spacing as Excel in 3/3 layouts, first line within 0,4 pt, left edge within 0,8 pt · clipboard with Excel both ways: values, formats, bold as in Excel (`check_calc_clipboard.py`) · F5 start 579 ms vs Excel 936 ms · F7 365 MB vs Excel 240 MB (`startup_calc.py`) | Step 7 (D-062, D-063) |
| 2026-10-13 | F1 66.1 ms (both apps two frames today; Step 1: one) · 0/120 late beyond that | RaabeWrite Step 2, real input: F1 66.3 ms, but 8/120 keystrokes two frames later than Word's worst (JS ≤ 4 ms per keystroke; the delay is in presentation, open) · F3 49/50 (the 50th the known Word glitch) · keyboard toggles 9/9 · mouse on palettes, lists, menus, gallery, dialogs, painter 13/13 | Step 2 |

## Running the harness

```
cd tools/feel-lab
.venv/Scripts/python.exe run.py fidelity      # no input needed, invisible
.venv/Scripts/python.exe run.py docx [f.docx] # no input needed, invisible (default: word-layout.docx)
.venv/Scripts/python.exe run.py formatting    # no input needed, invisible
.venv/Scripts/python.exe smoke_write.py       # no input needed, invisible
.venv/Scripts/python.exe run.py startup       # opens and closes windows
.venv/Scripts/python.exe run.py latency       # takes over keyboard and screen
.venv/Scripts/python.exe run.py behaviour     # takes over keyboard, mouse and screen
.venv/Scripts/python.exe keys_check.py        # takes over keyboard and screen (alignment toggles)
.venv/Scripts/python.exe ui_check.py          # takes over mouse and screen (ribbon controls, dialogs)
.venv/Scripts/python.exe catalogue.py [area]  # takes over keyboard, mouse and screen (behaviour catalogue, F3)
.venv/Scripts/python.exe catalogue_math.py [--word-only | --ours-only]  # keyboard and screen (formula editor vs Word)
.venv/Scripts/python.exe check_math_ui.py    # takes over mouse and keyboard (formula editor UI)
.venv/Scripts/python.exe clipboard_roundtrip.py  # clipboard with Word both ways (RaabeWrite window)
.venv/Scripts/python.exe compare_ui.py [--hover] # takes over mouse and screen (ribbon side by side)
.venv/Scripts/python.exe check_tables.py      # takes over mouse and screen (table handles, row heights)
.venv/Scripts/python.exe check_pictures.py    # takes over mouse, keyboard and clipboard (pictures)
.venv/Scripts/python.exe check_shapes.py      # invisible: every preset shape against Word's PDF (D-046)
.venv/Scripts/python.exe check_shapes_render.py  # RaabeWrite's drawing of Word's shapes against Word's (PDF)
.venv/Scripts/python.exe check_shapes_ui.py   # takes over mouse and keyboard (drawing, typing into shapes, Word reads it)
.venv/Scripts/python.exe show_docx.py         # per-line table of the last docx run
.venv/Scripts/python.exe devpage.py start      # dev build (npm run dev:write) in a headless InPrivate Edge:
.venv/Scripts/python.exe devpage.py shot x.png # screenshots, clicks, drags, keys, JavaScript; invisible
.venv/Scripts/python.exe compare_dev.py f.docx [--para N] [--recompute]  # line starts: Word vs the dev build
```

Rebuild the app after any frontend change (`npx tauri build --no-bundle` in
`apps/raabe-write`), otherwise the harness measures the old build. Test documents are
built by Word itself with `make_fixtures.py`.

## Interface fidelity

`capture_word_ui.py` reads Word's ribbon, backstage and status bar through UI Automation
(names, sizes, positions, keytips) and takes screenshots; the results stay local in
`captures/` because they show Microsoft's interface and the owner's recent files. RaabeWrite
is screenshotted at the same window size and compared side by side.

## Lessons

- Never restart a CSS/Web animation on every keystroke: it delayed the frame by one
  refresh. The caret blinks with a timer.
- A test is only valid if the pixels belong to the window under test: an earlier
  start-time run measured a window that was still open behind it.
- Test runs end the app abruptly, which leaves AutoRecover backups. The next run then opens
  with the recovery dialog over the document and every click lands on it. Test mode
  therefore uses its own recovery folder, cleared before each run.
- Word occasionally ignores a click; the behaviour test runs each Word case twice (three
  times on disagreement) and keeps the majority.
- Word's COM positions are whole screen pixels (0.5 pt at 150 % scaling) and its line
  boxes include paragraph spacing; exact baselines come from its PDF export instead.
- Word reports a range's formatting as uniform even when the spaces in it differ; the
  formatting test reads spaces one by one.
- Fonts arrive asynchronously (the font database loads after start-up): wait for
  `__probe.settled()` before measuring a freshly opened document.
- To split screen latency into "browser late" and "Windows late", `run.py latency --trace`
  also records the browser's Event Timing report per keystroke and a pipeline trace
  (`captures/traces/`); `--variants` adds RaabeWrite runs with other WebView2 presentation
  flags. `python -m feellab.presentation` measures event → presented frame without real
  input. On 2026-10-13 the browser presented every keystroke within two refreshes; the
  occasional four-frame delays seen on screen arise in Windows' composition.
