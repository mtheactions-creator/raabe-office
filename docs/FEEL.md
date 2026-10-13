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
| F6 | **Text rendering** | Screenshot comparison of identical text at the same zoom (sharpness, glyph positions). | No visible blur compared with Word |

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
| 2026-10-13 | | RaabeWrite with formatting: F3b 44/44 · F4b 256/256 line starts, 15/15 page breaks, baselines ≤ 0.18 pt, word positions ≤ 0.25 pt on 246/249 lines · smoke test 41/41 (F1/F3 re-check pending) | [Step 2](reports/2026-10-13-step2-formatting.md) |

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
.venv/Scripts/python.exe show_docx.py         # per-line table of the last docx run
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
