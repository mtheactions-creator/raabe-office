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
| F4 | **Layout fidelity** | Same document in Word and in our app; compare where every line and page breaks. Word's line breaks come from its COM interface. | ≥ 99 % identical line breaks for documents in our supported subset |
| F5 | **Start and load time** | Process launch → window shows a ready document; open a 30-page .docx. | Start < 1 s, open < 1 s |
| F6 | **Text rendering** | Screenshot comparison of identical text at the same zoom (sharpness, glyph positions). | No visible blur compared with Word |

## Conditions

- Development PC: Intel i5-8600K, 32 GB RAM, GTX 1070, Windows 10 Pro, 3840×2160 at
  150 % scaling. **The display currently runs at 30 Hz**, so screen-based timings (F1)
  are quantised to ~33 ms and need many samples; comparisons stay valid because both
  apps are measured on the same screen.
- Word: Microsoft 365, version 16.0.20430, 64-bit.
- The harness takes over keyboard, mouse and screen for a few minutes. It is only run
  when nobody is using the PC.

## Reference values

Measured results are recorded in `tools/feel-lab/results/` with date, versions and raw
samples, so we can track regressions over time. Reports that interpret them live in
`docs/reports/`.

| Date | Word | Prototype B | Report |
|---|---|---|---|
| 2026-10-13 | F1 36 ms · F5 1185 ms | F1 32 ms · F3 50/50 · F4 2000/2000 lines, 50/50 pages · F5 526 ms | [Step 0](reports/2026-10-13-step0-feel.md) |

## Running the harness

```
cd tools/feel-lab
.venv/Scripts/python.exe run.py fidelity    # no input needed
.venv/Scripts/python.exe run.py startup     # opens and closes windows
.venv/Scripts/python.exe run.py latency     # takes over keyboard and screen
.venv/Scripts/python.exe run.py behaviour   # takes over keyboard, mouse and screen
```

Rebuild the prototype app after any frontend change (`npx tauri build --no-bundle` in
`prototypes/`), otherwise the harness measures the old build.

## Lessons

- Never restart a CSS/Web animation on every keystroke: it delayed the frame by one
  refresh. The caret blinks with a timer.
- A test is only valid if the pixels belong to the window under test: an earlier
  start-time run measured a window that was still open behind it.
