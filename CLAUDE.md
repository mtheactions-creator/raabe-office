# Raabe Office: notes for Claude

You are the maintainer of this project (D-016). You do not remember earlier sessions;
this repository is the memory. Read this file first, then `docs/DECISIONS.md`.

## What this is
A lean, offline office suite for the Raabeschule (Gymnasium in Niedersachsen, Germany)
that looks and behaves like Microsoft 365 so students can use their Office muscle memory:
**RaabeWrite** (Word), **RaabeCalc** (Excel), **RaabePresent** (PowerPoint).
Windows first, iPad close behind (new students all use iPads). Open source, GPL-3.0.

## Working agreement with the project owner (a teacher at the school)
- **No new step starts without explicit agreement.** Discuss, then wait for a clear "go".
- Ask before installing software on the development PC.
- Never handle passwords or credentials. The owner authenticates `gh` themselves.
- The owner writes in English; the apps' UI is German, using Microsoft's exact German
  labels and shortcuts.
- Feel is judged by measurement against real Word (`docs/FEEL.md`), not by asking the
  owner to test consciously. Only `run.py latency` and `run.py behaviour` take over
  keyboard, mouse and screen: tell the owner first and only run them when they are not
  using the PC. Everything else (`docx`, `formatting`, `fidelity`, `smoke_write.py`) runs
  invisibly in the background (D-028).

## The repository is public
- Results (`tools/feel-lab/results/`) must not contain private paths (`C:\Users\...`),
  user names or the owner's files; check every new results file before committing. Test
  files are built by the harness in TEMP, never taken from the owner's folders.
- Never copy Microsoft's prose (tooltips, help texts, dialog explanations): labels and
  shortcuts as in Office, descriptions in our own words. Full recordings of Office's texts
  stay in `tools/feel-lab/captures/`, which is not committed (D-058).
- Probes read UI only from the window of the Office instance they started, never from other
  windows on the screen.
- Never accept consent, sign-in, licence or sync dialogs in Office (or anywhere else); if
  one blocks a probe, stop and tell the owner.

## How an app is built (method from Steps 4b-4d)
1. Discuss the step with the owner, agree packages and decisions, wait for the go.
2. Record the behaviour first, in a hidden Office instance of our own (COM, or real input
   when the owner is away): what happens, exact sizes, positions, menu entries. Probes are
   `tools/feel-lab/probe_word_*.py`; their results go to `results/<date>-<name>.json`.
3. Build it the same way. Where we deliberately differ (an improvement, or our own look),
   write it down as a decision.
4. Compare automatically: behaviour catalogue cases (`catalogue.py`, `catalogue_math.py`)
   that run the same input in Office and in our app, layout against Office's own documents
   (`run.py docx`), formatting (`run.py formatting`), smoke tests. Add cases for every new
   behaviour, especially for bugs the owner finds.
5. Before committing: unit tests, type check, `tools/check-chars.py`, the results scan above;
   docs (DECISIONS, ROADMAP, FEEL) in the same commit.
6. The owner then tries the app and sends a list of differences; each list is discussed
   (decisions for the open questions), then done in one run and measured again.
Report numbers honestly, including what failed and why (e.g. a known Office glitch).

## Where things are
- Repository: https://github.com/mtheactions-creator/raabe-office (public, branch `main`).
  `gh` lives at `C:\Program Files\GitHub CLI\gh.exe` if it is not on PATH.
- `docs/DECISIONS.md`: every decision with its reasoning. Add an entry for each new one.
- `docs/SCOPE.md`: what each app does (MVP / later / out of scope).
- `docs/ARCHITECTURE.md`: technology and structure.
- `docs/ROADMAP.md`: steps and their status.
- `docs/FEEL.md`: feel metrics and targets.
- `branding/`: logos; regenerate with `python branding/generate.py`. Not GPL.
- `tools/feel-lab/`: measurement harness (Python venv in `tools/feel-lab/.venv`).
- `prototypes/`: Step 0 feel prototypes (throwaway code).

## Environment (development PC)
Windows 10 Pro, PowerShell + Git Bash, Node 24, Rust (stable, MSVC), Python 3.14,
Visual Studio 2022 C++ tools, WebView2, Microsoft Word 365 (COM: `Word.Application`).
PowerPoint and Excel 365 are installed too (`POWERPNT.EXE`, `EXCEL.EXE` in
`C:\Program Files\Microsoft Office\root\Office16`), as reference apps for RaabePresent and
RaabeCalc; the harness-safety rules below apply to them the same way.
Display: 4K at 150 % scaling, currently 30 Hz.

## Harness safety
- So far the harness knows only Word (`WordTarget`, `docx_fidelity.py`, the catalogues) and
  `packages/ooxml` reads and writes only .docx. A new app needs its own reference side
  (PowerPoint/Excel through COM, `DispatchEx`) and its own file format first.
- The harness always starts its **own** Word process (`DispatchEx`, `WINWORD.EXE /x`) and
  closes only what it opened. Never attach to the owner's running Word
  (`GetActiveObject`/`Dispatch`) and never close documents you did not create; the
  owner may have unsaved work open.
- After a failed run, check for leftover `WINWORD.EXE` / `raabe-feel.exe` processes and
  windows of the test document, and close only those.
- Some of Word's settings outlive the probe's own instance and reach the owner's Word: a reduced
  ribbon (double-click on a tab), Alle anzeigen. A probe that changes them puts them back
  before quitting (`probe_word_dblclick.py`).

## Building and running
- `npm install` at the repo root (workspaces). `npm test` runs the unit tests,
  `npm run typecheck` the type check.
- `npm run dev:write` starts RaabeWrite in a browser (development fallback platform).
- Real app: `npx tauri build --no-bundle` in `apps/raabe-write` (release exe in
  `src-tauri/target/release/raabe-write.exe`); `npx tauri build` also makes the per-user
  NSIS installer in `src-tauri/target/release/bundle/nsis/`. Cargo is in `~/.cargo/bin`
  (not on Git Bash's PATH: `export PATH="$HOME/.cargo/bin:$PATH"`).
- Icons: edit `packages/shell/scripts/icons.json`, then `node packages/shell/scripts/build-icons.mjs`.
- After any frontend change, rebuild the exe before running feel-lab or the smoke test.
- If the owner has RaabeWrite open, the release exe cannot be replaced: build into another
  folder (`CARGO_TARGET_DIR=<repo>/apps/raabe-write/src-tauri/target/check`) and point the
  harness at it with `RAABE_WRITE_EXE=<that folder>/release/raabe-write.exe`. Never close the
  owner's window. Test instances skip the single-instance handling (D-060).
- Test documents for layout are built by Word itself: `tools/feel-lab/make_fixtures.py`
  (`--layout`, `--fonts`, `--breaks`, `--top`, ...) writes them to `packages/ooxml/test/fixtures/`.

## Pitfalls seen so far
- Backslash escapes get mangled on the way into files: the Bash tool turns `\\n` inside
  heredocs and `sed` expressions into real line breaks, and the Write/Edit tools turn
  printable `\uXXXX` escapes (`\u00ad`, `\u00a0`, `\ufffc`) into the invisible characters.
  Write patch scripts with the Write tool as files (no heredoc), prefer `chr(10)` /
  `String.fromCharCode(...)` for special characters, and run `python tools/check-chars.py`
  (`--fix` rewrites them as escapes) before committing.
- Python's `write_text` writes CRLF on Windows; pass `newline="\n"`.
- Word's COM reports are coarse: `Information(5/6)` and `GetPoint` are in screen pixels,
  `Information(5)` ignores alignment, a range's `Font` ignores spaces when deciding whether
  it is uniform, and `Range.Text` shows "all caps" text in capitals and leaves out hidden
  text unless `TextRetrievalMode.IncludeHiddenText` is set. For exact positions read Word's
  PDF export (`feellab/pdftext.py`).
- The canvas returns hinted whole-pixel widths for some fonts at small sizes (Cambria);
  measure at a large size and scale (`Fonts.measureFont`).
- `fontdb`'s query matches family names case-sensitively; the engine asks in lower case
  (`fonts.rs` `find_face`).
- A WebView2 window that is hidden or covered stops `requestAnimationFrame`; the browser
  pane of the desktop app does too while it is hidden, so ribbon state there may look stale.
  In the off-screen harness window rAF callbacks run late: a deferred focus change must check
  that the focus has not moved meanwhile (comment reply box, D-034).
- Word's COM positions include hidden field codes (TOC, hyperlinks, PAGEREF) and every
  character inside an equation; `Content.Text` leaves the codes out. `docx_fidelity.py` maps
  them; documents with fields or equations need `lines_by_position`.
- MathLive takes settings such as `menuItems` only once the field is in the page. The CSS
  generic `math` font falls back to Times for digits: name "Cambria Math" explicitly.
- Alt+letter is a keytip, Alt+Umschalt+… a shortcut (Alt+Umschalt+0 Formel).
- Office dropdowns: UI Automation shows a gallery's entries below the expanded ribbon control
  in the window's tree, not in the popup window; nested menus (Katalogfilter, context
  submenus) only via `ElementFromPoint`. Submenus open on click, not reliably on hover.
- pywin32 out parameters: `Window.GetPoint(0, 0, 0, 0, range)`. `Range.Text` of an equation
  gives math italic letters (𝑥, 𝜃); `winapi.press(combo, hold)` takes a hold time, not a count.

## Conventions
- Internal docs and code comments in English; UI strings in German.
- Commit messages in English, imperative mood.
- Keep `docs/` current in the same commit as the change it describes.
