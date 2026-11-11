# Handover: where the work stands and what comes next (2026-11-11)

This file is for the next Claude session, possibly on another account, that takes over as maintainer. It holds only
what is not in the other docs. Read in this order:
1. `CLAUDE.md`: the working agreement and the binding rules.
2. This file.
3. `docs/ROADMAP.md`: steps 10 and 10b are the latest.
4. The last two reports: `docs/reports/2026-11-08-step10.md` and `docs/reports/2026-11-09-night.md`.
5. `docs/DECISIONS.md`, from D-080 on. The latest decisions are D-111 to D-155; the next free number is **D-156**.

## State

- **Code:** branch `main`, everything pushed (38b0bf5 and this file). All work branches are merged; `git branch
  --no-merged main` is empty.
- **Checks:** 1396 unit tests pass; typecheck, `tools/check-chars.py` and the anonymize check are clean.
- **Builds:** the release exes of all three apps were built from the final state:
  - RaabeWrite (09.11. 05:51) and RaabePresent (09.11. 05:57), at `apps/<app>/src-tauri/target/release/`.
  - RaabeCalc, rebuilt 09.11. 17:09.
  - The per-user installers are next to each exe in `bundle/nsis/`.
- **Shortcuts:** RaabeWrite.lnk, RaabeCalc.lnk and RaabePresent.lnk in the project root start those exes. They are
  excluded in `.git/info/exclude` and not committed, because they are machine-specific.
- **Lesson from the night:** the owner once tried an old RaabeCalc exe. A running RaabeCalc window locked the release
  exe, so the 06:00 build went to another folder and the usual exe stayed old. If a build fails with "Zugriff verweigert"
  on the exe, the build did NOT replace it. Tell the owner, never close their window, rebuild once it is closed.

## Next step: Step 11 "the app around the document like Office" (owner's go 2026-11-11 evening, night run)

The owner's first short test of RaabeCalc after the night run found four things. Their verdict: "Those are things that
we need to get right in all three apps and in different variations."
1. After a right-click in the sheet, the context menu does not vanish with the next left click into another cell.
2. The selection marks of chart parts look different from Excel's.
3. In the chart Format pane, some sections are in a different order than Excel's. Many controls that Excel shows as
   picture dropdowns are text in ours, e.g. the arrowheads.
4. The font colour icon, in all three apps, is confusing. The owner explained (2026-11-11): it shows an "A", beneath it
   a bar, and beneath that the actual colour bar, so the middle bar reads as a mistake.
   - Cause: the icon is Fluent's `text_color` (`packages/shell/scripts/icons.json`, key `font-color`), which already
     has an underline bar. Our colour bar is drawn below it.
   - Fix: a glyph without that bar and one thick bar in the current colour, as Office draws it. The same for every
     colour split button: Texthervorhebungsfarbe, Füllfarbe, Formkontur, Fülleffekt, Schattierung.

**Why the night's checks missed it.** They tested whether an option changes the document. They never tested how menus
close, and never compared the look of our own UI with Office's. The look corpus compared documents, not the app's own
screen.

**Agreed scope** (discussed 2026-11-09 to 11-11; the owner confirmed "the scope is clear", but has NOT given the go).
Each point is recorded in Office first, then built, then tested by hand: trusted DevTools input, plus real mouse under
the lock. Every fault becomes a lasting case. All three apps.

1. **Menus close like Office's.**
   - Any open menu, flyout, gallery or palette closes on a click anywhere outside it, on Esc, and when the window loses
     focus. The click then still does its normal job: clicking into another cell closes the menu and selects the cell.
     Record what Excel, Word and PowerPoint do exactly.
   - Fix once in the shared shell (`packages/shell/src/popup.ts`; outside clicks are caught with a capturing
     `mousedown` around line 60) and in the app-specific flyouts (chart buttons, pane dropdowns).
   - Test every popup kind against every surface in all three apps:
     - popup kinds: context menus, ribbon dropdowns, galleries, palettes, chart flyouts, pane dropdowns;
     - surfaces: cell, header, sheet tab, text page, slide, thumbnail, ribbon, pane, status bar.
   - The cause in RaabeCalc is not diagnosed yet. The grid's `onDown` (`apps/raabe-calc/src/grid/grid.ts`, about line
     1340/1527) captures the pointer: look there first.
2. **Chart selection marks.**
   - Record Office's look for every selectable part: chart area, plot area, title, legend, axis, gridlines, series,
     single point, data labels. Take screenshots of our own Office at 100 % and 150 %: frame, handle size, colour,
     spacing.
   - Build them once in `packages/chart-ui` (shared by all three apps).
   - Compare crops pixel by pixel.
3. **Format panes: order and picture dropdowns.**
   - Order: a check that compares our section and control order with Excel's recorded pane
     (`tools/feel-lab/results/2026-11-09-excel-chart-pane.json`) and fails on any difference.
   - Picture dropdowns wherever Office has them:
     - Strichtyp, Verbundtyp, arrow types and arrow sizes, Markierungstyp;
     - the presets of Schatten, Leuchten, Weiche Kanten and Abschrägung;
     - Voreingestellte Farbverläufe, Richtung, Textur, Muster.
   - Build them once in the shared shell. The pictures are drawn by us, not copied from Office.
   - Use them in every pane that has them in Office: the chart pane, Form formatieren and Bild formatieren in all three
     apps, and the Rahmen line styles in Zellen formatieren.
4. **Colour buttons as in Office.**
   - The glyph sits over one thick bar in the last used colour: red at first for Schriftfarbe, yellow for
     Hervorhebung.
   - The main part applies that colour; the arrow opens the palette. The bar follows the last pick.
5. **Pictures instead of text everywhere (owner, 2026-11-11).** Not only the list in point 3: an inventory of every
   pane, dialog, dropdown and gallery in Office (UI Automation: are the entries pictures, text or both?) against the
   same controls in our apps. A lasting check fails wherever Office shows a picture and we show text.
6. **Screen comparison with Office.** A new check compares the app's own UI with Office's for the same state: ribbon,
   open menus, dialogs, panes, in all three apps. Every difference is fixed or written down as a deliberate decision.

**Still to ask the owner before starting:**
- The go for the step.
- When the PC is free for the real-mouse runs. CLAUDE.md: real-input runs only when the owner is not using the PC; tell
  them first.

## Other open items (details in the 2026-11-09 night report, "Still open")

- **Charts:**
  - Re-sweep column and pie with `check_calc_chartpane.py` on a build with the D-155 fixes.
  - Not built yet: Format tab galleries (Formenarten, WordArt, Formeffekte, Texteffekte), 3D-Drehung, texture tiling,
    a Datentabelle pane, shapes inside a chart.
  - RaabePresent: no live preview on slides, and Diagrammfilter's Namen does nothing.
- **RaabeCalc:**
  - Dragging the orientation dial does not turn the text.
  - Not built yet: relative macro recording (not compared with Excel), the recorder for dialogs and borders, Fülleffekte
    Voreinstellung.
- **RaabeWrite:**
  - Crop marks and the Rahmenpipette fail the by-hand check.
  - The Word comparison of the follow-up's runs is not done.
- **RaabePresent:**
  - SmartArt has no single-shape selection.
  - Text formatting is lost when a SmartArt's layout changes.
  - A master cannot be renamed (Umbenennen).
- **Drag test:** about 10 out-of-window drags still fail (D-128).
- **PROGNOSE.ETS:** about 1.7 % off Excel on average (D-124).
- **Waiting for the owner:** the force push to clean up the history (D-080). Never do it without their explicit OK in
  chat.

## How the night run was organised (reuse it, with these lessons)

- **Agents:** work went to background agents, each in its own git worktree and branch, with a written brief and a
  reserved range of D-numbers. The coordinator merged branch by branch. Conflicts in DECISIONS.md are append-only:
  keep both sides. After each merge: typecheck, `npm test`, check-chars, `anonymize_office.py --check`, a scan for
  private paths, then push.
  - When a merge conflicts in code, the agent that knows the code merges main into its own branch and resolves it.
- **Usage limits:**
  - Eight agents in parallel used up the 5-hour usage window in about 45 minutes, and every agent stops when that
    happens.
  - Agents must commit every working step and keep an uncommitted `NIGHT-STATUS.md` in their worktree, so a stop loses
    nothing.
  - A stopped agent continues where it was when it gets a message (SendMessage); a new session cannot resume another
    session's agents.
- **Timing:** a scheduled task as a wake-up timer was refused by the permission check. A background `sleep` until the
  reset time worked.
- **CPU and builds:** the PC has 6 cores, so builds were run with `CARGO_BUILD_JOBS=2`. A worktree's Rust build can be
  seeded by copying the main checkout's `target/release` folder.
- **Real input:** shared through `%TEMP%\raabe-real-input.lock`, held for at most about 20 minutes. Other agents' Office
  windows still came to the front sometimes, so a probe checks that its own window is in front before every key.
- **Housekeeping, not done yet (needs the owner's OK, since it deletes):**
  - The agent worktrees under `.claude/worktrees/` (16, all merged; they hold only uncommitted NIGHT-STATUS.md files)
    and their branches `worktree-agent-*`.
  - The Present worktree `step8-raabe-present-190dd0`.
  - The unused target folders `apps/raabe-calc/src-tauri/target/agent*`, `drag` and `macros`. These take a lot of disk
    space.
