# Handover: where the work stands and what comes next (2026-11-12)

This file is for the next Claude session, possibly on another account, that takes over as maintainer. It holds only
what is not in the other docs. Read in this order:
1. `CLAUDE.md`: the working agreement and the binding rules.
2. This file.
3. `docs/ROADMAP.md`: steps 10, 10b and 11 are the latest.
4. The last report: `docs/reports/2026-11-11-step11-night.md` (and `2026-11-09-night.md` before it).
5. `docs/DECISIONS.md`, from D-080 on. The latest decisions are D-156 to D-202 (D-167 to D-169, D-174 to D-179, D-184 and D-199
   were reserved and not used); the next free number is **D-203**.

## State

- **Code:** branch `main`, everything pushed. Tonight's agent branches (`worktree-agent-aaa061bda6b0729e4`,
  `worktree-agent-a09a757be7869ec48`) are merged.
- **Checks:** 1423 unit tests pass; typecheck, `tools/check-chars.py` and the anonymize check are clean.
- **Builds:** the release exes of all three apps were rebuilt from the final state of the night (12.11.), at
  `apps/<app>/src-tauri/target/release/`. The per-user installers in `bundle/nsis/` are still from 09.11.
- **Shortcuts:** RaabeWrite.lnk, RaabeCalc.lnk and RaabePresent.lnk in the project root start those exes. They are
  excluded in `.git/info/exclude` and not committed, because they are machine-specific.
- **Lesson:** a running RaabeCalc window locks the release exe. If a build fails with "Zugriff verweigert" on the exe,
  the build did NOT replace it. Tell the owner, never close their window, rebuild once it is closed.

## Step 11 is built (night 2026-11-11/12); next: the owner tries it

All five parts and the pictures sweep are done and measured (report: `docs/reports/2026-11-11-step11-night.md`). The
lasting checks, all in `tools/feel-lab`, rerun after any change to the shell, the ribbon or the panes:
- menus: `check_menuclose_calc.py`, `_write.py`, `_present.py` (DevTools input, invisible);
- colour buttons: `check_colorbuttons.py` (real input, crops);
- chart marks: `check_chart_selection.py`;
- pictures: `check_picture_controls.py` (needs the three dev servers), `paneorder.test.ts`, `panepictures.test.ts`;
- screen: `check_ribbon_geometry.py` (all main tabs; `--context` for the contextual tabs; `-v` for details),
  `check_ribbon_galleries.py` (inline galleries as the window narrows), `check_contextmenu_geometry.py` (real input),
  `check_pane_geometry.py`, `check_dialog_geometry.py`.

One decision the owner should hear about first: **a left click outside a menu in the document is swallowed** (it closes
the menu and does nothing else), because Excel, Word and PowerPoint all do that (recorded, D-156). The scope had assumed
the click would also select the cell. Right clicks and clicks on the ribbon do act.

Next step as usual: the owner tries the apps and sends a list of differences; discuss it, then do it in one run.

**Waiting for the owner:**
- The housekeeping deletes (old worktrees, branches, target folders; see the report). The automatic permission check
  refused them although the owner had allowed them; ask the owner to run them or allow them.
- Cleaning three October result files that hold an Office instruction sentence (see the report).

## Step 12 (2026-11-12, owner's requests after Step 11): done

- Gridlines: horizontal and vertical, major and minor, each on its own in every chart and app (D-201); the axis options show Excel's automatic Grenzen and Einheiten, our scaling checked against 270 recorded Excel cases (D-200, `packages/charts/test/axis-auto.test.ts`); the ribbon reduces groups before collapsing any (D-202).
- Designs: 46 of our own (Raabeschule, Allgemein, 19 Fächer, Anlässe) with four Varianten each, in RaabePresent and as Designs in RaabeCalc and RaabeWrite; slide charts follow the design (D-190 to D-198). Lasting check: `check_present_designs.py` (PowerPoint opens each design); `check_present_roundtrip.py` 22/22 + 58/58.
- Next: the owner tries the designs and charts and sends differences.

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
- **Housekeeping, not done yet (owner's OK given 2026-11-11, refused by the session's permission check):**
  - The agent worktrees under `.claude/worktrees/` (16, all merged; they hold only uncommitted NIGHT-STATUS.md files)
    and their branches `worktree-agent-*`.
  - The Present worktree `step8-raabe-present-190dd0`.
  - The unused target folders `apps/raabe-calc/src-tauri/target/agent*`, `drag` and `macros`. These take a lot of disk
    space.
