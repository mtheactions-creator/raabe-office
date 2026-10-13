# Decision log

Every decision that shapes the project gets an entry here: what was decided, why, and
what it rules out. Newest at the bottom. Never rewrite an old entry. If a decision is
reversed, add a new entry that supersedes it and mark the old one `Superseded by D-xxx`.

| ID | Date | Decision |
|---|---|---|
| D-001 | 2026-10-13 | Build our own suite instead of adopting ONLYOFFICE |
| D-002 | 2026-10-13 | Three apps: RaabeWrite, RaabeCalc, RaabePresent |
| D-003 | 2026-10-13 | Look and feel of current Microsoft 365 (German UI) |
| D-004 | 2026-10-13 | OOXML (.docx/.xlsx/.pptx) is the native file format |
| D-005 | 2026-10-13 | Fully local, no network access by default |
| D-006 | 2026-10-13 | Tauri 2 + TypeScript web UI (provisional, pending D-014) |
| D-007 | 2026-10-13 | Windows first, iPad close behind; touch-first design |
| D-008 | 2026-10-13 | Branding: logo variant A, all three apps pure Raabe red |
| D-009 | 2026-10-13 | GPL-3.0 for code; school logo and name excluded |
| D-010 | 2026-10-13 | Document tabs in one window (deliberate deviation from Word) |
| D-011 | 2026-10-13 | Default font Calibri 11, Carlito as metric-compatible fallback |
| D-012 | 2026-10-13 | Scope additions: comments, track changes, formula editor, school templates |
| D-013 | 2026-10-13 | "Feel" is measured objectively against real Word, not judged by gut |
| D-014 | 2026-10-13 | Editing-engine choice decided by the feel prototypes (Step 0) |
| D-015 | 2026-10-13 | Updates only on explicit user request; code signing later |
| D-016 | 2026-10-13 | Claude is the maintainer; the repository is the project's memory |
| D-017 | 2026-10-13 | Editing engine: controlled editing (prototype B); Tauri 2 confirmed |
| D-018 | 2026-10-13 | Window frame: Microsoft 365's light chrome, Raabe red as the accent |
| D-019 | 2026-10-13 | Shell built without a UI framework; one command registry for everything |
| D-020 | 2026-10-13 | Ribbon reference is Word at ~1330 CSS px; groups collapse in Word's order |
| D-021 | 2026-10-13 | Files: simplified saving, AutoRecover every 30 s, last tab closes the window |
| D-022 | 2026-10-13 | Originals with unsupported content are never overwritten |
| D-023 | 2026-10-13 | Development builds show planned commands greyed out |

---

### D-001 Build our own suite instead of adopting ONLYOFFICE
The project owner evaluated ONLYOFFICE Desktop Editors. Verdict: a convincing Word
alternative, much easier for Word users than LibreOffice, but it "feels like a web-based
document editor packaged as a desktop application". Selecting text, dragging objects,
adjusting tables and navigating long documents feel less immediate than in Word.
**Consequence:** fluid, native-feeling interaction is our primary quality bar, ahead of
feature count. ONLYOFFICE remains the benchmark we have to beat.

### D-002 Three apps
RaabeWrite (Word), RaabeCalc (Excel), RaabePresent (PowerPoint). They share one shell
(ribbon, backstage, file handling, dialogs). Build order: shell first, then RaabeWrite.

### D-003 Look and feel of current Microsoft 365, German UI
Students used Microsoft 365 until mid-2026, so the reference is the current M365 desktop
look with the classic (full) ribbon. Labels use Microsoft's exact German wording
("Start", "Einfügen", "Layout", "Referenzen", "Überprüfen", "Ansicht") and Word's German
keyboard shortcuts, because that is where the muscle memory lives.
**Not copied:** Microsoft icons, artwork, logos, product names. Icons come from openly
licensed sets (e.g. Fluent UI System Icons, MIT) or are drawn by us.

### D-004 OOXML is the native file format
Files must move freely between students, teachers, home PCs with Word, and LibreOffice.
We support a defined subset well. **Rule:** content we do not understand is preserved
untouched when saving and shown as a placeholder. We never silently destroy parts of a
teacher's document.

### D-005 Fully local
No accounts, no cloud, no telemetry, no automatic network calls. Translation models,
dictionaries and templates are bundled or installed locally. This also keeps the apps
GDPR-safe (DSGVO) without any data-protection paperwork.

### D-006 Tauri 2 + TypeScript (provisional)
Small installers (~10 MB), fast start, uses the WebView2 runtime built into Windows,
and the same code base can target iPadOS via Tauri's iOS support. The risk is the
"web app in a box" feel criticised in D-001. That risk is tested by D-014 before we
commit. Fallback if the prototypes fail: a native Windows stack, at the cost of a much
harder iPad port.

### D-007 Windows first, iPad close behind
All new students must buy iPads; grade 8 and up use Windows laptops/tablets (BYOD, parents
manage the devices). The iPad share grows every year. **Consequence:** every interaction
must work by touch from day one, the ribbon needs a touch mode, and the architecture must
leave room for Apple Pencil ink later. iOS builds run on GitHub's macOS runners (no Mac
needed locally).

### D-008 Branding
Logo variant A: the school logo (red crescent + red disc, white "r" with dark drop
shadow) with the dot replaced by the app letter: W, C, P. All three apps use pure Raabe
red `#ED1C24` (measured from the original); no per-app accent colours. Below 48 px a
simplified icon is used (red disc + large white letter), because the full logo is
illegible at 16-32 px. See `branding/README.md`.

### D-009 Licence
Code: GPL-3.0, so the suite stays open. The Raabeschule name and logo are excluded from
the licence; forks must rebrand. The school leadership approved publishing the name and
logo in the public repository.

### D-010 Document tabs
Several documents open as tabs in one window, like ONLYOFFICE. This is a deliberate
improvement over Word, requested by the project owner. Everything else follows Word.

### D-011 Default font
Microsoft 365 defaults to Aptos, which we cannot ship. Default is Calibri 11 (installed
with Windows, familiar from years of older Word versions). Where Calibri is missing
(iPad), we render with Carlito, which has identical metrics, so documents never reflow.
Same approach for Arial/Times/Courier (Liberation fonts) and Cambria (Caladea).

### D-012 Scope additions
Confirmed as needed: comments and track changes (teachers correct work digitally);
formula editor (high priority, maths and physics); Raabeschule templates (worksheet with
school header, Facharbeit template following the school's formal rules; the owner will
supply the material later).

### D-013 Feel is measured objectively
The project owner pointed out that consciously testing puts you in a different mental
mode than everyday use, so their judgement of subtle feel would be unreliable. Instead
we measure against Microsoft Word, which is installed on the development PC (M365,
16.0.20430, scriptable via COM): keystroke-to-screen latency, frame drops while
scrolling and dragging, a catalogue of micro-behaviours recorded from Word, layout
fidelity (line and page breaks) and start/load times. See `docs/FEEL.md`. The final
check is real use by a pilot class.

### D-014 Editing engine chosen by prototypes (decided in D-017)
Before writing product code we build two small editor prototypes with different
architectures, measure both with the D-013 harness, and pick the winner. Recorded as a
new decision when done.

### D-015 Updates and signing
Updates are only fetched when the user explicitly clicks "Nach Updates suchen", or come
through the Microsoft Store / App Store. Code signing (Microsoft Store or SignPath's free
open-source signing) is agreed but postponed until the first release.

### D-016 Maintenance
Claude maintains the code. Claude does not remember earlier sessions, so everything
needed to continue must live in this repository: `CLAUDE.md`, this decision log, the
architecture notes, and automated tests. Any session must be able to pick up the work
from the repository alone.

### D-017 Editing engine: controlled editing
Result of D-014, agreed by the project owner after the
[Step 0 report](reports/2026-10-13-step0-feel.md). Our own text model, line breaking,
pagination, caret and selection (prototype B); the browser only renders text and
delivers raw input through a hidden textarea. B matched Word in 50/50 behaviours,
2,000/2,000 line breaks and 50/50 page breaks, at Word's keystroke latency. Native
`contenteditable` (prototype A) matched only 31/50 behaviours and cannot paginate.
The same engine will serve text in RaabeCalc cells and RaabePresent text boxes.
Tauri 2 (D-006) is confirmed; it is no longer provisional.
**Consequence:** we build IME handling, accessibility (screen readers) and spell-check
underlines ourselves.

### D-018 Window frame: light chrome, red accent
Current Microsoft 365 no longer paints the title bar in the app colour: the frame is light
grey (#F0F0F0), the ribbon a white card, and Word's blue appears only as an accent (active
tab underline, window border, links). Students know this look, so we copy it and replace
the blue accent with Raabe red (D-008). Colours and sizes were measured from Word on the
development PC (`tools/feel-lab/capture_word_ui.py`). A red title bar would be a one-line
change in `packages/shell/src/theme.css` if the school prefers it.

### D-019 Shell without a UI framework
Title bar, ribbon, backstage, status bar and dialogs are plain DOM built from data
(`packages/shell`). Every action is a command in one registry (`Commands`) with Word's
German label, icon, shortcut and state; ribbon buttons, the Quick Access Toolbar, keyboard
shortcuts, keytips and the search box all run commands by id. No framework sits in the
input path (D-001: feel first).

### D-020 Ribbon reference width
The ribbon is defined as Word shows it in a window of about 1330 CSS px (a 2000 px window at
150 % scaling, typical for school laptops). Narrower windows collapse groups to single
buttons in Word's order (captured at several widths). Not yet done: Word's extra step on
very wide screens, where small buttons gain text labels.

### D-021 File handling
- "Speichern unter" (F12) opens the Windows dialog directly; Word's extra location page is
  dropped. Strg+O opens the backstage "Öffnen" page with the recent files, as in Word;
  Strg+F12 opens the Windows dialog directly.
- A new document is saved under its first line as the suggested file name, as in Word.
- AutoRecover backs up changed documents every 30 seconds (configurable in Optionen) to the
  app's local data folder; after a crash RaabeWrite offers to restore them.
- Saving is atomic (temporary file, then rename), so a failure never leaves a broken file.
- Closing the last document tab closes the window, like Word's X button.
- A .docx double-clicked in Explorer opens as a new tab in the running window.

### D-022 Unsupported content is protected (superseded by D-024)
When a file contains content RaabeWrite cannot keep yet (tables, images, character
formatting, lists, fields, comments ...), it is listed in a yellow bar above the document
and in "Informationen". Saving into the original file is refused; the user is asked to
save under a new name, so the original keeps everything. This is the D-004 rule in the
period before full round-trip preservation exists.

### D-023 Planned commands are visible
Agreed with the project owner: during development the ribbon shows every planned command;
those not built yet are greyed out with the tooltip "Noch nicht verfügbar". The pilot build
for students will show only working commands.

### D-024 Unsupported content is kept verbatim; saving to the original is allowed
Replaces D-022 (Step 2). The .docx layer now keeps everything it does not edit exactly as
Word wrote it: unknown paragraph and run properties, inline objects (fields, hyperlinks,
images, footnote references, revisions, equations, content controls) with their original
XML, anchors (bookmarks, comment ranges, floating drawings), whole blocks (tables, tables of
contents) as read-only placeholders, and every other package part byte for byte. Word
itself confirms the round trip (`tools/feel-lab/docx_word_check.py`: tables, images,
shapes, footnotes, comments, hyperlinks, revisions, equations, fields, bookmarks, lists and
headers survive an edit-and-save unchanged). So the yellow bar now only informs ("... bleiben
beim Speichern unverändert erhalten") and Speichern writes the original file. If the owner
prefers the old protection for the pilot, it is a small change in `apps/raabe-write/src/app.ts`.

### D-025 Layout follows Word's measured rules
Measured on 2026-10-13 against Word's own PDF output (exact glyph baselines) and its COM
interface, for twelve fonts and four test documents built by Word
(`tools/feel-lab/run.py docx`):
- Line height per font: the larger of the hhea sum (ascender + |descender| + line gap)
  and the Windows sum (winAscent + winDescent); math fonts (Cambria Math) use hhea only.
  winDescent lies below the baseline, the rest above it. A line spans the largest ascent
  and the largest descent of its fonts.
- "Mehrfach" multiplies and puts the extra space below the text, "Mindestens" puts it
  above, "Genau" places the baseline at 80 % of the line.
- The gap between paragraphs is the larger of "Abstand nach" and "Abstand vor".
  "Abstand vor" counts on the first page and is dropped at the top of every later page,
  whatever the break (natural, manual, "Seitenumbruch oberhalb").
- Characters advance by their design widths (not hinted pixel widths); the engine measures
  them at a large size and scales down, which is also what the DOM renders.
- Super- and subscript are drawn at 65 % of the size, rounded to half points.
- "Blocksatz": a line may overflow by up to 25 % of the total width of its spaces; the
  spaces are then narrowed (measured for 14-30 spaces per line, also on the last line).
Result: 256/256 line starts and 15/15 page breaks identical, baselines within 0.18 pt, and
horizontal positions within 0.25 pt on 246 of 249 lines.

### D-026 Formatting commands behave like Word's
Measured with Word's own commands (`tools/feel-lab/run.py formatting`, 44 cases):
- Character commands with the caret inside a word format the whole word (without the space
  after it; "Ausgeblendet" includes it). At a word boundary they set the formatting for
  what is typed next. Toggles on a mixed selection switch on.
- The paragraph mark changes only when the selection includes it. Text typed at the end of
  a paragraph also formats the mark. Enter inserts a new mark with the formatting at the
  caret for the first paragraph; the original mark moves on with the new paragraph, whose
  text then takes the mark's formatting.
- Schriftgrad vergrößern/verkleinern use Word's steps without 10.5 (11 → 10).
- Alignment shortcuts toggle (real keystrokes, `keys_check.py`): the active alignment again
  returns to left, and Strg+L on a left-aligned paragraph switches to Blocksatz.
- Shortcuts are the ones Word reports on the development PC (`Application.FindKey`, German
  keyboard): Strg+Umschalt+F/K/U/D, Strg+Umschalt+G = Großbuchstaben (formatting, not a case
  change), Strg+< / Strg+Umschalt+<, Strg+8 / Strg+9, Strg+# = tiefgestellt,
  Strg+Leertaste, Strg+Q, Strg+L/E/R/B, Strg+M, Strg+T, Strg+1/2/5, Umschalt+F3, Strg+Alt+1-3.
- Styles: Word 2016-2019's built-in set with German names. Applying a paragraph style
  removes direct paragraph formatting, and character styles toggle bold/italic/caps against
  the paragraph style (the OOXML rules; not yet checked against Word by the harness).

### D-027 Fonts
Installed fonts and their metrics come from the font files (Rust `fontdb` + `ttf-parser`;
the browser does not expose hhea/OS/2 values). Family names match without regard to case,
as in Word. A missing font falls back to a metric-compatible open font where one exists
(Calibri → Carlito, Cambria → Caladea, Arial/Times/Courier → Liberation). Known gap:
Aptos, the default of documents made with current Microsoft 365 templates, is only a cloud
font of Office (not installed in Windows) and has no metric-compatible substitute, so such
documents break lines differently. New RaabeWrite documents use Calibri 11 (D-011).

### D-028 Measurements without taking over the PC
The harness no longer needs the PC to itself except for latency (F1) and real-input
behaviour (F3): Word runs invisibly (its own process) and RaabeWrite runs off screen without
ever taking the focus (`RAABE_FEEL_BACKGROUND`), driven through the DevTools protocol.
Layout, formatting behaviour and the smoke test can therefore run while the owner works.

### D-029 Document structure for Step 3 and 4
The model follows Word's own (WordprocessingML), so editing, undo and the file format stay
simple:
- A document is a set of *stories*: the main text, each header and footer, each foot- and
  endnote and each comment. Editing works on the story with the caret; one undo step holds
  all stories.
- Tables are not nested objects: their cells' paragraphs are ordinary paragraphs of the
  story in row order, each pointing to its cell (table, row, cell). Caret, selection,
  search and spell check therefore work in tables unchanged. A selection across cells
  selects whole cells, as in Word. Nested tables are kept verbatim and shown read-only.
- A section ends at a paragraph (the paragraph carries the section's properties, as the
  paragraph mark does in Word); the last section belongs to the document.
- Lists are Word's numbering definitions (numbering.xml) referenced by paragraphs; labels
  are computed during layout.
- Fields that span paragraphs (Inhaltsverzeichnis) become start/end markers around ordinary
  text, hyperlinks become start/end markers around their text, tracked changes become a
  property of the text, comment ranges are markers. The file layer writes only complete
  pairs, so editing can never produce invalid XML.
- Pictures keep their original XML and relationship; when they are copied into another part
  the relationships are rewritten. Parts whose stories did not change are written back byte
  for byte.
- Layout is split: `layout.ts` breaks one paragraph into lines for a given width (with list
  label, page fields and room around pictures); `flow.ts` places paragraphs on pages:
  sections with page size and columns, headers and footers per page, tables (rows split at
  line boundaries, header rows repeated, vertical merges), footnotes at the page bottom,
  floating pictures with text wrapping, balanced columns before continuous section breaks.

### D-030 Page layout follows Word's measured rules (Step 3)
Measured on 2026-10-13 with six more test documents built by Word (`make_fixtures.py
--lists --tables --sections --notes --borders --images`) and `run.py docx`:
- Lists: the label stands at the first-line indent and is followed by its tab, space or
  nothing; the label's own font (Symbol for "•") raises the line if it is taller, but never
  deepens it.
- Tables: a cell's text starts after half the width of its left border plus the cell margin
  (0.19 cm); every row is as tall as its content plus the full width of its top border (the
  last row also its bottom border); a row height set in Word is the height between the
  borders. A row breaks across pages only when every cell can start on the page (with widow
  and orphan control in each cell); otherwise it moves whole. Header rows repeat.
- Pictures in a line stand on the baseline; "Mehrfach" line spacing adds its extra space for
  the text only, not for the picture. Text flows around pictures with square wrapping on the
  side with more room, and below pictures set to "Oben und unten".
- Continuous section breaks: the new section starts below the longest column, counting the
  "after" of a paragraph that ends there; columns are balanced unless a column break is in
  them. After a column break the paragraph mark starts the next column (unlike a page
  break, after which it stays on the line).
- Footnotes stand at the page bottom under a 2-inch separator line; a line whose note does
  not fit moves to the next page. Paragraph borders and shading take the border width plus
  its distance above and below the text.
Result: 690/690 line starts and every page break identical in ten Word documents, baselines
within 0.18 pt. The harness now finds Word's lines by position (Word's own line navigation
moves between table cells), maps table, section and picture characters between Word and
RaabeWrite, and reads baselines from the PDFs by position.
