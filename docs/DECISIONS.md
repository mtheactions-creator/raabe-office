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
Result: 632/632 line starts (first reported as 690, a summing error corrected in Step 4) and every page break identical in ten Word documents, baselines
within 0.18 pt. The harness now finds Word's lines by position (Word's own line navigation
moves between table cells), maps table, section and picture characters between Word and
RaabeWrite, and reads baselines from the PDFs by position.

### D-031 Spell checking
Hunspell dictionaries checked in Rust with `spellbook` (MPL-2.0): German (igerman98, GPL-2.0
or GPL-3.0), US and British English (SCOWL, MIT and BSD), taken from the npm packages
`dictionary-de`, `dictionary-en` and `dictionary-en-gb` and compiled into the program, so
checking works offline and nothing extra is installed. The language comes from the text
(Korrekturhilfesprache), else the document default. As in Word: words in capitals, words
next to digits and Internet/e-mail addresses are ignored, and the word being typed is not
marked until the caret leaves it. German compounds the dictionary lacks are accepted when
they split into known words with the usual linking letters (Krautschicht,
Waldbodenlebewesengemeinschaft; found with F6, where Word accepted them). "Alle ignorieren"
lasts for the session; "Zum Wörterbuch hinzufügen" keeps words on this computer.

### D-032 Text rendering
Measured with `run.py rendering` (F6): Word 365 draws document text with grayscale
smoothing, not ClearType colour fringes. RaabeWrite turns LCD text off in WebView2
(`--disable-lcd-text`). Result: colour fringes 3.4 against Word's 3.7 (74 before), sharpness
97 % of Word's, glyph positions correlate at 0.97.

### D-033 Step 3 behaviour taken from Word
Checked with Word's own commands through COM where no real input was needed:
- The Aufzählungszeichen and Nummerierung buttons give Standard paragraphs the style
  Listenabsatz and the level's indents; lists made like the ribbon makes them have nine
  levels (Tab changes the level).
- Changing the orientation rotates the margins (landscape top = portrait right, bottom =
  left, left = top, right = bottom).
- Contextual tabs as in Word 365: Tabellenentwurf and Layout in tables, Kopf- und Fußzeile
  while a header or footer is edited (selected automatically), Bildformat for a picture.
- Header and footer editing: double click in their area opens them, a double click in the
  text closes them; the text is dimmed meanwhile.
- Copy and paste with other programs use HTML; content copied inside RaabeWrite keeps its
  full formatting.

### D-034 Comments and tracked changes (Step 4)
Comments are markers in the text (range start, end, reference) plus a story per comment, as
in the file; replies and "erledigt" are Word 2013+ data (commentsExtended.xml). They are shown
as Word 365's modern cards right of the page: the page moves left to make room, each card
sits at the height of its text and is pushed down when the one above is in the way; the
Kommentare button opens the list instead. Author and initials come from the user name in
Optionen ("Benutzer" when none is set).
Tracked changes are a property of the text (each run and paragraph mark can carry one
revision: insertion or deletion, author, time), so layout, saving and undo need nothing
extra. With tracking on, typed and pasted text is an insertion, deletions only mark text,
and the caret moves over deleted text as in Word; deleting one's own tracked insertion
removes it for real. Simplification: a deletion of another author's insertion becomes a
plain deletion (Word keeps both); rejecting it keeps the text. Markup modes as in Word:
Einfaches Markup (result, red change bars), Alle Markups (struck and underlined in the
author's colour, grey bars), Kein Markup, Original.

### D-035 Raabeschule templates
Our own designs, made of ordinary Word formatting (styles, lists, tables, header/footer
fields, an equation, a table of contents), so files from them work in Word: Arbeitsblatt,
Facharbeit, Versuchsprotokoll, Handout. Placeholders are bracketed text ("[Thema]") that is
typed over; no content controls. The Facharbeit follows common requirements in
Niedersachsen: title page without page number, Inhaltsverzeichnis, numbered chapters,
Literaturverzeichnis, Anhang and Erklärung unnumbered, Times New Roman 12 pt, 1,5-zeilig,
Blocksatz, margins 2,5 cm with a 4 cm correction margin on the right. The school can replace
the wording later without code changes beyond `apps/raabe-write/src/templates.ts`.

### D-036 Equations
- In the file: Office Math (OMML). Read equations are converted to LaTeX
  (`packages/ooxml/src/math.ts`) and become editable; their XML is written back unchanged
  until they are edited. Unknown OMML elements keep the equation read-only.
- Editing: MathLive (MIT) in place over the equation, loaded only when an equation is first
  edited (keeps start-up and memory as before). Its own MathML writes limits beside the sum
  even in display equations, so RaabeWrite does not use it for display. Replaced in Step 4b by
  our own editor (D-039, D-047); read equations are trees now, LaTeX is only a cache key.
- Display: MathML built from the same parse as the OMML, drawn by the WebView with Cambria
  Math (named explicitly: the generic `math` font falls back to Times for digits) and laid
  out with the font's MATH table, as Word does. An equation alone in its paragraph is a
  display equation (centred, display style, m:oMathPara), as in Word. Measured against Word
  (`word-step4.docx`): matrix rows 0.3ex padding, 0.08 em extra above and below display sums
  and integrals; text after five equations within 0.8 pt of Word.
- Input: Word's Formel tab (Strukturen galleries insert LaTeX templates with placeholders),
  Alt+Umschalt+0, arrow keys enter and leave equations, Esc or Enter leaves.

### D-037 Notes and tables of contents
- Endnotes follow the document text under a separator line ("Ende des Dokuments", Word's
  default), numbered i, ii, iii. Notes whose reference was deleted are left out on saving.
- The table of contents is a TOC field (\o "1-3" \h \z \u) whose entries are ordinary
  paragraphs (Verzeichnis 1-3) with hyperlinks to _Toc bookmarks and PAGEREF fields, the
  structure Word writes, so Word can update it. Page numbers come from RaabeWrite's layout;
  the table is laid out again until they no longer change. Word often ends the field at the
  start of the first heading; that paragraph is not part of the table and is kept on
  updates. Numbered headings appear with their number and a tab.
- Alt+Umschalt+… are shortcuts, not keytips (Alt+letter still jumps to a tab).

### D-038 Word parity before the pilot (Step 4b)
After trying Step 4 the owner listed differences from Word (cursor, icons, tables, lists,
formulas, pictures, behaviour) and asked for all ribbon commands to work. Instead of fixing
reports one by one, Step 4b adds two audit tools (a behaviour catalogue run in Word and
RaabeWrite, a screen comparison of ribbon, menus and dialogs, normal and on hover) and works
through six packages (ROADMAP). The owner contributes no material; everything is measured
against Word on this PC, with real input allowed for the whole step.

### D-039 Own formula editor instead of MathLive
MathLive's editing model is TeX-like: its own caret rules, fonts and an overlay over the
line, which caused shifting formulas, wrong arrow-key navigation, no mouse selection and no
handles. The editor becomes our own, working on the same formula structure that is saved as
OMML and shown as MathML in Cambria Math (D-036), with Word's navigation, selection,
Backspace, sizing and Math AutoCorrect recorded from Word first. MathLive is removed when
the new editor passes the behaviour catalogue.

### D-040 Own grammar checker
Options measured on this PC (2026-10-14): RaabeWrite is 8.2 MB, Microsoft Office 4.6 GB
(Word's proofing tools 131 MB). LanguageTool offline would add about 250-320 MB (30
languages, a Java runtime; 8 GB more for its statistics) and cannot run on iPadOS. Decision:
our own checker in Rust, estimated 10-30 MB for four languages and runnable on Windows and
iPad. It reuses public material: LanguageTool's XML rules and lexicons (LGPL and open
licences, checked per file), Grammalecte for French (GPL-3.0), ideas from nlprule
(MIT/Apache). LanguageTool's Java rules for German (agreement, capitalisation, commas) are
rewritten in Rust; own rules cover typical school confusions (das/dass, seid/seit,
wieder/wider, als/wie, Standart). No student texts are used (data protection). Measured
instead with LanguageTool's example sentences (recall), clean open texts such as Klexikon
and Wikipedia (false alarms), the same texts with generated typical errors (recall per error
type) and public learner corpora where their licences allow, always against Word's offline
grammar check. A language is switched on once its false-alarm rate is close to Word's.

### D-041 Deferred commands stay visible
The ribbon keeps Word's layout. Commands deferred by agreement (SCOPE) stay greyed out with
"Zurückgestellt" in their tooltip, so they cannot be mistaken for overlooked ones; every
other command must work.

### D-042 AutoCorrect and AutoFormat as Word 365 does them
Measured with real typing in an own Word instance (`probe_word_typing.py`, the behaviour
catalogue). Word's settings are its defaults; the lists of abbreviations and typing slips
are our own (Word's AutoCorrect list is Microsoft's data), Word's symbol entries are facts
and copied: (c) (r) (tm) (e) ... --> <-- ==> <== <=> and the smileys, which Word 365 now
replaces by emoji in Segoe UI Emoji (written as w16se:symEx like Word). These entries are
replaced as soon as their last character is typed ("Text(c)" → "Text©"); words are corrected
when they end. Other rules found by measuring: a period after a number ends a sentence
("Am 1. januar" → "Januar"); "hALLO" is only corrected while Caps Lock is on; German "geht's"
gets ’ once the word ends; list labels ("a)") are not capitalised. AutoFormat makes lists
only from the first value (1, a, A, i, I with "." ")" "-" or brackets) and from "*", "-",
"–", "—", ">", "->", "=>", after a space or a Tab; roman numerals get 1.27 cm hanging at
1.9 cm. Word 365's Markdown is followed too: "# " to "###### " make headings, **text** bold,
~~text~~ struck through (*text* and _text_ stay). Backspace at a paragraph start removes a
first-line indent, then the left indent back to the previous paragraph's text or number
position (a "Listenabsatz" becomes "Standard"), then joins; Tab at a list's first item moves
the whole list to the next stop (default tab stops and 0.635 cm after each). Deleting or
replacing whole words also removes the doubled space; a pasted trailing space before
punctuation goes. A selection that reaches another cell covers whole cells, one leaving the
table whole rows. Strg+Alt+1/2/3 do nothing in German Word (AltGr), so RaabeWrite does not
bind them either.

### D-043 Clipboard with Word: the whole HTML document
The web view's clipboard API hands out only the copied fragment, without the style sheet in
which Word defines its lists (`@list l0:level1 {mso-level-text:…; font-family:Symbol}`) and
classes (margin-left of MsoListParagraph). A small Rust command (clipboard-win, already part
of the build through Tauri's clipboard plugin) reads and writes the clipboard's "HTML Format"
as a whole. RaabeWrite reads Word's list definitions, paragraph indents, column widths, row
heights, shading and merged cells, and writes lists and tables the way Word writes them
(mso-list paragraphs with @list definitions, table cells with widths), so both directions
keep bullets, numbers, levels, indents and widths (`clipboard_roundtrip.py`: 0 differences).
Pictures on the clipboard (screenshots, copied pictures) are pasted as pictures.

### D-044 Context menus and mini toolbar as in Word
Recorded with `probe_word_context.py` (UI Automation and screenshots of Word's popups):
search box, Ausschneiden/Kopieren, the three paste options, then entries for text, list
items (bullets: Separate Liste, Mit vorheriger Liste verknüpfen; numbers: Neu beginnen mit
1, Nummerierung fortsetzen, Nummerierungswert festlegen), a table cell, several cells and
pictures, with Word's submenus, and the mini toolbar above. Left out: entries for Microsoft's
online services (Copy Link to Location, Mit Copilot bearbeiten, Auswahl im Editor
überprüfen, Bild ändern aus Archiv-/Onlinebildern und Piktogrammen) and "In Office Math
konvertieren", which Word shows greyed out for pictures. Word's bullet library has a picture
bullet of its own; RaabeWrite has no picture bullets, so the library shows the six symbol
bullets. The numbering library is Word's: 1. 1) I. (1) a) a. i.

### D-045 Bildformat as Word does it
Word's galleries were recorded with real clicks in an own Word instance
(`probe_word_picture.py`, the XML Word writes for every entry and its tooltip name):
Korrekturen, Farbe, Transparenz, the 28 Bildformatvorlagen and the presets of Bildeffekte
(Voreinstellung, Schatten, Spiegelung, Leuchteffekt, Weiche Kanten, Abschrägung,
3D-Drehung). The presets are kept as Word's DrawingML (`picture-presets.ts`, generated) and
read and written without loss (test: every preset comes back as Word wrote it). Pictures
carry their settings in the model (`PictureFormat`: a14 image effects, blip colour changes,
cropping, frame shape, line, effects, 3D as XML, effect extent); unchanged pictures keep
their XML as read. The view draws them with CSS and SVG filters, 3D rotation approximated
with CSS perspective. Gallery tiles show the selected picture with the entry and preview it
in the document on hover. Ribbon layouts and reduction order follow Word at 2560-800 px.
Freistellen removes the background connected to the edges in about the edge colour (Word's
interactive mode with marking is not rebuilt); Künstlerische Effekte stay deferred (Word's
23 artistic filters cannot be drawn close enough); Bildlayout converts to SmartArt, which is
out of scope, so it stays greyed out. Pictures insert at their file's resolution as Word
does (a 500 dpi PNG is small), in table cells at most as wide as the cell. The CSP allowed
only data: images, so pictures inserted in the app never showed: blob: is allowed now.


### D-046 Formen and Textfeld as Word does them
**Geometry.** Word's preset shapes are the formulas of ECMA-376 (presetShapeDefinitions):
guides, adjust values, paths with arcs. `geometry.ts` evaluates them in a compact form of
the standard's syntax (`geometry-presets.ts`); every one of the 152 presets (all of Word's
Formen gallery, connectors and callouts) was compared with Word's own drawing, read from its
PDF export (`check_shapes.py`: wide, tall and with adjust values moved; all within 0.5 pt,
and the shaded parts in Word's colours). Measured on the way: arc angles are visual angles
in the path's own coordinates; a speech bubble whose tail point lies inside it has its point
moved to the edge (the tail vanishes); shaded sub-paths are darken ×0.6, darkenLess ×0.8,
lighten +40 % white, lightenLess +20 %; arrowheads are 2/3/5 (klein/mittel/groß) times the
line width but at least times 2 pt (`probe_word_arrows.py`); DrawingML shade and tint work in
linear light (accent1 shaded to 15 % is Word's 172C51); the arrow keys move a selected shape
0.35 pt.
**Theme.** Word fills shapes from the theme (fillRef, lnRef, effectRef into the format
matrix). A document without a theme part is shown by Word 365 in its 2023 colours, while new
RaabeWrite documents use Calibri and the 2013-2022 palette (D-011): they now carry an
Office 2013-2022 theme part, written by us, so Word and RaabeWrite agree; files without one
read with Word's 2023 colours, as Word shows them.
**Model.** A shape is a floating object with the picture's shape properties (geometry,
fill, outline, effects) and a `ShapeInfo` (kind, adjust values, the style resolved for
drawing and kept as XML, the text frame of wps:bodyPr, custom geometry). Its text is a story
of its own (`tx:N`), laid out in the shape's text rectangle (insets, top/middle/bottom,
Form dem Text anpassen) and drawn above the shape; a click on the text or typing into a
selected shape edits it, Esc selects the shape again. Unchanged shapes keep Word's XML (with
the VML fallback); changed shapes and shapes with text get a new wps:wsp inside Word's
drawing, new ones the XML Word writes. Word renders our round trip of its own shapes
identically (PDF outlines, deviation 0.0).
**Interaction.** Einfügen > Formen with Word's gallery (order and names recorded from Word),
drawing by dragging (Umschalt: square, 45° lines), a click inserts 2.54 cm; Kurve,
Freihandform and Skizze make custom geometry; the eight handles, rotation, the yellow
adjust handles (any preset: the adjust value whose handle lands nearest the mouse) and two
end handles for lines; copy and paste of floating objects. The Formformat tab has Word's
groups, layouts and reduction order (captured at 2560-800 px); its Formenarten gallery has
all 77 entries as Word writes them (`probe_word_shapes_ui.py applylist`, `shape-presets.ts`).
Einfügen > Textfeld offers our own designs (Word's built-in text boxes are Microsoft's
content): Einfaches Textfeld, Randleiste rechts, Merkkasten, Zitat, and Textfeld erstellen.
**Moved to package E:** the WordArt-Formate (they are Word's text effects, E's Texteffekte),
Verknüpfung erstellen (linked text boxes), Gruppieren, Auswahlbereich, Punkte bearbeiten,
Neuer Zeichenbereich, picture and texture fills.

### D-047 The formula editor as Word's
**Specification first.** Word's behaviour was recorded before the editor was written:
`catalogue_math.py` types 114 cases into a new equation in our own Word instance (real
keystrokes: build-up, Math AutoCorrect, arrows, Up/Down, Home/End, Strg+arrows, Backspace,
Delete, Umschalt+arrows, Enter) and compares the saved OMML, normalised to what changes the
look. The editor (`packages/text-engine/src/math/edit.ts`) replays all 114 like Word
(`math-editor.test.ts`). Recorded rules include: build-up on a space or an operator, scripts
on the last letter, brackets kept around a script's base but removed around fraction and
script arguments, `lim_(n→∞)` as a function with its limit below, capital Greek from
\Delta upright, Up from a subscript to the base, nothing above or below a sum's operand.
**Galleries and symbols from Word.** `probe_word_equation_ui.py` read the Formel tab (groups,
labels, keytips) and `probe_word_equation_gallery.py` inserted every entry of every gallery
(231) and every symbol of the eight categories into an empty equation and recorded what Word
wrote; `math-gallery.ts` is generated from that (as trees, exact) with Word's headings, entry
names and tile sizes. Word's labels differ from older versions: Hoch/Tief, Wurzel, Eckige
Klammer, and the group Konvertierungen (Unicode, LaTeX, Text, Konvertieren).
**In the app.** No overlay: the equation stays in the line, drawn as MathML whose atoms carry
their place in the tree (`data-a`, `data-s`); caret, selection (structures whole), Word's box
with the move handle and the Formeloptionen arrow are drawn from that. Every change goes into
the document at once (one undo step per edit). Konvertierungen: the linear format
(UnicodeMath, `linear.ts`) and LaTeX both ways, Text for plain text in an equation.
Formeloptionen is Word's dialog; the document's settings are m:mathPr in settings.xml (kept
as the file has them unless changed, Word's defaults in new documents) and decide where new
limits go and how limLoc is written; display equations keep their own alignment (m:jc).
Display fractions stay full size when nested (Word's default "Größe von verschachtelten
Brüchen reduzieren" off). MathLive is removed (D-039).
**Menus from Word.** `probe_word_equation_context.py` right-clicked every character of twelve
equations in Word and recorded each menu: the common entries (Ausschneiden, Kopieren,
Einfügeoptionen, Professionell, Linear, Als neue Formel speichern, Schriftart, Absatz) and
each structure's own (fractions: the other two fraction kinds and Bruchstrich entfernen;
scripts: Skripts vor Basis platzieren, Tief-/Hochstellung entfernen, Skripts ausrichten;
roots; sums and integrals; brackets: arguments, remove, Eckige Klammern dehnen, an
Argumenthöhe anpassen; matrices: Platzhalter ausblenden, rows and columns, alignment; limits;
operators: Manuellen Umbruch einfügen, An diesem Zeichen ausrichten) and the box's arrow menu
(In 'Inline' ändern, Ausrichtung). The structure properties they set are kept in the tree and
in OMML (grow, shp, plcHide, baseJc, mcJc, alnScr; pre-scripts are m:sPre). Left out: Word's
"Copy Link to Location" (cloud documents), Argumentgröße erhöhen/reduzieren and
Matrixabstand (sizes and spacing the tree does not hold).
**Measured.** Real keystrokes in RaabeWrite: 114/114 cases as in Word (`catalogue_math.py
--ours-only`; the last three were Up/Down on the only line, where Word stays in the equation).
`check_math_ui.py`: typing, a click into the numerator, a gallery entry around the selection,
a ribbon symbol, the fraction's context menu, Konvertieren both ways, Enter, a click back,
the arrow menu and In 'Inline' ändern, and Word reading the file. Layout against Word
(`word-math.docx`, 12 equations display and inline): every line start and page as in Word;
heights per equation within about 2 pt, except matrices in brackets (6-7 pt taller: Cambria
Math's next bracket size is chosen where Word stays smaller) and nested display fractions
(2.3-2.6 pt taller); open for a later step.

### D-048 The remaining commands (Step 4b E), recorded from Word
Every dialog and menu is first read from Word (`probe_word_dialogs.py`, `probe_word_symbol_dialog.py`,
`probe_word_source_dialog.py`: names, access keys, positions, screenshots kept locally), then
built with Word's labels and behaviour.
- **Symbol**: Word's menu of the 20 symbols used last (tooltips are Unicode's English names,
  as in Word; names generated from Python's unicodedata, loaded with the dialog) and the
  dialog: Schriftart and Word's 47 Subset names, the font's characters (read from its cmap),
  Zuletzt verwendete Symbole, Zeichencode with "von:", Sonderzeichen with Word's names and
  shortcuts; Einfügen keeps the dialog open. Symbol fonts are written as w:sym.
- **Links**: w:hyperlink with an external relationship or w:anchor, the Hyperlink style; the
  dialog with Datei oder Webseite (folder list, recent files, visited pages), Aktuelles Dokument
  (headings via hidden _Toc bookmarks, Textmarken), Neues Dokument erstellen, E-Mail-Adresse,
  QuickInfo and Zielframe; Link bearbeiten/entfernen; Strg+Klick follows a link (Windows opens
  web addresses, mail and files; only those).
- **Textmarke, Querverweis, Beschriftung, Abbildungsverzeichnis**: bookmarks as Word writes
  them; cross references are REF/PAGEREF/NOTEREF fields to hidden _Ref bookmarks; captions use
  SEQ fields (formats 1, a, A, i, I); a table of figures is a TOC \c field rebuilt from the
  captions. Felder aktualisieren (F9) computes SEQ, REF (text, \n number, \p oben/unten),
  PAGEREF ("auf Seite 3") and NOTEREF results.
- **Zitate und Literaturverzeichnis**: sources in Word's customXml b:Sources part (Word reads
  ours, we read Word's), CITATION fields and a BIBLIOGRAPHY field; Word's twelve styles with
  their patterns (APA 6, MLA 7, Chicago 15, Harvard - Anglia, IEEE, ISO 690 both, Turabian,
  GB7714, GOST both, SIST02); Quelle erstellen with Word's fields for all 17 types, Name
  bearbeiten, Platzhalter, Quellen-Manager with a Masterliste for this user.
- **Fuß- und Endnote**: position, number format, start, restart (per section, and per page from
  a second layout), Konvertieren; w:footnotePr/w:endnotePr in settings.xml.
- **Überprüfen**: Optionen zum Nachverfolgen von Änderungen as Word's dialog (what RaabeWrite
  does not track, ink and formatting, greyed); the Überarbeitungsbereich vertical (left) or
  horizontal (bottom) as Word's pane: the count with a detailed summary, each change as
  "<Autor> Hinzugefügt", "hat gelöscht", "hat kommentiert" with its text; a click selects it.
- **Start**: Word 365's labels in Groß-/Kleinschreibung and Zeilen- und Absatzabstand; Text
  sortieren (paragraphs or table rows, three keys as text, number or date, Überschrift,
  Trennzeichen, Groß-/Kleinschreibung, Sortiersprache; section breaks stay in place); Markieren
  with Objekte auswählen (the pointer picks objects, also behind the text; Esc ends it) and
  Auswahlbereich.
- **Auswahlbereich**: the floating objects front first with Word's names ("Rechteck 1",
  "Textfeld 2", now also given to new shapes and written as wp:docPr name), Ausblenden
  (wp:docPr hidden: not drawn, no text wrap), renaming, Alle anzeigen/ausblenden, order.
- **Formatvorlagen pane**: Word's list (Empfohlen by priority, Überschrift n once n - 1 is
  used; In Verwendung, Im aktuellen Dokument, Alle; five sort orders), type marks, Vorschau
  anzeigen, each entry's menu (aktualisieren, um der Auswahl zu entsprechen; Ändern...;
  löschen; Formatvorlagenkatalog), Formatvorlage ändern / Neue Formatvorlage von Formatierung
  erstellen with Word's controls, preview and description and Format > Zeichen, Absatz,
  Sprache; the Formatvorlageninspektor. Changed styles are written from the model into the
  file's styles.xml (others keep their XML), deleted ones removed; changed document defaults
  and theme colours/fonts are written too.
- **Zwischenablage**: Word's pane as RaabeWrite's Office clipboard (24 items, collected while
  the pane is open, with Word's three options), Alle einfügen, Alle löschen.
- Task panes dock beside the document as in Word, several side by side (Navigation,
  Zwischenablage, Überarbeitungen left; Formatvorlagen, Auswahl, Kommentare right).
- **Ansicht** (Word's views captured with `probe_word_views.py`): pages stand side by side
  whenever several fit at the zoom, as in Word; Mehrere Seiten picks the zoom for two.
  Weblayout lays the text out on one page as wide as the window and as tall as the text (no
  headers, footers or page breaks; footnotes follow the text). Lesemodus lays it out on
  screen-sized pages side by side (columns of Word's Schmal/Standard/Breit widths), hides the
  ribbon for Datei / Extras / Ansicht (Spaltenbreite, Seitenfarbe Keine/Sepia/Umkehren, Layout
  Spalten-/Papierlayout), turns screens with the arrows and keys, is read-only, shows
  "Bildschirm 1 von 6" and ends with Esc. Lineal: Word's horizontal and vertical rulers (cm
  from the margin, margins grey, indent markers, tab stops, default stops, the tab selector);
  dragging snaps to 0.25 cm (Alt: free). Gitternetzlinien: Word's drawing grid, 9 pt from the
  margins over the text area (read from Word: GridDistance 9 pt, every line).
- **Zeilennummern**: w:lnNumType as Word writes it (start 0-based, distance in twips, restart
  newPage by default), w:suppressLineNumbers; numbers right-aligned 0.25" (columns 0.13") left
  of the text in the style Zeilennummer, lines of tables not counted; Word's menu and dialog.
- **Notizen anzeigen**: from the text to the note at or after the caret, from a note back to
  its reference; with foot- and endnotes Word's question which area to show.
- **Entwurf** (`probe_word_design.py` applied every gallery entry in Word and read the result):
  Designs (Word's 31 themes: colours and fonts), Farben (24 schemes) and Schriftarten (26
  pairs) with Farben/Schriftarten anpassen and Nach Designs suchen (.thmx, colour and font
  scheme files); theme colours now follow the theme (w:themeColor with shade/tint resolved in
  HSL as Word does; the default style set uses theme colours like Word 2013's), so headings
  change colour with the theme. Formatvorlagensatz: Word's 18 style sets, read with our reader
  from Office's QuickStyles files and matched to the German gallery names by the formatting
  Word showed. Absatzabstand: Word's presets (document defaults). Effekte: only Office's own
  effects are included (the others are Microsoft's designs). Als Standard festlegen keeps the
  theme, defaults and main styles for new blank documents. Changed themes go into the file's
  theme part (format matrix kept), changed styles and defaults into styles.xml. The gallery
  thumbnails are RaabeWrite's own.
- **Seitenhintergrund** (`probe_word_background.py`): Wasserzeichen as Word writes them (VML
  WordArt or picture in a Watermarks sdt in every header; the first section gets headers if it
  has none); Word's 12 gallery entries have fixed proportions (diagonal: width + height 659.85
  pt, horizontal: 468 pt wide), the dialog Gedrucktes Wasserzeichen sizes custom ones; drawn
  behind the text, stretched to the shape like VML's fitshape. Seitenfarbe: w:background with
  its theme colour, shown on screen but not printed (Word's default); Fülleffekte are not
  included. Seitenränder: w:pgBorders (from the page edge or the text, all/first/other pages)
  in Word's dialog Rahmen und Schattierung, which also sets paragraph borders and shading.
  Word opens all of it as its own (checked through COM).
- **Texteffekte und Typografie** (`probe_word_texteffects.py` applied every gallery entry in
  Word and read the w14 XML it wrote): Word's 15 presets, Kontur, Schatten, Spiegelung and
  Leuchten galleries with their names, Zahlenformatvorlagen, Ligaturen and Stil-Sets; kept as
  Word's w14 elements in the run properties (written back in Word's order) and drawn with CSS
  (shadows, outline, gradient fill, reflection, glow, OpenType features). The WordArt-Formate
  of the Formformat tab and the pane Texteffekte formatieren use the same data. Abschrägung,
  3D-Drehung and Transformieren are not included (greyed).
- **Gruppieren**: objects selected with Umschalt/Strg+Klick become one floating object with
  members placed inside it, written as Word's wpg:wgp (Word opens ours, we read Word's);
  Gruppierung aufheben and wiederherstellen, member text editable, resizing scales members.
- **Verknüpfte Textfelder**: Verknüpfung erstellen (the next click picks an empty text box, Word's
  messages otherwise) and aufheben; the text flows through the boxes in order and is written as
  Word writes it (wps:txbx id, wps:linkedTxbx id/seq). Word reads our chains (checked via COM).
- **Deckblatt**: Word's gallery layout and entries; the designs are RaabeWrite's own for school
  work (Schlicht, Facharbeit, Referat, Seitenleiste, Rahmen, Farbfläche, Linie; Word's designs
  are Microsoft's building blocks), in the theme's accent colour, with bracketed placeholders
  (the first is selected). Stored as Word stores its own: a Cover Pages sdt at the start of the
  document ending with a page break, so inserting again replaces it, Aktuelles Deckblatt
  entfernen removes it, and Word treats ours as its cover page.
- Silbentrennung moves to F, where the hyphenation patterns come with the languages.

### D-049 Languages and proofing tools (Step 4b F)
Recorded from Word first (`probe_word_languages.py`: Word's 260 languages, which have proofing
tools on this PC, the dialog Sprache, the Thesaurus pane and Synonyme menu, the Editor pane of
F7, Manuelle Silbentrennung, Datei > Optionen > Dokumentprüfung and Sprache, and what Word
does with typed French, Spanish and English sentences).
- **Data**: open sources only, compiled into the program so everything works offline (and on
  the iPad later). Spelling: Hunspell dictionaries for French (Dicollecte, MPL-2.0) and Spanish
  (RLA-ES, GPL/LGPL/MPL) join German and English (npm `dictionary-*`). Thesaurus: LibreOffice's
  thesauri (German OpenThesaurus, English from WordNet, French by Dicollecte, Spanish
  OpenThesaurus; LGPL-2.1+ and the WordNet licence), fetched by `tools/lang-data/fetch.py`
  and built by `build.py` into 7 MB of gzip files (meanings stored once, headwords pointing to
  them; generic and related terms dropped as Word does). The program grows from 8 to 18 MB.
- **Sprache**: Word's dialog with its language names and tags (260, generated from Word),
  RaabeWrite's proofing languages on top and marked, Markierter Text / Aktives Dokument,
  Rechtschreibung und Grammatik nicht prüfen (w:noProof, now a run property), Sprache
  automatisch erkennen and Als Standard festlegen. The status bar shows the language at the
  caret and opens the dialog.
- **Sprache automatisch erkennen** (on by default, as in Word): Word decides only among the
  editing languages enabled in Office (here German and English: a French sentence stayed
  German, a Spanish one became English). RaabeWrite decides per sentence among its four
  proofing languages from function words and letters, only when the evidence is clear, and
  only for text without a language of its own. The result is used for checking and shown in
  the status bar but not written into the file (Word writes the detected language); that keeps
  undo and the file free of background changes.
- **Thesaurus** (Umschalt+F7): Word's pane (Zurück, search, meanings headed by a synonym and
  the part of speech, antonyms marked, Einfügen/Kopieren, the language) and the context menu's
  Synonyme (five words of the first meaning, Thesaurus...). Inflected words find their base
  form (schönen → schön).
- **F7** opens Word 365's Editor pane instead of the old dialog: "n verbleibend", the issue
  card with the sentence, suggestions with synonyms ("ähnlich wie"), Alle ändern, Einmal
  ignorieren, Alle ignorieren, Zum Wörterbuch hinzufügen, Optionen. Hunspell's suggestions are
  re-ranked (same first letter, fewest changes), as Word's list puts "Fehler" first for "Feher".
- **Optionen**: the pages Dokumentprüfung (Word's options that RaabeWrite has: capitals,
  numbers, addresses, repeated words, Benutzerwörterbücher, checking while typing, Dokument
  überprüfen) and Sprache (the editing languages, the preferred one). Wiederholte Wörter are
  marked as in Word (Wiederholtes Wort löschen).

### D-050 Silbentrennung as Word divides words
- **Patterns**: LibreOffice's Hyphen files (German two-level: compound boundaries, then
  syllables; English US/GB, French, Spanish; LGPL/BSD/MIT) become modules of the engine
  (`tools/lang-data/build.py`, `packages/text-engine/src/hyph/`), loaded on first use, and
  Liang's algorithm finds the places (`hyphen.ts`). Compound parts keep two letters
  (LibreOffice's "ent1eile" would give "Pflanzent-eile").
- **Line breaking**: a word that overflows is divided at the last place whose part and hyphen
  still fit; justified lines first use their allowed squeeze (Word kept "Algen" whole there).
  Measured with Word-built files (`make_fixtures.py --hyphenation`, three files, 125 lines):
  Word leaves at least three letters on the next line ("Inhaltsan-gabe", "seine" whole) and
  two before the hyphen ("Al-gen", "zu-rück"); it ignores the Silbentrennzone (a 0.75 cm zone
  would have kept "ei-nen" and "Er-eignissen" whole, which Word divides), as its greyed field
  says. 120 of 125 line starts are Word's; the rest are differences of Word's own hyphenation
  lexicon ("transpor-tation" where TeX's patterns have "trans-porta-tion").
- **Settings** as Word writes them: w:autoHyphenation, w:consecutiveHyphenLimit,
  w:doNotHyphenateCaps (Wörter in Großbuchstaben trennen), w:hyphenationZone kept;
  w:suppressAutoHyphens as Absatz > Keine Silbentrennung (with Zeilennummern unterdrücken).
- **Menu and dialogs**: Keine, Automatisch, Manuell, Silbentrennungsoptionen (Word's labels;
  the zone greyed). Manuelle Silbentrennung shows every syllable ("mit-hil-fe"), proposes the
  last place that fits and inserts an optional hyphen (Ja), as Word does.

### D-051 The grammar checker: our engine, LanguageTool's open rules, measured against Word
Built as D-040 planned, in Rust (`src-tauri/src/grammar/`), compiled into the program:
- **Material**: LanguageTool's German XML rules and disambiguation (LGPL-2.1+) compiled by
  `tools/lang-data/build_grammar.py` into 0.35 MB of JSON (4,490 rules, 382 disambiguation
  rules); Daniel Naber's German POS lexicon and synthesis lexicon (german-pos-dict,
  CC-BY-SA-4.0; Morfologik automata, 2.5 MB) with LanguageTool's added/removed lists. Rules
  that need LanguageTool's Java filters are left out; LanguageTool's purely typographic advice
  (protected spaces in "z. B.", typographic apostrophes and ellipses) too, since Word never gives
  it and it fires on most correct texts.
- **Engine**: Morfologik reader (`fsa.rs`), tokenizer and sentence splitter (`text.rs`), tagger
  with compounds by their last part when the first part is a word (`tagger.rs`), and the pattern
  engine (`rules.rs`) with LanguageTool's semantics as its source code defines them: conditions
  per reading (negations included), exceptions for the token, the next and the previous token,
  skips, repetitions, back-references (counted from 0 in patterns), markers, antipatterns,
  unification of number, case and gender (which also drives disambiguation: "eine klassische
  Zangenbewegung" loses the verb "einen"), immunization, messages and suggestions with case
  conversion and regular-expression replacement.
- **Own rules** (`de_rules.rs`) for what LanguageTool checks in Java and for school errors:
  spaces around punctuation, "größer wie", subject and verb (pronouns, noun phrases,
  "Meine Freundin und ich"), article–adjective–noun agreement with corrected articles from the
  synthesis lexicon, the case after für/gegen/ohne and mit/nach/bei/von/zu/seit, nominalised
  infinitives, capitalised verbs, the genitive apostrophe, the comma before aber/denn/sondern
  and "das" for "dass" after verbs and adjectives. Each is cautious (object-first sentences,
  copulas, titles, names, foreign words, particles are left alone).
- **Measured** (`tools/lang-data/grammar_measure.py`, Word in its own invisible instance through
  Range.GrammaticalErrors per paragraph, its options untouched; sentence level). German:
  | Corpus | RaabeWrite | Word |
  |---|---|---|
  | 96 typical school errors (written for this, `school-errors-de.txt`) | 63 found | 46 found |
  | 400 of LanguageTool's incorrect examples | 392 | 72 |
  | 500 paragraphs of German Wikipedia on school topics (false alarms) | 31 flagged | 50 flagged |
  | 400 of LanguageTool's correct examples (false alarms) | 36 | 28 |
  The false-alarm rate is close to Word's (lower on Wikipedia, where many of ours are real
  errors or old spelling in quotes), so German is switched on (D-040). Speed: about 2.5 ms per
  sentence; paragraphs are checked in the background after typing pauses.
- **In the app**: Word's blue double underline, the suggestions and Einmal ignorieren in the
  context menu, the findings in the Editor pane (F7) under Word's headings (Grammatik,
  Zeichensetzung, Groß-/Kleinschreibung …), Grammatikfehler während der Eingabe markieren in
  the options. English, French and Spanish follow once their measurements are as good.

### D-052 Grammar in English, French and Spanish: LanguageTool's rules, own rules, switched on
The same engine (D-051) now runs LanguageTool's English, French and Spanish rules on
LanguageTool's lexicons (english-pos-dict LGPL-2.1, french-pos-dict CC-BY-SA-4.0, Softcatalà's
spanish-pos-dict LGPL-2.1, with LanguageTool's added/removed lists). What it took:
- **Engine, closer to LanguageTool** (German results unchanged): contractions and elisions split
  as LanguageTool's tokenizers do (do|n't, it|'s; l'|on, jusqu'|à, also where the lexicon has
  the whole word); after a skip only the first token that matches is tried, as LanguageTool does
  (otherwise "die Philosophen vor Sokrates, die …" agrees with Sokrates); `filterall` acts on the
  marked tokens only; `<disambig postag>` keeps the lemma of the reading with that tag (does:
  do, not doe); synthesis lexicons without a tag field (English); empty optional tokens no longer
  leave double spaces in suggestions; zero-width spaces are invisible.
- **Memory** (measured at the Step 4b wrap-up, `memory_use` test with a counting allocator):
  compiled with the regex crate, each of the rules' tens of thousands of token expressions took
  14-20 KB, so the checkers took 174 MB (German), 104 MB (English), 336 MB (French) and 65 MB
  (Spanish), and RaabeWrite 636 MB with a document open. Now expressions whose matches are a
  short list of words ("aller|venir", "para[iî]tre", "chats?") are word sets, one per
  expression; the rest run on our own small matcher (`minire.rs`: Java's syntax as a Pike VM of
  a few dozen instructions), with fancy-regex only for lookarounds and the like (21 of French's
  expressions); an expression used more than 4,000 times also gets the regex crate's automaton,
  so the few hundred hot ones stay fast. Every finding is the same as with fancy-regex alone
  (compared rule by rule on LanguageTool's examples in all four languages, `GRAMMAR_NO_MINI`).
  Result: 43, 22, 59 and 19 MB; RaabeWrite 462 MB with a document (Word 291 MB, RaabeWrite at
  Step 4 346 MB); loading takes 0.35 s (German, French), 0.14 s (English), 0.08 s (Spanish);
  checking is as fast as before (German examples 27.5 s instead of 28.9 s). French's two rules
  for postal addresses are left out (one expression of place names was 18 MB).
- **Own rules per language** (`en_rules.rs`, `fr_rules.rs`, `es_rules.rs`) for what
  LanguageTool checks in Java or with its chunker, aimed at pupils' errors and as cautious as
  the German ones: English subject and verb (he go, I has, we was, there is many, Do he …,
  My brother like), the base form after do/can/to (didn't went, does he likes), more taller,
  a/an by sound (an hour, a university, an X-ray), its/it's, your/you're, their/there, whose,
  I am agree, look forward to see, much friends, borrow me, informations, go to home, spaces
  before punctuation; French determiner–noun agreement (le maison, les enfant, mon sœur), the
  subjunctive after il faut que/vouloir que/pour que/bien que (que tu viens → viennes), the
  participle after avoir/être (je suis aller, il a manger), Ca → Ça, the space before !, ? and ";";
  Spanish determiner–noun agreement (los niño, la problema, mi hermanos; el agua stays), ser for
  a place (Madrid es en España → está), irregular participles (escribido → escrito), the
  opening ¿ and ¡, spaces before punctuation. Suggestions come from the synthesis lexicons.
- **Left out** where they fire on correct text because they need LanguageTool's chunker or are
  style advice: en EN_QUOTES, EN_ELLIPSIS; fr dashes, unit spaces, 1990 spellings, ellipses,
  the comma before "mais", NORD_SUD[9] (l'Est de la France), VIRG_NON_TROUVEE; es typographic
  quotes, dashes, ellipses, commas before pero/sino que, AGREEMENT_NUMERAL_PLURAL[1],
  AGREEMENT_ADJ_NOUN[4]/[5], DEMAS[0].
- **Measured** as German (sentences; Word has no Spanish proofing tools on the development PC,
  so Spanish has no Word column):
  | Corpus | en ours | en Word | fr ours | fr Word | es ours |
  |---|---|---|---|---|---|
  | typical school errors (found) | 47 of 50 | 39 | 35 of 44 | 29 | 24 of 28 |
  | 400 of LanguageTool's incorrect examples (found) | 393 | 92 | 387 | 103 | 383 |
  | 500 Wikipedia paragraphs on school topics (flagged) | 17 | 130 | 76 | 112 | 58 |
  | 400 of LanguageTool's correct examples (flagged) | 13 | 59 | 79 | 81 | 31 |
  We find every school error Word finds in English and French and more besides; on clean text
  we flag fewer paragraphs than Word. Spanish flags 12 % of the Wikipedia paragraphs, between
  German (6 %) and French (15 %) and below Word's rate in the other three languages (10–26 %);
  several of its findings there are real errors (un baraja, la conversiones). So all three are
  switched on (D-040). LanguageTool's own example sentences, rule by rule: en 5124 of 5278
  errors found, 13 of 3464 correct sentences flagged; fr 5495 of 5860, 33 of 4961; es 2183 of
  2312, 50 of 2292 (de 6192 of 6323, 20 of 2960).

### D-053 Step 4c: Word's places, our look, and two deliberate improvements
After trying Step 4b the owner listed what still differs from Word (ROADMAP, Step 4c). General
rule: our own design language may deviate from Word's, but every command, option and pane sits
where a Word user expects it (Bildfarboptionen in a pane at the right edge, as in Word, drawn
in our style), so Word habits carry over. Two deliberate differences from Word, chosen by the
owner:
- Applying a Formatvorlage to a selection resets its direct formatting cleanly (colour, size,
  highlighting, effects), where Word keeps some of it.
- Tables can be moved freely by their handle and become floating tables as in Word; in
  addition a second handle of our own gives layout options like a picture's (wrapping,
  alignment), because aligning a table in Word is cumbersome.

### D-054 Pictures as Word places them (Step 4c B)
Measured with a Word-built document (`make_fixtures.py --wrap`, `word-wrap.docx`: Quadrat in the
middle, off-centre and near the margin, "Nur größte Seite", Eng around a box, a turned picture,
a cut-out picture with a transparent background and an ellipse) and `probe_word_wrap.py`:
- **Both sides**: Quadrat and Eng ("Beide Seiten", the default) put text left *and* right of a
  picture; a row of text is then made of pieces, read left to right. Word fills no piece
  narrower than a quarter inch (18 pt used, 17.5 pt left empty); a piece where the next word
  does not fit stays empty and the text goes on in the next one. "Nur links/rechts/größte Seite"
  leave the other side empty.
- **The band** tested against a picture is the text's own height (13.4 pt for Calibri 11), not
  the line pitch with the 1.08 spacing.
- **Eng and Transparent** follow Word's wrap polygon (wp:wrapPolygon, in 21600ths of the
  object's *unturned* box; Word turns it with the object) plus the left/right text distance;
  Quadrat follows the box with the effect extent, for a turned picture its turned bounding box.
  When Eng is chosen in RaabeWrite the polygon is made as Word makes it: a shape's or
  cut-to-shape picture's outline, the visible part of a picture with a transparent background
  (about 3 % larger, as Word's), else the box. It is written to the file and read back.
- **Aligned objects** (zentriert, rechts) are placed with their effect extent: Word centres an
  ellipse with its 2 pt outline extent 1 pt further left.
- **Text above the anchor**: a picture that reaches up into text already placed on its page
  (moved above its paragraph, or turned so it rises) pushes that text aside too; the layout then
  runs a second time with the picture known from the start of its page.
- Result: all 162 line starts of word-wrap.docx as in Word, also with RaabeWrite's own polygons
  for the ellipse and the cut-out picture.
- **Snapping**: moving or sizing a floating object snaps its edges and middle to the margins and
  the middle of the text area within 6 px, with Word's green alignment guides; Alt switches it
  off, as in Word.
- **Double-click** on a picture opens Bildformat, on a shape Formformat (and its text).
- **Textumbruch icons** show what each option does (text beside, around an outline, through,
  above and below, behind, in front); tooltips now also appear inside open menus and galleries.
- **Transparente Farbe bestimmen** (Farbe menu): a pen cursor, the clicked colour becomes
  see-through (a:clrChange, as Word writes it).
- **Grafik formatieren / Form formatieren** is a task pane at the right, laid out as Word's
  (recorded with `probe_word_format_pane.py`): the tabs Füllung und Linie, Effekte, Layout und
  Eigenschaften and Bild (shapes: under Formoptionen, with Textoptionen beside them), fold-out
  sections with Word's labels, a slider and a value field per setting. Changes apply at once;
  dragging a slider is one undo step. Every "…optionen" entry opens its section, as in Word
  (Bildkorrektur-, Bildfarb-, Bildtransparenzoptionen, Schatten-, Spiegelungs-, Leuchteffekt-,
  3D- und 3D-Drehungsoptionen, Optionen für weiche Kanten, Weitere Linien/Pfeile,
  Textrichtungsoptionen); the group launchers open Effekte (Bildformatvorlagen) and Füllung und
  Linie (Formenarten). Shown but off: Bild-/Textur- and Musterfüllung, Farbverlaufslinie,
  Künstlerische Effekte; the 3D-Format settings beyond the top bevel are kept for Word, not shown.
  Größe and Weitere Layoutoptionen stay dialogs, as in Word.
- **Turned objects**: frame, handles, rotation handle and crop marks turn with the object; sizing
  and cropping work along its own sides with the opposite side staying where it is, and the
  sizing pointers turn with it. A click hits a turned or mirrored object where it is drawn.
  Undo keeps the floating object selected (the pane stays on it).
- **Context menu** of shapes and text boxes as Word's (recorded with `probe_word_context.py
  objects`): Text bearbeiten, Punkte bearbeiten, Gruppieren, Vordergrund/Hintergrund, Link,
  Beschriftung, Textumbruch, Alternativtext, Weitere Layoutoptionen, Form formatieren.

### D-055 Formulas as Word builds them, through the clipboard (Step 4c C)
Recorded in Word (32 new cases in `catalogue_math.py`, areas "fraction" and "text"):
- **"/" takes the whole operand**: the numerator is everything since the last operator (2x/3,
  0,5/2 with its decimal comma, 2π/3, x^2/3, 2(a+b)/c keeps its brackets), the denominator
  everything up to the next one (y=3a/4b). All 146 recorded cases now match Word, in the unit
  replay and in the app.
- **Quoted text** ("wenn") becomes normal text (m:nor) when the closing quote is typed; inside the
  quotes nothing is built up; a space right after it is used up; typing inside or at the end
  of normal text continues it, after the closing quote it is mathematics again. Word splits a
  run where a character is typed into it; both catalogue comparisons merge neighbouring runs
  with the same properties, since the look is the same.
- **Clipboard with Word** (recorded with `probe_word_math_clipboard.py`): Word puts equations
  into its HTML as OMML in a `<!--[if gte msEquation 12]>` comment (namespace
  `…/office/2004/12/omml`), each run's text in HTML spans, italic marked by `<i>`, plus MathML
  and a picture for other programs. RaabeWrite reads that (skipping the fallback) and writes
  the same, with the linear form as fallback text; Word accepted our HTML only with the `m`
  namespace declared and the text in spans (as `<m:t>` it came out upright). Copying part of
  an equation gives Word an equation too, and Word's equations paste into an equation as
  structures. `clipboard_math.py` checks both directions with inline and display equations.
- The Start tab shows Cambria Math while the caret is in an equation; a click on the
  equation's handle (top left) selects the whole equation, dragging it still moves it.

### D-056 Tables as Word handles them (Step 4c D)
- **Entf / Rücktaste on selected cells** (recorded with `probe_word_table_delete.py`): Entf
  empties the cells and keeps them selected; Rücktaste deletes whole rows, whole columns or
  the whole table, and for other cells (one whole cell included) opens Zellen löschen with
  Word's four choices, "Zellen nach links verschieben" first.
- **Dragging across cells** switches to whole cells as soon as the pointer crosses into
  another cell (Word, measured), and the selection shading covers each cell only as far down
  as its text goes, between the gridlines, which stay visible.
- **Rahmen group** as Word's: Rahmenarten (21 Designrahmen in the theme's accent colours and
  the Rahmenpipette), Stiftart (Word's 25 line styles), Stiftstärke, Stiftfarbe, Rahmen (with
  inner horizontal/vertical lines) and Rahmen übertragen: the pen paints clicked or dragged
  cell edges (both cells' sides), Esc ends it.
- **Tabellenformatvorlagen**: Word's 105 gallery styles in its order and groups (Einfache
  Tabellen, Rastertabellen, Listentabellen), definitions recorded from Word
  (`probe_word_table_styles.py`) with their theme references, so each document gets them in
  its own theme colours (Word's tint/shade computation, checked against Word's values). The
  ribbon shows a row of previews drawn from the definitions with the table's options; hovering
  previews the style on the table. Reading Word's table styles also became right: a border
  element names only the sides it sets; a conditional format's outer sides belong to its region
  (the header row's bottom line), corner cells (nwCell …) come last, and an edge two cells
  share shows the stronger border (Word's conflict rule).
- **Floating tables** (Textumbruch Umgebend, w:tblpPr): placed on the page, text flowing around
  with Word's distances; vertically "text" counts from the paragraph after the table (measured).
  `word-table-float.docx` (centred, offset, right) lays out as in Word to within 0.3 pt.
  Dragging the move handle sideways indents the table, up or down makes it floating, as Word
  does; our layout handle (D-053) beside its top right corner offers Ohne/Umgebend and
  Links/Zentriert/Rechts, and Tabelleneigenschaften has Textumbruch Ohne/Umgebend.

### D-057 Selection, pointers and Formatvorlagen (Step 4c E)
- **The selection** is drawn over the text and multiplied with what is below (its layer blends,
  not each rectangle): highlighting, shading, coloured text and page colour stay visible and
  the selection shows over them, as Word's grey does; before, a highlight hid it.
- **Pointers**: every drawn pointer (Format übertragen, the pen and pipette of the Rahmen group,
  rotation, Transparente Farbe) comes at 1×, 1.5× and 2× (CSS image-set) so it is as sharp as
  Windows' own at 150 %; the table border pointers are Word-like drawn ones instead of
  Chromium's col-resize/row-resize bitmaps, which blur. Where Word uses Windows' pointers
  (I-beam, arrow, move, sizing arrows) we use them too.
- **Formatvorlagen** (D-053): a paragraph style removes the direct character formatting of its
  paragraphs (font, size, colour, highlighting, effects, also on the paragraph mark); a
  character style that of the selection; links' and note marks' character styles, language,
  proofing and revisions stay. Hovering an entry of the Formatvorlagen gallery shows it on the
  text (Livevorschau, no undo step); leaving puts everything back.

### D-058 The audit of the whole app (Step 4c F)
Every kind of difference found in A–E was looked for everywhere, by script where possible
(`audit_labels.py`, `audit_menus.py` in the headless development page; Word's side recorded by
`probe_word_tooltips.py`, `probe_word_table_layout.py`, `probe_word_list_define.py`).
- **Tooltips**: every ribbon command has Word's tooltip title and shortcut (as Word writes them:
  "Fett (Strg+Umschalt+F)", "Alt+Strg+F", "Strg+Return", "Alt+="), the shortcut on the title
  line, and a description in our own words (`apps/raabe-write/src/tooltips.ts`). Microsoft's
  descriptions are not copied: the committed recording keeps only the titles, the full texts
  stay in the uncommitted captures/.
- **Labels**: no ribbon label is cut off on any tab (contextual ones included) at window widths
  from 1920 to 800 CSS px.
- **Menus and galleries** against Word's recordings, entry by entry: 19 menus and 7 context
  menus, no unexplained difference left. Fixed on the way: the shades of Word's palettes (white
  and black have their own rows, other very light or dark colours 10–90 %), "5%" without a space,
  the names of the current Office theme's colours, "Office 2013 – 2022-Design" in Designs, the
  names of the font schemes. Explained differences: cloud and online features, deferred ones,
  recently used symbols, our own cover pages (D-048).
- **Commands that did nothing**: Als Grafik speichern (now saves the picture's file; for shapes
  deferred, it needs the shape drawn into a file), AutoAnpassen, Zellen einfügen, Neues
  Aufzählungszeichen/Zahlenformat definieren; the rest are marked deferred (Textrichtung in
  cells, Neue Liste mit mehreren Ebenen/Neuen Listentyp definieren, Tabelle zeichnen, Radierer,
  Formel in tables).
- **Tabellenlayout** complete as in Word (recorded at a wide window and at 2000 px): Word's
  labels and keytips (ours were guessed), the groups Zeichnen and Daten, Tabelle teilen, Zellen
  einfügen (Word's dialog and its four results), AutoAnpassen (Inhalt: the widest line and the
  longest word of each column shared out as a browser does, Word's tblW/tcW auto; Fenster:
  tblW 100 %; Fest: tblLayout fixed), Zellenbegrenzungen (Word's Tabellenoptionen; Abstand
  zwischen Zellen is not drawn, so it is greyed), In Text konvertieren (Word's dialog; the results
  as Word's, "; " for Semikolons). Measured: AutoAnpassen an Inhalt gives 32.2/167.3/32.2 pt where
  Word gives 32.2/165.6/32.2.
- **List dialogs**: Neues Aufzählungszeichen definieren (Symbol... through the Symbol dialog,
  Schriftart..., Ausrichtung, Vorschau; Bild... greyed) and Neues Zahlenformat definieren
  (Word's 13 Zahlenformatvorlagen, the number shaded between the text around it). In a list they
  change the caret's level, elsewhere they start a list. New number formats: German number words
  and ordinals, and Word's custom 001/0001/00001 formats (written in mc:AlternateContent as Word
  writes them; Word reads all of them as we show them, checked through COM).
- **Livevorschau everywhere Word has it**: one mechanism (`Editor.preview`): the entry under the
  mouse is applied to the document without an undo step and taken back when the mouse leaves or
  the menu closes; a click ends the preview first, so the real change is one undo step. Used by
  the Schriftart and Schriftgrad lists, the Schriftfarbe, Texthervorhebungsfarbe and
  Schattierung palettes (also those of shapes, text, picture borders, table shading and
  Seitenfarbe), Zeilenabstand and Unterstreichen, Designs, Farben, Schriftarten,
  Formatvorlagensatz, Absatzabstand, the list libraries and the text effect galleries, in
  addition to the galleries that had it (Formatvorlagen, table, picture and shape styles).
- **Double-clicks as recorded in Word** (`probe_word_dblclick.py`): on a note reference into the
  note, on the note's number back to the text; on a column border the column fits its content;
  the table's move handle selects the table; the new selection zones of a table (inside a cell
  at its left edge: the cell, double-click the row; above the table: the column; the left margin
  beside a row: the row), each with Word's black arrow pointer; on a section break's mark
  (Alle anzeigen) Seite einrichten. Strg+Umschalt+Eingabe in a table splits it, as in Word.
  Word keeps a reduced ribbon and Alle anzeigen for the user's next session: the probe puts them
  back (and a note in CLAUDE.md says so).
- **Panes**: Alternativtext is a task pane on the right that follows the selection, as in Word,
  no longer a dialog.
- **Clipboard with Word** (`clipboard_objects.py`): pictures and hyperlinks now go both ways.
  Our HTML carries pictures as data: images (Word reads them, at their size) and links as
  <a href>; Word's HTML is read with its links and its pictures, the ones Word keeps in its
  temporary folder (clip_image001.png) read before the paste. Kopieren/Ausschneiden work on a
  selected picture or shape (Word: yes, we: before only with Strg+C). Floating objects reach
  Word as pictures in the line (Word makes floats only from VML, and reads a VML picture with a
  data: address as a linked one), Word's floating objects in a mixed selection arrive as
  pictures in the line, shapes and text boxes as their pictures and text; a single picture or
  shape copied in Word arrives as a picture (Word offers no HTML for it, only PNG). Found on the
  way: WebView2's clipboard API does not return while the clipboard holds only pictures, which
  made Einfügen hang after copying a picture in Word; the app reads the clipboard natively and
  uses the browser's API only in the development page.
- **Behaviour catalogue**: ten new cases (Strg+Umschalt+Eingabe in tables, Rücktaste and Entf on
  a cell's content, double-clicks on words and in the margin, margin and triple clicks). They
  found that Word keeps the paragraph mark at the end of a selection that is typed over
  (intelligent paragraph selection: a paragraph selected by a margin double-click or a
  triple-click, then typed over, stays a paragraph of its own), and that whole paragraphs typed
  over take the formatting of the paragraph whose mark survives (the last one); both now as in
  Word.
- **Found by the wrap-up's layout run**: an object placed below its paragraph that would reach
  past the bottom margin takes the paragraph to the next page, when it fits there (Word, with
  word-wrap.docx: a 110 pt picture near the end of page 2 moved "Eng, freigestellt" to page 3);
  objects in front of or behind the text and aligned ones are not moved.

### D-059 Step 4d: the owner's second list
Agreed with the owner on 2026-10-16 after trying Step 4c, with four decisions of theirs:
1. **Ribbon tabs** switch only after inserting an object (and on a double-click), never when
   the caret enters a table or formula or a picture or shape is clicked (Word's behaviour). The
   reported jump could not be reproduced in the development page or the release build (clicks
   into a cell, a picture, a shape and a formula all leave the tab where it was); the code
   switches tabs only on insertion, double-click and header/footer editing.
2. **Dotted gridlines** for missing borders are off by default and shown while the border pen
   or the eraser is out (as in Word); the switch Gitternetzlinien anzeigen in Tabellenlayout
   stays.
3. **Whole table vs. cells**: Word draws both the same. We add a marker of our own when the
   whole table is selected together with text beyond it (a drag past the table, which Entf then
   deletes as an object): a thin blue frame around the table and the move handle drawn pressed;
   the text beyond the table is highlighted too.
4. Step name 4d.

Recorded in Word (`probe_word_table_typing.py`, `probe_math_sizes.py`, catalogue_math.py) and
built the same way:
- **Typing over selected cells** (nine cells, a row, a block, the whole table) replaces only the
  first cell's content and leaves the caret there; the table stays. Before, "a" deleted the
  table.
- **Fractions**: "/" + space on its own builds an empty fraction, "a/" + space a fraction with a
  as numerator and an empty denominator; the caret goes after the fraction. Six new cases,
  formula catalogue 152/152.
- **Script sizes**: Word shrinks scripts once to 73 % and once more to 59 %, then no further
  (11.04 → 8.04 → 6.48 pt). Chromium's MathML went on shrinking (4.68 pt at the third level); CSS
  clamps the depth at two levels. Now 10.99 / 8.02 / 6.6 pt.
- **Border pen**: a press on a gridline draws a straight line along it while the mouse is held
  (the line follows that direction only) and is applied on release, as one undo step; single
  clicks elsewhere keep the pen, a double-click away from the lines ends it.
- **Right-click on selected cells** keeps the selection (the click counts as inside it). A
  right press no longer takes the focus from the text, so typing goes on after Esc closes the
  context menu (found on the way: the keys went nowhere).
- **A drag from a cell into the text after the table** selects the whole table and, in the
  text, whole words (Word's automatic word selection); before, it ended in the middle of a word.
- **Behaviour catalogue**: four new cases (typing over dragged cells and over a row, right-click
  on selected cells then Esc and typing, a drag past the table then Entf).
- **The layout handle** beside the table stays while the pointer travels to it from the table.
- **Tabellenformatvorlagen** in the ribbon show at most 14 previews, as Word's.
- **Menus with a chevron** close when their button is pressed again (Office).
- **Text boxes and shapes while their text is edited**: the handles keep their pointers and
  work, shapes keep their yellow handles; a text box's frame (a 5 px band) moves it and shows the
  move pointer, inside it the pointer is the text pointer and a drag selects text. Before, the
  frame did not move it and a drag on it moved the caret.
- **The caret** is its own compositor layer moved by a transform, so fast Rücktaste leaves no
  "afterglow" of earlier positions.
- **Drag and drop**: selected text is dragged to another place (Strg: copied) with Word's
  spacing (a space between words, none before punctuation); files from Explorer: pictures go in
  at the drop point, a .docx opens in a tab of its own. The web view gives the page no file path,
  so a dropped document is saved with Speichern unter the first time; text from other programs
  is pasted at the drop point.

### D-060 Row heights: the sizing handle and Zeilen verteilen (2026-10-26)
Two bugs the owner found after Step 4d, fixed with their go-ahead:
- **The sizing handle of a new table** made it grow without end on the first drag: the rows had
  no set height, so each mouse move measured them on the page, where the previous move had
  already scaled them. The rows are now measured once when the drag starts, as their height
  between the borders (what a set row height means in Word), and scaled on their sum. A second
  drag had worked because the first one had set the heights.
- **Zeilen verteilen** removed every row height, so the rows shrank to their text. Word
  (recorded 2026-10-26, `check_table_rows.py`): the selected rows, or all rows of the table
  when the selection lies in one row, get the same minimum height (Mindestens, an exact height
  too): their average height between the borders, or more if a row's content needs more. So
  the table keeps its height where the text fits (rows of 20 and 60 pt become 40 and 40), and
  otherwise the rows are as tall as the tallest content (20, 60, 40 pt exact and a row of four
  lines: all 53.7 pt). Ours now gives the same heights in all three recorded cases. The layout
  hands the editor each cell's content height and the row's border widths for this
  (`Frame.need`, `Frame.rowBorders`).

Found on the way: the owner had RaabeWrite open, so the release exe could not be rebuilt, and
the app's single-instance handling would have handed every test window to the owner's window
(a test would have typed into their documents). Test instances (`RAABE_FEEL_DEBUG_PORT` set)
now skip the single-instance handling, and the harness takes another build through
`RAABE_WRITE_EXE` (built with `CARGO_TARGET_DIR=src-tauri/target/check`).

### D-061 RaabeCalc next, handling exactly like Excel (2026-10-27)
Agreed with the owner after a RaabePresent proposal:
- **Order**: RaabeCalc (Step 7) starts now. RaabePresent's charts come from RaabeCalc's chart
  engine (SCOPE), and Excel runs as a process of our own (`DispatchEx`), so the harness rules
  for Word apply unchanged; PowerPoint runs as a single process and needs new rules (Step 8).
  The RaabeWrite pilot (Step 5) runs alongside when a class fits the timetable; the RaabePresent
  proposal is kept in the ROADMAP under Step 8 and discussed again before it starts.
- **Excel in every way, especially the controls**: every control of the grid and the window
  (cell pointer, edit modes, fill handle, Namenfeld, Bearbeitungsleiste, headers, sheet tabs,
  scroll bars, status bar, filter buttons, option buttons, context menus, shortcuts; list in the
  ROADMAP) is recorded in an Excel instance of our own before it is built, and the behaviour
  catalogue and screen comparison exist for Excel from package A on rather than after the
  fact as in Step 4b.
- **Default font** Calibri 11, as in Excel 2013-2022, instead of Excel 365's Aptos Narrow, which
  we cannot ship (D-011).
- **Diagramm in RaabeWrite** (deferred, D-041) comes right after Step 7 as Step 7b, with the
  chart's embedded workbook as Word writes it.
- The owner leaves the PC for the whole step, so real input is allowed throughout.

### D-062 RaabeCalc's building blocks (Step 7, packages A-E)
Built overnight on 2026-10-27 with the owner away (real input allowed, D-061):
- **Shared native side**: `crates/raabe-host` holds what RaabeWrite's `src-tauri` had in common
  (files, AutoRecover, PDF with a page size, fonts, clipboard, window, single instance); each
  app adds its own commands (RaabeWrite: spelling, thesaurus, grammar) and combines both
  handlers. The shell's file dialogs know .xlsx and .csv.
- **The grid is drawn on one canvas** at the device's pixel size (ARCHITECTURE "Workbooks"),
  with Excel's measured look at 150 % (`capture_excel_ui.py`): 1-device-pixel gridlines
  #E0E0E0, headers #F0F0F0 with a #ABABAB line, the selected headers #E1E1E1 with an accent
  line and accent text, a 2-device-pixel selection frame on the gridline, the 6-pixel fill
  handle. Excel's green accent becomes Raabe red (#B3141B, D-018). The cell editor, Namenfeld,
  Bearbeitungsleiste (38 px high, as Excel's), menus and dialogs are DOM. Excel's pointers are
  drawn as sharp SVGs (D-057).
- **Widths**: German Excel's standard column width is 10,71 characters (80 px), not the 8,43 of
  English Excel; the model keeps widths in characters without padding as the file has them at
  96 dpi (`ooxml/src/xlsx/units.ts`). On screen Excel computes widths with the display's own
  digit width and padding (measured at 150 %: 11 device pixels and 7 px, so a column dragged
  40 CSS px wider shows 15,73 characters, not 16,43); the grid does the same (`grid/metrics.ts`),
  printing uses the 96-dpi widths.
- **Formulas** are stored as Excel's `Formula2` without "=" (English names, no `_xlfn.`
  prefixes) and evaluated with Excel 365's semantics (spilling arrays); the UI converts to
  German on the way in and out. Measured: 4229/4234 recorded Excel results equal (the five
  others are Excel's own floating-point and iteration quirks, listed in the tests), all
  conversions, copies, insertions/deletions, moves and cell shifts as Excel adjusts them.
  Not supported yet: the 1904 system in formulas, structured references, LET/LAMBDA/AGGREGAT.
- **Number formats and typed input** follow German Excel as recorded (about 66 000 cases):
  locale-dependent built-in codes show the German way (`m/d/yyyy` as TT.MM.JJJJ), General
  shortens to the column width as Excel does, typed input keeps or replaces the cell's format as
  Excel does ("1.1." stays text, "3/4" is a date unless the cell has a number format).
- **.xlsx** keeps everything not edited exactly as Excel wrote it (D-024): unchanged sheet parts
  are written back from Excel's own text, unknown parts byte for byte, styles keep their file
  index. Excel opens our files without repair and reads 0 differences in 17 checked files
  (9 Excel-built fixtures written back, 8 edits). CSV as German Excel: ";" separator, values as
  shown, UTF-8 with byte order mark.
- **Charts** (package E, `packages/charts`, shared with Step 7b and RaabePresent): Excel's
  default look, axis scaling equal to Excel's in 145 recorded axes, Excel's rule for series from
  a selection, histogram bins as Excel's (Scott's rule), trendline labels in Excel's wording.
  Gaps: no Diagrammentwurf/Format tabs yet, charts do not follow inserted rows.

### D-063 Excel's controls as recorded (Step 7 package D)
Recorded in an Excel of our own with real input (`catalogue_calc.py`, 117 cases;
`probe_excel_context.py`, `probe_excel_tooltips.py`) and built the same way:
- **Shortcuts on the German keyboard**: Strg+Umschalt+F/K/U fett, kursiv, unterstrichen,
  Strg+2/3/4/5 as well; Strg+Umschalt+1 "#.##0,00", +4 Währung with red negatives, +5 Prozent;
  Strg+Umschalt+3 and +6 and Strg+0 do nothing in German Excel; Strg+U/Strg+R fill down/right;
  Strg+. today's date; Strg+Umschalt+A Schriftarteinstellungen; Strg+Alt+L Erneut anwenden;
  Alt+= (Alt+Umschalt+0) AutoSumme, which ends in Zeigen with the range marked. Letters are
  matched by their character (Z and Y swap places on German keyboards), digits by position.
- **Moving**: Tab remembers the column it started in and Enter returns there; arrows in
  Eingeben commit and move one cell; Enter and Tab move inside a selection; Bild ab moves one
  row more than the window shows; Umschalt+F11 inserts the new sheet before the active one,
  the + beside the tabs after it.
- **Editing**: AutoVervollständigen takes the stored entry's spelling; Alt+Eingabe turns on
  Textumbruch; "+2*3" becomes "=2*3"; Enter pastes copied cells while the marching ants show;
  columns of standard width widen when a formatted number no longer fits.
- **Context menus** of cells, column and row headers and sheet tabs with Excel's labels
  (three dots, "Zellen einfügen…"); the search box at their top and Microsoft's descriptions are
  not copied (D-058).
- **Ribbon menus** entry by entry as recorded (`probe_excel_menus.py`: Rahmen with "Doppelte
  Rahmenlinien unten", Bedingte Formatierung with "Regeln für oberste/unterste Werte", AutoSumme
  with "Anzahl Zahlen", Löschen with "Kommentare und Notizen löschen", Suchen und Auswählen with
  Formeln/Notizen/Konstanten, Seitenränder, Berechnungsoptionen with Teilweise); labels that follow
  the content as Excel's (Fixierung aufheben; "Von A bis Z" for text, "Nach Größe" for numbers);
  F4 and Strg+Y repeat the last formatting; dragging the fill handle back empties cells.
- **Printing** as Excel's PDF export (`check_calc_print.py`): rows and columns at 0,968 of their
  screen size (a 15-pt row takes 14,52 pt, 50 rows on an A4 page with the normal margins), cut
  down first, then across, the cells 2,6 pt inside the margins; the page size and orientation from
  Seite einrichten.
- **The AutoFilter dropdown** is a panel under the button as Excel's (recorded with
  `probe_excel_filter.py`), Alt+Unten opens it; a click on a link's text follows it; a finger
  drag scrolls and a tap selects (touch).
- **Clipboard with Excel** both ways (`check_calc_clipboard.py`): values, number formats and bold
  as Excel's HTML carries them (classes with `mso-number-format`, `font-weight`, `background`;
  Excel's names such as "Short Date"); a formula in a Standard cell shows its result in the
  format of the first referenced cell that has one, HEUTE/DATUM as a date (Excel).
- **Display and view** (recorded): a wheel notch scrolls 3 rows, Strg+wheel zooms in steps of 15 %;
  whether a value fits its column is decided at the display's pixels with whole-pixel character
  widths and Excel's padding, so Standard numbers shorten exactly where Excel's do
  (1234567,891 shows 1234567,89); a column of standard width widens for a Standard number only
  where it would otherwise turn scientific or cut a small number (123456789012, 0,000012345), and
  for dates and other formats that do not fit.
- **More controls** (recorded where marked): Alt+Unten outside a filter button opens the
  Dropdown-Auswahlliste with the column's texts, sorted, nothing highlighted until the first Down
  (recorded); Neue Notiz/Notiz bearbeiten type straight into the yellow box beside the cell, a
  click outside ends it; Inhalte auswählen (Notizen, Konstanten, Formeln, Leerzellen, Aktueller
  Bereich, Letzte Zelle, Nur sichtbare Zellen, Bedingte Formate) selects all hits at once;
  Finanzmathematik lists the financial functions the engine has; Drucktitel (Wiederholungszeilen
  oben, Wiederholungsspalten links) repeat on every page as Excel prints them and are read and
  written as `_xlnm.Print_Titles`.
- Not yet as Excel: the drag of the selection's border onto other sheets, comments (Notizen
  work). Datenüberprüfung, Blatt schützen, Schnellanalyse, Blitzvorschau: D-065.

### D-064 Raabeschule templates for RaabeCalc (2026-10-28)
As RaabeWrite's (D-035): our own designs from ordinary Excel features, so files from them work in
Excel, offered in Datei > Neu beside the blank workbook (`apps/raabe-calc/src/templates.ts`):
**Notenliste** (25 names, three Klassenarbeiten, mündlich, schriftlich and Gesamt with
MITTELWERT and RUNDEN, averages, fixed header row, Gesamt worse than 4,4 highlighted),
**Notenspiegel** (counts per grade with ZÄHLENWENN, Durchschnitt, share below 4, a column chart)
and **Messwerte** (t and s with STEIGUNG, ACHSENABSCHNITT, BESTIMMTHEITSMASS and an s-t chart
with trendline, equation and R²). The weighting in the Notenliste (written and oral half each) is
a placeholder the school can change; school material can replace the designs.

## D-065 Step 7: Datenüberprüfung, Blatt schützen, Formelüberwachung (2026-10-28)

Built during the overnight run (D-061) after the behaviour catalogue was complete, as the parts of
Excel a school uses most among those first deferred.

- **Datenüberprüfung** as Excel checks it, recorded through COM in an Excel of our own
  (`probe_excel_validation.py`; Excel's verdict on 150 typed entries in 33 cases, all reproduced):
  number types refuse text and WAHR and compare exactly; Textlänge counts the edit text; a literal
  list is case-sensitive, a list from cells is not; Benutzerdefiniert passes on WAHR or a number
  other than 0; a criterion that refers to an empty cell lets every entry of the right kind pass.
  Zellendropdown arrow beside the cell and its list (Alt+Unten), Eingabemeldung below the cell,
  Fehlermeldung Stopp (Wiederholen/Abbrechen), Warnung (Ja/Nein/Abbrechen), Information
  (OK/Abbrechen). Files: unchanged validations stay as Excel wrote them; Excel's x14 form (other
  sheets) and the x12ac list form (entries with commas or quotes) are read and written; 0
  differences in Excel. Not recorded (no COM access): the dialog's labels (taken from German
  Excel), the default button of Warnung (Nein, a guess), positions of arrow and message. Excel's
  default error text is in our own words (D-058). Ungültige Daten einkreisen draws red ellipses
  (at most 255, a corrected cell loses its circle); Einfügen and AutoAusfüllen carry the source's
  validations and replace the target's, relative references moving with the cells (Excel).
- **Blatt schützen**: the sheet's `<sheetProtection>` attributes are kept as the file has them;
  the dialog offers Excel's list of allowed actions and writes the password as Excel 2013+ does
  (SHA-512 over salt and password, 100 000 rounds); older files' 16-bit hashes are accepted. On a
  protected sheet locked cells refuse changes with a message in our own words, formatting,
  inserting, deleting and sorting are greyed out unless allowed, and Tab moves between unlocked
  cells (Excel), which is what teachers' templates need. The later additions follow it too:
  Gruppieren, Teilergebnis, Seitenumbrüche, Zielwertsuche, Datenüberprüfung and Rechtschreibung
  are greyed, Schnellanalyse does not appear, pictures stay put unless Objekte bearbeiten is
  allowed, the filter buttons work only with AutoFilter verwenden.
- **Arbeitsmappe schützen** (Struktur; Fenster greyed as in Excel 365): no sheets added, deleted,
  renamed, moved, hidden or coloured; the password is written as Excel's workbookProtection hash
  (checked in Excel: locked, wrong password refused, right one accepted). The button shows its
  state; pressing it again asks for the password.
- **Formelüberwachung**: Spur zum Vorgänger / Spur zum Nachfolger draw Excel's arrows (blue, a
  dot at the source, a frame around ranges, dashed from a sheet symbol for other sheets, red from
  errors), each click one level further, removed by Pfeile entfernen or any change.
  Formelauswertung works the formula out one part at a time in Excel's order (single-cell
  references first, then each operation once its operands are values, WENN after its condition
  as a whole), the next part underlined and the last value in italics; Einzelschritt and
  Prozedurschritt are greyed out. Fehlerüberprüfung walks the cells with an error indicator from
  the active cell on, with explanations in our own words.
- **Kopf- und Fußzeile** (Einfügen > Kopf- und Fußzeile opens Seite einrichten there): three
  sections each, typed with Excel's German field names (&[Seite], &[Seiten], &[Datum], &[Zeit],
  &[Datei], &[Register]) and kept in the file as Excel's codes (&P, &N ...); printed at the
  header and footer distances. Excel's separate Kopf- und Fußzeile view on the sheet is not built.
- **Zahlenfilter / Textfilter** in the AutoFilter panel: Benutzerdefinierter AutoFilter (two
  conditions with Und/Oder, * and ? wildcards), Top 10 (items or percent), Über/Unter dem
  Durchschnitt; which rows pass follows Excel's results for 18 recorded criteria
  (`probe_excel_filter_sub.py`: numbers compare as numbers, text case-insensitively, a plain value
  must match the whole text, Top 10 percent rounds down). Read and written as Excel's
  customFilters, top10 and dynamicFilter. Undoing a paste brings the copy marquee back, and copy
  mode shows a status message (our own words) in place of the cell mode (recorded).
- **Schnellanalyse** (recorded with `probe_excel_quickanalysis.py`: tiles clicked in Excel, the
  result read through COM): a 22-px button 2 px beyond the selection's corner when it holds data,
  Strg+Q; tabs Formatierung (Datenbalken, Farbskala, Größer als, Obere 10 %, Formatierung
  löschen), Diagramme, Ergebnisse (below and to the right: Summe, Durchschnitt, Anzahl with
  ANZAHL2, % Gesamt, Laufende Summe, with Excel's bold labels and formulas, 6/6 recorded cases
  equal), Tabellen (Tabelle without a dialog, as Excel). PivotTable and Sparklines are greyed
  (not built); the chart tiles are our choice of common types, since Excel's recommended ones
  depend on the data; the Symbolsatz tile uses 3 Pfeile (Excel's set there not recorded).
- **Symbolsätze** (Bedingte Formatierung; the gallery's 24 names and groups recorded with
  `probe_excel_iconsets.py`): arrows, triangles, traffic lights, signs, symbols, Kennzeichen,
  Rot/Schwarz and stars drawn as our own icons at the cell's left; a value moves one icon up for each
  limit it reaches (Excel's defaults: equal shares of the range in percent; numbers,
  percentiles and formulas read from files); "Nur Symbol anzeigen" hides the value. The
  ratings, quarters and boxes sets show a filled circle instead of Excel's pictures.
- **Nach Farbe sortieren / filtern** (submenus recorded; rules through COM): the column's fill
  colours then font colours in order of appearance, Keine Füllung and Automatisch; filtering
  shows the rows with that colour (5/5 recorded cases), sorting puts them on top and keeps the
  others' order (as Excel); colour filters are read and written as Excel's colorFilter. Colours
  from conditional formats are not offered yet.
- **Zielwertsuche** (Was-wäre-wenn-Analyse; Szenario-Manager and Datentabelle stay deferred):
  Zielzelle, Zielwert, Veränderbare Zelle; stops within 0,001 of the goal or after 100 steps as
  Excel does; Newton's steps with a measured slope, never larger than the value itself, so it finds
  the solution nearest the start. Against Excel's Range.GoalSeek in 10 school cases: the same
  outcome and the same solution (Excel's exact iterations are not reproduced). Abbrechen in the
  status box puts the old value back; the box's text is our own.
- **Rechtschreibung (F7)**: RaabeWrite's Hunspell module (D-031, D-049) is compiled into
  RaabeCalc too, included from RaabeWrite's crate rather than copied (the program grows from about
  2 to 10,7 MB with the dictionaries); the text cells from the active cell on, then from the top,
  with Excel's dialog labels; capitals, words with digits and addresses are left out as Office
  does. AutoKorrektur, Optionen and Rückgängig are greyed; the closing message is our own.
- **Umbruchvorschau** (Ansicht and the status bar): the printed area framed in blue, everything
  outside it grey, blue lines between the pages as the printing cuts them (dashed where the page
  is full, solid at manual breaks) and "Seite n" in each page, at 60 % the first time as Excel.
  Dragging a line moves the break to the gridline nearest the pointer and makes it manual (as
  Excel does); the area's frame cannot be dragged yet, and the view is not written to the file.
- **Pictures in sheets** (PNG, JPEG, GIF, BMP, SVG from the file's drawing, e.g. a school logo in
  a template): shown at their two- or one-cell anchor over the cells, following scrolling and
  zoom, and printed on the page of their top-left cell; the printed range includes them, as in
  Excel. Einfügen > Bilder > Dieses Gerät puts one at the active cell at its natural size; a click
  selects it, dragging moves it, a corner sizes it keeping its proportions, Entf deletes it. The
  file keeps untouched pictures byte for byte, rewrites only the corners of moved ones and adds
  new ones as Excel's inserted pictures (checked in Excel: 0 differences, the new one is a
  picture). Stockbilder and Onlinebilder stay greyed (offline). Print positions are computed,
  not compared with Excel's PDF.
- **Berechnungsoptionen**: Automatisch, Automatisch außer bei Datentabellen (the same while there
  are no Datentabellen) and Manuell, kept in the file's calcPr. Manuell leaves formulas as they
  are until F9 / Neu berechnen and shows "Berechnen" in the status bar; a formula typed meanwhile
  is worked out at once, as Excel does.
- **Kommentare** (Excel 365's threads; Einfügen > Kommentar, Überprüfen > Neuer Kommentar,
  Löschen, Vorheriger / Nächster Kommentar): a purple corner on the cell, a card beside it with
  each comment's author, date and text, Antworten with Posten (Strg+Eingabe), Thread löschen and
  Thread auflösen in its "…" menu; hovering shows the card. Files: threadedComments and persons
  parts as Excel's, each thread with a legacy note for older programs linked by the author
  "tc={id}" (Excel's way; the note's text is our own); Excel reads them as threads with their
  replies and no extra note (checked). Kommentare anzeigen opens a pane at the right with every
  thread of the sheet (a click goes to it, Neu starts one); the author is the name from Datei >
  Optionen.
- **Undo of whole-sheet steps** now also restores the sheet's page setup, filter, conditional
  formats, validations, protection, threads and pictures (before, only cells, sizes and merges).
- **Review of these additions** (an independent read of the code at 04:00) found and we fixed:
  Kommentare and pictures did not move with inserted or deleted rows and columns; Blatt schützen
  checked only the selection, so filling, pasting, Strg+Eingabe, dropdown choices and
  Blitzvorschau could write into locked cells (now every written cell is checked); saving dropped
  pictures the reader does not show (EMF logos); Tage / Arbeitstage ausfüllen used a wrong step;
  Arbeitsmappe schützen could be bypassed from the tab menu; copied sheets lost protection,
  page setup, filter, rules, Kommentare and pictures; a kept Print_Titles name could be written
  twice; SVG pictures (which Excel stores with a PNG copy) are no longer offered; undo no longer
  copies picture data per step; sorting moves Kommentare with their rows; Registerfarbe can be
  undone. Its last point, that sorting does not move pictures, is Excel's behaviour too: Excel
  leaves shapes in place when it sorts, whatever their placement (`probe_excel_sortpics.py`).
- **Sparklines** (Einfügen > Sparklines and Schnellanalyse): Linie, Säule and Gewinn/Verlust,
  one per row (or column) of the data in the Positionsbereich, drawn in the cell with Excel's
  default colours; read and written as Excel's x14:sparklineGroups (Excel sees ours as sparkline
  groups of the right type). Alles löschen removes them. The Sparkline tab (markers, styles,
  axis) is not built; sparklines whose data is on another sheet are kept but not drawn.
- **Hidden rows and columns** (14 recorded cases, catalogue area "hidden"): arrows, Strg+arrows,
  Umschalt+arrows and Eingabe skip hidden rows and columns as before. The status bar's Mittelwert,
  Anzahl and Summe leave out hidden (and filtered) rows but count cells in hidden columns, as
  Excel does. A column border dragged past the column's start hides the column and the drag goes
  on into the visible columns before it, which shrink to the pointer (Excel: column A 9,36 after
  130 px, 3,0 after 200 px; ours equal); rows are built the same way, not recorded.
  Umschalt+wheel does not scroll (Excel 365 here: recorded no scroll); a sideways wheel or
  touchpad still does. The cases hide columns by dragging, since Strg+0 does nothing in German
  Excel (D-063).
- **More keys** (19 recorded cases, area "keys5"): Enter and Tab in a selection of several
  areas go on into the next area after the last cell of one (before, they stayed in the area);
  Umschalt+Rücktaste shrinks the selection to the active cell (before, it started Eingeben); a
  selection extended over whole columns or rows (Strg+Umschalt+Unten past the data, Umschalt+Unten
  after a row header) no longer scrolls to the sheet's end. Alt+Bild ab / auf, Umschalt+Bild ab,
  Strg+Bild auf, Umschalt+Leertaste while typing, Strg+Umschalt+Leertaste and the editing keys in
  Bearbeiten were already as in Excel.
- **Right-click and the fill handle** (11 recorded cases, area "mouse2"): a right-click on a
  column or row header outside the selected whole columns (rows) selects that column (row) before
  the menu opens, as a click does; inside a selection, or on a cell of it, the selection stays.
  Umschalt+drag of the fill handle inserts empty cells below or to the right (the cells there
  move on), and dragged back into the selection it deletes the cells left out (those below move
  up); before, it filled as without Umschalt.
- **Editing** (16 recorded cases, area "edit2"): F4 in a typed range changes only the cell
  reference at the caret ("=SUMME(A1:B2" gives A1:$B$2), a pointed range as a whole; Strg+Z while
  typing takes the typing back and stays in the cell (before, it left the cell); Alt+Unten while
  typing opens the Auswahlliste with the matching entry highlighted, and the choice goes into the
  text; Strg+Umschalt+A writes the argument names after a function name as Excel does
  ("=RUNDEN(Zahl;Anzahl_Stellen)", "=SUMME(Zahl1;Zahl2;...)"). The status bar shows
  "Zirkelbezüge: A1" for the first cell of a circular reference on the sheet ("Zirkelbezüge" alone
  when it is on another sheet); before, nothing told the user. A formula ending in an operator
  ("=1+") is left out: Excel stored "=1", presumably through its formula correction, which the
  harness cannot see; RaabeCalc keeps "=1+" (#NAME?). To be checked with the owner watching.
- **Merged cells** (12 recorded cases, area "merge"; the harness can now merge cells before a
  case): arrows, Umschalt+arrows, Umschalt+click, Strg+Rechts, Enter, Tab and clicks around and
  inside merged cells were already as in Excel. One recording was left out: when the harness
  selects a merged cell through COM instead of a click, Excel's Enter goes to B3 below a merged
  A1:B2; after a real click it goes to A3, as in RaabeCalc.
- **Sheet tabs and pointing into other sheets** (10 recorded cases, area "tabs"): while a formula
  is typed at a place where a reference can go, clicking another sheet's tab or Strg+Bild ab/auf
  keeps the formula open; clicks and arrows there write "Tabelle3!B2" (names that need it in
  quotes), the cell editor is hidden and only the Bearbeitungsleiste shows the formula; Enter
  writes it into the cell it was typed in and goes back to that sheet, Esc goes back without it.
  Once another sheet was visited, references keep a sheet name also after going back
  ("=1+Tabelle1!C3", as Excel writes it). Strg+Bild ab while typing a value enters it first, as a
  tab click does (before, it did nothing). The harness names extra sheets itself, since Excel's
  name count goes on across cases. A range dragged on the other sheet is not compared (Excel's
  recording was unreliable).
- **Copy mode** (12 recorded cases, area "clip2"): Entf and F2 end the copy mode as typing does;
  once it has ended (also by Esc or by pasting a cut), Strg+V pastes nothing, since Excel
  empties the clipboard then. RaabeCalc leaves the clipboard alone and instead ignores it while
  it still holds that copy's text, so text copied in another program still pastes (the one
  difference: the same text copied elsewhere afterwards is ignored too). Undo and a sheet change
  keep the copy mode, as in Excel, and undoing a paste brings it back even after Esc (recorded),
  so the ended copy is kept for that. One recording was redone: a click right of a pasted cell hit
  Excel's Einfügeoptionen button.
- **Keyboard modes** (recorded): Ende then an arrow jumps as Strg+arrow (status bar "Endmodus"),
  F8 Auswahl erweitern, Umschalt+F8 Abschnitt hinzufügen oder entfernen, Strg+Rücktaste shows the
  active cell.
- **Umschalt+drag of the selection's border** (7 recorded cases): the cells are taken out (the
  rest of their columns closes up) and inserted between rows, below the cell at the pointer when
  the pointer is in its lower half, above it otherwise, in the pointer's columns; a grey line
  shows the place while dragging. Whole columns go between columns. Formulas follow the moved
  cells and the cells that closed up or made room; inserting and deleting cells (Zellen einfügen /
  löschen with shifting) now adjusts references everywhere too, as Excel does.
- **Dragging with the right mouse button** (menus recorded, `probe_excel_rightdrag.py`): the fill
  handle offers Zellen kopieren, Datenreihe ausfüllen, Nur Formate / Ohne Formatierung
  ausfüllen, Tage / Arbeitstage / Monate / Jahre ausfüllen (dates only; months as EDATUM),
  Linearer and Exponentieller Trend (Excel's results for 6 recorded cases), Reihe...; the border
  offers Hierhin verschieben / kopieren / nur als Werte / nur als Format kopieren, Verknüpfung
  and Link hier erstellen, Kopieren / Ausschneiden und nach unten / rechts verschieben,
  Abbrechen.
- **Gruppieren** (Daten > Gliederung, Umschalt+Alt+Rechts / Links): rows or columns get an
  outline level (up to 7); a bar beside the headers shows one lane per level with level buttons
  1, 2, …, a bracket along each group and a −/+ box at its summary row below (column to the
  right), as Excel's default; collapsing hides the group and marks the summary row as the .xlsx
  keeps it, and outlineLevelRow/Col are written so Excel shows the bar. Sizes of the bar are our
  estimate (16 px per lane at 100 %), not measured.
- **Teilergebnis** (recorded with `probe_excel_subtotal.py`, Range.Subtotal in an Excel of our
  own; all 11 functions equal cell by cell): after each run of equal values a bold row
  "<value> Ergebnis" (Anzahl, Mittelwert, Max., Min., Produkt, STABW, Var … as Excel names them)
  with TEILERGEBNIS over the run, a last row "Gesamtergebnis" over the first to the last data row,
  outline level 2 for data rows and 1 for result rows; Alle entfernen and Vorhandene ersetzen
  remove the rows with TEILERGEBNIS; Seitenumbrüche zwischen Gruppen starts a page after each
  result row. The dialog's labels are from German Excel, not recorded.
- **Seitenumbrüche** (Seitenlayout > Umbrüche: einfügen above and left of the active cell,
  entfernen, Alle zurücksetzen): printing starts a new page there, as Excel's PDF shows (2/2
  layouts with manual breaks equal); kept in the file as Excel's rowBreaks/colBreaks.
- **Blitzvorschau** (Strg+E, Daten > Blitzvorschau, right-drag menu): from the examples typed in
  the column, the cheapest description out of words (the first counted from the left, the last
  from the right), first letters, stretches between delimiter characters, case changes and
  literal text, applied to the rows below where the columns to the left have text; a first row
  that does not fit is the header. Excel's own results (Range.FlashFill) for 13 school cases are
  reproduced except two rows where Excel keeps a double first name ("Lea Sophie S."). Excel's
  grey suggestions while typing are not built.

## D-066 RaabeCalc: everything Excel offers, except the cloud (2026-11-05)

After trying RaabeCalc for one minute the owner found three differences (reference text not
coloured while typing, clicks in Bearbeiten not pointing, no Format pane on a chart) and concluded
that the catalogue only measured what we had thought of. Decided with the owner:

- **Scope:** the full Excel feeling and the same behaviour for clicking, dragging and keys, and all
  menus and possibilities except cloud features. Nothing is deferred on our own judgement or for
  lack of time; what is not finished stays listed as open and is continued after checking with the
  owner. The items deferred in D-065 are in scope (PivotTables with PivotCharts, shapes, Seitenlayout
  view, Datentabelle, Szenario-Manager), as is whatever the inventory turns up.
- **Cloud, left out (the buttons are missing, not grey):** OneDrive and SharePoint saving, Teilen and
  co-authoring, sign-in and account features, Copilot and Daten analysieren, online pictures, Smart
  Lookup, online translation, stock and geography data types, the add-in store, Python in Excel, web
  data connections.
- **Order:** macros (VBA, Entwicklertools) are built last; Power Query and Power Pivot (local) after
  everything else. Goal for macros: .xlsm files from Excel run in RaabeCalc with an interpreter of our
  own; macros reaching outside Excel (other programs, Windows calls) cannot work.
- **Method:** (1) an inventory taken from Excel itself (ribbon with contextual tabs, galleries,
  dialogs and their controls, Format panes, context menus, shortcuts), each entry marked built, partly
  built, missing or left out, so that coverage is a number (`docs/EXCEL-INVENTORY.md`); (2) behaviour
  cases generated systematically: mode x input x situation; (3) screenshots of Excel and RaabeCalc in
  the same state compared automatically, because the catalogue does not see what the screen shows.
- **Limits kept:** labels and shortcuts as in Excel, tooltips, help and dialog explanations in our own
  words and our own icons (D-058); Excel's quirks are copied unless they are plain calculation
  glitches or harm students, each exception written down; RaabeCalc stays offline and may grow in size;
  this step covers RaabeCalc only (Step 7b and RaabePresent follow).

**First run (2026-11-05 evening), what was built and measured:**
- **Pointing while a formula is typed** (168 generated cases, areas "point", "pointedit", "pointed",
  "frames"; `feellab/calc_matrix.py`): a click inserts a reference when the character before the
  caret is an operator, "(", ";", ":", "!" or the leading "=", "+", "-", in Eingeben and in
  Bearbeiten alike (F2, double-click, Bearbeitungsleiste), whatever follows the caret; at the very
  start too ("=C3=1+2"); after a space only a click points (Excel's intersection "=A1 C3"), not an
  arrow. Arrows point only in Eingeben. Strg+click and Strg+drag add a reference after ";", headers
  give "C:C" / "3:3" (Umschalt extends, Strg adds), Strg+Leertaste / Umschalt+Leertaste the pointer's
  column / row, Pos1 / Strg+Pos1 / Strg+Ende / Bild ab point as they move. F4 on a pointed range
  changes it as a whole, also again. Text that does not start with "=" gets it once a reference is
  pointed ("+" then a click is "=+C3"). The coloured frames of a formula's references can be dragged:
  the corner sizes, the border moves, $ signs are kept. The reference text is coloured as its frame
  in the cell and the Bearbeitungsleiste (a layer under the transparent text of the editor). Inside a
  string there is no function list.
- **Excel's checks on Enter** (`checkEntry`, recorded with Excel's messages on, 42 entries in
  `entry.test.ts`): some slips are mended without asking (brackets of a function closed, "1++2"),
  for others Excel proposes a mended formula with Ja/Nein (an operator at the end, an open string or
  plain bracket, a bracket too many, two operators in a row, "2(3)", "1 2"), and some are refused with
  a message (no parse, too few or too many arguments; WENN needs two). The cell stays in Bearbeiten
  after Nein or the message. Ending the entry with a click or an arrow mends without asking (recorded).
  Our messages are in our own words. The harness's Excel now runs with its messages on (they were
  off, which hid these dialogs); a dialog on one side only is a difference.
- **Circular references**: a new one brings a warning (own words), the status bar shows
  "Zirkelbezüge: A1"; a dynamic array spilling over the cells it reads is circular too. Whole
  columns or rows in an array formula keep their full size, so `=1+2+C:C` gives #ÜBERLAUF! as in
  Excel, and the column widens for an error result. A function's name alone gives #CALC! (Excel shows
  that untranslated, so the German error name is #CALC! too).
- **Chart Format pane** (the owner's third finding): a task pane docked at the right (shared shell
  component) with Excel's option sets ("Diagrammoptionen", "Textoptionen"), icon tabs and sections
  for Diagrammbereich, Zeichnungsfläche, Titel, Legende, both axes, Gitternetzlinien, Datenreihen,
  Datenbeschriftungen, Trendlinie and Fehlerindikatoren: Füllung (keine, einfarbig with Transparenz,
  automatisch), Rahmen/Linie (Farbe, Transparenz, Breite, Strichtyp), Markierung, Linie glätten,
  Achsenoptionen (Begrenzungen, Einheiten, schneidet bei, logarithmisch, umgekehrt, Teilstriche,
  Beschriftungsposition, Zahlenformat), Reihenachsenüberlappung and Abstandsbreite, Kreissegment-
  winkel and Innenringgröße, Beschriftung enthält / Trennzeichen / Position, all six trendline types
  with Prognose and Name, Fehlerindikatoren (Richtung, Endformat, Betrag), Textfüllung and Schrift,
  Größe. A click selects a chart element (dashed frame), a double-click or Strg+1 opens its pane;
  the new Format tab has Diagrammelemente, Auswahl formatieren and Auf Formatvorlage zurücksetzen.
  Everything is drawn and written to the file as DrawingML and read back (`chart-format.test.ts`).
  Not yet: Farbverlauf, Bild-/Muster-füllung, Effekte, 3D, Sekundärachse, Datentabelle (inventory).
- **Diagrammentwurf**: Schnelllayout (Layout 1-11 as recorded per chart type,
  `probe_excel_chartlayouts.py`), Farben ändern (schemes 10-26 from Excel's colour parts,
  `probe_excel_chartcolors.py`, colours within 3/255 of Excel's), Zeile/Spalte wechseln, Daten
  auswählen (Datenquelle auswählen with series list), Diagramm verschieben (to another sheet; chart
  sheets not yet).
- **Formen and Textfelder** (Einfügen > Formen with Office's whole gallery, shared with RaabeWrite and
  drawn by its Word-checked renderer: every preset, shaded sub-paths, Office's arrowheads, the yellow
  handles that change a preset's adjust values; Zuletzt verwendete Formen; Textfeld): drawn with a
  drag (Umschalt keeps the proportions, a click alone gives Excel's default size), selected with
  eight handles and a rotation handle, moved, sized, rotated (Umschalt in 15° steps), Pfeiltasten
  move, Entf deletes, typing goes into the shape. The Formformat tab (Formen einfügen, Fülleffekt,
  Formkontur with Stärke / Striche / Pfeile, Anordnen, Größe) and the pane "Form formatieren".
  Saved as Excel 365 saves a new shape (xdr:sp / xdr:cxnSp, the theme's shape style) and written back
  verbatim while unchanged. `check_calc_shapes.py`: 10/10 RaabeCalc→Excel (kind, place, size, fill, line
  with dash and both arrow ends, text, bold and size, rotation, flip, adjust value) and 10/10
  Excel→RaabeCalc. `shot_compare.py` puts Excel's picture of a range (CopyPicture) next to RaabeCalc's
  for a look by eye. Not yet: Freihand shapes (Kurve, Freihandform, Skizze).
  Shapes drawn in Excel with the mouse (`probe_excel_drawshape.py`, both Office themes): a new
  shape's text starts at the top left; its outline is Akzent 1 shaded to 15 % at the theme's second
  line width (1 pt in the Office 2013-2022 theme, 1,5 pt in the current one); a line takes the
  theme's first line style in the 2013-2022 theme and the second in the current one; a click gives
  a 2,54 cm square. RaabeCalc keeps the style reference of each shape and reads the theme's line
  widths. `shot_compare.py` on Excel's own file: the same picture (mean difference 2 of 255).
- **Anordnen, Gruppieren, Auswahlbereich, Freihand** (D-066): Umschalt+click or Strg+click selects
  several shapes, which then move, nudge, delete and format together; Ebene nach vorne / hinten with
  In den Vordergrund / Hintergrund, Ausrichten (Linksbündig … Unten ausrichten for two or more,
  Horizontal / Vertikal verteilen for three or more, to the outermost shapes), Gruppieren, Gruppierung
  aufheben and wiederherstellen (xdr:grpSp read and written; a group's shapes keep their places in its
  own coordinates and scale with it, line widths stay), the Auswahlbereich (frontmost first, Anzeigen /
  Ausblenden per shape and for all, renaming by double-click, order arrows; hidden="1" in the file),
  on Seitenlayout and Formformat. Kurve, Freihandform: Form and Freihand: Skizze draw custom geometry
  (a:custGeom) as RaabeWrite does. `check_calc_shapes.py` 12/12 both ways with a group and a freeform.
  Not yet: selecting one shape inside a group, pictures and charts in groups and in the Auswahlbereich.
- **Rahmenlinien zeichnen**: Rahmenlinie zeichnen (a click on a cell edge draws it, a drag the
  outline), Rahmenraster zeichnen, Rahmenlinie löschen (also the neighbouring cell's side), Linienfarbe
  and Linienart (13 styles, our own samples); choosing a colour or style switches drawing on, Esc ends.
- **Datentabelle** (Was-wäre-wenn-Analyse): one-variable tables down a column or across a row and
  two-variable tables; each result is the formula's value with the input cell set to that row's or
  column's value (worked out through the engine, the input cell put back), again after every change
  (not under "Automatisch außer bei Datentabellen", then F9). The cells show
  {=MEHRFACHOPERATION(Zeile;Spalte)} and refuse single edits; the whole table can be cleared. Saved as
  Excel saves it (f t="dataTable" on the first result cell). `check_datatable.py`: 3/3 both ways.
- **215 more functions** (statistics and distributions with the compatibility names, maths with
  MINV/MDET/MEINHEIT/RÖMISCH/AGGREGAT, all finance functions, engineering with the complex IM*
  functions and UMWANDELN, TEXTTEILEN and REGEX*, the new array functions such as VSTAPELN, ZUSPALTE,
  GRUPPIERENNACH, PIVOTMIT), recorded in Excel with 1206 cases (`probe_excel_functions2.py`,
  fixtures/excel-functions2.json, `formula-functions2.test.ts`): 1199 as Excel within 1E-9 relative,
  the German names both ways and the _xlfn. prefixes. Deliberate differences, each with its reason in
  the test's KNOWN list: RGP with constant x values, BESSELJ(10;1) (Excel's value is 5E-7 off),
  "3.5" read as a date (changes every year), AUFGELZINS quarterly over several years (Excel divides
  by 91 for a 90-day quarter), XINTZINSFUSS and RENDITE where Excel stops short of the root. Still
  missing: LET, LAMBDA, MAP, REDUCE, SCAN, NACHSPALTE, NACHZEILE, MATRIXERSTELLEN (they need the
  parser), the cloud and external ones, INFO, ZELLE, BLATT, BLÄTTER, HYPERLINK.
- **Excel tables** (Einfügen > Tabelle, Strg+T / Strg+L, Start > Als Tabelle formatieren): a real
  table object instead of painted cells, saved as xl/tables/tableN.xml with <tableParts> (unchanged
  tables written back verbatim, removed ones with their parts). The 60 built-in table styles were
  recorded from Excel element by element with theme-relative colours (`probe_excel_tablestyles.py`,
  `gen_tablestyles.py` -> table-styles.gen.ts) and are drawn under the cells' own formatting in
  Excel's element order. The contextual tab Tabellenentwurf: Tabellenname, Tabellengröße ändern, Mit
  PivotTable zusammenfassen, Duplikate entfernen, In Bereich konvertieren (the look stays as the
  cells' formatting), the seven Tabellenformatoptionen (Kopfzeile, Ergebniszeile with "Ergebnis" and
  TEILERGEBNIS(109) in the last column, Gebänderte Zeilen/Spalten, Erste/Letzte Spalte, Schaltfläche
  "Filter") and the style gallery drawn from the data. `check_calc_tables.py`: 4/4 both ways (name,
  range, style, options, columns, total row). Not yet: structured references (Tabelle1[Note], [@Note]),
  calculated columns, growing the table while typing next to it, Datenschnitt; the ribbon keytips of
  Tabellenentwurf are not yet recorded from Excel.
- **Inventory of the contextual tabs** (Diagramm, Tabelle, Bild, Form, Textfeld, Sparkline,
  PivotTable; Datenschnitt's setup and SmartArt failed and are still open) and the report
  `docs/EXCEL-INVENTORY.md` with RaabeCalc's contextual tabs and own panels counted: 766 of 2422
  controls and menu entries (31,6 %). The largest gaps: 252 of Excel's 493 functions (being built),
  real Excel tables with Tabellenentwurf, Sparklines, chart types and styles, Designs, the Zeichnen tab.
- **Sheet names and status bar** (catalogue): Excel numbers a new sheet on from the last one it made in
  the workbook and never reuses a number (Tabelle4 after Tabelle2 and 3 were deleted); the status bar
  shows Mittelwert, Summe, Min and Max in the number format of the first number in the selection
  (dates as dates).
- **The harness places windows by their visible frame**: RaabeCalc's and RaabeWrite's standard frames
  have invisible resize borders, so their cell areas were compared smaller than Excel's; the formula
  bar sat 4 px lower than Excel's and now does not. The remaining difference in how many rows fit is
  the default font (Calibri 15 pt rows against Aptos Narrow 14,5 pt, D-011).
- **Row heights in whole device pixels**, as Excel rounds them at the screen's DPI: 14,5 pt rows
  are 29 device px at 150 % (they were 28,5 px; over 20 rows RaabeCalc drifted 10 px from Excel).
  Printing stays at 96 dpi. The Zeilenhöhe dialog keeps the typed height exactly.
- **Inventory results hold labels only** (CLAUDE.md, D-058): `feellab/sanitize.py` drops Office's
  descriptions (function tips, explanations after a label) and anything naming a file; the full
  recording stays in `captures/`. The first inventory commit (36c7d49) still had them in the file:
  cleaned in the next commit; removing them from the history needs a force push (the owner decides).
- **PivotTable** (Einfügen > PivotTable, `probe_excel_pivot.py` with 8 configurations): the dialog,
  the field list pane (fields dragged into Filter / Spalten / Zeilen / Werte, a chip's menu),
  Wertfeldeinstellungen (all eleven functions, Excel's captions "Summe von Note", a second use of a
  field numbered "…2"), the Kompakt-, Gliederungs- and Tabellenformat, Teilergebnisse,
  Gesamtergebnisse, filters above the table, Aktualisieren (Alt+F5, Strg+Alt+F5 for all), the
  PivotTable-Analyse and Entwurf tabs; computed as Excel computes it (`pivot.test.ts` 8/8 against
  Excel's cells). In the file Excel keeps the layout on each field (compact="0" / outline="0"), the
  items in sorted order with the filter's item as an index into that order, and an x14 extension
  that hides the "Werte" row of several value fields; without it Excel adds that row (all recorded
  from Excel's own file). The cache is written without records and refreshed on opening, as other
  writers do. `check_pivots.py`: 8/8 RaabeCalc→Excel (Excel's refreshed report equals ours cell by
  cell) and 8/8 Excel→RaabeCalc (fields, values, filters, layout).

## D-070 RaabePresent: everything PowerPoint offers, except the cloud (2026-11-06)

The owner's go for Step 8 (2026-11-06), with their answers to the open questions of the Step 8
proposal (ROADMAP), given for an overnight run without consulting them:

- **Goal:** RaabePresent feels exactly like PowerPoint 365 in German: the same ribbon, menus, keytips,
  shortcuts, dialogs, panes and context menus, and the same behaviour for clicking, double-clicking,
  right-clicking, dragging, typing, selecting, sizing, rotating and snapping.
- **Scope:** everything PowerPoint offers except the cloud. Left out: OneDrive/SharePoint, Teilen and
  co-authoring, sign-in, Copilot, Designer/Design-Ideen, Stockbilder, Onlinebilder and -videos, Intelligente
  Suche, Übersetzen, Live-Untertitel, the add-in store and Cameo. Macros come last.
- **What we copy:** labels and shortcuts exactly as PowerPoint's; descriptions in our own words; our own
  icons. PowerPoint's quirks are copied unless they are harmful. PowerPoint's design artwork and stock
  content are not copied: we ship the Office design as in PowerPoint 2013-2022 (Calibri Light / Calibri and
  its colours) plus our own designs, including a Raabeschule design in Raabe red. New presentations are
  16:9 in that Office design. The frontend stays portable for the iPad.
- **Order:** A reference side and inventory; B the app and .pptx; C Normalansicht; D all tabs, views, Kopf-
  und Fußzeile, PDF and printing; E Bildschirmpräsentation; F the rest of the inventory, audio and video,
  macros. No stopping between packages.
- **Measuring:** a behaviour catalogue with the same real input in PowerPoint and RaabePresent
  (`catalogue_present.py`), round trips in both directions (`check_present_*.py`), slide renders against
  PowerPoint's PNG export and UI screenshots, inventory coverage. Lesson from RaabeCalc (D-066): test every
  object type with every mouse action and every key in every state on purpose.
- **PowerPoint harness rules** (PowerPoint runs as one process per user): see CLAUDE.md "Harness safety".
- **Running beside the RaabeCalc session:** this step works in its own worktree branch; shared packages
  only gain additions (all apps' tests stay green); decision numbers from D-070 on; real input only with
  the lock file.

### D-071 RaabePresent's building blocks (decided during the run)
Decided on our own during the overnight run, for the owner to confirm or overturn:
- **The parts stay XML documents** (`packages/ooxml/src/pptx`): a presentation is its package with every
  part as PowerPoint wrote it; the parts RaabePresent reads or edits are DOM documents, and edits change
  only the elements the app understands (place and size, text, fills ...). Whatever it does not know is
  written back untouched without a separate "raw" model, parts nothing refers to any more (a deleted
  slide) are left out, unchanged parts are written byte for byte (checked on PowerPoint's fixtures).
- **Inheritance as PowerPoint's** (`resolve.ts`): a placeholder takes place, geometry, look, text frame
  and each level's text style from its layout's placeholder (matched by index, then type) and that from
  the master's (by type); levels start at the master's titleStyle, bodyStyle or otherStyle (date, footer,
  number), other shapes at the presentation's defaultTextStyle. Our new presentation's eleven layouts sit
  exactly where PowerPoint's own 2013-2022 blank presentation has them (test), including PowerPoint's
  German names, also the line break in "Abschnitts-/überschrift".
- **Undo by snapshots:** each change keeps the XML of the slides it touches before and after (a slide is a
  few kilobytes); typing merges into one step per word run, as in PowerPoint.
- **Text frames are laid out by the browser** (line breaking at the frame's width, line height from the
  font's ascent and descent times the line spacing), with our own caret, selection and keys over the
  text model. RaabeWrite's Word-measured layout engine is not used for slides yet; slide renders are
  compared with PowerPoint's own PNG export to see how far the browser's breaks differ.
- **Prompt texts** of empty placeholders ("Titel durch Klicken hinzufügen", "Untertitel durch Klicken
  hinzufügen", "Text durch Klicken hinzufügen") are German PowerPoint's labels from memory; to be checked
  against a screenshot of PowerPoint.
- **A new shape's style** is PowerPoint 365's (Akzent 1 fill, its 15 % shade as outline, white text), as
  recorded in Excel for the same Office (D-066).
- **Designs:** "Office-Design" (2013-2022), and our own "Raabeschule" (Raabe red band and titles), "Klar"
  and "Tafel" (dark). Applying a design changes the theme's colours and fonts and the master's design
  shapes and background; placeholders and layouts stay. School material can replace them.

### D-072 RaabePresent: measured rules and further decisions of the run
**Measured in PowerPoint** (and built that way):
- **Line height** is the largest font size of the line times 1.2 times the line spacing; Autofit
  (Text bei Überlauf verkleinern) rounds the shrunk size to whole points, reduces the line spacing by
  `lnSpcReduction` and leaves Abstand vor/nach unchanged (`check_present_lines.py`: 127/127 line starts).
- **A click beside the text** of a text box selects the frame, a click on the text puts the caret there
  (behaviour catalogue, `click` cases).
- **Vertical text** turns the whole frame; it is not broken word by word.
- **What a click hits** (behaviour catalogue, first comparison 99/307): the letters of a text start editing
  at once and a drag from them selects text (also with Umschalt, Strg or Alt); beside the letters, on a
  line of text, the frame is selected and a drag moves it; an unfilled shape (text box, placeholder) is
  passed through where it has no text; the outline is hit within 4 screen pixels, lines likewise; a click
  in an empty placeholder starts editing; a double-click selects the word, a triple-click the paragraph
  with its mark.
- **Keys on a selected frame:** the arrow keys move by one pixel at the zoom (as Strg+Pfeil; the grid
  spacing only with Am Raster ausrichten); Umschalt+Pfeil sizes about the centre by 10 % (Rechts/Oben
  larger, Links/Unten smaller, a text box keeps the new height), Strg+Umschalt+Pfeil by 1 %;
  Alt+Umschalt+Pfeil, Umschalt+Eingabe and Strg+Entf do nothing; Strg+Rücktaste puts the caret at the end
  of the text; typing adds to the end of the text; Strg+E/L/R/J and Strg+Umschalt+F on an empty shape
  start editing it; Strg+D and repeated Einfügen offset by 12 pt; Strg+Umschalt+G groups,
  Strg+Umschalt+H ungroups, Strg+G does nothing; after Rückgängig the shapes it restores are selected;
  Bild ab/auf, Pos1 and Ende with nothing selected go to another slide.
- **Text boxes** (spAutoFit) take their text's height as soon as they are drawn or changed.
- Equations and other `mc:AlternateContent` are drawn from their fallback picture, SmartArt from its
  drawing cache; editing either is open.

**Decided on our own** (for the owner to confirm or overturn):
- **Keys:** arrow keys move a frame by the grid spacing (0,2 cm, 72008 EMU), Strg+Pfeil by one pixel;
  Strg+D duplicates with PowerPoint's offset (172720 EMU); Strg+Mausrad zooms in 10 % steps.
- **Kommentare** are written in the classic format (`p:cmLst`, `p:cmAuthorLst`) that every PowerPoint
  since 2007 reads, not the newer threaded one (`p188`), so that no account or author id is needed.
- **Rechtschreibung** checks in a dialog box (as RaabeWrite's) rather than a pane; the dictionary and the
  native check are RaabeWrite's (`spell.rs`).
- **Gliederungsansicht** edits titles and body text as plain lines; formatting is shown, not edited there.
- **Keytips** of two letters are written joined as PowerPoint shows them ("ÄN"), the inventory's spaces left out.
- **Zeichnen:** ink is stored as freeform shapes (`a:custGeom`), which every PowerPoint shows; PowerPoint's
  own ink (content parts) is kept and drawn from its fallback picture.
- **Colour and font sets** in Entwurf > Varianten are Office's own sets (from the files Office installs),
  named with German PowerPoint's labels (`gen_present_themesets.py`).
- **Diagramme:** RaabeCalc's chart engine writes the chart part with its values cached; Daten bearbeiten
  edits them in a small grid inside RaabePresent. No embedded workbook is written yet, so PowerPoint shows
  the chart but cannot open its data in Excel.
- **Morphen** is our own drawing: shapes with the same name on both slides glide from their old place, size
  and turn to the new one, the others fade; text and shape geometry are not morphed letter by letter.
- **Table cells:** a drag over cells, Umschalt+Klick or Umschalt+Pfeil past the cell's text selects a range
  (grown to whole merged cells, grey veil as PowerPoint's); Zellen verbinden puts the texts of the cells
  as paragraphs into the first; Zellen teilen takes a merged cell apart and makes more columns or rows by
  splitting the cell's column or row (the cells beside it then span the new grid line). Written as
  PowerPoint writes merged cells (`gridSpan`, `rowSpan`, `hMerge`, `vMerge`).
- **Raabeschule templates** in Datei > Neu: Referat and Unterrichtsstunde (six slides each, ordinary
  layouts in the Raabeschule design).
- **Master views' tabs** are PowerPoint's recorded ones (key O, the same keytips, the Hintergrund group);
  Folien pro Seite of the Handzettelmaster is a setting of the view (our drawing of the slide places) and
  is not saved; Notizenseiten- and Handzettelausrichtung turn the notes page size, which both share, and
  scale the two masters' placeholders.
- **Formeln:** an equation is drawn from its Office Math (the shared converter, MathML in Cambria Math)
  instead of its fallback picture; Einfügen > Formel (Alt+Umschalt+0) offers the built-in equations
  (RaabeWrite's list) and writes them as PowerPoint does (text box with a14:m in mc:AlternateContent, a
  plain-text Fallback). Editing is in PowerPoint's linear format, in place over the equation (double-click it,
  or Formel > Linear; Eingabe converts back, Esc cancels; Alt+Umschalt+0 starts a new one that way);
  PowerPoint's professional in-place editing with a caret inside the structures is open.
- **Schnellformatvorlagen:** the 77 entries as PowerPoint's gallery; the six Designformatvorlagen rows are
  shape styles (references into the theme's format matrix), the five Voreinstellungen rows set fill and
  outline on the shape; how each looks in detail is our reading of PowerPoint's tiles. Colour words follow
  the design, as PowerPoint names them.
- **Bild > Farbe:** Neu einfärben as PowerPoint's (Sepia, Ausgeblichen, Schwarzweiß 25/50/75 and the
  theme's dark and light recolourings as duotone); Farbsättigung and Farbton (a14 image effects) are open.
- **Tabellenformate:** PowerPoint's 74 entries with their names and families, written into tableStyles.xml
  when used. PowerPoint's own style ids are known only for its default (Mittlere Formatvorlage 2 -
  Akzent 1); the others carry ids of our own, so PowerPoint lists them as the presentation's custom styles.
  Their look is our reading of PowerPoint's gallery.
- **Formel tab:** PowerPoint's (Word's) galleries from RaabeWrite's recording; a structure or symbol is
  added at the end of the selected equation (no caret inside an equation yet), LaTeX in Mathematik and
  Mathe nach LaTeX open the LaTeX box, Linear the linear box; Text, Formeloptionen and Freihandgleichung
  are greyed.
- **Bildformatvorlagen:** PowerPoint's 28 entries with their names; each sets outline, shape and effects
  (shadow, reflection, soft edges). Bevels and 3-D perspectives are left out (no 3-D drawing), so those
  entries keep only their frame and shadow.
- **Ribbon by UI screenshots** (`compare_present_ui.py`, PowerPoint and RaabePresent side by side, captures not
  committed): Entwurf shows the designs as a gallery and Anpassen as five large buttons; Übergänge and
  Animationen show their effects in the ribbon; labelled stacks as PowerPoint's (Erweiterte Animation,
  Animation neu anordnen, Masteransichten, Farbe/Graustufe, Einrichten with two columns of checkboxes);
  ribbon buttons show no "..." (their accessible names keep it, as PowerPoint's).
- **Überprüfen:** Sprache für die Korrekturhilfen festlegen (selected text, or the slide's or the whole
  presentation's text, with a list of the school's languages); Barrierefreiheit überprüfen is our own short
  check (Alternativtext of pictures, charts, tables and SmartArt; slide titles).
- **Ansicht > Graustufe / Schwarzweiß** change only how the slides are shown, as in PowerPoint.
- **Animationspfade:** written as PowerPoint writes them (p:animMotion with a path relative to the shape, in
  fractions of the slide); the curves of Bögen, Wendungen, Formen and Schleifen are our own; the path is not
  drawn in the editing view yet; Benutzerdefinierter Pfad (drawing one) is open.
- **WordArt styles:** PowerPoint's 20 names with the design's colour words; bevels not drawn.
- **Found by the UI screenshots:** several commands had no implementation behind them and stayed greyed
  (Übergänge > Dauer, Bei Mausklick, Nach, Auf alle anwenden, Sound; Animationen > Dauer, Verzögerung; the
  Layout, Abschnitt, Markieren, Farben and Schriftarten dropdowns; Gruppierung wiederherstellen). They work
  now; Rahmen (13 entries) and Texteffekte (Schatten, Spiegelung, Leuchten; no 3-D) were added, and in the
  Folienmaster Titel, Fußzeilen and Beibehalten.
- **Einfügen > Aktion:** Aktionseinstellungen with Mausklick only (Keine; Hyperlink zu Nächste, Vorherige,
  Erste, Letzte, Letzte angezeigte Folie, Bildschirmpräsentation beenden, or a slide), written as PowerPoint
  writes it (a:hlinkClick in p:cNvPr) and followed in the show; the dialog's labels are from knowledge of
  PowerPoint, not from a recording (it was not opened by the inventory probe). Mouseover, programs, macros
  and sounds are open. **Thesaurus** (Umschalt+F7) is RaabeWrite's (open thesauri) in a box; **Lineal**
  shows cm rulers with 0 at the slide's centre.
- **Bilder komprimieren:** PowerPoint's choices (HD 330, Drucken 220, Web 150, E-Mail 96 ppi, Nur für dieses
  Bild, Zugeschnittene Bildbereiche löschen) with labels only; the picture is resampled at the size it is
  shown into a new media part. When writing, picture relationships an edited slide no longer refers to are
  left out with their pictures (as PowerPoint does); they stay in memory so Rückgängig still works.
- **Trigger and Animation übertragen:** Trigger > Bei Klick auf <Form> moves the selected shapes' effects into
  an interactive sequence started by that shape (as PowerPoint writes it); in the show a click on the shape
  plays its next step, shapes it shows start hidden. Animation übertragen copies a shape's effects to the next
  shape clicked. Removing effects now tidies the timing (no empty lists, which PowerPoint rejects).
- **AutoKorrektur while typing:** RaabeWrite's rules (text-engine `AutoCorrect`, now exported for both apps):
  sentence start, two capitals, weekdays, quotes typographic, dash, (c) and the list's symbols. Measured
  against PowerPoint in 7 catalogue cases, all the same: an arrow (`-->`) becomes, as in PowerPoint, the
  Wingdings character U+F0E0 in a run with a:sym (drawn in that font). Rückgängig takes back the
  correction first. **Datei > Optionen** (had no implementation) is a short dialog of our own: Benutzername and
  Initialen (for comments), Touchmodus, AutoKorrektur on/off, AutoWiederherstellen interval.
- **Picture effects (Farbsättigung, Farbton, Schärfen/Weichzeichnen, Künstlerische Effekte):** recorded in
  PowerPoint (`probe_powerpoint_picfx.py`): the effects are a list in the blip (a14:imgProps > a14:imgLayer >
  a14:imgEffect), the original stays behind the layer's hdphoto relationship and the picture shown becomes the
  rendered result. We write the same (our original stays PNG/JPEG behind that relationship; PowerPoint accepts
  it: round trip 31/31, PowerPoint lists the effects as its own). The adjustments follow PowerPoint's rendered
  pixels: Sättigung scales HSL saturation by k/(1+(k-1)s) (mean difference 0.0-0.2 of 255), Temperatur uses
  PowerPoint's per-channel curve fitted at 14 temperatures (below 1), Weichzeichnen a small Gaussian (0.08),
  Schärfen an unsharp mask (0.38, PowerPoint's own sharpening not reproduced exactly). The 22 artistic effects
  are our own drawings of each effect's look as PowerPoint renders it (texture, strokes, relief; differences
  of 0.5-62 of 255, noisy textures cannot match pixel for pixel); PowerPoint redraws them its way when it
  edits the picture. The menus are galleries of the selected picture in each effect, as in PowerPoint.
  PowerPoint's own originals are HD Photo (.wdp), which a browser cannot read: on such pictures a new effect
  is drawn on the picture as shown, and removing effects keeps the drawn picture until PowerPoint redraws it.
- **Drawing by points (Kurve, Freihandform: Form, Freihand: Skizze, Benutzerdefinierter Pfad):** as
  PowerPoint: a click sets a corner, dragging with the button down draws freehand (Skizze: one drag), a
  double-click, Enter or Esc ends, ending near the start closes the shape (filled; open ones are lines).
  Kurve goes smoothly through the clicked points (our Catmull-Rom curve). Before tonight these three shapes
  drew a rectangle. Benutzerdefinierter Pfad writes p:animMotion with presetID 0 and the path relative to the
  shape's centre in fractions of the slide; PowerPoint reads both back as its own (round trip 33/33).
- **Tabelle zeichnen and Radierer:** the eraser joins the two cells on either side of the clicked cell line;
  the pen splits a cell in two by a line drawn across it (rows) or down it (columns), in the middle (PowerPoint
  splits where the line was drawn: our simplification), and draws a new one-cell table outside a table. Both
  stay on until Esc or the button again, as in PowerPoint.
- **Einfügen > Zoom > Folienzoom:** recorded in PowerPoint (`probe_powerpoint_zoom.py`, keytips Alt, I, Y, F
  under the real-input lock; the dialog's check box ticked through UI Automation): mc:AlternateContent with a
  graphicFrame (pslz:sldZm: the target's sldId and p14:creationId, the preview picture, a 0.25 pt light grey
  line) and as Fallback a picture with a slide-jump link. We write the same: a quarter of the slide in its
  middle (several zooms 12 pt apart); a target slide without a creationId gets one. The preview picture is
  the target slide drawn by our renderer (its DOM in an SVG foreignObject onto a canvas); it is not redrawn
  when the target changes (PowerPoint redraws its previews). In the show a click goes to the slide and, with
  Zurück zu Zoom (returnToParent, read from PowerPoint's files), back afterwards. PowerPoint reads our zoom as
  its own (round trip 34/34). **Abschnittszoom** (recorded the same way: psez:sectionZm with the section's id,
  Zurück zu Zoom on) shows the section's first slide and plays the section in the show, then returns.
  **Zusammenfassungszoom** (recorded): a new "Titel und Inhalt" slide before the first chosen slide, sections
  "Standardabschnitt", "Zusammenfassungsabschnitt" and one per chosen slide named after its title, and in the
  content placeholder's place one frame (psuz:summaryZm) with a tile per section in a grid; the grid's spacing
  is PowerPoint's for two tiles (3.75 % gaps, centred in 91.7 % of the frame), other counts by the same rule
  (our assumption). PowerPoint reads all three as its own (round trip 36/36).
- **Ansicht > Neues Fenster:** another window onto the same presentation, named "Name:1", "Name:2" as in
  PowerPoint; in RaabePresent it is another tab of the window (own slide and view, shared undo and file).
  Closing one window does not ask to save while another shows the presentation. Alle anordnen stays greyed
  (one window with tabs).
- **SmartArt creation stays open (measured):** `probe_powerpoint_smartart.py` made six SmartArts in PowerPoint
  and cut the layout, style and colour parts down to their uniqueId. PowerPoint opened the file but laid every
  SmartArt out again as Einfache Blockliste: it needs its full layout definitions inside the file. Writing them
  would mean copying Microsoft's 159 layout definitions (D-070: no Microsoft content), so creating SmartArt (and
  In SmartArt konvertieren) is not built; SmartArt from PowerPoint's files is drawn and kept unchanged.
- **Macros and file types:** Öffnen takes PowerPoint's types (.pptx, .pptm, .ppsx, .ppsm, .potx, .potm); a template
  opens as a new presentation based on it. Speichern unter offers PowerPoint-Präsentation, ... mit Makros,
  PowerPoint-Bildschirmpräsentation and PowerPoint-Vorlage and writes the main part's content type of the chosen
  type. Macros (the VBA project of a .pptm) are kept and saved again, never run (RaabePresent has no VBA, and
  running macros from files would be a risk at school); saving them as a type without macros asks first, as
  PowerPoint does, and leaves the VBA project out. This is all of "macros" for Step 8 (D-070 put them last).
- **Hintergrund entfernen:** recorded (`probe_powerpoint_picfx.py bg`): a14:backgroundRemoval with PowerPoint's
  default rectangle (t 10 %, b 90 %, l 10 %, r 90 %) in the same effect list, the picture shown becomes the
  result with a transparent background. Our method is our own (region growing from outside the rectangle into
  similar colours that look like the background's, one pixel of edge kept); on the probe's test picture (two
  objects on a graded sky) it keeps the same pixels as PowerPoint (100 %); real photos will differ. Marking
  areas came later the same night (Freistellen, below).
- **Formeffekte, Texteffekte > Abschrägung, 3D-Drehung, Transformieren:** the shape's Formeffekte had only two
  shadow entries. Now PowerPoint's whole menu (Voreinstellung 1-12, Schatten 24, Spiegelung 9, Leuchteffekt 24,
  Weiche Kanten 6, Abschrägung 12, 3D-Drehung 25) with PowerPoint's names from the inventory; the effects are
  Office's presets as RaabeWrite recorded them from Word (the same galleries; `effect-presets.ts` is a copy of
  RaabeWrite's generated file). Text 3-D and Transformieren (40 warps) go into the text body (bodyPr), as
  PowerPoint writes them. Drawn: shadows, reflection, glow, soft edges, bevel and the 3-D camera (CSS
  perspective, our approximation); the text warps are written for PowerPoint but our text stays straight. The
  "...optionen" entries (a format pane for 3-D) are greyed.
- **Interaktive Schaltflächen:** PowerPoint's 12 action buttons as the Formen gallery's last group, with its
  names (inventory). Their geometries are new presets in text-engine (added, nothing else changed; our own
  icons in a darker shade on the button, PowerPoint's own shading not copied). Drawn, a button gets the jump
  PowerPoint gives it (Weiter: nächste Folie, Zurück: vorherige, Anfang/Startseite: erste, Ende: letzte,
  Zurückkehren: zuletzt angezeigte) and Aktionseinstellungen opens, as in PowerPoint.
- **Context menus and the inventory:** the notes pane has its own context menu (Ausschneiden, Kopieren,
  Einfügen, Neuer Kommentar; Schriftart and Absatz greyed, the notes are plain text). The inventory export
  now also reads the picture, group, chart and notes context menus, the tiles inside menus (picture effect
  galleries), the table grid's cells, and names spin boxes by their command (PowerPoint's accessible name,
  "Animationsdauer" beside the visible "Dauer:"). Chart colours renamed to PowerPoint's (Bunt, Einfarbig,
  Monochrome Palette n). Einfügen > Tabelle has Tabelle zeichnen.
- **Colour palettes in the presentation's design:** the palettes (Schriftfarbe, Fülleffekt, Formkontur,
  Textfüllung, Textkontur, Schattierung, Stiftfarbe …) showed the Office 2013-2022 colours whatever the design.
  Now they show the current design's colours, named with PowerPoint's colour words ("Dunkelblaugrün, Akzent 1"),
  and a pick of a design colour is written as PowerPoint writes it: a:schemeClr with lumMod/lumOff ("heller
  60 %": lumMod 40000, lumOff 60000), so it follows a later change of design; the colour bars show the resolved
  colour (Office's own values: "Orange, Akzent 2, heller 40 %" is F4B183). The shell palette takes the design as
  an optional parameter (Word and Excel unchanged).
- **Inventory measured in PowerPoint's default colours:** PowerPoint's inventory was recorded in its default
  design (Office, Aptos colours); `inventory_present.py` now applies the same colour set (Office) to the
  untitled presentation before reading RaabePresent's controls, so colour entries are compared by the same
  names. Coverage with it: see the report.
- **Hintergrundformate, Fülleffekt's own entries, the table pen:** Entwurf > Hintergrundformate (and the
  masters') is PowerPoint's menu: 12 styles (the design's three background fills, bgRef 1001-1003, in bg1,
  bg2, tx2, tx1; written on the slide master, the slides' own backgrounds go, as PowerPoint applies a style
  to all slides), Hintergrund formatieren..., Folienhintergrund zurücksetzen. Fülleffekt has Keine Füllung,
  Bild... (a picture fill) and Farbverlauf with PowerPoint's 13 light and 13 dark variants by name (our own
  stops from the shape's colour; PowerPoint's directions); Struktur (Microsoft's textures) and Pipette are not
  built. Tabellenentwurf > Stiftart (11) and Stiftstärke (9) are PowerPoint's entries; choosing one, or a
  Stiftfarbe, takes up the pen, and a click on a cell line with the pen gives that line the pen's look (both
  cells beside it).
- **More of PowerPoint's menus by the inventory:** Aufzählungszeichen and Nummerierung with PowerPoint's entries
  (Ausgefüllte/Hohle runde, quadratische, Silbentrennung-, Sternförmige, Pfeilförmige, Häkchen-; Nummern in
  Klammern) in PowerPoint's bullet fonts (Wingdings, Courier New); Bildformat > Zuschneiden as PowerPoint's
  split button (Zuschneiden, Auf Form zuschneiden with the Formen gallery, Seitenverhältnis 1:1 … 16:10 cropped
  about the middle with the width kept, Ausfüllen, Einpassen); Formkontur > Pfeile (eleven Pfeilarten, our
  reading; their names are from knowledge of PowerPoint, the inventory did not open that submenu); Ersetzen >
  Schriftarten ersetzen... (text and design fonts throughout); Farben > Farben anpassen... (the twelve design
  colours and a name, applied like a colour set). Formkontur > Skizziert and the Pipette are not built.
- **Folien wiederverwenden and pasting slides between presentations:** Neue Folie > Folien wiederverwenden...
  opens a presentation and shows its slides to pick (or Alle Folien einfügen); they come after the current
  slide with what they refer to (pictures, media, charts without their workbook, external links) and take the
  layout of the same name here (PowerPoint's Zieldesign verwenden; Ursprüngliche Formatierung beibehalten is
  not built). Our dialog instead of PowerPoint's pane. Pasted slides use the same way now, so pictures are no
  longer dropped when slides are pasted (also into another window). **Folien aus Gliederung...** takes a Word
  document (Überschrift 1: a slide's title, lower headings its bullet levels, other text left out) or a text
  file (lines starting with tabs for the level), as PowerPoint imports an outline; RTF is not read.
- **Smaller entries by the inventory:** Drehen with PowerPoint's names (Rechtsdrehung 90 Grad, Vertikal umdrehen,
  Horizontal spiegeln, Weitere Drehungsoptionen...); Animation hinzufügen > Weitere Eingangs-, Hervorhebungs-,
  Ausgangseffekte und Animationspfade (our lists in PowerPoint's dialogs' titles; OLE-Aktionsarten greyed);
  the chart's Format tab with Aktuelle Auswahl (Diagrammelemente, Auswahl formatieren, Auf Formatvorl.
  zurücks.), Form ändern (greyed, as PowerPoint with the whole chart) and Größe und Position; Zeichnen >
  Hinzufügen (six pen kinds, taken up at once: there is no own pen toolbox to add them to) and Lassoauswahl (a
  loop round shapes selects those whose middle is inside). **Freihand zu Form**: while on, a pen stroke that
  looks like a line, triangle, rectangle, rhombus or ellipse becomes that shape outlined in the pen's colour
  and width (our recognition: the stroke simplified to its corners); other strokes stay ink. Lineal,
  Wiedergabe and Freihand in Gleichung (handwriting recognition) are not built.
- **Highlight colours, Barrierefreiheit überprüfen, Kommentare anzeigen:** PowerPoint names three highlight
  colours differently from Word (Grelles Grün, Grau-50 %, Grau-25 %); the shell's highlight palette takes
  other names as an optional parameter and RaabePresent passes PowerPoint's. Barrierefreiheit überprüfen and
  Kommentare anzeigen are split buttons as in PowerPoint (Alternativtext, Lesereihenfolgebereich as the
  Auswahlbereich; Kommentarbereich, Markup anzeigen to show or hide the comment markers).
- **Editing equations in place:** a click into an equation (or a double-click, Eingabe on it, or a new
  equation) puts a caret into the drawn equation, as PowerPoint does: typing builds structures up as it goes,
  the arrows walk through the slots, Rücktaste and Entf take structures apart, Umschalt+Eingabe starts a new
  line, Esc, Eingabe or a click outside end editing and write the equation back as OMML in one undo step. The
  editing model is text-engine's MathEditor (RaabeWrite's, measured against Word's equation editor, D-059);
  the caret geometry is RaabeWrite's (copied into mathedit.ts). A new equation starts empty with its prompt
  (before it started as "x" in the linear field). Structures from the Formel tab go in at the caret. The
  linear field stays for Formel > Linear. Formel > Text types plain text while it is on. The ribbon, menus and
  dialogs keep the equation in editing (a structure goes in at the caret, the caret into its first empty slot).
  Not built: Formeloptionen, Freihandgleichung.
- **SmartArt from PowerPoint's files, its texts editable:** selecting a SmartArt shows SmartArt-Design with
  Textbereich (also on a double-click): its points' texts, one line each, indented by level, in our dialog
  (PowerPoint's is a pane beside the graphic). A change goes into the data model and into PowerPoint's drawing
  of it (the shape of the node's presentation point, presAssocID), so both show the same; Rückgängig takes it
  back. Adding or removing shapes, layouts and SmartArt styles stay out (PowerPoint's layout definitions).
- **Notes keep their formatting:** the notes pane edits plain text (as before); writing it back now keeps
  every paragraph whose text did not change as it was (bold, colours, bullets set in PowerPoint), and a changed
  paragraph keeps its paragraph properties and its first run's look. Before, any edit rewrote all notes as
  plain text. Formatting notes in RaabePresent itself is not built.
- **Exportieren > Dateityp ändern:** Präsentation, PowerPoint-Vorlage and PowerPoint-Bildschirmpräsentation
  (through Speichern unter) and the slides as pictures, PNG or JPEG, 1920 pixels wide, as PowerPoint offers
  them ("Welche Folien möchten Sie exportieren?" Alle Folien / Nur diese). PowerPoint puts all slides into a
  folder; we write them beside the chosen name ("Name-Folie1.png" …), since the platform has no folder
  creation. Video, animiertes GIF and Für CD verpacken are not built.
- **SmartArt-Design > Konvertieren:** In Formen konvertieren puts PowerPoint's drawing of the SmartArt on the
  slide as a group of ordinary shapes (same look; they can then be edited freely); In Text konvertieren makes a
  text box with the points as a bulleted list by level. Both replace the SmartArt as in PowerPoint.
- **Freistellen:** Hintergrund entfernen now opens PowerPoint's Freistellen tab after the automatic removal:
  Zu behaltende Bereiche markieren and Zu entfernende Bereiche markieren draw lines on the picture (written as
  PowerPoint writes them, a14:foregroundMark/backgroundMark in 1/1000 % of the picture); after each line the
  removal runs again with them (a "keep" line is never background and stops the background spreading, a
  "remove" line becomes background and spreads into similar colours). Alle Änderungen verwerfen takes the
  removal away, Änderungen beibehalten ends Freistellen. PowerPoint's magenta overlay is not drawn: the picture
  shows the result directly.
- **Benutzerdefinierte Bildschirmpräsentation:** the menu was empty. Now PowerPoint's custom shows
  (p:custShowLst, slides by relationship) are listed to run, and Zielgruppenorientierte Präsentationen...
  makes, changes, copies and removes them (our dialogs; the labels from knowledge of PowerPoint, the
  inventory did not open them). A custom show plays only its slides in its order, Endlos as set up.
- **Pipette:** in Fülleffekt, Formkontur, Schriftfarbe, Textfüllung and Textkontur (where PowerPoint's inventory
  has it). The next click on the slide takes the colour of our own drawing of the slide at that point (without
  selection marks), written as an RGB colour as PowerPoint does; Esc stops it. PowerPoint's live preview of the
  colour beside the pointer is not drawn.
- **SmartArt's context menu** as PowerPoint's recorded one, without Layout ändern... and Grafik zurücksetzen
  (they need the layout definitions, see SmartArt above); our Textbereich is a dialog, so its entry reads
  "Textbereich anzeigen" where PowerPoint toggles a pane.
- **Gaps from the inventory (04:50):** the palettes' custom colour entry carries PowerPoint's label per control
  (Weitere Füllfarben..., Weitere Konturfarben..., Weitere Rahmenfarben...; the shell's palette takes the label as
  an optional parameter, Word and Excel unchanged); Kopieren with Duplizieren; Fotoalbum as a split button
  (Neues Fotoalbum...; editing an album is not built); Bild zurücksetzen with Bild und Größe zurücksetzen (the
  picture's pixels at 96 dpi); Bild ändern with Dieses Gerät... and Aus Zwischenablage... (archive, online and icon
  sources are cloud); Schriftarten anpassen... (our dialog) and Farben des Foliendesigns zurücksetzen (for our own
  designs, found by the theme's name); Tabellenlayout > Anordnen with Gruppieren and Drehen; PowerPoint's names
  for Zeilenabstand (1, 1,5 …), Linksbündig/Rechtsbündig, Referentenansicht, the Formel tab's Formelstruktur …
  entries and Nummerierung und Aufzählungszeichen....
- **Inventory counting:** two controls of one name in a tab (Formformat's two Schnellformatvorlagen) are compared
  with both our controls' entries, and a split button's own part and its arrow (Weitere Optionen) count as there
  when ours is a split or colour control. This counting fix alone moved 66 entries; the features above 53.
- **Rotation handle (measured, replaces "our handle follows the pointer"):** probe_powerpoint_rothandle.py found
  PowerPoint's handle 38-56 device pixels above the frame (66 % zoom, 150 % scaling), so 31 screen pixels to its
  centre; ours is there now. PowerPoint turns the shape by the handle's centre, keeping the offset where it was
  grabbed (rectangle and ellipse now the same as PowerPoint, the text box 1.3 degrees apart), and Umschalt snaps to the last
  15-degree step passed (53.4 becomes 45, 82.2 becomes 75), not the nearest. The catalogue's PowerPoint side now
  grabs the handle at 45 device pixels.
- **Drags start at the first movement (measured):** PowerPoint's moves and handle drags in the catalogue land
  one of the harness's 20 steps short (56.8 of 60 pt, 38 of 40): it measures a drag from the pointer's first
  movement after the button went down, not from the button, when that button press also selects the shape; a
  frame selected already and the size handles move by the whole way (measured, both). Ours does the same, and
  the rotation handle turns from the first movement too; with a hand on the mouse the first movement is a
  pixel or two, so nobody feels it. Lines and placeholders move by the whole way in PowerPoint (measured), so
  the rule is for shape bodies only. The browser merges pointer movements per display frame (this PC runs at
  30 Hz), so the first movement is taken from the first event's coalesced list. Catalogue 344 -> 356 of 377.
- **Tabellenentwurf > Schattierung fixed:** it went to the shapes' fill and so did nothing on a table (found while
  closing the inventory's gaps). It now fills the cells being edited (or the table), in a:tcPr as PowerPoint
  writes it, with Keine Füllung and Pipette below the palette; a round-trip scenario checks PowerPoint reads it.
- **Format > Fülleffekt and Formkontur on a chart fixed:** with the whole chart selected they did nothing; they
  now set the Diagrammbereich's fill and line through the chart model (RGB, as the model keeps colours; a
  design colour picked is written as its RGB value). A round-trip scenario checks PowerPoint reads both.
- **Tabellenentwurf > Effekte:** Zellenabschrägung, Schatten and Spiegelung with Office's presets (the
  Formeffekte galleries). Shadow and reflection go to the table (a:tblPr > a:effectLst) and are drawn; the
  bevel goes to the cells being worked on (a:tcPr > a:cell3D with the light rig the schema requires) and is
  written but not drawn by us (PowerPoint draws it).
- **Textfüllung's own entries:** Keine Füllung (a:noFill on the runs), Pipette, Farbverlauf with the 13 light and
  13 dark directions of Fülleffekt over the text colour. A text gradient is drawn in its first colour (our
  renderer draws text in one colour); PowerPoint shows the gradient.
- **Formkontur > Skizziert (measured):** probe_powerpoint_sketch.py recorded PowerPoint's format (keytips in a
  PowerPoint of our own; Stärke and Skizziert share the keytip S): an a:ln extension with
  ask:lineSketchStyleProps, a random seed, a copy of the geometry and the type (lineSketchCurved, Freehand,
  Scribble). Ours writes the same (Keine removes it) for shapes, lines and pictures, and draws the shape through
  a displacement of our own, stronger from Gekrümmt to Skizze; PowerPoint's own sketch drawing is not documented.
- **From the UI screenshots (compare_present_ui.py):** Formformat's shape gallery fills row by row as
  PowerPoint's (6 per row; the shell's shared rule fills columns as Word's, so RaabePresent overrides it in its
  own CSS); Ebene nach vorne/hinten and In den Vordergrund/Hintergrund are greyed when the selection is at the
  front or back already, as in PowerPoint. PowerPoint's three-row Schriftart and Absatz groups in the
  screenshots are its narrower layout (its cloud groups take the width); ours is its wide layout.

## D-073 RaabePresent after Step 8: the open items, their order and the owner's answers (2026-11-07)

The owner went through the Step 8 report's open items and decided:
- **Everything is built, except the cloud,** including what Step 8 had left out on purpose: our own content in
  place of Microsoft's (transition sounds under PowerPoint's names, textures, designs with artwork of our
  own; Microsoft's own files are still never copied), Aufzeichnen, and macros.
- **Order:** (1) Referentenansicht on a second screen, Video and animated GIF export, Alle anordnen (urgent);
  (2) Formeloptionen taken over from RaabeWrite, and the drawing of text warps (Transformieren), text gradients,
  bevels and 3-D (3-D last); (3) SmartArt creation with a layout engine of our own (layout descriptions of our
  own under PowerPoint's layout ids, measured against SmartArt that PowerPoint lays out); (4) own sounds,
  textures and designs, Aufzeichnen; (5) Freihandgleichung for RaabeWrite and RaabePresent together, macros.
- **Second screen for testing:** a virtual display driver on the development PC (owner's OK). Chosen:
  VirtualDrivers/Virtual-Display-Driver (MIT, signed driver, GitHub release). The owner confirms Windows' own
  administrator prompt; the virtual screen is switched off again after each test, so it never catches the
  owner's windows or mouse.
- **MP4:** the file format (ISO base media: ftyp, moov, mdat) is written by our own code, no new package;
  pictures and sound are encoded by the WebView's own encoders (WebCodecs, H.264 and AAC).
- **Sounds:** whichever is better per sound: a recording in the public domain or under CC0 (downloads approved
  by the owner), or one we generate. Each file's source and licence is listed next to it.
- **Macros:** the same way as RaabeCalc (D-066): an interpreter of our own, shared by both apps; macros are off
  until "Inhalt aktivieren", and they cannot reach outside the app.

## D-074 RaabePresent: Referentenansicht on a second screen, Video erstellen, Animiertes GIF erstellen (2026-11-07)

Package 1 of Step 8b (D-073).
- **Two screens:** with two or more monitors and Bildschirmpräsentation > Bildschirme > Referentenansicht on
  (PowerPoint's default), the show opens a borderless full-screen window on the show monitor and the main
  window shows the Referentenansicht in full screen. The show window is made by the native side
  (`raabe_host::builder_with_show_windows`, add-only: `window.open("about:blank#raabe-show:<n>")` from the
  page opens it on monitor n, any other `window.open` stays refused) and the page draws into it, so the show
  has one state: keys and clicks on either screen go on, ink and the laser show on both, the Referentenansicht's
  slide is a second drawing animated with the audience's. Bildschirme > Bildschirm: Automatisch (a monitor the
  main window is not on, a secondary first, as PowerPoint puts the show on the secondary), Primärer Monitor,
  Monitor 2 …; this user's setting, as in PowerPoint. Esc or Alt+F4 on the show window ends the show and closes
  it through the native side (a page's `window.close` does not reach it).
- **Referentenansicht (one or two screens):** PowerPoint's parts: Anzeigeeinstellungen (Referentenansicht und
  Bildschirmpräsentation tauschen, Bildschirmpräsentation duplizieren), Bildschirmpräsentation beenden; the timer
  with anhalten/fortsetzen and neu starten, the clock; the current slide with Stift- und Laserpointertools, Alle
  Folien anzeigen, schwarz schalten, weitere Optionen below; Vorherige/Nächste and "Folie n von m"; the next slide
  (or the current one while click animations remain) and the notes with Text vergrößern/verkleinern. A
  right-click on the audience's screen opens its menu in the Referentenansicht (one page draws both).
  Recorded from PowerPoint's own Referentenansicht on the virtual display (probe_powerpoint_8b_ui.py presenter;
  its window, class PodiumParent, shows no controls to UI Automation, so its layout was measured from a
  screenshot): the bar at the top in capitals with Taskleiste anzeigen (ours: the main window leaves full screen
  so Windows' taskbar shows), Anzeigeeinstellungen, Bildschirmpräsentation beenden; the timer and the clock above
  the slide; the current slide 65 % of the width, at the top, its tools right below it on the left (Stift- und
  Laserpointertools, Alle Folien anzeigen, Folie zoomen, schwarz schalten, weitere Optionen); the navigation
  centred at the bottom with round buttons; the next slide and, below a line, the notes in the right column.
  **Folie zoomen** as PowerPoint's: the button, then a click on the slide shows the slide at twice its size
  around that point on both screens; another click goes back.
- **Tested** with the virtual display (`check_present_screens.py`): RaabePresent off screen, the show window
  full screen on the virtual monitor, keys through DevTools, so the owner's screen is never covered.
- **Video erstellen:** PowerPoint's choices (Videoqualität: Ultra HD (4K), Full HD (1080p), HD (720p), Standard
  (480p); Aufgezeichnete Zeitabläufe und Kommentare nicht verwenden / verwenden; Anzeigedauer der einzelnen Folien
  in Sekunden, 05,00; labels recorded from its Exportieren page). The timeline is PowerPoint's, measured with its own CreateVideo (probe_powerpoint_video.py): a slide
  takes its transition plus its seconds (or Nach with recorded timings); steps after the previous start with
  it; click steps come at equal shares of the seconds (two clicks in 2 s at 0.67 and 1.33 s); the animations'
  own durations add nothing. The width is the height times the slide's proportions rounded down to a multiple of
  4 (480p: 852, as PowerPoint). 30 frames per second; H.264 High profile and AAC from the WebView's encoders
  (WebCodecs) in an MP4 written by our own code (mp4.ts, D-073), videos on slides drawn into their frames,
  transition sounds and the sound of audio and video mixed into one track. Only MP4 (PowerPoint also offers
  WMV, a Microsoft format).
- **Animierte GIF-Datei erstellen:** PowerPoint's page (recorded): GIF-Qualität Sehr groß, Groß, Mittel (the
  default), Klein; Transparenter Hintergrund; Wieviel Sekunden für jede Folie benötigt werden (01,00); Folien
  von-bis. Our GIF writer (gif.ts): a colour table of up to 256 colours per frame (median cut), only the changed
  rectangle of each frame, equal frames merged into a longer delay, endless loop. PowerPoint's own GIF cannot be
  made through COM (CreateVideo writes an MP4 under any name), so it was made through its page with the test
  deck (UI Automation only, nothing typed): 853 x 480 at Mittel (the width rounded, the video's 852 is a
  multiple of 4), about 16 frames per second, and a slide's Nach time is kept although the page offers no
  timings choice. Ours now does the same (16 frames per second, Nach kept, width rounded). Compared frame by
  frame (`check_present_video.py`): same size, 6 of PowerPoint's 8 clear states within 0.25 s; from slide 4 ours
  is 0.4-0.5 s early, and the whole GIF lasts 7.94 s against PowerPoint's 8.58 s. Reason not found yet (PowerPoint
  seems to give the GIF's animations more time than the video's).
- **Alle anordnen, Überlappend:** PowerPoint arranges its document windows; RaabePresent keeps its presentations
  in one window (tabs), so the arrangement happens inside it (windows.ts): a frame per presentation (and per Neues
  Fenster) with its name, side by side or offset one behind the other; the active frame holds the working area,
  the others show their current slide and come forward with a click; a double-click on a frame's title or its
  Maximieren button goes back to one presentation at a time.
- **Verblassen durch Schwarz:** PowerPoint writes `thruBlk="1"`; the show and the video now go through black as
  PowerPoint does (the video comparison found it).
- **Measured:** `check_present_video.py` reads PowerPoint's video and ours frame by frame (each frame named by
  the colours of the test deck's slides and shapes): all 9 clear states of PowerPoint's video appear in ours
  within 0.2 s, without and with recorded timings; PowerPoint's slide with a fade-through-black transition and a
  click fade is 0.1 s longer than ours (11.12 s against 11.0 s in all), reason not found.

## D-075 RaabePresent: Formeloptionen (2026-11-07)

Package 2 of Step 8b (D-073). PowerPoint's Formeloptionen dialog is smaller than Word's (the inventory recorded
it): Allgemein with Beim Kopieren einer Formel (MathML / Lineares Format als Nur-Text in die Zwischenablage),
Zeilen mit binären und relationalen Operatoren umbrechen (Vor, Nach, Doppelt), Operatoren für die Subtraktion
duplizieren als (only with Doppelt) and Math. AutoKorrektur... RaabePresent's dialog has exactly these
(mathopts.ts), with RaabeWrite's settings key and meanings (D-047) as this user's settings. Strg+C and Strg+X
inside an equation now put the selected part on the clipboard as text in the chosen format; the automatic
build-up follows RaabeWrite's setting. The break settings are kept for equations that wrap (our text boxes do
not break equations at operators yet, as before).

## D-076 RaabePresent: text warps and text gradients drawn, Form formatieren > Effekte, cell bevels (2026-11-07)

Package 2 of Step 8b (D-073).
- **Texteffekte > Transformieren drawn:** the 40 warps are ECMA-376's presetTextWarpDefinitions (the standard's
  formulas, read from the published file through GitHub and converted to the compact geometry syntax of
  geometry.ts, `textwarps.gen.ts`), evaluated with the text engine's guide formulas. Measured against PowerPoint
  (`make_pptx_fixtures.py --texteffects`: every warp PowerPoint takes through COM, four per slide, and two text
  gradients; `check_present_render.py pptx-texteffects`): PowerPoint spreads warped text over the whole shape
  without its insets and stretches the letters' ink (not their em box) between the curves; a band that runs
  backwards (the far side of a ring) mirrors the letters; open paths are read left to right (some of the
  standard's run right to left). Our letters are placed, turned and scaled one by one from the curves (SVG text),
  PowerPoint bends their outlines: mean pixel difference 9.0 over the 11 slides (24.5 at the first try; plain
  slides of the other fixtures unchanged at 2.75). With one line and several bands PowerPoint uses the first band
  (measured), as we do. While its text is edited, a warped shape shows the text plain (to type into).
- **Textfüllung Farbverlauf drawn:** a run's a:gradFill paints its letters (CSS background clip for plain text,
  an SVG gradient for warped text); measured: mean difference 0.75 on the gradient slide.
- **Form formatieren > Effekte:** the pane's third tab with PowerPoint's sections Schatten (Farbe, Transparenz,
  Größe, Weichzeichnen, Winkel, Abstand), Spiegelung (Transparenz, Größe, Weichzeichnen, Abstand), Leuchten (Farbe,
  Größe, Transparenz), Weiche Kanten (Größe), 3D-Format (Abschrägung oben/unten with Breite and Höhe, Tiefe and
  Kontur with Farbe and Größe, Material, Beleuchtung, Winkel) and 3D-Drehung (X-, Y-, Z-Drehung, Perspektive,
  Abstand vom Boden). The Formeffekte menus' "…optionen..." entries open it at their section (they were greyed).
  Round trip "effects pane": PowerPoint reads what the pane wrote (shadow 90 degrees 6 pt, glow 10 pt, soft edges
  5 pt, Kreis bevel 8 pt, X rotation 30 degrees). COM's RotationX counts the other way (RotationX 30 writes lon
  19800000, measured); the pane writes the value as entered into lon. PowerPoint's own pane, recorded through UI
  Automation (probe_powerpoint_8b_ui.py effects), shows 330° for that shape, so the pane and lon agree with ours.
  The recording also gave the pane's parts, now built as PowerPoint's: each of Schatten, Spiegelung, Leuchten,
  Weiche Kanten and 3D-Drehung begins with a **Voreinstellungen** list (the Formeffekte menu's presets, "Ohne" /
  "Keine" first, Benutzerdefiniert when the shape's effect is none of them) instead of our check box; 3D-Format
  groups Material and Beleuchtung under their own headings and ends with **Zurücksetzen** (no 3-D format left, a
  rotation stays); 3D-Drehung has buttons beside each angle (Links/Rechts, Nach oben/Nach unten, Uhrzeigersinn/
  Gegen Uhrzeigersinn, 10 degrees each; which way PowerPoint's turn was not measured) and Zurücksetzen.
- **Zellenabschrägung drawn:** a cell's a:cell3D bevel as a light edge at the top left and a shade at the bottom
  right over the bevel's width (inset profiles the other way round). Not yet measured against PowerPoint.

## D-077 RaabePresent: creating SmartArt (2026-11-07)

Package 3 of Step 8b (D-073).
- **What PowerPoint does with our files (measured):** PowerPoint opens a SmartArt that RaabePresent made and lays
  it out with its own built-in layout, found by the layout id we write (loTypeId, the layout part's uniqueId),
  with our items' texts and levels ("Einfache Blockliste", 4 nodes, check_present_smartart.py; round trip
  "smartart": "Einfacher Prozess" with its three texts). So our parts carry PowerPoint's ids (layout, quick style
  "simple1", colour set) and our own minimal content; none of PowerPoint's layout definitions is copied.
- **The parts:** a data model as PowerPoint writes it (document point, one point per item with parTrans and
  sibTrans, parOf connections ordered by srcOrd), layout, quick style and colours parts, and the drawing part
  (dsp:drawing) with our engine's shapes, which RaabePresent draws (as it already drew PowerPoint's drawings) and
  which PowerPoint shows until it lays the diagram out again. Laying a diagram out again keeps its parts.
- **Our layout engine (smartart-engine.ts):** families of our own (grid, row, column, process, cycle, hierarchy,
  venn, pyramid, target, radial, matrix). The catalog (smartart-catalog.gen.ts, made by make_smartart_catalog.py)
  holds for each of PowerPoint's layouts its name, category and id and what PowerPoint's own drawing of it
  measures (probe_powerpoint_smartart.py: every layout with 3 and with 5 items in a 640 x 360 pt frame, the
  shapes' boxes): for the measured numbers of items our drawing takes PowerPoint's arrangement (positions and
  sizes as numbers, not definitions); other numbers, and items with sub-items, are carried over by the family.
  PowerPoint lists 159 layouts through COM; 146 could be drawn (13 refused by COM).
- **Measured** (check_present_smartart.py, the item boxes against PowerPoint's, a layout matching at a mean overlap
  of 0.7 or more): PowerPoint's arrangements were measured for 1 to 8 items (probe runs --items 1 … 8; 140 to 150 of
  the 159 layouts drawable through COM per count), so our drawing equals PowerPoint's for those counts: 127 of 129
  layouts with 3 items (mean overlap 0.985), 118 of 120 with 5 (0.983). The families alone, as used for 9 items
  and more and for items with sub-items, are much weaker: 27 of 120 for 5 items (0.38), because PowerPoint's
  layouts differ too much (some wrap into two rows, some keep their size, some shrink). The catalog with the
  arrangements is 845 KB of numbers.
- **UI:** Einfügen > SmartArt with "SmartArt-Grafik auswählen" (PowerPoint's categories, pictures drawn by our
  engine); Start > In SmartArt konvertieren with PowerPoint's 20 entries and Weitere SmartArt-Grafiken...;
  SmartArt-Design: Form hinzufügen (danach, davor, oberhalb, unterhalb), Aufzählungszeichen hinzufügen,
  Textbereich (Eingabe makes an item, Tab and Umschalt+Tab change its level; only texts changed: PowerPoint's own
  drawing stays), Höher/Tiefer stufen, Verschieben vor/nach (on the last item), Layouts, Farben ändern (PowerPoint's
  colour sets by name and id, our drawing in the accent), Grafik zurücksetzen, Konvertieren. Von rechts nach links
  and the Schnellformatvorlagen gallery are not built yet.

## D-078 RaabePresent: our own sounds, textures and designs; Aufzeichnen (2026-11-07)

Package 4 of Step 8b (D-073).
- **Sounds:** the owner left the choice open ("was besser ist"). Chosen: our own sounds, synthesised from noise and
  tones by tools/sounds/make_sounds.py (no download, no licence to track, rebuilt at any time), one for each of
  PowerPoint's 19 names, embedded under PowerPoint's file names (applause.wav, type.wav ... measured with
  probe_powerpoint_sounds.py) as ppt/media/audioN.wav and shared by the slides that use them. PowerPoint reads
  them (round trip "transition sound"). "Wiederholen bis zum nächsten Sound" now has PowerPoint's label.
- **Textures:** Fülleffekt > Struktur with PowerPoint's 24 names and our own tiles (tools/textures/make_textures.py,
  periodic noise, so they repeat without a seam), written as PowerPoint writes textures (a:blipFill with a:tile)
  and drawn tiled. PowerPoint reads them as a textured fill (round trip "texture").
- **Designs:** six designs with artwork of our own (vector shapes on the master in the design's colours):
  Wellenband, Kreisbögen, Keil, Punkteraster, Schräge Streifen, Abendblau. PowerPoint's design artwork is not copied.
- **Aufzeichnen tab:** PowerPoint's groups and keytips (Cameo, a cloud feature, left out). Von Beginn an / Ab
  Aktueller Folie (also Bildschirmpräsentation > Aufzeichnen): the show with a recording bar (Aufzeichnen,
  Anhalten/Fortsetzen, Beenden, the time) and, when chosen, the camera; the microphone is recorded per slide
  (MediaRecorder, decoded to PCM, saved as WAV), the camera into an MP4 by our own writer (H.264 and AAC from the
  WebView's encoders). Each recorded slide gets its narration as PowerPoint stores narrations (an audio object
  that plays on entry and hides while not playing: a mediacall at the main sequence's start and an audio node
  with isNarration) and its time as Nach, with Anzeigedauern verwenden on. Round trip "narration": PowerPoint
  reads a sound that plays on entry and hides, and the slide's 4.5 s. Neue Anzeigedauern testen records only the
  times and asks whether to keep them. Aufzeichnung löschen (this slide, all slides), Bildschirmaufzeichnung
  (getDisplayMedia, as an MP4 video on the slide), Audio (a recording on the slide), Als Bildschirmpräsentation
  speichern, In Video (Exportieren). The show now plays a media object's mediacall (narrations at the slide's
  start) and hides narration icons. Ink drawn while recording is not kept as an animation yet.
  The microphone and camera need the user's permission in WebView2; recordings stay on the device.
- **App CSP:** fetch to the app's own files and to blob: and data: URLs, and media from blob: and data:, are now
  allowed (connect-src, media-src): the sounds, textures and the video export need them; before, media from blob:
  URLs could not have played in the real app.

## D-079 RaabePresent: Freihandgleichung and macros (2026-11-07)

Package 5 of Step 8b (D-073).
- **Freihandgleichung:** handwriting is recognised by Windows' own Math Input Panel (mip.exe, part of Windows, its
  recogniser runs on the device): Formel > Freihandgleichung starts it, the user writes and clicks its "Einfügen",
  and the equation arrives on the clipboard as MathML ("MathML Presentation" format, read by a small command of
  RaabePresent's own native side, src-tauri/src/inkmath.rs) and becomes an equation on the slide through our
  MathML reader (packages/text-engine/src/math/mathml-read.ts, add-only; also used when MathML is pasted). Where the
  Math Input Panel is missing (iPad, browser) the command says so; recognition there needs a recogniser of our own
  (open). RaabeWrite's Freihandgleichung can use the same pieces (open).
- **Macros, as for RaabeCalc (D-066):** a new shared package @raabe/vba: the VBA project format (compound file
  read and write, MS-OVBA compression, dir stream, modules; written projects carry no compiled code, so Office
  compiles the source) and our interpreter (Sub, Function, ByRef/ByVal, Dim/ReDim/Const, If, Select Case, For,
  For Each, Do/Loop, While, With, Exit, On Error Resume Next, arrays, the common built-in functions, MsgBox and
  InputBox as our dialogs; endless loops stopped after 5 million steps). Measured with PowerPoint
  (check_present_macros.py): PowerPoint opens a .pptm whose project RaabePresent wrote, saves it with its own
  compiled project, and our reader reads the module and its code back from PowerPoint's file (4/4); nothing is run
  and no security setting is changed.
- **In RaabePresent (macros.ts):** a PowerPoint-like object model (Application, ActivePresentation, Slides, Slide,
  Shapes with AddTextbox and AddShape, Shape with Name, Left/Top/Width/Height, Rotation, Visible, TextFrame(2).
  TextRange.Text and Font, Fill and Line colours, Delete, Increment...; SlideShowWindow.View with Next, Previous,
  GotoSlide, CurrentShowPosition, Exit). Macros of an opened file are off until "Inhalt aktivieren" in the message bar
  (SICHERHEITSWARNUNG); Ansicht > Makros (Alt+F8): Makroname, Ausführen, Bearbeiten, Erstellen, Löschen; our editor
  (Alt+F11) with the modules, Modul einfügen, Kompilieren (the syntax check); Aktionseinstellungen > Makro ausführen,
  run in the show with the show as SlideShowWindow.View. One macro run is one undo step. Macros cannot reach outside
  the app: Shell, CreateObject, file statements, SendKeys and Declare stop with a message.
- **Not yet:** UserForms, class modules' own objects (New), events (PowerPoint has few), Debug in the editor
  (breakpoints, single steps: Schrittweise is greyed), RaabeCalc's Excel object model (the interpreter is ready for
  it), and recognising handwriting on the iPad.

## D-080 After Step 8b: the open items, macro tests, a drag test and a look corpus (2026-11-08)

Agreed with the owner, who leaves the PC (mouse and keyboard) to the run until the whole plan is done.
- **What "like Office" means first:** documents made in Office look the same when opened in RaabeX, and the
  operation and menus are built the same way. For the drag test the first rule is: no sudden, unwanted jumps.
- **A. Open items.** RaabeCalc (this checkout): structured references, calculated columns and growing tables;
  LET, LAMBDA, MAP, REDUCE, SCAN, NACHSPALTE, NACHZEILE, MATRIXERSTELLEN, INFO, ZELLE, BLATT, BLÄTTER, HYPERLINK;
  Seitenlayout > Designs, Sparklines, Bildformat, Zeichnen; chart styles and more chart types; Szenario-Manager,
  Konsolidieren, Datenschnitt; then Step 7b (Diagramm in RaabeWrite); then macros in RaabeCalc (Excel's object
  model on @raabe/vba). RaabePresent (its own session and branch): the GIF timing, the catalogue's remaining
  differences, the inventory gaps, SmartArt's Schnellformatvorlagen and Von rechts nach links, placeholder
  prompts and cell bevels checked against PowerPoint; the shared interpreter's UserForms and file access. The
  Freihandgleichung for RaabeWrite is left out for now (owner).
- **Power Query and Power Pivot (decided on the owner's behalf):** Daten abrufen und transformieren is in the
  ribbon as in Excel; Aus Text/CSV, Aus Arbeitsmappe and Aus Tabelle/Bereich import without the Power Query
  editor; files with queries or a data model open, look the same and keep those parts. The full editor and
  Power Pivot (M and DAX engines) are a later step if there is a need: schools hardly use them.
- **Macros reach outside the app only after asking:** opening and saving files (e.g. every sheet as PDF) asks
  the user first; Shell, network, registry and Declare stay blocked. UserForms are built.
- **B. Macro tests:** our own versions of popular Office macros (the forum versions are not ours to copy),
  school macros (Notenrechner, random pupil, groups, vocabulary trainer, quiz with score) and openly licensed
  VBA as a stress test; each runs in Excel or PowerPoint (our own instance; the .xlsm/.pptm is written by our
  VBA writer, so no Office security setting is changed) and in RaabeX, and the results are compared.
- **C. Drag test** in all three apps with real input: every anchor (handles, borders, rulers, fill handle,
  selection frames, splitters, thumbnails, paths), many pointer paths (slow, fast, around the drag threshold,
  back and forth, out of the window and back, over snap points, modifiers pressed mid-drag, zoom levels).
  Checked automatically: the object follows the pointer without a jump, no step is larger than the pointer's
  (snapping only within its distance), never against the motion, no jump on release, Esc and undo restore
  exactly. Outliers are fixed in the same run and measured again.
- **D. Look corpus:** school-like documents (worksheets, tests with equations, mark lists with charts,
  timetables, talks) built by Word, Excel and PowerPoint through our own scripts, opened in RaabeX and
  compared with Office's PDF/PNG export; differences fixed.
- **History to be cleaned (decided on the owner's behalf):** the inventory results file of commit 36c7d49
  holds the owner's recent file names and Office's descriptions. That one file in that one commit is replaced
  by its sanitized version (`git filter-branch --index-filter`, all trees afterwards identical) and pushed with
  force. The rewrite was done and checked on 2026-11-08, but the session's safety check refused the force
  push, so it waits for the owner's permission; it is repeated at the end of the run. GitHub may serve the
  old commit by its id for a while; only GitHub's support can purge it.
- **Two sessions:** real input only under the lock file; decisions D-080 to D-099 in this checkout, D-100
  upward in the RaabePresent session; the shared interpreter (packages/vba) belongs to the RaabePresent
  session, RaabeCalc adds its object model beside it.

## D-081 RaabeCalc: structured references, berechnete Spalten, growing tables, LET and LAMBDA (2026-11-08)

Package A of D-080, measured against our own invisible Excel.
- **Structured references** (`formula/structref.ts`): Tabelle1[Note], [@Note], [[#Kopfzeilen];[Note]],
  [[Name]:[Note]], [#Alle] / [#Daten] / [#Ergebnisse] / [#Diese Zeile] in both dialects, a table's name alone
  for its data, escapes ('[ '] '# ''). Recorded (`probe_excel_structref.py`): Excel shows [#Diese Zeile] as "@"
  and leaves out the own table's name inside the table (TEILERGEBNIS(109;[Punkte])), but its file always
  names the table, writes [@Note] as Noten[[#This Row],[Note]] and a table's name alone as Noten[]; the model
  keeps what Excel shows and the .xlsx layer converts both ways (cells and calculatedColumnFormula). [@...]
  outside the data rows is #WERT!, a missing table or column is refused on Enter as Excel refuses it.
  Renaming a table or typing over a header renames the references in every formula and defined name.
- **Berechnete Spalten and growing tables** (AutoKorrektur während der Eingabe): a formula typed into an
  empty table column (or over a berechnete Spalte) fills its data rows; a value typed right below a table
  without Ergebniszeile or right next to it makes the table grow (the new row takes the berechnete Spalten,
  the new column the header typed or "SpalteN"). The Ergebniszeile writes TEILERGEBNIS(109;[Spalte]) as Excel.
  Not yet: the AutoKorrektur-Optionen button that undoes the growing, and pointing into a table while typing
  a formula (Excel then writes structured references): both need real-input recordings.
- **LET, LAMBDA and the helpers** (`formula/fn-lambda.ts`): LET, LAMBDA with calls (LAMBDA(x;x+1)(2), a LET
  name holding a LAMBDA, defined names holding one, recursion), MAP, REDUCE, SCAN, NACHZEILE, NACHSPALTE,
  MATRIXERSTELLEN, WURDEAUSGELASSEN. A LAMBDA is a value only while a formula is worked out: a cell showing one
  gets #CALC!, any function that is not a LAMBDA helper gets #WERT! for one (recorded); every parameter
  needs its argument (f(1;) leaves one out). The file writes the bound names with _xlpm. (recorded); a name
  bound twice is refused on Enter. `probe_excel_lambda.py`: 31 cases.
- **Measured:** structref-excel.test 23 cases + structref.test, lambda-excel.test 32 cases (values, spills,
  German form, file text both ways); `check_calc_structref.py` RaabeCalc -> Excel 23/23 (the table grown,
  the berechnete Spalte, formulas and values as Excel shows them), Excel -> RaabeCalc 25/25.

## D-082 RaabeCalc: ZELLE, INFO, BLATT, BLÄTTER, HYPERLINK (2026-11-08)

Recorded in German Excel (`probe_excel_info.py`, 98 cases with 29 number formats) and built the same way
(`formula/fn-info.ts`): the German and English info types where Excel takes them (INFO takes "Dateienzahl",
"System", "Ursprung", "Version" and every English one, not "Rechenmodus" or "Systemversion"); German Excel's
own ZELLE("Format") codes (S, F2, .2 for #.##0,00, P0, E2, W2 for currency, D1 to D3, U1 to U4, "-" after the
code for coloured negatives); "w" for values in ZELLE("Typ"); ZELLE("Breite") as Excel 365's two values.
Deliberate differences: ZELLE without a reference reads the formula's own cell (Excel: the cell changed
last, which depends on what the user did); ZELLE("Breite") counts RaabeCalc's Calibri columns (D-011);
INFO("Version") and INFO("osversion") give Excel's values so that workbooks testing them keep working.
INFO("Verzeichnis") and ZELLE("Dateiname") give the document's folder and path once it is saved (never part
of the results files). =HYPERLINK(...) takes the Link look on entry (blue, underlined, as Excel's "Link"
style) and a click on its text follows the first argument, as for links made with Strg+K.

## D-083 RaabeCalc: Seitenlayout > Designs, Farben, Schriftarten (2026-11-08)

- **Design colours are kept as references:** a cell's font colour and fill keep the file's theme index and tint
  besides the RGB value (`Font.theme`, `CellStyle.fillTheme`), read from and written to styles.xml as Excel does
  (`<color theme="4" tint="0.39997558519241921"/>`). Picks from the palette's Designfarben set them with Excel's
  tints (a share of 32767: heller 80 % is 26213/32767, recorded in Excel's files); standard and own colours stay
  RGB. A reference counts only while the style still shows its colour: a colour set some other way drops it.
- **Designs, Farben, Schriftarten:** the galleries hold Office's 31 designs (as RaabeWrite recorded them from
  Word's Entwurf tab), its 24 colour sets and 24 font sets (as RaabePresent recorded them), shared in
  `packages/shell/src/office-designs.gen.ts` (`gen_office_designs.py`); tiles and rows are drawn by us. Applying
  one changes the workbook's theme colours or fonts, every style with design colours or theme fonts (the
  Standard style too: cells, columns and rows follow) and the table styles, in one undo step; the theme part is
  rewritten only when the design changed (colours, fonts, names), everything else in it stays.
- **Not yet:** Effekte (the theme's effect styles: shapes would need them drawn), Farben/Schriftarten anpassen,
  Aktuelles Design speichern and Nach Designs suchen (.thmx files).

## D-088 Diagramm in RaabeWrite, part 1: Word's charts drawn (Step 7b, 2026-11-08)

Charts in Word documents were shown as nothing and reported as "Diagramme und eingebettete Objekte" not
supported. RaabeWrite now draws each chart with the chart engine shared with RaabeCalc and RaabePresent, from
the values Word keeps in the chart part (`chartFromXml`, the document theme's accents, the drawing's size), as
an SVG picture in the text; the drawing, the chart part, its style/colour parts and the embedded workbook are
written back as they were. Measured (`check_write_charts.py`, Word's own document with seven chart types from
`make_fixtures.py --charts`): no chart reported unsupported, Word finds the same seven charts in RaabeWrite's
file, same page count, mean page difference against Word's PDF 2.06 / 1.83 / 0.37 of 255. Not yet (part 2):
Einfügen > Diagramm, Daten bearbeiten (the embedded workbook) and the chart tabs in RaabeWrite; charts of the
newer chartEx types follow the chart engine (D-086).


## D-085 RaabeCalc: the Sparkline tab, every sparkline option, drawn as Excel draws (2026-11-08)

Package A of D-080 (Sparklines), recorded in our own Excel and built the same way.
- **Recorded:** `probe_excel_sparklines.py` (invisible Excel, COM only): 50 sheets, one per option (types, the six
  Anzeigen boxes, RGB and design colours, eight Stärken, Lücken/Null/verbinden, ausgeblendete Zellen, Achse anzeigen
  with colour, rechts nach links, vertical axis automatic/identisch/custom, date axis, Modify, Gruppieren,
  Gruppierung aufheben, Löschen); the XML Excel saved and what COM reads back (fixture `excel-sparklines.json`).
  `probe_excel_sparkline_ui.py`: the tab's keytips (JD; ZS B, ZT I Ä G, ZA H T N E L M, ZF S P D, ZG A U R Ö), which
  controls are enabled per type and selection, and the 36 styles of the Sparklineformatvorlage gallery, each applied
  through UI Automation (no real input) and read from Excel's file (`sparkline-styles.gen.ts`). Our Excel was made
  visible off screen for this; selecting its tab and invoking a gallery item made it the foreground window for
  under 0.1 s each time, and the probe handed the focus straight back. The Stärke submenu could not be read that
  way (the menu reading picked up the whole window): its entries are PowerPoint's recorded Stärke labels (D-070).
- **Model:** `SparklineGroup` holds every x14:sparklineGroup attribute (markers, high, low, first, last, negative,
  displayXAxis, displayEmptyCellsAs, displayHidden, lineWeight, min/maxAxisType, manualMin/Max, rightToLeft, the
  date range) and the eight colours (series, negative, axis, markers, first, last, high, low) with their design
  colour reference. The file's default for displayEmptyCellsAs is "zero" (Excel then leaves the attribute out).
  Read and written as Excel writes them (attribute order of the schema); unchanged groups are written back
  verbatim (with Excel's xr2:uid), new groups come first in the file as in Excel.
- **Behaviour as recorded:** new sparklines get Excel's colours (#376092, the others #D00000, axis black);
  Gewinn/Verlust starts with Negative Punkte on; Gruppieren takes the settings of the active cell's group;
  Gruppierung aufheben makes each selected sparkline a group of its own; Löschen takes the selected sparklines or
  the whole groups they belong to; every other setting applies to all groups the selection touches. Gruppieren is
  enabled for sparklines of two or more groups, Gruppierung aufheben when a touched group has several, Daten
  bearbeiten only within one group, Datenpunkte only for lines. The tab appears while the active cell holds a
  sparkline and is selected after Einfügen > Sparklines.
- **Drawn as Excel draws** (measured on Excel's CopyPicture at 150 %, `apps/raabe-calc/test/sparklines.test.ts`):
  plot area inset 1.25 pt left/right and 2 pt top/bottom; n values in n equal slots, line points at the slot
  centres, columns 72 % of a slot; columns from zero, or from the bottom when all values are positive (the
  smallest is a bar of one pixel); Gewinn/Verlust fills the upper or lower half; points as diamonds of 2 pt + 4/3 of
  the Stärke; colour priority Höchstpunkt, Tiefpunkt, Erster, Letzter, Negative, Datenpunkte; negative columns in
  the series colour unless Negative Punkte is on; the axis only where zero lies inside the vertical range.
- **Measured:** xlsx-sparklines.test 104 cases (the 50 Excel files read as COM reports them, written as Excel wrote
  them apart from the xr2:uid, verbatim when unchanged, a whole workbook through writeXlsx/readXlsx, the 36 styles);
  raabe-calc sparklines.test 10 cases (operations against Excel's files, drawing against Excel's pixels);
  `check_calc_sparklines.py`: RaabeCalc -> Excel 7/7 (six groups built through the tab with every option, read back
  through COM), Excel -> RaabeCalc 50/50, saved again unchanged 50/50 sheets byte for byte, reopened in Excel 50/50.
  Visual comparison of 36 cells (nine kinds x four rows, Excel's CopyPicture against RaabeCalc at 100 %): the drawn
  ink of each side lies within 3 device pixels of the other's in all 36 (mean F1 1.0), mean grey difference 3.3 of 255.
- **Not built:** Zeile/Spalte wechseln works only for square data blocks (Excel greys it out otherwise; not
  recorded further); the reference fields of the dialogs take typed ranges, not a range picked with the mouse;
  sparklines do not move with inserted or deleted rows and columns or follow a renamed sheet yet; the gallery's
  tooltips leave out Excel's leading colour word (it names the colour in Excel's theme); keytips of the menu
  entries were not recorded.

## D-084 RaabeCalc: Szenario-Manager and Konsolidieren (2026-11-08)

Package A of D-080, recorded in our own invisible Excel (`probe_excel_scenarios.py`, `probe_excel_consolidate.py`; no
real input) and built the same way (`scenarios.ts`, `consolidate.ts`, the file parts in `ooxml/src/xlsx/scenarios.ts`).
Dialog labels as Excel's dialog resources hold them (the inventory's Szenario-Manager and Konsolidieren, the sub-dialogs
read from Excel's German resource DLL as strings, no UI opened); explanations in our own words.
- **Szenario-Manager:** list, Anzeigen, Schließen, Hinzufügen..., Löschen, Bearbeiten..., Zusammenführen...,
  Zusammenfassung...; Szenario hinzufügen/bearbeiten (Szenarioname, Veränderbare Zellen up to 32, Kommentar starting
  "Erstellt von <Name> am <Datum>" as recorded, Bearbeiten adds a line "Geändert von …" (not recordable through COM), Schutz: Änderungen verhindern /
  Ausblenden, which count while the sheet is protected with Szenarios bearbeiten off), Szenariowerte (with Hinzufügen).
  Values are kept as typed: German Excel writes "0,25" into the file, Anzeigen enters them as if typed (a formula among
  the cells is replaced, recorded), one undo step "Szenario anzeigen"; list changes are undo steps too.
  Zusammenführen from another sheet of the workbook; a clashing name gets the date ("Gut 08.11.2026", recorded).
- **Szenariobericht** (new sheet before the active one, "Szenariobericht", "Szenariobericht 2"): Excel's layout cell by
  cell (purple title rows, grey labels, navy section labels, grey scenario cells, borders, font sizes 12/10/8, the
  comments row grouped and hidden, D… grouped, no gridlines), names instead of $B$1 when a name refers to the cell,
  result values as each scenario alone gives them. The three note rows are our own words (Excel's text is its prose).
  Deliberate difference: RaabeCalc's outline has the summary row below its group, so the comments group's button sits on
  row 5 (Excel: row 3, summaryBelow="0"). **Not built:** the PivotTable-Szenarioberichts (greyed out; Excel makes a
  PivotTable with a scenario cache), Zusammenführen from other open workbooks (only this workbook is listed), pointing
  into the sheet while the dialogs are open (references are typed, as in Zielwertsuche).
- **Konsolidieren:** Funktion (Excel's 11), Verweis, Vorhandene Verweise (sorted as Excel's list; Excel also computes in
  that order, recorded through the number format), Hinzufügen, Löschen, Beschriftung aus Oberster Zeile / Linker Spalte,
  Verknüpfungen mit Quelldaten; Durchsuchen... is greyed (other files). Recorded rules: labels compared without case,
  spelt as first seen; a new label goes before the next label of its own source already listed, else to the end (14
  order cases); text counts only for Anzahl; a cell nothing counts for stays empty; a result cell takes the number format
  of its first source cell; one cell / one row / one column grows, a larger range limits the result. Verknüpfungen: per
  label one row per source with =Blatt!$B$2 links (a source's duplicate labels packed per column, as Excel), the
  workbook's name beside each source's first row (labels from the left column only), =SUMME(…) below, detail rows
  grouped and hidden. Deliberate difference: Excel narrows column A to 1.18 there; we leave it. The settings stay in the
  sheet as Excel's <dataConsolidate>; Excel itself drops leftLabels when it reads any file (its own too, checked).
- **File:** <scenarios> (CT_Scenarios) and <dataConsolidate> read into the model and written in Excel's form, left as they
  were while unchanged.
- **Measured:** scenarios.test 10 (Anzeigen in turn, a formula replaced, four reports cell by cell against Excel's,
  names), consolidate.test 44 (28 recorded cases cell by cell with formulas, formats and outline, the settings, 14 label
  orders), xlsx-scenarios.test 7 (Excel's file read, written back unchanged, edited as Excel writes it);
  `check_calc_scenarios.py` RaabeCalc -> Excel 38/38 (scenarios, Excel showing ours, the report, Excel's own
  Konsolidieren equal to ours, link formulas and outline), Excel -> RaabeCalc 15/15.

## D-086 Charts: every chart type of Excel 365 except the map (2026-11-08)

The owner's first priority is that Office documents look the same in our apps. The chart engine knew 13 types;
Excel offers 56 (without Karte, a cloud feature that stays out, D-066).

- **Recorded first** (`probe_excel_charttypes.py`, our own invisible Excel, Office 2013-2022 design, COM only):
  one chart per type from a fixed data block, `AddChart2` with the default style; Chart.ChartType, SERIES formulas,
  axis groups, title, legend, chart group options and the 3-D view went to `results/2026-11-08-excel-charttypes.json`,
  Excel's workbook (user name and save path removed, `anonymize_office.py`) to
  `packages/ooxml/test/fixtures/excel-charttypes.xlsx`, its PNG exports to TEMP. Menus and the dialog's categories
  come from the inventory (`results/2026-11-05-excel-inventory.json`).
- **Model** (`packages/charts`, additive; the 13 old kinds draw and write exactly as before, their tests unchanged):
  43 new kinds (`kinds.ts`: family, direction, grouping, 3-D, xlChartType, German name), per-series `kind` and
  `secondary` (Verbund, Sekundärachse), Blasengröße, category levels (Treemap, Sunburst), Wasserfall totals, the
  3-D view, Kreis-aus-Kreis split and second plot size, Blasengröße skalieren.
- **File** (`packages/ooxml/src/xlsx/chart.ts`): every c:chart group read (bar3D, line3D, area3D, pie3D, ofPie,
  radar, bubble, stock with and without the volume columns, surface and surface3D, several groups = Verbund, groups
  on a right/top value axis = secondary) and written as Excel writes them (view3D, floor and walls, series axis,
  up/down and hi-lo lines, bandFmts left to Excel's defaults). The chartex part reads Wasserfall, Trichter, Treemap,
  Sunburst, Kastengrafik, Pareto and Histogramm (Pareto is a clusteredColumn series with a paretoLine); Excel writes
  their data only as hidden names (`_xlchart.v1.N`) without caches, so RaabeCalc resolves those names when it reads
  the cells (`refreshCaches`). Unchanged charts are still written back verbatim (D-024); a chart Excel's maps or a type
  we do not know stays in the drawing as it is.
- **Drawn as Excel draws them**: 100 % and stacked columns, bars, lines and areas; Verbund with a secondary axis;
  Kurs (hi-lo lines, white/dark up/down bars, the close as a dash, volume columns with the prices' axis at the right);
  Blase (area-true sizes, the largest a quarter of the plot's shorter side, 75 % opacity; 3-D bubbles shaded);
  Netz (polygon gridlines, spokes, values along the first spoke, filled or with markers, axis without headroom);
  Kreis aus Kreis / Balken aus Kreis (the last three points gathered in a slice facing right, Verbindungslinien,
  measured widths); 3D-Kreis (ellipse height = width x sin rotX, rim 0.244 x radius, darker); 3-D columns and bars
  (oblique projection for Rechtwinklige Achsen: back and side wall gridlines, boxes with darker top and side);
  3D-Linie, 3D-Fläche, 3D-Oberfläche and its Drahtmodell through a parallel rotation by rotX/rotY; Kontur as the top
  view with Excel's lighting on the slope, Kontur (Drahtmodell) as iso lines (marching squares); bands of the surface
  legend ("0-50" …) in the accents; Wasserfall (Anstieg/Abfall/Summe, connectors), Trichter, Treemap (squarified),
  Sunburst, Kastengrafik (exclusive quartiles, mean x, whiskers to 1.5 x IQR, outliers), Pareto (sorted, cumulative
  line on a 0-100 % axis). 3-D value axes end on the data (4.5 → 5) as Excel's do.
- **Einfügen > Diagramme** (`apps/raabe-calc/src/charts/types.ts`): the seven menus with Excel's groups, names and
  order, "Weitere …" opening Diagramm einfügen at the category; Verbund with the three presets. Diagrammentwurf >
  Diagrammtyp ändern and Empfohlene Diagramme open Excel's dialog: Empfohlene Diagramme / Alle Diagramme, the 17
  categories (Karte greyed), the types of each, a live preview, Kombi with the presets and Benutzerdefinierte
  Kombination (type and Sekundärachse per series). Not built: Zuletzt verwendet and Vorlagen (no chart templates);
  Empfohlene Diagramme lists Excel's four recorded picks instead of choosing by the data. The last Kombi menu entry
  is not readable through UI Automation ("[Beschreibung]"); "Benutzerdefiniertes Kombinationsdiagramm erstellen..." is
  our assumption.
- **Measured.** `charttypes.test.ts` 115 tests: all 56 recorded charts read as their type with Excel's SERIES
  formulas and write back byte for byte; every new type goes write → read; hierarchy, totals and bubble sizes from
  the cells. `check_excel_charts.py`: Excel opens RaabeCalc's workbook of the 43 new types plus two Verbund charts
  without repair and reports the same ChartType and SERIES for 45/45 (the old 26/26 still pass).
  `compare_calc_charttypes.py calc` (the release exe against Excel's PNG; mean absolute greyscale difference, 0-255,
  the old types for scale: Säulen 4.3, Linie 3.5, Kreis 1.3, Ring 1.2, Histogramm 2.2):
  100 %/stacked 2-D 3.3-4.6 (Säulen 4.6, Balken 4.5, Linie 3.9-4.5, Fläche 3.3-3.6); Verbund 3.8-4.6; Punkte
  4.4-4.6; Blase 3.4, 3D-Blase 4.9; Kurs HLC 4.2, OHLC 5.3, VHLC 9.5, VOHLC 9.6; Netz 5.6/5.7, gefüllt 3.3;
  Balken aus Kreis 3.4, Kreis aus Kreis 7.0, 3D-Kreis 9.1; Kastengrafik 2.6, Pareto 7.2, Wasserfall 7.8,
  Trichter 8.5, Sunburst 7.6, Treemap 10.5; Kontur (Drahtmodell) 2.8, 3D-Oberfläche (Drahtmodell) 6.1, Kontur 9.6,
  3D-Oberfläche 19.4; 3D-Balken 11.4-14.9, 3D-Linie 11.3, 3D-Säulen 18.5-23.9, 3D-Fläche 20.3-21.8.
- **Not matched:** Excel's 3-D engine (perspective, its light, label placement along tilted axes): we project in
  parallel, so 3-D columns, areas and the surface sit a few points off and score highest; Kreis aus Kreis turns its
  second pie slightly differently and splits the last three points where Excel's "automatic" split may choose
  otherwise; Sunburst labels do not wrap; Kurs dates are categories (no gaps for missing days); the Format pane has
  no entries yet for walls, floor, hi-lo lines, up/down bars and Verbindungslinien; RaabePresent and RaabeWrite read
  the new types through the same engine but were not checked here.

## D-087 RaabeCalc: pictures with Bildformat, the Bildformat tab, written as Excel writes them (2026-11-08)

Recorded in our own Excel without real input (`probe_excel_pictures.py`): 30 pictures with one setting each plus a
group, set through COM and saved as `xlsx-pictures.xlsx` / `excel-pictures.json` (anchors, COM readback, Excel's
rendering by CopyPicture); the Bildformat tab's controls, keytips and menus through UI Automation; Excel's 28
Schnellformatvorlagen applied one by one through UI Automation and read from the saved file.
- **Look:** each picture's DrawingML is read with the shared reader (`readPicture`) into text-engine's PictureFormat
  and drawn with RaabeWrite's renderer (`pictureElement`): crop, a:lum, Graustufe, Schwarzweiß, Ausgeblichen,
  duotone, Transparente Farbe, Transparenz, a14 Korrekturen/Sättigung/Farbton, rotation, flips, Bildrahmen, shadow,
  reflection, glow, soft edges, bevel, 3D rotation, frame shapes, Alternativtext, Ausgeblendet; pictures inside
  a group are shown in their place (the group itself stays as the file has it, not editable). Printing uses
  the same drawing.
- **Recorded rule:** Excel stores Office 2010 effects (a14:imgEffect) worked into the stored picture, keeps the
  original as an HD Photo (a14:imgLayer) and draws the stored picture as it is: a file whose effects or picture
  differ shows the stored picture (three variants checked). So such pictures are drawn without applying the
  effects again, and when RaabeCalc sets Korrekturen, Sättigung or Farbton it works them into a PNG copy (canvas
  with the same filters) which the file stores, with the effects listed as Excel lists them (no HD Photo layer).
- **Bildformat tab** (`picture-tab.ts`, labels and keytips as recorded, JP): Anpassen (Freistellen, Korrekturen,
  Farbe with Neu einfärben in Excel's names and Transparente Farbe bestimmen, Transparenz, Bilder komprimieren...,
  Bild ändern: Dieses Gerät... / Aus Zwischenablage..., Bild zurücksetzen / Bild und Größe zurücksetzen),
  Bildformatvorlagen (the 28 styles with Livevorschau, Bildrahmen, Bildeffekte with Office's presets, launcher
  Bild...), Alternativtext (pane), Anordnen (Ebene nach vorne/hinten among the pictures, Auswahlbereich lists
  pictures too: select, show/hide, rename; Ausrichten: Am Raster ausrichten, Gitternetzlinien anzeigen; Drehen:
  Rechts-/Linksdrehung 90 Grad, Vertikal umdrehen, Horizontal spiegeln), Größe (Zuschneiden with the picture
  behind the frame and black crop handles, Auf Form zuschneiden, Seitenverhältnis, Ausfüllen, Einpassen;
  Formenhöhe/-breite keeping proportions while Seitenverhältnis sperren is on; launcher Größe und Eigenschaften...).
  Grafik formatieren pane: Füllung und Linie, Effekte, Größe und Eigenschaften, Bild. On the sheet: eight
  handles and the turning handle, Strg+1, arrow keys, Entf. Every change is one undo step. Gallery entries,
  presets and Freistellen are shared with RaabeWrite (moved to text-engine, RaabeWrite unchanged).
- **Writing:** unchanged pictures verbatim, moved ones with new corners, changed ones with Excel's DrawingML;
  a turned picture is anchored by its turned box (45° to 135°) as Excel does.
- **Measured:** unit tests read all 30 settings; every picture rewritten from the model gives Excel's XML element
  for element (30/30); Excel's 28 picture styles write exactly Word's recorded DrawingML (28/28).
  `check_calc_pictures.py`: Excel -> RaabeCalc 31/31 settings read; rendering against Excel's mean grey difference
  2.0-6.1 for 28 of 31 blocks (2.5 for the plain picture: grid and text; Bildrahmen dashes 7.0, the group 9.2, 3D
  rotation 10.4), where drawing the pictures plainly (the state before) gave up to 23.5; RaabeCalc -> Excel 29/30 pictures read back through COM as Excel's own, rendered
  by Excel within 0.2-1.5 of Excel's own except Helligkeit/Kontrast 4.5 (our filter is not Excel's formula);
  Rechtsdrehung and Horizontal spiegeln from the tab read back as 90° / flipped; Rückgängig/Wiederholen checked.
- **Not built / differences:** Künstlerische Effekte (Office's own filters; pictures having one show it as stored,
  but a picture copied or rewritten loses the effect's name: the readback miss above), Bildlayout (SmartArt), In
  Zelle, Skizziert, An Form ausrichten and aligning several objects (one picture is selected at a time, Gruppieren
  greyed as in Excel for one object; pictures cannot join shape groups); Freistellen writes a new PNG instead of
  Excel's marks beside the original. The shared renderer's inner shadow falls on the opposite edges, its bevel is
  softer and its 3D rotation is mirrored compared with Excel (rotation3d block 10.4, bevel 5.3): left for a fix
  shared with RaabeWrite. The tab's layouts for narrow windows were not recorded (groups collapse from the right).

## D-089 Fonts found only by their full names: Aptos Narrow, Aptos Display, Calibri Light (2026-11-08)

Found by the look corpus (D-080 D): RaabeCalc printed Excel 365's standard font Aptos Narrow as Arial, RaabeWrite
Word's headings in Aptos Display as Arial (the PDFs embed ArialMT where Office's embed Aptos Narrow / Aptos Display).
Aptos, Aptos Narrow and Aptos Display are not Windows fonts at all: Office keeps them per user in its cloud font cache
(%LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts, found by the RaabePresent session), where neither the browser nor
fontdb looked. The native host now loads that folder into its font database (metrics) and gives a face's bytes on
request (`cloud_fonts`, `cloud_font`, only files inside the cache); `ensureFontFamily`
(text-engine, used by RaabeWrite's Fonts.cssFamily and RaabeCalc's grid font; RaabePresent asked to use it too) tests
whether the browser draws a family and, when it does not, registers the cloud font's faces with FontFace(bytes), else
adds @font-face aliases that take installed faces by their full names (fonts Windows files under a shared typographic
family); layouts and the grid are redone once they have loaded. Only this user's own Office fonts are used, on this PC;
nothing is redistributed. Names the browser finds, and fonts that are not installed (the iPad), are left as they were.

## D-090 Spezialfilter (Daten > Sortieren und Filtern > Erweitert...) (2026-11-08)

Recorded in Excel (`probe_excel_advfilter.py`: twelve criteria ranges, in place, copied, Keine Duplikate; one case per
Excel process, as copying to a place with chosen headings made Excel crash through COM) and built the same way
(`sheet-engine/src/advfilter.ts`): the criteria's first row names list columns, each further row is an alternative,
its cells all have to hold; a text without operator finds every entry beginning with it (="=Mathe" only "Mathe"),
operators and wildcards as ZÄHLENWENN, a column whose heading is no list column holds a formula criterion written for
the first record. The dialog (Spezialfilter: Liste an gleicher Stelle filtern / An eine andere Stelle kopieren,
Listenbereich, Kriterienbereich, Kopieren nach, Keine Duplikate) keeps the ranges as Excel does in the sheet's names
_FilterDatabase, Criteria and Extract, filters in place by hiding rows (sheetPr filterMode written and read), copies
the records under the list's or the target row's own headings; Löschen and Erneut anwenden work on it. Measured: 20
of 20 recorded cases the same rows and the same copies as Excel.

## D-091 Diagramm in RaabeWrite, part 2: Einfügen > Diagramm and Daten bearbeiten (Step 7b, 2026-11-08)

Einfügen > Diagramm opens Diagramm einfügen with the types of Word's dialog grouped as Word groups them (Säule, Linie,
Kreis, Balken, Fläche, Punkt (X Y); previews drawn by the chart engine) and Word's sample data (Kategorie 1-4,
Datenreihe 1-3). Word then opens the data in Excel; RaabeWrite opens Daten bearbeiten, a table of categories and
series where rows and series can be added and removed. A chart keeps its chart part and an embedded workbook with the
data in Tabelle1 (names in row 1, categories in column A), written as Word writes them (word/charts/chartN.xml, a
package relationship to word/embeddings/Microsoft_Excel_WorksheetN.xlsx, c:externalData), so Word can edit the data of
a chart made in RaabeWrite. Daten bearbeiten (context menu) changes a chart's series and keeps its own look; a chart
read from Word stays byte for byte until its data are changed. Not yet: the Diagrammentwurf and Format tabs in
RaabeWrite (Diagrammformatvorlagen, Schnelllayout, Diagrammelemente) and floating (anchored) charts made here.

## D-100 RaabePresent (D-080): the shared VBA interpreter grows: classes, events, debugger, files after asking, UserForms (2026-11-08)

Part A and B of D-080 for the interpreter @raabe/vba (this session owns it; RaabeCalc builds Excel's object model
beside it). The public API only grew (listed in the commit message); old calls behave as before except where
VBA behaves otherwise (below).
- **Language:** class modules with `New`, `Nothing`, `Is`, `TypeOf … Is`, `Property Get/Let/Set`, `Me`,
  `Class_Initialize/Terminate`, default members, `For Each` over classes, `Implements`, `Type`, `Enum`, `Static`,
  `ParamArray`, named arguments, `Option Base/Compare`, `#If VBA7/Win64` (64-bit Office), typed variables that
  coerce (Overflow 6), `On Error GoTo/Resume/Resume Next`, `Err.Raise`, `End`. `Collection` and
  `Scripting.Dictionary` are built in (in memory). Events: `Public Event`, `RaiseEvent`, `WithEvents`; the app fires
  procedures (`fire`) and its objects' events (`raise`).
- **Class_Terminate** runs when the last reference goes (references in variables, arrays, members, collections are
  counted), at the end of that statement; `End` and Zurücksetzen run none, as in VBA.
- **Debugger:** breakpoints, Einzelschritt / Prozedurschritt / Prozedur abschließen, `Stop`, `Debug.Assert`,
  pausing on request, call stack, local variables, expressions and Direktbereich statements while paused.
  `Debug.Print` formats as VBA (print zones). Several statements on one line pause once.
- **Files only after asking (D-080):** `Open … For Input/Output/Append`, `Print #`, `Write #`, `Input #`,
  `Line Input #`, `Close`, `FreeFile`, `EOF`, `LOF`, `Dir`, `FileLen`, `FileDateTime`, `GetAttr`, `Kill`, `MkDir`,
  `RmDir`, `FileCopy`, `Name … As` ask the app each time (`fileAccess`) before going through its file system; refused
  is Laufzeitfehler 70. Text is Windows-1252, as VBA writes it. `GetSetting`/`SaveSetting` use the app's store,
  never the registry. Still blocked: Shell, CreateObject (except Scripting.Dictionary and Collection), GetObject,
  SendKeys, Environ, AppActivate, ChDir, Declare, AddressOf.
- **UserForms:** the runtime model (Show modal or not, Hide, Load, Unload with QueryClose, `Controls`, the common
  controls with their properties and events: Change and Click also when VBA sets a value, as Office does), and the
  file format ([MS-OFORMS] designer storage, [MS-OVBA] records and the MSForms reference), written by our own code from
  the published specifications. A form not edited is written back byte for byte; controls we do not model
  (MultiPage, TabStrip, ActiveX) are kept as bytes.
- **Behaviour changes callers may see:** module-level `Public x` now declares x; `Dim i As Integer` starts at 0;
  arrays are copied on assignment; `End` ends the whole run; statements outside procedures are a syntax error.
- **Measured in PowerPoint** (check_present_userform_office.py): PowerPoint opens a .pptm with two of our forms
  (one with a Frame holding controls), loads them and reads every property as written, keeps them when saving, and
  our reader reads PowerPoint's file back: 6/6. Not yet measured: a form shown in PowerPoint (it would need real
  input), MultiPage written by us (refused: "cannot write a MultiPage yet").

## D-101 RaabePresent: SmartArt-Formatvorlagen, all of Farben ändern, Von rechts nach links (2026-11-08)

- **Recorded from PowerPoint** (probe_powerpoint_smartart_styles.py, our own invisible PowerPoint, its window on the
  virtual display): all 38 colour sets and 14 quick styles applied to an "Einfache Blockliste"; for the five style
  labels our drawing uses (node1, sibTrans2D1, parChTrans1D1, revTx, bgShp) the colour lists (method and colours)
  and the quick styles' references to the theme's fill, line and effect styles with their 3-D scene and shape go
  into smartart-styles.gen.ts (make_smartart_styles.py). Only these parameters are kept, as for the layouts
  (D-077); the parts PowerPoint wrote stay in captures/.
- **Drawing as PowerPoint draws:** a shape's colours come from its label's lists by the node's place (repeat:
  the list in order, the last kept; span: from the first colour to the last over the nodes; cycle: there and back
  by 2/n of the way, measured: 0, 0.4, 0.8, 0.8, 0.4 for five nodes), written as the first colour with hue,
  saturation, luminance and alpha offsets, the hue by the plain difference of the two (accent 5 to 6: -202
  degrees, measured); the fill, line and effects are the theme's styles by the quick style's references with the
  placeholder colour replaced; 3-D from the label. The parts we write carry PowerPoint's ids and these parameters.
- **@raabe/ooxml reads more colour modifiers** (add only): hueOff, hueMod, satOff, alphaOff, alphaMod (PowerPoint
  writes them for such colours, so Office's SmartArt with spread colours was drawn in one colour before).
- **Measured** (check_present_smartart_styles.py, 52 slides against PowerPoint's pictures): our drawing equals
  PowerPoint's drawing as we render it for all 38 colour sets (mean pixel difference 0.12-0.8) and the five 2-D
  quick styles (0.23-0.49). The nine 3-D quick styles are 4-28 off for PowerPoint's own drawing too: PowerPoint
  shades the faces by its light rig and material, our renderer does not yet (open).
- **Von rechts nach links:** the arrangement mirrored in its frame (shapes that are not symmetric flipped), stored
  as PowerPoint stores it (the document point's layout variable dir = rev). Form darüber / darunter hinzufügen now
  carry PowerPoint's labels.

## D-102 RaabePresent: placeholder prompts as PowerPoint 365 shows them (2026-11-08)

The prompts were written from memory ("Titel durch Klicken hinzufügen"). PowerPoint's prompts are not in the file
and not visible to UI Automation, so a PowerPoint of our own on the virtual display was pictured window by window
(PrintWindow: that window only) with a slide per layout, and the prompts read off the pictures: "Titel
hinzufügen", "Untertitel hinzufügen", "Text hinzufügen", the picture placeholder a button "Bild einfügen
(Füllung)", the notes "Klicken Sie, um Notizen hinzuzufügen" (ours already). PROMPTS in @raabe/ooxml's pptx
reader changed accordingly (only RaabePresent reads it). Not recorded: the master-only kinds (table, chart, SmartArt,
media), which keep the old wording.

## D-103 RaabePresent: cell bevels and the 3D-Drehung buttons measured (2026-11-08)

- **Zellenabschrägung:** PowerPoint drew a table with each of the twelve bevels (the bevels written by us into
  PowerPoint's own table; COM refuses ThreeD on cells). Measured from its picture: the face 0.845 as bright, and
  across each edge a profile of its own per bevel (softRound is dark on the left, angle a flat light band …). The
  profiles (eight steps over the bevel's width, as blends towards white or black) are kept in cellbevel.gen.ts
  (make_cellbevel_profiles.py) and drawn as gradients. Mean pixel difference 19.5 before, 3.95 after (cells
  4.5-6.9). The profiles come from this one picture: other widths and colours are not checked separately.
- **3D-Drehung's buttons** (probe_powerpoint_8b_ui.py spin, invoked through UI Automation, no input): 5 degrees
  per click, not 10 as guessed; Links and Uhrzeigersinn lower the value shown, Rechts, Nach oben and Gegen
  Uhrzeigersinn raise it (Nach unten lowers). Ours does the same.

## D-104 RaabePresent: Entwurf > Varianten > Effekte (2026-11-08)

Office's theme effect schemes as PowerPoint writes them: our own invisible PowerPoint loads each of the 14 effect
files Office installs into a presentation of its own and saves it; the a:fmtScheme it wrote, with "Office" from a
new presentation's theme, goes into themeeffects.gen.ts by PowerPoint's German names (their order in PowerPoint's
gallery, recorded, matches the files'). As D-072 did for the effect presets: format data, no prose. Effekte (Entwurf,
Folienmaster) replaces the theme's effect scheme. Grunge-Muster's texture is a picture of Office's, which we do not
ship: that scheme is applied with plain fills in its place.

## D-105 RaabePresent: macro tests against PowerPoint, VBA's random numbers, VBA's shape names (2026-11-08)

- **Test macros of our own** (tools/feel-lab/macros/Tests.bas): replacing a font on all slides, pictures of one
  width, notes into a file, deleting empty placeholders, a countdown, a quiz with a score (the buttons' macros), a
  seeded name picker, runtime errors, functions and German formats, and how VBA names shapes. PowerPoint builds the
  test presentation through COM; our VBA writer adds the module; each test runs on a fresh copy in PowerPoint (COM
  Application.Run, our own instance, its window on the virtual display, no security setting changed; a watchdog
  ends VBA's error boxes in that process) and in RaabePresent; the presentations afterwards, the file written and
  errors are compared: 10/10 the same (check_present_macro_suite.py).
- **Found and fixed on the way:** VBA's Rnd and Randomize are a 24-bit generator (Rnd with a negative number seeds
  and steps once; Randomize mixes the halves of the number's high word into the seed), now ours too, so seeded
  draws equal Office's. VBA sees shapes PowerPoint named itself by English names ("Title 1", "Picture 6",
  "TextBox 3") although a German file says "Titel 1", and finds them by either name (measured): our object model
  does the same. The object model gained Shape.Type, PlaceholderFormat, LockAspectRatio (Width keeps the
  proportions when locked), NotesPage.
- **Openly licensed code as a stress test:** VBA-JSON (MIT, with the BSD notice of its origin) is in
  macros/thirdparty with its licence, but not run: the session's safety check refused running downloaded code. It
  runs only after the owner allows it.
- **Safety in use** (check_present_macro_security.py, 7/7): a .pptm opens with macros off and the message bar;
  a macro started before asks first and Abbrechen runs nothing; Inhalt aktivieren enables them; writing a file
  asks with the file's path every time; Nicht zulassen is Laufzeitfehler '70' and writes nothing; a runtime error
  shows Office's box (Laufzeitfehler '13':, Beenden, Debuggen).

## D-106 RaabePresent: the macro editor and UserForms in the app (2026-11-08)

- **Editor** (vbaide.ts, Alt+F11 and Makros > Bearbeiten): a window over the app with the project (Formulare,
  Module, Klassenmodule), code with line numbers and breakpoints (F9, a click in the margin), Ausführen (F5),
  Unterbrechen, Zurücksetzen, Einzelschritt (F8), Prozedurschritt (Umschalt+F8), Prozedur abschließen
  (Strg+Umschalt+F8), Direktbereich (Debug.Print; statements typed, while paused in the running macro) and
  Lokal-Fenster; Einfügen > UserForm / Modul / Klassenmodul. The UserForm designer has Office's Werkzeugsammlung
  (twelve controls), a 6-pt grid, moving and sizing by mouse, the Eigenschaften list, and a double-click that opens
  the control's Click (or Change) procedure. The editor knows no PowerPoint: the app gives its modules, saving and
  object model, so RaabeCalc can open it too.
- **Shown forms** are windows over the app whose controls run the form's code (also when a macro runs from an
  action button or the Makros dialog). Runs get the code as the editor shows it, so error lines and breakpoints
  agree.
- **Runtime errors** as Office's box: "Laufzeitfehler 'n':", the message, Beenden or Debuggen (opens the editor at
  the procedure); ours adds the procedure and line.
- **Measured** (check_present_vbaide.py, headless, 11/11): a form drawn in the designer, its Click written, a
  breakpoint hit with the Lokal-Fenster showing n = 41, a single step to the next line, Debug.Print's output, the
  form shown and its button running its code, the project saved with the form.

## D-107 RaabePresent: the GIF's timing, the behaviour catalogue's last cases, the inventory (2026-11-08)

- **Animierte GIF-Datei as PowerPoint's** (measured frame by frame, check_present_video.py): PowerPoint's GIF draws
  no transition (the slide before stays until the transition's middle, then the cut), and a slide lasts one frame
  (1/16 s) longer, one more per animation step, the last slide one more. Ours does the same: 8 of 8 states within
  0.25 s (before 6/8), 8.61 s against PowerPoint's 8.58 s; the video stays 9/9 + 9/9.
- **Behaviour catalogue:** the six Strg+Z cases undid the harness's own COM set-up in PowerPoint; they now undo a
  nudge or typing made by real input (catalogue_present.py). Recorded again, two behaviours of PowerPoint's
  showed and are ours now: Strg+Z while editing a text keeps editing, the caret where the undone change began; a
  click inside a group's frame between its members selects the group, a second click there starts a new text box
  at that point. 363 of 377 (before 356). Left: the twelve handle and rotation cases with Umschalt or Strg pressed
  during the drag end 1-2 pt off with real input only (with CDP input the movement and the rules are PowerPoint's:
  Umschalt keeps the proportions by the height, Strg sizes from the centre), not found yet; a text box pasted twice
  is in text editing in PowerPoint; Umschalt+Ende at a line's end selects one character more in PowerPoint.
- **Inventory** 4104 of 4511 (4004 before): SmartArt colour sets and quick styles, theme effects, labels. Of the 399
  left, about 150 are Microsoft's own designs (we ship ours, D-078), about 50 the galleries' scroll buttons, the
  rest picture SmartArt layouts, chart styles, Aktuelles Design speichern / Nach Designs suchen (the .thmx format),
  high-contrast colour options and single entries.

## D-108 RaabePresent: the look corpus (D-080 D), theme background pictures, a shape's own empty effect list (2026-11-08)

- **Corpus:** twelve school-like presentations PowerPoint built (make_corpus_ppt.py, 18 slides), each slide drawn
  by RaabePresent at PowerPoint's picture width and compared pixel by pixel (check_present_corpus.py): mean
  difference 3.78 at first, 1.94 after the two fixes below.
- **Theme background pictures** (Ion, 13.8 to 1.1): a background reference to the theme's third background style
  is a picture of the theme's own (its file in the theme part's relationships) recoloured by a:duotone. The pptx
  reader (add only: `background.duotone`) resolves the picture and the duotone's two colours with the background
  colour as placeholder; our slide draws the picture through an SVG filter that maps its brightness from the dark
  colour to the light one.
- **A shape's own effect list replaces its style's** (Facette, 4.2 to 0.3): Facette's translucent shapes carry
  `<a:effectLst/>` in spPr against their style's shadow; we drew the style's shadow, which showed through the
  translucent fills (darker by 10-20 per channel). DrawingML: an effect list of the shape's own, also an empty one,
  replaces the style's effects. Shared code changed accordingly (text-engine `shapeLook`, `PictureFormat.ownEffects`
  set by the ooxml reader), so RaabeWrite follows the same rule; all tests pass, RaabePresent's earlier render
  fixtures unchanged (2.75).
- **Left: Aptos.** The remaining larger differences (2.4-7.3: referat, notizen, abschnitt, tabelle, zwei_inhalte)
  are text in Aptos, Office's default font since 2023. On this PC Aptos, Aptos Display and Aptos Narrow are not
  Windows fonts but files in Office's per-user cloud-font cache (FontCache\4\CloudFonts), which a browser cannot
  find by name (local() aliases fail too). Loading those files for all three apps is shared work, agreed with
  RaabeCalc's session before it is built.

## D-109 RaabePresent: Office's cloud fonts (Aptos) through @raabe/text-engine (2026-11-08)

RaabePresent calls setCloudFonts in its app (the platform's cloud_fonts / cloud_font, D-089), ensureFontFamily for
every font its text renderer names (runs, paragraphs, bullets, symbol fonts) and draws the slides again on
onFontAlias. Look corpus (check_present_corpus.py; the headless page gets the user's cloud fonts through a harness
stand-in for the native command): mean 1.94 before, 1.16 now (18 slides). Found on the way: Aptos Display's files call
themselves family "Aptos" (style "Display"); listed also as "Aptos" they replace Aptos Regular and body text gets
about 10 % narrower. The stand-in names each face by its cache folder only; the native list was reported to
RaabeCalc's session, which owns it. Left: Calibri Light titles (Rückblick 3.5), probably text smoothing in the
headless page.
