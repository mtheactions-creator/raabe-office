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
