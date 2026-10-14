# Architecture

> Status: editing engine decided: **controlled editing** (D-017). This file describes the
> intended structure and is updated as decisions land.

## Principles

1. **Feel first.** Every keystroke is handled synchronously on the UI thread with a
   minimal path to the screen. Heavy work (spell check, file parsing, layout of
   off-screen pages, translation) runs in workers and never blocks input.
2. **OOXML is the document.** The in-memory model maps cleanly to .docx/.xlsx/.pptx.
   Unknown XML parts and elements are kept verbatim and written back on save (D-004).
3. **Offline by construction.** The app has no network code paths except an explicit
   "check for updates" action (D-005, D-015).
4. **One code base, two platforms.** Windows (WebView2) and iPadOS (WKWebView) share
   everything except a thin platform layer: file dialogs, printing, fonts, OS voices.
5. **Touch is a first-class input.** Every interaction has a pointer *and* a touch path;
   hit targets follow the ribbon's touch mode sizes.

## Technology

| Layer | Choice | Why |
|---|---|---|
| App shell | Tauri 2 (Rust) | ~10 MB installers, fast start, WebView2 on Windows, iOS support |
| UI | TypeScript, no heavy framework in the editing hot path | Predictable latency; the ribbon and dialogs may use a small component layer |
| File formats | Own OOXML reader/writer (zip + XML), in TypeScript | Full control over fidelity and preservation of unknown parts |
| Spell check | Hunspell dictionaries (igerman98, en_US/en_GB) checked in Rust (`spellbook`) | Offline, open licences, fast (D-031) |
| Formula editor | MathLive (MIT) for editing; own LaTeX ↔ OMML conversion and MathML display in Cambria Math | D-036 |
| Charts | Own renderer (shared by RaabeCalc and RaabePresent) | Must match Excel chart look |
| Fonts | System fonts, metrics read from the font files (Rust fontdb + ttf-parser); Carlito/Caladea/Liberation as metric-compatible fallbacks | D-011, D-027 |

## Repository layout

npm workspaces; TypeScript packages are consumed as source (no build step between them).

```
apps/
  raabe-write/        RaabeWrite: app.ts (documents, files, backstage pages), command-list.ts
                      (every command with Word's label), ribbon-def.ts (tabs and groups),
                      features/ (lists, tables, page setup, headers/footers, pictures, find,
                      spelling, references, review, math: their commands, galleries, dialogs
                      and contextual tabs), templates.ts (Raabeschule templates, D-035),
                      src-tauri/ (Rust: file access, AutoRecover, PDF, single instance, fonts,
                      spell checking)
  raabe-calc/         later
  raabe-present/      later
packages/
  text-engine/        model.ts (paragraphs, runs, objects, tables, sections), styles.ts (style
                      sheet and resolution), fonts.ts (metrics, measuring), layout.ts (line
                      breaking, D-025), flow.ts (pages, D-029), numbering.ts (list labels),
                      doc.ts (stories and edits), edits.ts (list, table, section operations),
                      school.ts (notes, tables of contents, comments), revisions.ts (tracked
                      changes, D-034), autocorrect.ts (AutoKorrektur, D-042), images.ts
                      (picture sizes and resolution),
                      html.ts (clipboard HTML as Word writes it, D-043), view.ts (DOM
                      rendering, caret, selection),
                      editor.ts (input and commands, D-026), probe.ts (feel-lab)
  shell/              commands, ribbon, keytips, title bar (QAT, document tabs, search),
                      backstage, status bar, dialogs, popups, platform layer, theme.css,
                      icons.gen.ts (generated from Fluent UI System Icons, MIT)
  ooxml/              .docx reading and writing; everything not edited is kept verbatim (D-024);
                      math.ts: LaTeX ↔ OMML and MathML for equations (D-036)
  (planned) spell/, charts/
branding/             logos (excluded from the GPL, see branding/README.md)
docs/                 decisions, scope, architecture, feel metrics, roadmap, reports
prototypes/           Step 0 feel prototypes (kept for reference; the harness uses the real app)
tools/feel-lab/       measurement harness against Microsoft Word, Word UI capture, smoke test
```

## Platform layer

`packages/shell/src/platform.ts` is the only place that touches the operating system: file
dialogs, reading and writing files, AutoRecover storage, clipboard (on Windows the whole
"HTML Format" document through `src-tauri/src/clipboard.rs`, D-043), PDF export, window
controls and files opened from Explorer. `platform-tauri.ts` implements it for Tauri
(Windows now, iPadOS later); `platform-browser.ts` is a development fallback for
`npm run dev:write` in a normal browser.

## Documents (Steps 2-4)

- A document holds stories (main text, headers, footers, notes, comments), styles, list
  definitions, sections and comment data (`doc.ts`, D-029). A paragraph is text plus runs of
  character properties (`CharProps`), paragraph properties (`ParaProps`) and the formatting of
  its paragraph mark; it may belong to a table cell, end a section, carry floating pictures
  and markers (bookmarks, comment ranges, field and hyperlink edges). Objects in the text
  (fields, pictures, note references, equations) are U+FFFC with their original XML.
  Unknown property elements are kept as raw XML and written back in schema order.
- Formatting is resolved through the style sheet (document defaults → table style →
  paragraph style chain → list level → character style → direct formatting), cached per
  object.
- `layout.ts` breaks one paragraph into lines for a width: characters are measured with the
  canvas at a large size (design widths), vertical metrics come from the platform
  (`fontMetrics`). `flow.ts` places paragraphs on pages (sections, columns, tables, headers
  and footers, footnotes, floating pictures); every line gets page coordinates and a frame
  (column, cell, header …) used for drawing and hit testing. `view.ts` renders lines as
  absolutely positioned spans whose baselines are placed exactly, plus borders, shading and
  pictures.
- `packages/ooxml`: `read.ts` (all parts into the model), `write.ts` (model back into the
  package; unchanged parts stay byte for byte), `package.ts` (relationships, content types,
  media), `props.ts` (run/paragraph properties, borders, shading), `numbering.ts`.

## Testing

- `npm test`: unit tests (Vitest + jsdom) for the text engine and .docx package, including a
  document saved by Word itself.
- `tools/feel-lab/smoke_write.py`: drives the real app off screen through the DevTools
  protocol (no real input) and checks files with Word: typing, undo, styles, formatting
  shortcuts, save, Word round trip of formatting and of a table kept verbatim, keytips,
  search, backstage, PDF, AutoRecover after a crash, close prompts, ribbon collapsing.
- `tools/feel-lab/smoke_step3.py`: Step 3 in the real app (lists, tables, sections, headers
  and footers, pictures, find and replace, spelling, clipboard HTML, print preview), files
  checked with Word.
- `tools/feel-lab/smoke_step4.py`: Step 4 in the real app (notes, table of contents,
  comments, tracked changes, equations, templates); Word checks the files and updates the
  table of contents to compare page numbers. `capture_step4.py` takes screenshots (local).
- `tools/feel-lab/docx_word_check.py`: edits and saves Word-made documents with rich content
  and has Word confirm nothing was lost.
- `tools/feel-lab/run.py`: the feel metrics against Word (docs/FEEL.md), including layout
  fidelity on .docx files (`docx`) and formatting behaviour (`formatting`).
- `python tools/check-chars.py`: finds invisible characters in sources (see CLAUDE.md).

## Editing engine candidates (Step 0)

| | A: native editing | B: controlled editing |
|---|---|---|
| Input and caret | Browser `contenteditable` | Hidden textarea for input; our own caret and selection |
| Layout | Browser text layout | Our own line breaking and pagination; lines rendered as positioned DOM text |
| Strengths | Native caret, IME, accessibility and spell-check squiggles for free | Full control: Word's exact selection behaviour, exact line and page breaks, pagination |
| Risks | Hard to paginate; browser behaviour differs from Word and is hard to override | Everything must be built ourselves (IME, accessibility); more code |

Canvas rendering (the approach ONLYOFFICE takes) is deliberately not a candidate: canvas
text loses the system's subpixel text rendering and accessibility, which contributes to
the "web editor" feel described in D-001.
