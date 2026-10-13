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
| Spell check | Hunspell dictionaries (igerman98, en_US/en_GB) via WebAssembly | Offline, open licences |
| Formula editor | MathLive (MIT) for input, MathML ↔ OMML conversion | |
| Charts | Own renderer (shared by RaabeCalc and RaabePresent) | Must match Excel chart look |
| Fonts | System fonts, metrics read from the font files (Rust fontdb + ttf-parser); Carlito/Caladea/Liberation as metric-compatible fallbacks | D-011, D-027 |

## Repository layout

npm workspaces; TypeScript packages are consumed as source (no build step between them).

```
apps/
  raabe-write/        RaabeWrite: app.ts (documents, files, backstage pages), command-list.ts
                      (every command with Word's label), ribbon-def.ts (tabs and groups),
                      src-tauri/ (Rust: file access, AutoRecover, PDF, single instance)
  raabe-calc/         later
  raabe-present/      later
packages/
  text-engine/        model.ts (paragraphs, runs, objects), styles.ts (style sheet and
                      resolution), fonts.ts (metrics, measuring), layout.ts (line breaking and
                      pagination, D-025), doc.ts (edits), view.ts (DOM rendering, caret,
                      selection), editor.ts (input and commands, D-026), probe.ts (feel-lab)
  shell/              commands, ribbon, keytips, title bar (QAT, document tabs, search),
                      backstage, status bar, dialogs, popups, platform layer, theme.css,
                      icons.gen.ts (generated from Fluent UI System Icons, MIT)
  ooxml/              .docx reading and writing; everything not edited is kept verbatim (D-024)
  (planned) spell/, charts/
branding/             logos (excluded from the GPL, see branding/README.md)
docs/                 decisions, scope, architecture, feel metrics, roadmap, reports
prototypes/           Step 0 feel prototypes (kept for reference; the harness uses the real app)
tools/feel-lab/       measurement harness against Microsoft Word, Word UI capture, smoke test
```

## Platform layer

`packages/shell/src/platform.ts` is the only place that touches the operating system: file
dialogs, reading and writing files, AutoRecover storage, clipboard, PDF export, window
controls and files opened from Explorer. `platform-tauri.ts` implements it for Tauri
(Windows now, iPadOS later); `platform-browser.ts` is a development fallback for
`npm run dev:write` in a normal browser.

## Documents and formatting (Step 2)

- A paragraph is text plus runs of character properties (`CharProps`), paragraph properties
  (`ParaProps`) and the formatting of its paragraph mark. Objects the engine does not edit
  (fields, links, images, notes, revisions, equations) sit in the text as U+FFFC with their
  original XML; bookmarks and comment ranges are zero-width anchors; tables and other blocks
  are read-only placeholder paragraphs. Unknown property elements are kept as raw XML and
  written back in schema order.
- Formatting is resolved through the style sheet (document defaults → paragraph style chain
  → character style → direct formatting), cached per object.
- Layout measures characters with the canvas at a large size (design widths) and takes
  vertical metrics from the platform (`fontMetrics`), so it matches both Word and the DOM.
  Lines are rendered as absolutely positioned spans whose baselines are placed exactly.

## Testing

- `npm test`: unit tests (Vitest + jsdom) for the text engine and .docx package, including a
  document saved by Word itself.
- `tools/feel-lab/smoke_write.py`: drives the real app off screen through the DevTools
  protocol (no real input) and checks files with Word: typing, undo, styles, formatting
  shortcuts, save, Word round trip of formatting and of a table kept verbatim, keytips,
  search, backstage, PDF, AutoRecover after a crash, close prompts, ribbon collapsing.
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
