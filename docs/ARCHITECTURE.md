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
| Fonts | System fonts; Carlito/Caladea/Liberation bundled as metric-compatible fallbacks | D-011 |

## Repository layout (planned)

```
apps/
  raabe-write/        Tauri app: Word
  raabe-calc/         Tauri app: Excel
  raabe-present/      Tauri app: PowerPoint
packages/
  shell/              ribbon, backstage, dialogs, tabs, status bar, theming
  ooxml/              zip + XML reading/writing, preservation of unknown parts
  text-engine/        text model, line layout, pagination, selection, input
  spell/              Hunspell (WASM) + dictionaries
  charts/             chart model and renderer
  icons/              icon set
branding/             logos (excluded from the GPL, see branding/README.md)
docs/                 decisions, scope, architecture, feel metrics, roadmap
prototypes/           Step 0 feel prototypes (throwaway)
tools/feel-lab/       measurement harness against Microsoft Word
```

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
