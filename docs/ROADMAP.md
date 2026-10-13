# Roadmap

Each step ends with something that can be checked. No step starts without the project
owner's explicit go-ahead.

| Step | Content | Status |
|---|---|---|
| 0 | Dev environment, repository, documentation, logos, measurement harness against Word, two feel prototypes, report and engine decision (D-014) | done 2026-10-13: engine decided (D-017), published on GitHub ([report](reports/2026-10-13-step0-feel.md)) |
| 1 | Shared shell: window, document tabs, ribbon (incl. touch mode), backstage, file dialogs, AutoSave/recovery, theming in Raabe red | done 2026-10-13 (installer 0.1.0); real-input re-check: F1 32 ms vs Word 34-36 ms, F3 49/50 |
| 2 | RaabeWrite core: text model, layout, pagination, selection, character/paragraph formatting, styles, undo/redo, .docx read/write for that subset | done 2026-10-13 ([report](reports/2026-10-13-step2-formatting.md)): layout identical to Word in four test documents, formatting behaviour vs Word's own commands, lossless .docx; real-input re-check done (F1 tail: Windows composition, see report) |
| 3 | RaabeWrite documents: lists, tables, images, header/footer, page numbers, page setup, spell check, find/replace, print, PDF | |
| 4 | RaabeWrite school features: footnotes, table of contents, comments, track changes, formula editor, Raabeschule templates | |
| 5 | Pilot: one class uses RaabeWrite in lessons; fix what real use reveals | |
| 6 | iPad build of RaabeWrite (touch polish, Files app integration, distribution) | |
| 7 | RaabeCalc | |
| 8 | RaabePresent | |
| 9 | Optional features: read aloud, translator, dictation, ink | |
