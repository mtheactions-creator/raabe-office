# Step 8 report: RaabePresent, overnight run

Date: 2026-11-06/07 · Measured with `tools/feel-lab` against PowerPoint 365 (German, 16.0) on the
development PC · Raw data in `tools/feel-lab/results/` (2026-11-06 and 2026-11-07) · Branch
`claude/step8-raabe-present-190dd0` (not merged; the owner merges)

The owner agreed on 2026-11-06 that Step 8 runs overnight without questions in between (D-070):
RaabePresent should feel exactly like German PowerPoint 365, everything except the cloud, measured
against PowerPoint. The method was Step 4b's and Step 7's: PowerPoint's behaviour recorded first in a
PowerPoint of our own (invisible through COM, or with real input under the lock), then built the same
way and compared automatically. PowerPoint runs as one process per user, so every probe first checked
that no POWERPNT.EXE was running and worked only on presentations it created (CLAUDE.md, D-070).

## Numbers

| Measure | Result | Target |
|---|---|---|
| P1 behaviour catalogue (`catalogue_present.py`, same real input in PowerPoint and RaabePresent) | run 1: 99/307 · run 2: 270/336 · run 3: 305/336 · with 34 caret-key cases: **337/370** (91 %) | 100 % |
| P2 round trips with PowerPoint (`check_present_roundtrip.py`) | fixtures 11/11 the same, scenarios 27/27 the same (incl. table merge/split, table style, equation, motion path, click action, text box height) | all the same |
| P3 slide rendering against PowerPoint's PNG export (28 slides) | mean 2.75 of 255 (equations now drawn from Office Math: 1.13) | below 3 |
| P4 line starts against PowerPoint (`check_present_lines.py`) | 127/127 (Office 2013-2022 design); Aptos design 112/116 (Aptos is not installed on this PC) | 100 % |
| P5 inventory coverage (`inventory_report_present.py`) | 2767 of 4503 controls and entries (61 %; first count tonight 1757). Names of colour entries depend on the design (PowerPoint was recorded in its Aptos design, we start in Office 2013-2022), so many colour entries count as missing though they are there | everything but the cloud |
| Unit tests, type check | 607 passed (5 skipped), type check clean, all apps | green |
| UI screenshots (`compare_present_ui.py`, `compare_present_ui_context.py`) | all main tabs and the format and table tabs side by side with PowerPoint (kept in `captures/`, not committed); differences fixed, see below | for looking at |

### What the catalogue taught (D-072)
The first comparison showed how differently PowerPoint decides what a click hits, and most of the
work after it went there:
- The letters of a text start editing and a drag from them selects text, by whole words once it leaves
  the first one; beside the letters on a line the frame is selected and dragged; an unfilled text box or
  placeholder is passed through where it has no text; outlines and lines are hit within 4 pixels; an
  empty placeholder starts editing on any click.
- Frame keys: arrows move one pixel (the grid only with Am Raster ausrichten), Umschalt+arrows size by
  10 % about the centre, Strg+Umschalt by 1 %, Strg+D and repeated pasting offset 12 pt, Strg+Umschalt+G
  groups and Strg+Umschalt+H ungroups (Strg+G does nothing in German PowerPoint), typing on a selected
  shape adds to the end of its text, Strg+Rücktaste goes into the text at its end.
- Text boxes take their text's height at once (a cache bug kept them at their drawn height before);
  their side handles are hidden when they are short, n and s change nothing on them.

### Remaining differences, and why
Of the 33 cases that differ: 6 rotation cases where the harness does not grab PowerPoint's rotation handle
(PowerPoint reports no rotation); 6 Strg+Z cases where PowerPoint's undo includes the harness's own COM
set-up; 10 drags and handle drags 1.5-3.4 pt apart (PowerPoint's ellipse drags end 3.4 pt short of the
pointer, its textbox drags do not; not understood yet); 2 handle drags with Umschalt on a text box; the
group-member case (PowerPoint made a new text box where the harness clicked twice into the group's empty
middle); Umschalt+Ende (PowerPoint's COM counts the paragraph mark); a pasted text box that PowerPoint
puts into text editing.

### What the UI screenshots showed
Side by side with PowerPoint (same window size), the ribbon was brought to PowerPoint's arrangement:
designs, effects, shape, table and picture styles as galleries in the ribbon; labelled stacks where
PowerPoint has them; no "..." on ribbon buttons. They also showed commands that had no implementation and
stayed greyed: Übergänge > Dauer, Bei Mausklick, Nach, Auf alle anwenden, Sound; Animationen > Dauer and
Verzögerung; several dropdowns (Layout, Abschnitt, Markieren, Farben, Schriftarten, Diagramm). All work now.

## What was built tonight

- **Reference side (package A):** `feellab/powerpoint.py` (a PowerPoint of our own, ends only its own
  process), 22 fixtures built by PowerPoint with JSON and 56 PNGs, layouts and AutoShapes recorded, the
  full inventory (21 tabs, 439 controls, 5771 menu and gallery entries, 13 context menus, 114 dialogs and
  panes, 4 backstage pages) and its report `docs/POWERPOINT-INVENTORY.md`.
- **The app (B-F):** .pptx kept as DOM parts with PowerPoint's inheritance; Normalansicht with the slide
  panel, placeholders, text editing, shapes, pictures, tables (cell ranges, Zellen verbinden/teilen,
  Tabellenformate), charts, SmartArt and equations drawn; Foliensortierung, Gliederung, Notizenseite,
  Folienmaster, Notizenmaster, Handzettelmaster; designs (Office 2013-2022, Raabeschule, Klar, Tafel) with
  Office's 24 colour and 24 font sets; Übergänge (49, Morphen as our own drawing), Animationen with the
  Animationsbereich; the show with PowerPoint's keys, pens, laser, blank screens and Referentenansicht;
  Kommentare, Links, Kopf- und Fußzeile, Zeichnen, Video and Audio, PDF and printing with handouts;
  Raabeschule templates Referat and Unterrichtsstunde.
- **Galleries as PowerPoint's:** Schnellformatvorlagen (77), Tabellenformate (74), Bildformatvorlagen (28),
  WordArt (20), picture Farbe and Korrekturen, Textkontur and Formkontur with Stärke and Striche, Rahmen,
  Texteffekte, Textrichtung, Zellenbegrenzungen, the Formel tab with its structures and symbols,
  Einfügen > Formel with the built-in equations.
- **More of PowerPoint tonight:** Animationspfade, Übergänge > Sound, Einfügen > Aktion, Fotoalbum,
  Thesaurus (Umschalt+F7), Sprache, Barrierefreiheit überprüfen, Lineal, Graustufe/Schwarzweiß,
  Fenster wechseln, Handzettelmaster, table cell ranges with Zellen verbinden/teilen, Morphen.
- **Exe and installer:** `apps/raabe-present/src-tauri/target/release/raabe-present.exe` and the NSIS
  installer in `target/release/bundle/nsis/` (built at the end of the run).

## Decisions taken on my own (for the owner to confirm or overturn)
D-071 and D-072 hold them with reasons. In short:
1. The .pptx parts stay XML documents; everything not understood is written back unchanged.
2. Undo keeps the XML of the changed slides before and after each step.
3. Text is laid out by the browser with PowerPoint's measured line height (1.2 × size × spacing).
4. Prompt texts of empty placeholders from memory of German PowerPoint (to be checked).
5. New shapes in PowerPoint 365's style; designs Office 2013-2022, Raabeschule, Klar, Tafel.
6. Comments in the classic format every PowerPoint reads (no account needed), not the threaded one.
7. Rechtschreibung in a dialog box (RaabeWrite's), not a pane; Gliederungsansicht edits plain lines.
8. Keytips written joined as PowerPoint shows them ("ÄN").
9. Ink stored as freeform shapes; PowerPoint's ink kept and drawn from its fallback picture.
10. Colour and font sets from the files Office installs, with German PowerPoint's names.
11. Charts without an embedded workbook (PowerPoint shows them but cannot open their data in Excel).
12. Morphen as our own drawing (same-named shapes glide, others fade).
13. Table cell ranges, Zellen verbinden/teilen as described in D-072.
14. Handzettelmaster's Folien pro Seite is a view setting, not saved; orientation shared with the notes.
15. Equations edited in PowerPoint's linear format in a box (no caret inside an equation yet); the Formel
    tab adds structures at the end of the selected equation.
16. Tabellenformate other than PowerPoint's default carry ids of our own (PowerPoint lists them as custom
    styles); the look of the 74 table styles, 77 shape styles and 28 picture styles is our reading of
    PowerPoint's galleries; picture styles without 3-D.
17. Where the catalogue could not measure (rotation handle, see above) our handle follows the pointer.
18. Keep-awake: the tool to keep the PC awake was refused by the session's safety classifier; the PC
    stayed awake on its own settings.
19. The behaviour rules measured tonight (D-072): what a click hits, word-wise drag selection, frame and
    caret keys, Strg+Umschalt+G/H for grouping, 12 pt offsets. Where PowerPoint was not measured I chose
    the same as the measured neighbour (Pfeil unten on the last line stays like Pfeil oben on the first).
20. Aktionseinstellungen, Fotoalbum, Thesaurus, Barrierefreiheit and Sprache are our own dialogs with
    PowerPoint's labels where recorded; Aktion's labels are from knowledge of PowerPoint, not recorded.

## Not done (open)
- SmartArt: creating and editing (In SmartArt konvertieren, SmartArt-Design); existing SmartArt is drawn.
- Editing inside an equation with a caret (PowerPoint's in-place editor); the Formel tab's Text,
  Formeloptionen and Freihandgleichung.
- Benutzerdefinierter Pfad (drawing a motion path), Künstlerische Effekte, Bild > Farbsättigung and Farbton,
  Bildkorrektur Schärfen/Weichzeichnen, 3-D effects (Abschrägung, 3D-Drehung, Transformieren),
  Designer-like features (cloud).
- PowerPoint's own transition sounds (Applaus ...: Microsoft content); [Ohne Sound], [Vorherigen Sound
  anhalten] and one's own WAV file work.
- Trigger, Animation übertragen, Folienzoom (Einfügen > Zoom), Tabelle zeichnen/Radierer, Diagrammelement
  hinzufügen, Bilder komprimieren, Hintergrund entfernen, SmartArt creation, Neues Fenster/Alle anordnen.
- Second-monitor Referentenansicht (one screen works), Aufzeichnen (out of scope by D-070), macros (last).
- The PowerPoint designs with artwork (not copied, D-070).

## How it was run
Real input only under `%TEMP%\raabe-real-input.lock`; PowerPoint started by our probes only when none was
running, presentations only in TEMP, PowerPoint quit only when nothing else was open. One finding worth
keeping: the inventory probe had read the owner's recent file names into the backstage page names; the
sanitizer now drops names that are file names (keys included) before anything is committed, and the
committed inventory was cleaned and checked.
