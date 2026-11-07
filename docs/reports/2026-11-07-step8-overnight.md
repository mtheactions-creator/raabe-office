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
| P1 behaviour catalogue (`catalogue_present.py`, same real input in PowerPoint and RaabePresent) | run 1: 99/307 · run 2: 270/336 · run 3: 305/336 · with caret keys: 337/370 · with AutoKorrektur: **344/377** (91 %) | 100 % |
| P2 round trips with PowerPoint (`check_present_roundtrip.py`) | fixtures **11/11** the same, scenarios **@RT@** the same (first count tonight 24) | all the same |
| P3 slide rendering against PowerPoint's PNG export (28 slides) | mean 2.75 of 255 (equations drawn from Office Math: 1.13) | below 3 |
| P4 line starts against PowerPoint (`check_present_lines.py`) | 127/127 (Office 2013-2022 design); Aptos design 112/116 (Aptos is not installed on this PC) | 100 % |
| P5 inventory coverage (`inventory_report_present.py`) | **@INV@** of 4511 controls and entries (first count tonight 1757). Names of colour entries depend on the design (PowerPoint was recorded in its Aptos design, we start in Office 2013-2022), so many colour entries count as missing though they are there | everything but the cloud |
| Picture effects against PowerPoint's own rendered pictures (`results/2026-11-07-present-picfx-compare.json`) | Sättigung 0.0-0.2 of 255, Temperatur below 1, Weichzeichnen 0.08, Schärfen 0.38; Hintergrund entfernen: the same pixels as PowerPoint on the probe's picture; artistic effects 0.5-48 (our own drawing, noisy textures) | for looking at |
| Unit tests, type check | **@TESTS@** passed (5 skipped), type check clean, all apps | green |
| UI screenshots (`compare_present_ui.py`, `compare_present_ui_context.py`) | all main tabs and the format and table tabs side by side with PowerPoint (kept in `captures/`, not committed); differences fixed | for looking at |

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
- Text boxes take their text's height at once; their side handles are hidden when they are short.
- AutoKorrektur while typing, as PowerPoint: 7 of 7 the same, including the arrow PowerPoint writes as
  a Wingdings character.

### Remaining differences, and why
Of the 33 cases that differ: 6 rotation cases where the harness does not grab PowerPoint's rotation handle
(PowerPoint reports no rotation); 8 Strg+Z cases where PowerPoint's undo includes the harness's own COM
set-up; 12 drags and handle drags 1.5-3.4 pt apart (PowerPoint moves about 94 % of the pointer's way in
these cases; not understood yet, perhaps its grid); the group-member case (PowerPoint made a new text box
where the harness clicked twice into the group's empty middle); Umschalt+Ende (PowerPoint's COM counts the
paragraph mark); a pasted text box that PowerPoint puts into text editing.

## What was built

- **Reference side (package A):** `feellab/powerpoint.py` (a PowerPoint of our own, ends only its own
  process), 22 fixtures built by PowerPoint, layouts and AutoShapes recorded, the full inventory
  (21 tabs, 439 controls, 5771 menu and gallery entries, 13 context menus, 114 dialogs and panes, 4
  backstage pages) and its report `docs/POWERPOINT-INVENTORY.md`. Tonight's new probes:
  picture effects, colour temperature, background removal, Folienzoom/Abschnittszoom/Zusammenfassungszoom
  (real input under the lock: PowerPoint has no COM method for zooms) and SmartArt.
- **The app (B-F):** .pptx kept as DOM parts with PowerPoint's inheritance; Normalansicht with the slide
  panel, placeholders, text editing, shapes, pictures, tables, charts, SmartArt and equations drawn;
  Foliensortierung, Gliederung, Notizenseite, Folienmaster, Notizenmaster, Handzettelmaster; designs
  (Office 2013-2022, Raabeschule, Klar, Tafel) with Office's 24 colour and 24 font sets; Übergänge (49),
  Animationen with the Animationsbereich; the show with PowerPoint's keys, pens, laser, blank screens and
  Referentenansicht; Kommentare, Links, Kopf- und Fußzeile, Zeichnen, Video and Audio, PDF and printing.
- **Galleries as PowerPoint's:** Schnellformatvorlagen (77), Tabellenformate (74), Bildformatvorlagen (28),
  WordArt (20), Bild > Farbe and Korrekturen, Textkontur and Formkontur, Rahmen, Formeffekte and
  Texteffekte (all of PowerPoint's entries now, with Abschrägung, 3D-Drehung and Transformieren),
  Textrichtung, Zellenbegrenzungen, the Formel tab, Einfügen > Formel.
- **Added after midnight:** AutoKorrektur; Datei > Optionen; Künstlerische Effekte (22), Farbsättigung,
  Farbton, Schärfen/Weichzeichnen and Hintergrund entfernen in PowerPoint's own file format; Benutzer-
  definierter Pfad; Kurve, Freihandform and Skizze (they drew rectangles before); Tabelle zeichnen and
  Radierer; Folienzoom, Abschnittszoom, Zusammenfassungszoom; Neues Fenster; .pptm (macros kept, never
  run), .ppsx and .potx; Trigger, Animation übertragen, Diagrammelement hinzufügen, Form ändern, Bilder
  komprimieren, Fotoalbum, Aktion, Thesaurus, Sprache, Barrierefreiheit überprüfen.
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
15. Equations edited in PowerPoint's linear format in a box (no caret inside an equation yet).
16. Tabellenformate other than PowerPoint's default carry ids of our own; the looks of the table, shape and
    picture styles are our reading of PowerPoint's galleries.
17. Where the catalogue could not measure (rotation handle) our handle follows the pointer.
18. Keep-awake: the tool to keep the PC awake was refused by the session's safety classifier; the PC
    stayed awake on its own settings.
19. The behaviour rules measured tonight (D-072): what a click hits, word-wise drag selection, frame and
    caret keys, Strg+Umschalt+G/H for grouping, 12 pt offsets.
20. Aktionseinstellungen, Fotoalbum, Thesaurus, Barrierefreiheit, Sprache and Optionen are our own dialogs
    with PowerPoint's labels where recorded.
21. Artistic picture effects are our own drawing of each effect's look (PowerPoint's filters are not
    documented); the adjustments follow PowerPoint's measured pixels. Our original picture stays PNG/JPEG
    behind PowerPoint's HD-Photo relationship (PowerPoint accepts it).
22. Hintergrund entfernen is applied at once with PowerPoint's default rectangle; the Freistellen tab
    (marking areas) is not built.
23. Tabelle zeichnen splits a cell in the middle (PowerPoint: where the line was drawn).
24. Zoom previews are drawn by us when the zoom is made and not redrawn later; Zusammenfassungszoom's grid
    for other counts than the measured two follows the same spacing (assumption).
25. Neues Fenster is another tab of the same window; Alle anordnen stays greyed.
26. Macros: kept and saved again, never run; saving them as a type without macros asks first.
27. SmartArt creation not built: PowerPoint needs its full layout definitions in the file (measured), which
    would mean copying Microsoft's 159 definitions.
28. Formeffekte presets taken from RaabeWrite's Word recordings (the same Office galleries, PowerPoint's
    names checked against the inventory); text warps (Transformieren) are written but not drawn by us.

## Not done (open)
- SmartArt: creating and editing (see 27); existing SmartArt is drawn and kept.
- Editing inside an equation with a caret (PowerPoint's in-place editor); the Formel tab's Text,
  Formeloptionen and Freihandgleichung.
- Drawing the text warps (Transformieren); the 3-D options panes; Hintergrund entfernen's Freistellen tab.
- PowerPoint's own transition sounds (Microsoft content); [Ohne Sound], [Vorherigen Sound anhalten] and
  one's own WAV file work.
- Second-monitor Referentenansicht (this PC has one screen; one screen works), Alle anordnen,
  Aufzeichnen (out of scope by D-070), running macros (never).
- The PowerPoint designs with artwork (not copied, D-070).

## How it was run
Real input only under `%TEMP%\raabe-real-input.lock`; PowerPoint started by our probes only when none was
running, presentations only in TEMP, PowerPoint quit only when nothing else was open. No sign-in, consent
or other dialog was clicked; the zoom probes ticked a check box in PowerPoint's own Zoom dialog of our own
presentation. One finding worth keeping: the inventory probe had read the owner's recent file names into
the backstage page names; the sanitizer now drops names that are file names (keys included) before
anything is committed, and the committed inventory was cleaned and checked.
