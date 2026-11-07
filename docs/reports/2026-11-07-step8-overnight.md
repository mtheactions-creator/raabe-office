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

## Numbers (final run, 06:06-06:21, one build, commit 76a6c51)

| Measure | Result | Target |
|---|---|---|
| P1 behaviour catalogue (`catalogue_present.py`, the same real input in PowerPoint and RaabePresent) | run 1: 99/307 · run 2: 270/336 · run 3: 305/336 · with caret keys: 337/370 · with AutoKorrektur: 344/377 · rotation and drag start measured: 355/377 · final: **356/377** (94 %) | 100 % |
| P2 round trips with PowerPoint (`check_present_roundtrip.py`) | fixtures **22/22** (both designs), scenarios **54/54** the same (first count tonight 24; every feature added after midnight got a scenario PowerPoint checks) | all the same |
| P3 slide rendering against PowerPoint's PNG export | mean **2.75 (28 slides)** of 255 | below 3 |
| P4 line starts against PowerPoint (`check_present_lines.py`) | **127/127 (Aptos design 112/116)** (Office 2013-2022 design) | 100 % |
| P5 inventory coverage (`inventory_report_present.py`) | **3895 (86 %)** of 4511 controls and entries (first count tonight 1757). Measured in PowerPoint's default colours (see below) | everything but the cloud |
| Picture effects against PowerPoint's own pictures (`results/2026-11-07-present-picfx-compare.json`) | Sättigung 0.0-0.2 of 255, Temperatur below 1, Weichzeichnen 0.08, Schärfen 0.38; Hintergrund entfernen the same pixels as PowerPoint on the probe's picture; artistic effects 0.5-48 (our own drawing) | for looking at |
| Speed, 149 slides with pictures (dev build) | open 70 ms, change slide 23 ms, duplicate 20 slides 89 ms | for looking at |
| Unit tests, type check | **623** passed (5 skipped), type check clean, all apps | green |

### What the catalogue taught (D-072)
- What a click hits: the letters of a text start editing and a drag from them selects text, by whole words
  once it leaves the first one; beside the letters the frame is selected and dragged; an unfilled text box is
  passed through where it has no text; outlines and lines are hit within 4 pixels.
- Frame keys: arrows move one pixel, Umschalt+arrows size by 10 % about the centre, Strg+Umschalt by 1 %,
  Strg+D and repeated pasting offset 12 pt, Strg+Umschalt+G/H group and ungroup (Strg+G does nothing).
- Text boxes take their text's height at once; AutoKorrektur as PowerPoint (7 of 7 cases).

- The rotation handle (probe_powerpoint_rothandle.py): PowerPoint's sits 31 screen pixels above the frame,
  turns by the handle's centre and with Umschalt snaps to the last 15-degree step passed. The "94 %" drags:
  PowerPoint measures a drag that selects the shape from the pointer's first movement (the harness's first of
  20 steps is lost); frames selected already and size handles move by the whole way. Both are copied now.

### Remaining differences, and why
Of the 21 cases that differ: 10 size-handle cases 0.5-2 pt apart (with Umschalt or Strg, where PowerPoint
counts a little less of the pointer's way; and the Strg copy of a text box, which PowerPoint places 6 pt
higher) and one edge drag 1.8 pt apart; 6 Strg+Z cases where PowerPoint's undo includes the harness's own COM
set-up (a harness matter); the text box rotation 1.3 degrees apart; Umschalt+Ende (PowerPoint's COM counts
the paragraph mark); a pasted text box that PowerPoint puts into text editing; the click on a group member
(the harness's group set-up differs between the two sides).

### The inventory, and how it is measured
PowerPoint was recorded in its default design (Office, Aptos colours: "Dunkelblaugrün, Akzent 1"). Our
untitled presentation starts in Office 2013-2022, so colour entries had other names. The harness now applies
the Office colour set before reading our controls; with that, and with the palettes now showing the design's
colours (a real fix, below), coverage rose from 65 % to 80 %. After 05:00 the gaps the report listed were
closed where they were not cloud or Microsoft content (86 %); 66 of those entries came from a counting fix
(two controls of one name, the parts of a split button), the rest from features. The rest is mostly what is left out on purpose:
SmartArt creation (21 + 38 + 43 + 14 entries), PowerPoint's own designs (32 + 39 × 3), theme effects (15 × 4),
transition sounds (21), Microsoft textures, chart and SmartArt styles, the Aufzeichnen tab and cloud features.

## What was built

- **Reference side (package A):** `feellab/powerpoint.py`, 22 fixtures built by PowerPoint, the full
  inventory (21 tabs, 439 controls, 5771 menu and gallery entries, 13 context menus, 114 dialogs) and
  `docs/POWERPOINT-INVENTORY.md`. New probes tonight: picture effects, colour temperature, background
  removal, the three Zoom kinds (real input under the lock: COM has no method for zooms), SmartArt.
- **The app (B-F):** .pptx kept as DOM parts with PowerPoint's inheritance; Normalansicht, Foliensortierung,
  Gliederung, Notizenseite, the three masters; designs (Office 2013-2022, Raabeschule, Klar, Tafel) with
  Office's colour and font sets; Übergänge (49), Animationen with the Animationsbereich, Trigger and paths;
  the show with PowerPoint's keys, pens, laser and Referentenansicht; comments, links, actions, video and
  audio, PDF and printing; Raabeschule templates.
- **Added after midnight (each written as PowerPoint writes it and checked by PowerPoint where it is a file
  matter):** AutoKorrektur; Datei > Optionen; picture effects (Künstlerische Effekte, Farbsättigung, Farbton,
  Schärfen/Weichzeichnen, Hintergrund entfernen with Freistellen); Benutzerdefinierter Pfad; Kurve, Freihandform and Skizze
  (they drew rectangles before); Tabelle zeichnen, Radierer and the table pen; Folienzoom, Abschnittszoom,
  Zusammenfassungszoom; Neues Fenster; .pptm (macros kept, never run), .ppsx, .potx; Formeffekte and
  Texteffekte complete (Abschrägung, 3D-Drehung, Transformieren); Interaktive Schaltflächen; palettes in the
  presentation's design colours (written as theme colours); Hintergrundformate; Fülleffekt (Keine Füllung,
  Bild, Farbverlauf); Zuschneiden menu (Auf Form zuschneiden, Seitenverhältnis, Ausfüllen, Einpassen);
  Formkontur > Pfeile; Schriftarten ersetzen; Farben anpassen; Folien wiederverwenden and Folien aus
  Gliederung; pasted slides keep their pictures; **equations edited in place with a caret** (RaabeWrite's
  editor model); SmartArt texts editable (Textbereich) and SmartArt In Formen / In Text konvertieren; notes
  keep their formatting when edited; Freihand zu Form and Lassoauswahl; Exportieren as PNG or JPEG; Benutzerdefinierte Bildschirmpräsentation;
  the Pipette; Formkontur > Skizziert (PowerPoint's format recorded by a probe); Tabellenentwurf > Effekte;
  Textfüllung (Keine Füllung, Farbverlauf); Schriftarten anpassen and Farben des Foliendesigns zurücksetzen; Bild
  und Größe zurücksetzen, Bild ändern > Aus Zwischenablage; SmartArt's context menu; the rotation handle and the
  start of a drag as measured in PowerPoint.
- **Two bugs found and fixed after 05:00:** Tabellenentwurf > Schattierung and Fülleffekt/Formkontur on a chart
  did nothing; both now work and a round-trip scenario checks each in PowerPoint.
- **Exe and installer:** `apps/raabe-present/src-tauri/target/release/raabe-present.exe` and the NSIS
  installer `target/release/bundle/nsis/RaabePresent_0.1.0_x64-setup.exe` (10.1 MB, built 08:00-08:03 from commit
  a62f33b; the code is that of 76a6c51, on which all numbers were measured).

## Decisions taken on my own (for the owner to confirm or overturn)
D-071 and D-072 hold them with reasons. In short:
1. The .pptx parts stay XML documents; everything not understood is written back unchanged.
2. Undo keeps the XML of the changed parts before and after each step.
3. Text is laid out by the browser with PowerPoint's measured line height.
4. Prompt texts of empty placeholders from memory of German PowerPoint (to be checked).
5. New shapes in PowerPoint 365's style; designs Office 2013-2022, Raabeschule, Klar, Tafel.
6. Comments in the classic format every PowerPoint reads (no account needed).
7. Rechtschreibung in a dialog box (RaabeWrite's), not a pane; Gliederungsansicht edits plain lines.
8. Keytips written joined as PowerPoint shows them ("ÄN").
9. Ink stored as freeform shapes; PowerPoint's ink kept and drawn from its fallback picture.
10. Colour and font sets from the files Office installs, with German PowerPoint's names.
11. Charts without an embedded workbook (PowerPoint shows them but cannot open their data in Excel).
12. Morphen as our own drawing.
13. Table cell ranges, Zellen verbinden/teilen as in D-072.
14. Handzettelmaster's Folien pro Seite is a view setting.
15. Equations: edited in place with RaabeWrite's equation editor model; the linear field stays for Linear.
16. The looks of the table, shape and picture styles are our reading of PowerPoint's galleries.
17. Where the catalogue could not measure (rotation handle) our handle follows the pointer.
18. Keep-awake: the tool to keep the PC awake was refused by the session's safety classifier; the PC stayed
    awake on its own settings.
19. The behaviour rules measured tonight (D-072).
20. Aktionseinstellungen, Fotoalbum, Thesaurus, Barrierefreiheit, Sprache, Optionen, Textbereich, Folien
    wiederverwenden are our own dialogs with PowerPoint's labels where recorded.
21. Artistic picture effects are our own drawing (PowerPoint's filters are not documented); the adjustments
    follow PowerPoint's measured pixels; our original picture stays PNG/JPEG behind PowerPoint's HD-Photo
    relationship (PowerPoint accepts it).
22. Hintergrund entfernen: our own region growing (the same pixels as PowerPoint on the probe's picture), with
    PowerPoint's Freistellen marks; no magenta overlay.
23. Tabelle zeichnen splits a cell in the middle (PowerPoint: where the line was drawn).
24. Zoom previews are drawn by us and not redrawn later; Zusammenfassungszoom's grid for other counts than
    the measured two follows the same spacing (assumption).
25. Neues Fenster is another tab of the window; Alle anordnen stays greyed.
26. Macros: kept and saved again, never run; saving them as a type without macros asks first.
27. SmartArt creation not built: PowerPoint needs its full layout definitions in the file (measured), which
    would mean copying Microsoft's 159 definitions. Existing SmartArt: texts editable, convertible.
28. Formeffekte presets from RaabeWrite's Word recordings (the same Office galleries, PowerPoint's names
    checked against the inventory); text warps are written but not drawn by us.
29. Interaktive Schaltflächen: our own icons (new presets in text-engine, added only).
30. Palettes show the design's colours; picks are written as theme colours (schemeClr, lumMod/lumOff). The
    shell's palette takes the design as an optional parameter (Word and Excel unchanged).
31. The inventory is measured with PowerPoint's default colour set applied to the untitled presentation.
32. Formkontur > Pfeile: eleven Pfeilarten by name from knowledge (the inventory did not open that submenu).
33. Exported slide pictures go beside the chosen name ("Name-Folie1.png"), not into a folder.
34. Notes are still edited as plain text, but unchanged paragraphs keep their formatting.
35. Pipette takes the colour of our own drawing of the slide; no live preview beside the pointer.
36. Skizziert is drawn by a displacement of our own (PowerPoint's sketch drawing is not documented).
37. Text gradients and table cell bevels are written as PowerPoint writes them but drawn plainly by us.
38. Drags: a shape body moves from the pointer's first movement (PowerPoint's way, measured); the rotation
    handle sits 31 px above the frame and Umschalt snaps to the last 15-degree step passed (measured).
39. The inventory counts two controls of one name together and the parts of a split button (66 entries).
40. Colours picked for a chart's Diagrammbereich are written as RGB (the chart model keeps RGB colours).

## Not done (open)
- SmartArt: creating it, layouts, styles, adding shapes (see 27).
- Formeloptionen and Freihandgleichung; drawing the text warps; the 3-D options panes.
- PowerPoint's own transition sounds and textures (Microsoft content); its designs with artwork (D-070).
- Second-monitor Referentenansicht (this PC has one screen; one screen works), Alle anordnen, Video and
  animated GIF export, Aufzeichnen (out of scope by D-070), running macros (never).
- Formatting notes in RaabePresent itself; the Pipette's live preview; drawing text gradients and cell bevels.

## How it was run
Real input only under `%TEMP%\raabe-real-input.lock`; PowerPoint started by our probes only when none was
running, presentations only in TEMP, PowerPoint quit only when nothing else was open. No sign-in, consent or
other dialog was clicked; the zoom probes ticked a check box in PowerPoint's own Zoom dialog on our own
presentation. The SmartArt probe needed a visible PowerPoint (SmartArtLayouts is not there without a window);
it took the lock as well. One finding worth keeping: the inventory probe had read the owner's recent file
names into the backstage page names; the sanitizer drops names that are file names before anything is
committed.
