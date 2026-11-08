# Step 10 report, RaabePresent's part (D-080), 2026-11-08

D-080 gave RaabePresent's session the open items of Step 8b, the shared VBA interpreter (UserForms, file access
after asking, the gaps D-079 listed) and the macro tests against PowerPoint. The owner left the PC to the run;
open questions were decided here and are listed below. All PowerPoint work ran in PowerPoint processes of our own
(invisible or on the virtual display), only while no other POWERPNT.EXE ran; no Office security setting was
changed; real input only under the lock file, in short runs.

## Numbers

| What | Result |
|---|---|
| GIF against PowerPoint's own GIF (check_present_video.py) | 8/8 states within 0.25 s (6/8 before); 8.61 s against 8.58 s |
| Video against PowerPoint's video | 9/9 + 9/9 (unchanged) |
| Behaviour catalogue (catalogue_present.py, real input) | 363/377 (356 before) |
| SmartArt colour sets, our drawing against PowerPoint's (check_present_smartart_styles.py) | 38/38 within 0.12-0.8 mean pixel difference |
| SmartArt quick styles | 2-D 5/5 within 0.23-0.49; 3-D 9 at 4-28, PowerPoint's own drawing as far off in our renderer |
| Cell bevels against PowerPoint (check_present_cellbevel.py) | mean 3.95 (19.5 before), cells 4.5-6.9 |
| Macro tests, the same as PowerPoint (check_present_macro_suite.py) | 10/10 |
| Our UserForms in PowerPoint (check_present_userform_office.py) | 6/6 |
| Macro editor in use (check_present_vbaide.py) | 11/11 |
| Macro safety in use (check_present_macro_security.py) | 7/7 |
| Inventory coverage | 4104 of 4511 (4004 before) |
| Unit tests | 683 passed before the merge of main (vba 40) |

## What was built

1. **GIF timing (D-107).** PowerPoint's GIF draws no transition and gives each slide one frame more, one more per
   animation step and the last slide one more; ours does the same.
2. **Catalogue (D-107).** The Strg+Z cases undo real input now (they undid the harness's set-up in PowerPoint).
   Two behaviours recorded and built: Strg+Z keeps editing the text; a second click into a selected group's empty
   area starts a text box.
3. **SmartArt (D-101).** All 38 colour sets and 14 quick styles, recorded from PowerPoint and drawn as PowerPoint
   draws them (spread and cycled colours with PowerPoint's own steps), Von rechts nach links, PowerPoint's labels
   for Form darüber / darunter hinzufügen. @raabe/ooxml reads five more colour modifiers (add only).
4. **Prompts (D-102)** as PowerPoint 365 shows them: "Titel hinzufügen", "Untertitel hinzufügen", "Text
   hinzufügen", the picture placeholder's "Bild einfügen (Füllung)".
5. **Cell bevels and the 3D-Drehung buttons (D-103)** measured and built: the face's lighting and each bevel's edge
   profile; 5 degrees per click, PowerPoint's directions.
6. **Entwurf > Effekte (D-104)** with Office's 15 theme effect schemes as PowerPoint writes them.
7. **The VBA interpreter (D-100).** Classes with New, events, WithEvents, the debugger, files and folders only after
   asking every time, settings in the app, Collection and Scripting.Dictionary, UserForms (runtime and Office's file
   format), VBA's own random generator, named arguments to the app's objects (asked for by RaabeCalc's session).
8. **The macro editor (D-106)** over the app: project, code with breakpoints, single steps, Lokal-Fenster,
   Direktbereich, the UserForm designer; shown forms run their code; Office's runtime error box.
9. **Macro tests (D-105)** of our own against PowerPoint, through COM on both sides' fresh copies.

## Decisions taken on the owner's behalf

1. PowerPoint's recorded parameters (SmartArt colour lists and quick style references, the theme effect schemes,
   the cell bevel profiles) are kept as generated tables, as D-072/D-077 did; the parts PowerPoint wrote stay in
   captures. Grunge-Muster is applied without its texture picture (Office's picture is not shipped).
2. The GIF follows PowerPoint's timing rule found from one test deck (one frame more per slide and step); other
   decks may show more of PowerPoint's rule.
3. VBA sees shapes PowerPoint named itself in English ("Title 1") and finds them by either name, as measured in
   PowerPoint; our object model does the same.
4. Class_Terminate runs when the last reference goes; a UserForm is a class module with a form (MS-OVBA); form
   values not measured in Office (see D-100 and the interpreter's comments) were chosen from the specification.
5. File access asks every time with the file's path; the safer answer (Nicht zulassen) is the default button.
6. The editor is written so that RaabeCalc can open it (it takes the app's modules, saving and object model).

## Not done (open)

- **Openly licensed VBA as a stress test:** VBA-JSON (MIT) is in tools/feel-lab/macros/thirdparty with its licence,
  but the session's safety check refused running downloaded code, so it was not run, neither in PowerPoint nor in
  RaabePresent. It needs the owner's permission (a permission rule for running it, or the owner's go in chat).
- Catalogue: the twelve handle and rotation cases with a modifier pressed during a real-input drag (1-2 pt), the
  text box pasted twice, Umschalt+Ende at a line's end.
- 3-D lighting of faces (SmartArt 3-D quick styles; generally shapes with bevels): PowerPoint shades by light rig
  and material, ours does not.
- Inventory: picture SmartArt layouts, chart styles (shared with RaabeCalc's chart work), Aktuelles Design
  speichern / Nach Designs suchen (.thmx), high-contrast colour options, gallery scroll buttons, Microsoft's designs
  (on purpose, D-078).
- Prompts of the master-only placeholder kinds; the cell bevel profiles checked at other widths and colours.
- Macros: a UserForm shown in PowerPoint (needs real input), MultiPage written by us, Binary/Random files.
- The history rewrite of D-080 (force push) is the other session's and waits for the owner.

## For RaabeCalc's session (named arguments and the rest asked for)

In @raabe/vba now: named arguments for the app's objects (`get/call(name, args, named?)`, lower-case keys; with
`argNames` they are also put into place), built-ins' own named parameters (MsgBox Prompt:=, Title:= …), Optional with
defaults, ParamArray, IsMissing, Enum, Type … End Type, GoTo label, On Error GoTo label / Resume / Resume Next,
Err.Number/Description/Raise/Clear. Application.Run belongs to the app's object model: the app calls
`runtime.run(name, args)` from its Application object (RaabePresent's model does not have it yet). The editor
(apps/raabe-present/src/vbaide.ts) can be opened with RaabeCalc's modules and host.
