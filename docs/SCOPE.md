# Scope

What each app does, ordered by priority. **M** = first usable version (MVP),
**L** = later, **X** = deliberately out of scope. The guiding question is always:
*does a student (or a teacher correcting student work) actually use this?*

## Shared shell (all three apps)

| Feature | Prio | Notes |
|---|---|---|
| Ribbon with Microsoft's German tab and command names | M | Classic full ribbon; collapsible (Strg+F1) |
| Touch mode for the ribbon | M | Larger targets, required for iPad |
| Quick Access Toolbar (Speichern, Rückgängig, Wiederholen) | M | Fixed set at first, customisable later |
| Datei / backstage: Neu, Öffnen, Speichern unter, Als PDF exportieren, Drucken | M | No Konto, Teilen, Info pages |
| Save via native OS dialog, without Word's location picker | M | |
| AutoSave every 30 s + crash recovery | M | Local only |
| Document tabs in one window | M | D-010 |
| Recent files list, pinned files | M | |
| Templates gallery (Neu) incl. Raabeschule templates | M | Done in Step 4: Arbeitsblatt, Facharbeit, Versuchsprotokoll, Handout (D-035); school material can replace them |
| Search box "Suchen" (Alt+M) that finds and runs commands | M | Done in Step 1 |
| Keytips (Alt shortcuts) with Word's German letters | M | Done in Step 1 |
| Dark mode | L | After the pilot (agreed) |
| Spell check German + English (Hunspell dictionaries) | M | Done in Step 3 (D-031) |
| Find and replace | M | Done in Step 3: Navigationsbereich, Suchen und Ersetzen, Gehe zu |
| Zoom slider and view buttons in the status bar | M | |
| Print with preview | M | Done in Step 3 |
| Optionen dialog | M | A handful of settings only |
| Read aloud (Vorlesen) with offline OS voices | L | Helpful for students with dyslexia (LRS) |
| Translator (offline, Firefox Translations / Bergamot models) | L | de, en, fr, es; no Latin |
| Dictation (offline, whisper.cpp) | L | Optional download, needs good hardware |
| Grammar check | L | LanguageTool is heavy; evaluate later |
| Apple Pencil ink (Zeichnen tab) | L | Architecture must allow it from day one |
| Accounts, cloud, sharing, co-authoring, version history | X | D-005 |
| Copilot, Designer, online pictures, stock images | X | Replaced by templates and a bundled icon set |
| Macros / VBA, add-ins | X | |

## RaabeWrite (Word)

| Feature | Prio | Notes |
|---|---|---|
| Character formatting (font, size, bold/italic/underline, colour, highlight, super/subscript, strikethrough) | M | Done in Step 2, incl. Schriftart dialog, Format übertragen, Groß-/Kleinschreibung |
| Paragraph formatting (alignment, indents, line and paragraph spacing, borders, shading) | M | Done in Steps 2 and 3 (Absatz dialog, tabs, page-break options, Rahmen, Schattierung) |
| Bulleted and numbered lists, multilevel lists | M | Done in Step 3, incl. Word's auto-list on "1. " / "- " and heading numbering |
| Styles (Formatvorlagen): Standard, Überschrift 1-3, Titel, Zitat … | M | Style gallery in Start tab; done in Step 2 (Word 2016-2019 set) |
| Tables: insert, resize, merge/split, borders, shading, header row | M | Done in Step 3 (Tabellenentwurf, Layout) |
| Images: insert, resize, rotate, crop, text wrapping (inline, square, tight, behind, in front) | M | Done in Step 3 except rotate and crop |
| Shapes and text boxes | M | Basic set |
| Header, footer, page numbers | M | Done in Step 3 |
| Page setup: margins, orientation, size, columns, page/section breaks | M | Done in Step 3 |
| Footnotes | M | Done in Steps 3 and 4: foot- and endnotes, Nächste Fußnote |
| Table of contents (automatic, from headings) | M | Done in Step 4: Automatische Tabelle 1/2, Text hinzufügen, Tabelle aktualisieren (D-037) |
| Cover page / title page | M | Via templates (Facharbeit, Step 4) |
| Comments | M | Done in Step 4: cards beside the page, replies, erledigt, Kommentare pane (D-034) |
| Track changes (Änderungen nachverfolgen) | M | Done in Step 4: insertions, deletions, paragraph marks, Annehmen/Ablehnen, Markup modes (D-034) |
| Formula editor | M | Done in Step 4: Formel tab with all Strukturen, MathLive input, Cambria Math display, OMML in .docx (D-036) |
| Word count, character count | M | |
| AutoCorrect (basic), smart quotes „…“ | M | German quotation marks |
| Symbols and special characters | M | |
| Hyperlinks | M | |
| Captions, cross-references | L | |
| Citations and bibliography | L | Schools teach this manually; maybe a simple version |
| Mailings (Sendungen), envelopes, labels | X | |
| SmartArt | X | Preserved in files, shown as image placeholder |
| Fields beyond page number / TOC / date | X | Preserved |
| Entwurf tab: Designs, Farben, Schriftarten, Absatzabstand | L | Theme support |
| Entwurf tab: Wasserzeichen, Seitenfarbe, Seitenränder | L | |
| Word count exactly like Word (standalone dashes are not words) | M | Done in Step 1 |

## RaabeCalc (Excel)

| Feature | Prio | Notes |
|---|---|---|
| Grid, multiple sheets, rename/insert/delete/move sheets | M | |
| Cell formatting: font, fill, borders, alignment, wrap, merge | M | |
| Number formats: number, currency, percent, date, time, fraction, scientific | M | German locale (Komma as decimal separator) |
| Fill handle / AutoFill (series, formulas) | M | |
| Formulas with German function names | M | See function list below |
| Relative/absolute references (F4), named ranges | M | |
| Sort, filter (AutoFilter) | M | |
| Freeze panes, hide rows/columns, column width/row height | M | |
| Charts: column, bar, line, pie, scatter (with trendline + equation + R²), area | M | Maths and physics |
| Conditional formatting (basic: highlight rules, data bars, colour scales) | M | |
| Print area, fit to page, gridlines/headings on print | M | |
| CSV import/export | M | |
| Comments / notes on cells | L | |
| Data validation (dropdown lists) | L | |
| PivotTables | L | Maybe a simplified version |
| What-if analysis, Solver | X | |
| Power Query, Power Pivot, data model | X | |
| Macros / VBA | X | |

**MVP function list** (German name / English name):
SUMME/SUM, MITTELWERT/AVERAGE, MIN, MAX, ANZAHL/COUNT, ANZAHL2/COUNTA,
ANZAHLLEEREZELLEN/COUNTBLANK, WENN/IF, UND/AND, ODER/OR, NICHT/NOT, WENNFEHLER/IFERROR,
ZÄHLENWENN/COUNTIF, ZÄHLENWENNS/COUNTIFS, SUMMEWENN/SUMIF, SUMMEWENNS/SUMIFS,
MITTELWERTWENN/AVERAGEIF, SVERWEIS/VLOOKUP, WVERWEIS/HLOOKUP, XVERWEIS/XLOOKUP,
INDEX, VERGLEICH/MATCH, RUNDEN/ROUND, AUFRUNDEN/ROUNDUP, ABRUNDEN/ROUNDDOWN,
GANZZAHL/INT, ABS, WURZEL/SQRT, POTENZ/POWER, REST/MOD, PI, EXP, LN, LOG, LOG10,
SIN, COS, TAN, ARCSIN/ASIN, ARCCOS/ACOS, ARCTAN/ATAN, BOGENMASS/RADIANS, GRAD/DEGREES,
MEDIAN, MODUS.EINF/MODE.SNGL, STABW.S/STDEV.S, STABW.N/STDEV.P, VAR.S, VAR.P,
RANG.GLEICH/RANK.EQ, KGRÖSSTE/LARGE, KKLEINSTE/SMALL, KORREL/CORREL, STEIGUNG/SLOPE,
ACHSENABSCHNITT/INTERCEPT, BESTIMMTHEITSMASS/RSQ, ZUFALLSZAHL/RAND,
ZUFALLSBEREICH/RANDBETWEEN, HEUTE/TODAY, JETZT/NOW, DATUM/DATE, JAHR/YEAR, MONAT/MONTH,
TAG/DAY, WOCHENTAG/WEEKDAY, TEXT, VERKETTEN/CONCATENATE, TEXTKETTE/CONCAT, LINKS/LEFT,
RECHTS/RIGHT, TEIL/MID, LÄNGE/LEN, GROSS/UPPER, KLEIN/LOWER, GLÄTTEN/TRIM,
FINDEN/FIND, SUCHEN/SEARCH, WECHSELN/SUBSTITUTE, KOMBINATIONEN/COMBIN, FAKULTÄT/FACT,
BINOM.VERT/BINOM.DIST, NORM.VERT/NORM.DIST.

## RaabePresent (PowerPoint)

| Feature | Prio | Notes |
|---|---|---|
| Slides with layouts (Titel, Titel und Inhalt, Zwei Inhalte, Nur Titel, Leer …) | M | |
| Designs (themes) with colour variants | M | A curated set, incl. a Raabeschule design |
| Text boxes, bullet levels, formatting | M | |
| Images, shapes, tables, charts (from the RaabeCalc chart engine) | M | |
| Align, distribute, group, arrange (bring to front …) | M | Smart guides while dragging |
| Slide sorter, sections | M | |
| Transitions (a curated set of ~10) | M | |
| Animations: entrance, emphasis, exit (curated set), animation order | M | |
| Speaker notes, presenter view, slide show with pen/laser pointer | M | |
| Export as PDF | M | |
| Video and audio insertion (local files) | L | |
| Morph transition | L | Students like it; technically demanding |
| Slide master editing | L | |
| Recording, live captions, rehearse timings | X | |
| SmartArt, 3D models | X | Preserved in files |
