# Step 7 report: RaabeCalc, overnight run

Date: 2026-10-27/28 · Measured with `tools/feel-lab` against Excel 365 (German, 16.0) on the
development PC · Raw data in `tools/feel-lab/results/` (2026-10-27)

The owner asked for the whole of RaabeCalc to be built in one run, without questions in
between, with the PC to the harness for the night (D-061). The method was the one of Step 4b:
Excel's behaviour recorded first in Excel instances of our own (invisible through COM, or
visible with real input), then built the same way and compared automatically. Four packages ran
in parallel (formula engine, number formats and typed input, .xlsx, charts) while the app, the
grid and the controls were built and measured against Excel.

## What was built

- **A. Foundation and Excel reference:** `feellab/excel.py` (Excel of our own, ends only its own
  process), Excel's ribbon, window and backstage recorded (`capture_excel_ui.py`), the native side
  shared with RaabeWrite moved into `crates/raabe-host` (RaabeWrite's smoke tests unchanged:
  41/41, 28/28, 43/43), `apps/raabe-calc` with Excel's tabs, groups, keytips, tooltip titles and
  shortcuts (`probe_excel_tooltips.py`, own descriptions).
- **B. Grid and .xlsx:** one canvas at the device's pixels with Excel's measured look (gridlines,
  headers, selection frame, fill handle), Excel's pointers, column widths computed with the
  display's digit width as Excel does at 150 % (D-062), frozen panes, merged cells, borders,
  fills, rotated text, conditional formats, AutoFilter buttons, notes, error indicators.
  .xlsx: everything not edited stays as Excel wrote it; 0 differences in Excel over 17 files.
- **C. Formula engine:** German names and separators, Excel 365 semantics with spilling arrays,
  the dependency graph (one typed value recalculates only what depends on it), 4229/4234
  recorded Excel results equal.
- **D. Excel's controls:** Eingeben/Bearbeiten/Zeigen, pointing at cells while typing, F4,
  AutoSumme, AutoVervollständigen, function AutoComplete, Strg+arrow jumps, Tab/Enter rules, the
  fill handle (series, Strg, double-click, dragging back), moving and copying by the border,
  header sizing and AutoAnpassen, the sheet tabs, the Namenfeld, the status bar statistics, the
  AutoFilter panel, context menus and ribbon menus with Excel's recorded labels, Excel's German
  shortcuts as recorded (D-063).
- **E. Charts:** the shared chart engine (`packages/charts`) with Excel's default look and axis
  scaling, charts in the grid (move, size, Diagrammelemente, Diagrammentwurf tab), written to and
  read from .xlsx.
- **F. Formats, output:** Zellen formatieren, number formats in German conventions, printing and
  PDF with Excel's default margins, Seite einrichten, CSV as German Excel.

- **G. After the catalogue was complete** (D-065), each recorded in or checked against Excel
  where it could be:
  - Datenüberprüfung (built by a parallel agent; Excel's verdict on 150 entries reproduced) with
    Ungültige Daten einkreisen; Blatt schützen and Arbeitsmappe schützen with Excel's password
    hash (Excel accepts our passwords and refuses wrong ones).
  - Formelüberwachung: Spur zum Vorgänger/Nachfolger, Formelauswertung, Fehlerüberprüfung.
  - Filters: Zahlen-, Text- and Farbfilter, Top 10, Durchschnitt (Excel's rows for 18 + 5
    criteria); Nach Farbe sortieren.
  - Printing: Drucktitel, Kopf- und Fußzeile, Seitenumbrüche, Umbruchvorschau with draggable
    breaks, pictures on the page.
  - Data tools: Schnellanalyse (Excel's formulas for 6 recorded Ergebnisse), Blitzvorschau (13
    of Excel's FlashFill results), Gruppieren and Teilergebnis (11 functions equal cell by cell),
    Zielwertsuche (10 cases, same solutions), Symbolsätze, Sparklines, Berechnungsoptionen incl.
    Manuell.
  - Rechtschreibung (F7) with RaabeWrite's dictionaries; Kommentare (Excel 365's threads, Excel
    reads ours as threads); pictures shown, printed, inserted, moved and sized.
  - Controls: Umschalt+drag (7 recorded cases), the right-button drag menus (recorded), Excel's
    keyboard modes (Ende, F8, Umschalt+F8, Strg+Rücktaste), Alt+Unten, Inhalte auswählen,
    Notizen typed in place, Finanzmathematik.
  - An independent review of these additions at 04:00 found nine problems (protection gaps,
    threads and pictures not following inserted rows, a fill-step bug, pictures dropped on save
    …); all are fixed or turned out to be Excel's behaviour too (D-065 lists them).
  - Later, 79 more control cases (hidden rows and columns, multi-area Enter/Tab,
    Umschalt+Rücktaste, right-click on headers, Umschalt+fill handle, F4 in a typed range, Strg+Z
    and Alt+Unten while typing, merged cells, a column border dragged past its start, pointing
    into another sheet by clicking its tab as in =SUMME(Tabelle2!B2)) found sixteen behaviours
    that differed from Excel, among them no "Zirkelbezüge" in the status bar; all are fixed.

## Measured

| What | Result |
|---|---|
| Behaviour catalogue (real input, Excel vs RaabeCalc) | 266/266 cases |
| Formulas (Excel's results) | 4229/4234 (5 are Excel's own floating-point/iteration quirks) |
| Number formats and typed input | ≈ 66 000 recorded cases equal (≈ 210 listed differences) |
| .xlsx round trip checked by Excel | 0 differences, 0 repair prompts, 21 files (incl. validation, protection, threads, pictures, outline, icon sets) |
| Chart axis scaling | 145/145 axes, 1914/1914 interval limits, 28/28 histogram bins |
| Speed, 20 000 rows (`perf_calc.py`) | typed value 12 ms, scroll redraw 6,4 ms median / 21 ms worst, full recalculation 0,26 s |
| Printing (PDF from both) | pages and lines per page equal in 8/8 layouts (Drucktitel, Kopf-/Fußzeile, manual breaks); header within 0,14 pt, footer within 0,01 pt |
| Clipboard with Excel, both ways | values, number formats, bold equal |
| Templates opened in Excel | 3/3 without repair, results equal |
| Smoke test of the later features (`smoke_calc.py`) | Excel opens the workbook and sees every feature |
| Start / memory (`startup_calc.py`) | 573 ms vs Excel 960 ms / 362 MB vs Excel 238 MB |
| RaabeWrite after the shared changes | smoke tests 41/41, 28/28, 43/43 |
| Program size | 10,7 MB, installer 2,9 MB (the spelling dictionaries add 8,6 MB to the program) |

## Not done / known gaps

- Shapes in sheets (kept in the file, not shown), PivotTables, the
  Seitenlayout view, Datentabelle and Szenario-Manager: deferred, greyed with "Zurückgestellt".
- Diagramm in RaabeWrite (Step 7b) is the next step.
- The owner has not tried the app yet (planned after package D, now after the run).
- Not measured against Excel: the outline bar's sizes, Schnellanalyse's chart tiles, the
  positions of pictures on printed pages, the dialogs of Datenüberprüfung and Teilergebnis
  (labels from German Excel, not recorded).
- The scheduled task meant to resume work after a session limit was refused by the permission
  check; the run did not need it.
