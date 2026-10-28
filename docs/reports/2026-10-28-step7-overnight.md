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

- **G. After the catalogue was complete** (D-065): Datenüberprüfung (built by a parallel agent,
  Excel's verdict on 150 entries reproduced), Blatt schützen with Excel's password hash, Spur zum
  Vorgänger/Nachfolger, Formelauswertung, Fehlerüberprüfung, Drucktitel, Kopf- und Fußzeile,
  Ungültige Daten einkreisen, Zahlen- und Textfilter (Top 10, Durchschnitt), Schnellanalyse,
  Symbolsätze, Blitzvorschau, Umschalt+drag and the right-button drag menus, Inhalte auswählen,
  Notizen typed in place, the Dropdown-Auswahlliste (Alt+Unten), Finanzmathematik, Excel's
  keyboard modes (Ende, F8, Umschalt+F8, Strg+Rücktaste).

## Measured

| What | Result |
|---|---|
| Behaviour catalogue (real input, Excel vs RaabeCalc) | 142/142 cases |
| Formulas (Excel's results) | 4229/4234 (5 are Excel's own floating-point/iteration quirks) |
| Number formats and typed input | ≈ 66 000 recorded cases equal (≈ 210 listed differences) |
| .xlsx round trip checked by Excel | 0 differences, 0 repair prompts, 17 files |
| Chart axis scaling | 145/145 axes, 1914/1914 interval limits, 28/28 histogram bins |
| Speed, 20 000 rows (`perf_calc.py`) | typed value 11 ms, scroll redraw 6 ms median / 14 ms worst, full recalculation 0,26 s |
| Printing (PDF from both) | pages, lines per page, spacing equal in 3/3 layouts; positions within 0,8 pt |
| Clipboard with Excel, both ways | values, number formats, bold equal |
| Templates opened in Excel | 3/3 without repair, results equal |
| Start / memory | 579 ms vs Excel 936 ms / 365 MB vs Excel 240 MB |
| RaabeWrite after the crate move | smoke tests 41/41, 28/28, 43/43 |

## Not done / known gaps

- Outline (Gruppierung), comments (threaded), shapes and
  pictures in sheets (kept in the file, not shown), PivotTables, colour filters, Seitenlayout
  and Umbruchvorschau views: deferred, greyed with "Zurückgestellt".
- Diagramm in RaabeWrite (Step 7b) is the next step.
- The owner has not tried the app yet (planned after package D, now after the run).
