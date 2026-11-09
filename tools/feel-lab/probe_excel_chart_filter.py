"""How German Excel 365 saves what the chart buttons change (D-115, D-116), COM only, our own invisible Excel:
a series and a category hidden by Diagrammfilter (Series.IsFiltered, ChartGroup.FullCategoryCollection(i).IsFiltered),
Ausgeblendete und leere Zellen (DisplayBlanksAs, PlotVisibleOnly), the Datentabelle, the title overlaid, and the
Diagrammformatvorlagen 1-14 of a column, line, bar and pie chart (Chart.ChartStyle). Each case is one sheet of
one workbook saved to TEMP (excel-chart-filter.xlsx); its chart XML is copied to
packages/ooxml/test/fixtures/excel-chart-filter/ for the tests, and what COM reports goes to
results/2026-11-09-excel-chart-filter.json.

    .venv/Scripts/python.exe probe_excel_chart_filter.py
"""
from __future__ import annotations

import json
import re
import sys
import tempfile
import zipfile
from pathlib import Path

from feellab import uia, winapi
from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
ROOT = HERE.parents[1]
RESULT = HERE / "results" / "2026-11-09-excel-chart-filter.json"
FIX = ROOT / "packages" / "ooxml" / "test" / "fixtures" / "excel-chart-filter"
TMP = Path(tempfile.gettempdir())
DATA = [("Monat", "Nord", "Süd", "West"), ("Jan", 3, 4, 2), ("Feb", 5, 2, 3), ("Mrz", 2, 6, 4), ("Apr", 4, 3, 5)]
TYPES = {"column": 51, "line": 4, "bar": 57, "pie": 5, "lineMarkers": 65, "scatter": -4169, "area": 1}


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    x = ExcelTarget(visible=False)
    out: dict = {"cases": {}}
    path = TMP / "excel-chart-filter.xlsx"
    try:
        wb = x.wb
        first = wb.Worksheets(1)
        first.Name = "Daten"
        for r, row in enumerate(DATA, 1):
            for c, v in enumerate(row, 1):
                first.Cells(r, c).Value = v
        first.Range("C4").ClearContents()  # an empty cell for the blanks cases

        cases: list[tuple[str, str, object]] = []
        styled: list = []

        def case(name: str, kind: str = "column"):
            def deco(fn):
                cases.append((name, kind, fn))
                return fn
            return deco

        @case("plain")
        def _(ch):
            pass

        @case("series2-filtered")
        def _(ch):
            ch.FullSeriesCollection(2).IsFiltered = True

        @case("cat3-filtered")
        def _(ch):
            ch.ChartGroups(1).FullCategoryCollection(3).IsFiltered = True

        @case("series-and-cat-filtered")
        def _(ch):
            ch.FullSeriesCollection(1).IsFiltered = True
            ch.ChartGroups(1).FullCategoryCollection(1).IsFiltered = True

        @case("line-series2-filtered", "line")
        def _(ch):
            ch.FullSeriesCollection(2).IsFiltered = True

        @case("line-cat2-filtered", "line")
        def _(ch):
            ch.ChartGroups(1).FullCategoryCollection(2).IsFiltered = True

        @case("pie-cat2-filtered", "pie")
        def _(ch):
            ch.ChartGroups(1).FullCategoryCollection(2).IsFiltered = True

        @case("scatter-series1-filtered", "scatter")
        def _(ch):
            ch.FullSeriesCollection(1).IsFiltered = True

        @case("area-series1-filtered", "area")
        def _(ch):
            ch.FullSeriesCollection(1).IsFiltered = True

        @case("bar-series3-filtered", "bar")
        def _(ch):
            ch.FullSeriesCollection(3).IsFiltered = True

        @case("blanks-zero", "line")
        def _(ch):
            ch.DisplayBlanksAs = 2  # xlZero

        @case("blanks-span", "line")
        def _(ch):
            ch.DisplayBlanksAs = 3  # xlInterpolated

        @case("hidden-shown")
        def _(ch):
            ch.PlotVisibleOnly = False

        @case("datatable-keys")
        def _(ch):
            ch.HasDataTable = True
            ch.DataTable.ShowLegendKey = True

        @case("datatable-nokeys")
        def _(ch):
            ch.HasDataTable = True
            ch.DataTable.ShowLegendKey = False

        @case("title-overlay")
        def _(ch):
            ch.ChartTitle.IncludeInLayout = False

        @case("legend-top")
        def _(ch):
            ch.Legend.Position = -4160  # xlLegendPositionTop

        @case("errorbars-stderr")
        def _(ch):
            for i in range(1, ch.FullSeriesCollection().Count + 1):
                ch.FullSeriesCollection(i).ErrorBar(1, 1, 4)  # xlY, xlBoth, xlStError

        @case("labels-inend")
        def _(ch):
            ch.SetElement(203)  # msoElementDataLabelInsideEnd

        @case("axes-off")
        def _(ch):
            ch.HasAxis(1, 1)
            ch.SetElement(348)  # msoElementPrimaryCategoryAxisNone
            ch.SetElement(352)  # msoElementPrimaryValueAxisNone

        for kind in ("column", "line", "bar", "pie"):
            for n in range(1, 15):
                cases.append((f"style-{kind}-{n}", kind, n))

        for i, (name, kind, fn) in enumerate(cases):
            ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            ws.Name = f"C{i + 1}"
            sh = ws.Shapes.AddChart2(-1, TYPES[kind], 20, 20, 360, 216)
            ch = sh.Chart
            src = first.Range("A1:B5") if kind == "pie" else first.Range("A1:D5")
            ch.SetSourceData(src)
            info: dict = {"sheet": ws.Name, "kind": kind}
            try:
                if isinstance(fn, int):
                    info["gallery_index"] = fn
                    info["default_style"] = ch.ChartStyle
                    styled.append((info, sh, fn))
                else:
                    fn(ch)
            except Exception as ex:  # noqa: BLE001
                info["error"] = str(ex)[:200]
            out["cases"][name] = info
        # The gallery's n-th entry, applied through the ribbon (UI Automation, no input): Excel off screen, made
        # visible; the focus goes back to the window that had it after every step.
        if styled:
            fg0 = winapi.foreground()

            def give_back() -> None:
                if fg0 and winapi.foreground() != fg0:
                    try:
                        winapi.activate(fg0, 0)
                    except Exception:  # noqa: BLE001
                        pass
            winapi.place_quietly(x.hwnd, -4000, 0, 2400, 1300)
            x.app.Visible = True
            import time
            time.sleep(1.0)
            give_back()
            root = uia.element_for(x.hwnd)
            names_by_kind: dict = {}
            for info, sh, n in styled:
                wb.Worksheets(info["sheet"]).Activate()
                sh.Select()
                time.sleep(0.5)
                tab = uia.find(root, lambda d: d["type"] == "TabItem" and d["name"] == "Diagrammentwurf", 25)
                uia.select(tab)
                give_back()
                time.sleep(0.5)
                more = uia.find(root, lambda d: d["type"] == "Button" and d["name"] == "Schnellformatvorlagen", 30)
                uia.invoke(more)
                give_back()
                time.sleep(0.8)
                items = []
                f = uia._uia.GetRootElement().FindAll(uia.U.TreeScope_Children, uia._uia.CreatePropertyCondition(uia.U.UIA_ProcessIdPropertyId, x.pid))
                for i in range(f.Length):
                    w = f.GetElement(i)
                    allf = w.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
                    for j in range(allf.Length):
                        e = allf.GetElement(j)
                        try:
                            d = uia.describe(e)
                        except Exception:  # noqa: BLE001
                            continue
                        if d["type"] == "ListItem" and d["name"].startswith("Formatvorlage"):
                            items.append((e, d))
                seen = {}
                for e, d in items:
                    seen.setdefault(d["name"], e)
                names_by_kind[info["kind"]] = list(seen)
                target = seen.get(f"Formatvorlage {n}")
                if target is None:
                    info["error"] = "gallery item not found"
                    winapi.press("esc") if winapi.window_pid(winapi.foreground()) == x.pid else None
                    continue
                try:
                    uia.invoke(target)
                except Exception:  # noqa: BLE001
                    uia.select(target)
                give_back()
                time.sleep(0.5)
                info["style"] = sh.Chart.ChartStyle
                print(info["kind"], n, info["style"], flush=True)
            out["gallery"] = names_by_kind
            x.app.Visible = False
            give_back()
        # Read back what COM reports for the filter cases.
        for name, info in out["cases"].items():
            ch = wb.Worksheets(info["sheet"]).ChartObjects(1).Chart
            try:
                info["series"] = [{"name": ch.FullSeriesCollection(i).Name, "filtered": bool(ch.FullSeriesCollection(i).IsFiltered),
                                   "formula": ch.FullSeriesCollection(i).Formula}
                                  for i in range(1, ch.FullSeriesCollection().Count + 1)]
                fc = ch.ChartGroups(1).FullCategoryCollection()
                info["categories"] = [{"name": str(fc(i).Name), "filtered": bool(fc(i).IsFiltered)} for i in range(1, fc.Count + 1)]
                info["visible_series"] = ch.SeriesCollection().Count
                info["blanks"] = ch.DisplayBlanksAs
                info["plotVisibleOnly"] = bool(ch.PlotVisibleOnly)
            except Exception as ex:  # noqa: BLE001
                info["read_error"] = str(ex)[:200]
        x.app.DisplayAlerts = False
        wb.SaveAs(str(path), 51)
    finally:
        x.close()

    # Map sheets to chart parts and copy each chart XML.
    FIX.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path) as z:
        wbxml = z.read("xl/workbook.xml").decode("utf-8")
        rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        sheets = dict(re.findall(r'<sheet name="([^"]+)" sheetId="\d+" r:id="(rId\d+)"', wbxml))
        targets = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels)) | {k: v for v, k in re.findall(r'Target="([^"]+)"[^>]*Id="(rId\d+)"', rels)}
        for name, info in out["cases"].items():
            sheet = "xl/" + targets[sheets[info["sheet"]]].lstrip("/").replace("xl/", "")
            srels = z.read(sheet.replace("worksheets/", "worksheets/_rels/") + ".rels").decode("utf-8")
            drawing = re.search(r'Target="\.\./drawings/(drawing\d+\.xml)"', srels).group(1)
            drels = z.read(f"xl/drawings/_rels/{drawing}.rels").decode("utf-8")
            chart = re.search(r'Target="\.\./charts/(chart\d+\.xml)"', drels).group(1)
            xml = z.read(f"xl/charts/{chart}").decode("utf-8")
            (FIX / f"{name}.xml").write_text(xml, encoding="utf-8", newline="\n")
            st = re.search(r'Target="(style\d+\.xml)"', z.read(f"xl/charts/_rels/{chart}.rels").decode("utf-8"))
            if st:
                sxml = z.read(f"xl/charts/{st.group(1)}").decode("utf-8")
                m = re.search(r'<cs:chartStyle [^>]*id="(\d+)"', sxml)
                info["style_id"] = int(m.group(1)) if m else None
    RESULT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print("written", len(out["cases"]), "cases")


if __name__ == "__main__":
    main()
