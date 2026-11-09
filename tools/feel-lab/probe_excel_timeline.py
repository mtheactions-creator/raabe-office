"""Records Excel's Zeitachse (timeline, D-121) through COM in our own invisible Excel 365: a table of dated marks, a
PivotTable over it and timelines on its date field, as Excel makes and saves them.

Cases: a timeline with Excel's defaults (Slicers.Add with only the sheet: size and place), one filtered to a range of
months (TimelineState.SetFilterDateRange), the four levels (Jahre, Quartale, Monate, Tage), the Anzeigen options
(Kopfzeile, Auswahlbezeichnung, Bildlaufleiste, Zeitebene), a style, a caption, two PivotTables connected to one
timeline, ClearDateFilter. For each: the timeline's and its cache's properties, the PivotTable's report, Excel's
rendering (Shape.CopyPicture, to TEMP/raabe-timeline/). Also the built-in timeline styles' elements.

    .venv/Scripts/python.exe probe_excel_timeline.py

Writes packages/ooxml/test/fixtures/excel-timeline.xlsx (anonymized, save path removed), excel-timeline.json and
results/2026-11-09-excel-timeline.json.
"""
from __future__ import annotations

import datetime as dt
import json
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path

from feellab.excel import ExcelTarget
from probe_excel_slicers import BORDERS, clean_fixture, rgb, safe

HERE = Path(__file__).parent
ROOT = HERE.parents[1]
FIX = ROOT / "packages" / "ooxml" / "test" / "fixtures"
TMP = Path(tempfile.gettempdir())
SHOTS = TMP / "raabe-timeline"
RESULT = HERE / "results" / "2026-11-09-excel-timeline.json"
NEUTRAL = "Raabe Office Test"
FACH = ["Mathe", "Deutsch", "Englisch", "Biologie"]
# Marks from Aug 2025 to Sep 2026 (school year and a bit), every ~17 days.
START = dt.date(2025, 8, 18)
DATA = [["Datum", "Fach", "Note"]] + [[START + dt.timedelta(days=17 * i), FACH[i % 4], 1 + (i * 7) % 6] for i in range(24)]
# XlTableStyleElementType values tried for the timeline elements (those that answer are kept).
ELEMENT_RANGE = range(0, 45)


def serial(d: dt.date) -> int:
    return (d - dt.date(1899, 12, 30)).days


def timeline_info(sl) -> dict:
    sc = sl.SlicerCache
    ts = safe(lambda: sc.TimelineState)
    vs = safe(lambda: sl.TimelineViewState)
    sh = sl.Shape
    d = lambda v: str(v)[:19] if v is not None else None  # noqa: E731
    return {
        "name": sl.Name, "caption": sl.Caption, "cache": sc.Name, "sourceName": safe(lambda: sc.SourceName),
        "cacheType": safe(lambda: sc.SlicerCacheType), "top": sl.Top, "left": sl.Left, "width": sl.Width, "height": sl.Height,
        "style": safe(lambda: sl.Style.Name if not isinstance(sl.Style, str) else sl.Style), "displayHeader": safe(lambda: sl.DisplayHeader),
        "state": None if ts is None else {
            "startDate": d(safe(lambda: ts.StartDate)), "endDate": d(safe(lambda: ts.EndDate)), "filterType": safe(lambda: ts.FilterType),
            "singleRange": safe(lambda: ts.SingleRangeFilterState), "value1": d(safe(lambda: ts.FilterValue1)), "value2": d(safe(lambda: ts.FilterValue2))},
        "view": None if vs is None else {
            "level": safe(lambda: vs.Level), "showHeader": safe(lambda: vs.ShowHeader), "showSelectionLabel": safe(lambda: vs.ShowSelectionLabel),
            "showTimeLevel": safe(lambda: vs.ShowTimeLevel), "showScrollbar": safe(lambda: vs.ShowHorizontalScrollbar)},
        "shape": {"name": sh.Name, "top": sh.Top, "left": sh.Left, "width": sh.Width, "height": sh.Height,
                  "tl": sh.TopLeftCell.Address.replace("$", ""), "br": sh.BottomRightCell.Address.replace("$", ""), "placement": safe(lambda: sh.Placement)},
    }


def shot(sl, name: str) -> None:
    from PIL import Image, ImageGrab
    SHOTS.mkdir(exist_ok=True)
    for _ in range(6):
        try:
            sl.Shape.CopyPicture(1, 2)
            time.sleep(0.3)
            img = ImageGrab.grabclipboard()
            if isinstance(img, Image.Image):
                img.convert("RGB").save(SHOTS / f"excel-{name}.png")
                return
        except Exception:  # noqa: BLE001
            pass
        time.sleep(0.4)


def report(pt) -> dict:
    rng = pt.TableRange2
    return {"at": rng.Address.replace("$", ""), "values": [[None if v is None else (str(v)[:10] if hasattr(v, "year") else v) for v in r] for r in rng.Value]}


def style_elements(wb) -> dict:
    out = {}
    for i in range(1, wb.TableStyles.Count + 1):
        st = wb.TableStyles(i)
        if not safe(lambda: st.ShowAsAvailableTimelineStyle, False):
            continue
        rec: dict = {"local": safe(lambda: st.NameLocal), "builtIn": safe(lambda: st.BuiltIn)}
        for k in ELEMENT_RANGE:
            e = safe(lambda: st.TableStyleElements(k))
            if e is None or not safe(lambda: e.HasFormat, False):
                continue
            d: dict = {}
            pat = safe(lambda: e.Interior.Pattern)
            d["pattern"] = pat
            if pat not in (None, -4142, 4000):
                d["fill"] = {"rgb": rgb(safe(lambda: e.Interior.Color)), "theme": safe(lambda: e.Interior.ThemeColor), "tint": safe(lambda: e.Interior.TintAndShade)}
            if pat == 4000:
                g = safe(lambda: e.Interior.Gradient)
                d["gradient"] = {"degree": safe(lambda: g.Degree), "stops": [{"pos": safe(lambda i=i: g.ColorStops(i).Position), "rgb": rgb(safe(lambda i=i: g.ColorStops(i).Color)),
                                                                              "theme": safe(lambda i=i: g.ColorStops(i).ThemeColor), "tint": safe(lambda i=i: g.ColorStops(i).TintAndShade)}
                                                                             for i in range(1, (safe(lambda: g.ColorStops.Count, 0) or 0) + 1)]}
            d["font"] = {"rgb": rgb(safe(lambda: e.Font.Color)), "theme": safe(lambda: e.Font.ThemeColor), "tint": safe(lambda: e.Font.TintAndShade),
                         "bold": safe(lambda: e.Font.Bold), "size": safe(lambda: e.Font.Size), "colorIndex": safe(lambda: e.Font.ColorIndex)}
            b = {}
            for side, bi in BORDERS.items():
                bd = safe(lambda: e.Borders(bi))
                ls = safe(lambda: bd.LineStyle)
                if ls in (None, -4142):
                    continue
                b[side] = {"style": ls, "weight": safe(lambda: bd.Weight), "rgb": rgb(safe(lambda: bd.Color)), "theme": safe(lambda: bd.ThemeColor), "tint": safe(lambda: bd.TintAndShade)}
            if b:
                d["borders"] = b
            rec[str(k)] = d
        out[st.Name] = rec
    return out


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    out: dict = {"data": [[str(v) if isinstance(v, dt.date) else v for v in r] for r in DATA], "cases": {}}
    x = ExcelTarget(visible=False)
    path = TMP / "raabe-timeline-excel.xlsx"
    try:
        wb = x.wb
        ws = wb.Worksheets(1)
        ws.Name = "Noten"
        for r, row in enumerate(DATA, 1):
            for c, v in enumerate(row, 1):
                cell = ws.Cells(r, c)
                if isinstance(v, dt.date):
                    cell.Formula = f"=DATE({v.year},{v.month},{v.day})"
                    cell.Value = cell.Value2
                    cell.NumberFormat = "dd.mm.yyyy"
                else:
                    cell.Value = v
        src = ws.Range(ws.Cells(1, 1), ws.Cells(len(DATA), 3))
        lo = ws.ListObjects.Add(1, src, None, 1)
        lo.Name = "Noten"
        sheets = {}

        def pivot(name: str, sheet: str, cache=None):
            pws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            pws.Name = sheet
            sheets[sheet] = pws
            pc = cache or wb.PivotCaches().Create(1, "Noten")
            pt = pc.CreatePivotTable(pws.Range("A3"), name)
            pt.PivotFields("Fach").Orientation = 1
            pt.AddDataField(pt.PivotFields("Note"), None, -4112)  # Anzahl
            return pws, pt

        def record(case: str, ws_, tls, pts, extra=None):
            rec = {"sheet": ws_.Name, "timelines": [timeline_info(t) for t in tls], "reports": [report(p) for p in pts]}
            if extra:
                rec.update(extra)
            out["cases"][case] = rec
            for t in tls:
                shot(t, f"{case}-{t.Name}")
            print(case, json.dumps(rec, ensure_ascii=False, default=str)[:400], flush=True)

        # 1. Defaults: Slicers.Add with only the sheet (Excel's size and place).
        pws, pt = pivot("PivotStandard", "Standard")
        sc = wb.SlicerCaches.Add2(pt, "Datum", SlicerCacheType=2)
        t1 = sc.Slicers.Add(pws)
        record("default", pws, [t1], [pt])

        # 2. Filtered to Oct 2025 .. Jan 2026, months.
        pws, pt = pivot("PivotFilter", "Filter")
        sc = wb.SlicerCaches.Add2(pt, "Datum", SlicerCacheType=2)
        t2 = sc.Slicers.Add(pws, Top=10, Left=300)
        sc.TimelineState.SetFilterDateRange("01.10.2025", "31.01.2026")
        record("months", pws, [t2], [pt])
        out["cases"]["months"]["pivotHidden"] = {it.Name: bool(it.Visible) for it in pt.PivotFields("Fach").PivotItems()}

        # 3. The levels and a filter at each.
        for lvl, name in ((0, "Jahre"), (1, "Quartale"), (3, "Tage")):
            pws, pt = pivot(f"Pivot{name}", name)
            sc = wb.SlicerCaches.Add2(pt, "Datum", SlicerCacheType=2)
            t = sc.Slicers.Add(pws, Top=10, Left=300, Width=400, Height=120)
            t.TimelineViewState.Level = lvl
            if lvl == 0:
                sc.TimelineState.SetFilterDateRange("01.01.2026", "31.12.2026")
            elif lvl == 1:
                sc.TimelineState.SetFilterDateRange("01.01.2026", "30.06.2026")
            else:
                sc.TimelineState.SetFilterDateRange("04.09.2025", "21.09.2025")
            record(name, pws, [t], [pt])

        # 4. Anzeigen options off, a style and a caption.
        pws, pt = pivot("PivotOptionen", "Optionen")
        sc = wb.SlicerCaches.Add2(pt, "Datum", SlicerCacheType=2)
        t = sc.Slicers.Add(pws, Name="Zeit", Caption="Zeitraum der Noten", Top=10, Left=300, Width=330, Height=100)
        vs = t.TimelineViewState
        vs.ShowHeader = False
        vs.ShowSelectionLabel = False
        vs.ShowHorizontalScrollbar = False
        vs.ShowTimeLevel = False
        t.Style = "TimeSlicerStyleDark2"
        record("options", pws, [t], [pt])

        # 5. Two PivotTables of one cache, one timeline connected to both.
        pws, pt = pivot("PivotA", "Verbunden")
        pt2 = pt.PivotCache().CreatePivotTable(pws.Range("E3"), "PivotB")
        pt2.PivotFields("Note").Orientation = 1
        pt2.AddDataField(pt2.PivotFields("Fach"), None, -4112)
        sc = wb.SlicerCaches.Add2(pt, "Datum", SlicerCacheType=2)
        t = sc.Slicers.Add(pws, Top=200, Left=10)
        sc.PivotTables.AddPivotTable(pt2)
        sc.TimelineState.SetFilterDateRange("01.02.2026", "31.05.2026")
        record("connected", pws, [t], [pt, pt2])

        # 6. Filter löschen.
        pws, pt = pivot("PivotGeloescht", "Geloescht")
        sc = wb.SlicerCaches.Add2(pt, "Datum", SlicerCacheType=2)
        t = sc.Slicers.Add(pws, Top=10, Left=300)
        sc.TimelineState.SetFilterDateRange("01.03.2026", "31.03.2026")
        record("beforeClear", pws, [t], [pt])
        sc.ClearDateFilter()
        record("cleared", pws, [t], [pt])

        out["styles"] = style_elements(wb)
        out["defaultTimelineStyle"] = safe(lambda: wb.DefaultTimelineStyle.Name if not isinstance(wb.DefaultTimelineStyle, str) else wb.DefaultTimelineStyle)
        for prop in ("Author", "Last author"):
            safe(lambda: setattr(wb.BuiltinDocumentProperties.Item(prop), "Value", NEUTRAL))
        if path.exists():
            path.unlink()
        wb.SaveAs(str(path), 51)
    finally:
        x.close()

    with zipfile.ZipFile(path) as z:
        out["parts"] = sorted(z.namelist())
    fixture = FIX / "excel-timeline.xlsx"
    fixture.write_bytes(path.read_bytes())
    subprocess.run([sys.executable, str(HERE / "anonymize_office.py"), str(fixture)], check=False)
    clean_fixture(fixture)
    (FIX / "excel-timeline.json").write_text(json.dumps(out, ensure_ascii=False, indent=1, default=str), encoding="utf-8", newline="\n")
    RESULT.write_text(json.dumps({"cases": out["cases"], "defaultTimelineStyle": out.get("defaultTimelineStyle"), "parts": out["parts"]}, ensure_ascii=False, indent=1, default=str), encoding="utf-8", newline="\n")
    print("written", fixture.name)


if __name__ == "__main__":
    main()
