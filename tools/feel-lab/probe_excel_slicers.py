"""Records Datenschnitte (slicers) in German Excel 365 (D-092), COM only, our own invisible Excel.

Cases, one sheet each: a table with a slicer added with Excel's defaults; slicers that filter (one item
off, two slicers filtering together, the four Elemente-ohne-Daten settings); a slicer's settings
(Beschriftung, Kopfzeile, Spalten, Höhe/Breite der Schaltflächen, Formatvorlage, Sortierung); a slicer
on a number column with an empty cell; a slicer of a PivotTable. For each: the slicer's and its cache's
properties, every item (Selected, HasData, in the order shown), the rows the table hides, the
PivotTable's report; the slicer styles' elements as Excel reports them; Excel's rendering of each
slicer (Shape.CopyPicture, the clipboard for a moment) to TEMP/raabe-slicers/excel-<name>.png.

    .venv/Scripts/python.exe probe_excel_slicers.py

Writes packages/ooxml/test/fixtures/excel-slicers.xlsx (anonymized) and excel-slicers.json.
"""
from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path

from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
ROOT = HERE.parents[1]
FIX = ROOT / "packages" / "ooxml" / "test" / "fixtures"
TMP = Path(tempfile.gettempdir())
SHOTS = TMP / "raabe-slicers"
DATA = [
    ["Name", "Klasse", "Fach", "Note"],
    ["Anna", "8a", "Mathe", 2], ["Ben", "8a", "Mathe", 3], ["Cem", "8b", "Mathe", 1], ["Dora", "8b", "Deutsch", 4],
    ["Anna", "8a", "Deutsch", 1], ["Ben", "8a", "Englisch", 2], ["Cem", "8b", "Deutsch", 3], ["Emil", "8c", "Englisch", 5],
]
# XlTableStyleElementType for slicers.
SLICER_ELEMENTS = {"wholeTable": 0, "headerRow": 1, "unselectedItemWithData": 28, "unselectedItemWithNoData": 29,
                   "selectedItemWithData": 30, "selectedItemWithNoData": 31, "hoveredUnselectedItemWithData": 32,
                   "hoveredSelectedItemWithData": 33, "hoveredUnselectedItemWithNoData": 34, "hoveredSelectedItemWithNoData": 35}
BORDERS = {"left": 7, "top": 8, "bottom": 9, "right": 10}
NEUTRAL = "Raabe Office Test"


def safe(f, d=None):
    try:
        return f()
    except Exception:  # noqa: BLE001
        return d


def rgb(n) -> str | None:
    if n is None or n < 0 or n > 0xFFFFFF:
        return None
    n = int(n)
    return f"{n & 0xFF:02X}{(n >> 8) & 0xFF:02X}{(n >> 16) & 0xFF:02X}"


def fill(ws, at: str = "A1", data=DATA):
    r0 = ws.Range(at).Row
    c0 = ws.Range(at).Column
    for r, row in enumerate(data):
        for c, v in enumerate(row):
            if v is not None:
                ws.Cells(r0 + r, c0 + c).Value = v
    return ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + len(data) - 1, c0 + len(data[0]) - 1))


def slicer_info(sl) -> dict:
    sc = sl.SlicerCache
    items = []
    for i in range(1, sc.SlicerItems.Count + 1):
        it = sc.SlicerItems(i)
        items.append({"name": it.Name, "caption": safe(lambda: it.Caption), "selected": bool(it.Selected), "hasData": bool(safe(lambda: it.HasData, True))})
    # The order shown: VisibleSlicerItems lists what is on the buttons (data first unless set otherwise).
    shown = [safe(lambda i=i: sc.VisibleSlicerItems(i).Name) for i in range(1, (safe(lambda: sc.VisibleSlicerItems.Count, 0) or 0) + 1)]
    sh = sl.Shape
    return {
        "name": sl.Name, "caption": sl.Caption, "cache": sc.Name, "sourceName": safe(lambda: sc.SourceName),
        "top": sl.Top, "left": sl.Left, "width": sl.Width, "height": sl.Height,
        "rowHeight": sl.RowHeight, "columnWidth": sl.ColumnWidth, "columns": sl.NumberOfColumns,
        "style": safe(lambda: sl.Style.Name if not isinstance(sl.Style, str) else sl.Style), "displayHeader": sl.DisplayHeader,
        "locked": safe(lambda: sl.Locked), "disableMoveResize": safe(lambda: sl.DisableMoveResizeUI),
        "sortItems": safe(lambda: sc.SortItems), "sortCustomLists": safe(lambda: sc.SortUsingCustomLists),
        "crossFilter": safe(lambda: sc.CrossFilterType), "showAllItems": safe(lambda: sc.ShowAllItems),
        "shape": {"name": sh.Name, "top": sh.Top, "left": sh.Left, "width": sh.Width, "height": sh.Height,
                  "tl": sh.TopLeftCell.Address.replace("$", ""), "br": sh.BottomRightCell.Address.replace("$", ""), "placement": safe(lambda: sh.Placement)},
        "items": items, "shown": shown,
    }


def hidden_rows(ws, rng) -> list[int]:
    return [r for r in range(rng.Row, rng.Row + rng.Rows.Count) if ws.Rows(r).Hidden]


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


def style_elements(wb) -> dict:
    out = {}
    names = [f"SlicerStyleLight{i}" for i in range(1, 7)] + [f"SlicerStyleOther{i}" for i in range(1, 3)] + [f"SlicerStyleDark{i}" for i in range(1, 7)]
    for name in names:
        st = safe(lambda: wb.TableStyles(name))
        if st is None:
            out[name] = None
            continue
        rec: dict = {"local": safe(lambda: st.NameLocal), "slicer": safe(lambda: st.ShowAsAvailableSlicerStyle)}
        for el, k in SLICER_ELEMENTS.items():
            e = safe(lambda: st.TableStyleElements(k))
            if e is None or not safe(lambda: e.HasFormat, False):
                continue
            d: dict = {}
            pat = safe(lambda: e.Interior.Pattern)
            d["pattern"] = pat
            if pat == 4000:  # linear gradient
                g = safe(lambda: e.Interior.Gradient)
                d["gradient"] = {"degree": safe(lambda: g.Degree), "stops": [{"pos": safe(lambda i=i: g.ColorStops(i).Position), "rgb": rgb(safe(lambda i=i: g.ColorStops(i).Color)),
                                                                              "theme": safe(lambda i=i: g.ColorStops(i).ThemeColor), "tint": safe(lambda i=i: g.ColorStops(i).TintAndShade)}
                                                                             for i in range(1, (safe(lambda: g.ColorStops.Count, 0) or 0) + 1)]}
            elif pat not in (None, -4142):
                d["fill"] = {"rgb": rgb(safe(lambda: e.Interior.Color)), "theme": safe(lambda: e.Interior.ThemeColor), "tint": safe(lambda: e.Interior.TintAndShade)}
            d["font"] = {"rgb": rgb(safe(lambda: e.Font.Color)), "theme": safe(lambda: e.Font.ThemeColor), "tint": safe(lambda: e.Font.TintAndShade), "bold": safe(lambda: e.Font.Bold), "colorIndex": safe(lambda: e.Font.ColorIndex)}
            b = {}
            for side, bi in BORDERS.items():
                bd = e.Borders(bi)
                ls = safe(lambda: bd.LineStyle)
                if ls in (None, -4142):
                    continue
                b[side] = {"style": ls, "weight": safe(lambda: bd.Weight), "rgb": rgb(safe(lambda: bd.Color)), "theme": safe(lambda: bd.ThemeColor), "tint": safe(lambda: bd.TintAndShade)}
            if b:
                d["borders"] = b
            rec[el] = d
        out[name] = rec
    return out


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    out: dict = {"data": DATA, "cases": {}}
    x = ExcelTarget(visible=False)
    path = TMP / "raabe-slicers-excel.xlsx"
    try:
        wb = x.wb
        THEME = r"C:\Program Files\Microsoft Office\root\Document Themes 16\Office 2013 - 2022 Theme.thmx"
        # The workbook's own theme (Excel 365's default "Office"): kept as Excel makes new workbooks.
        out["theme"] = safe(lambda: wb.Theme.ThemeColorScheme.Colors(5).RGB)
        _ = THEME
        sheets = []

        def sheet(name: str):
            ws = wb.Worksheets(1) if not sheets else wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            ws.Name = name
            sheets.append(ws)
            return ws

        def table(ws, name: str, data=DATA):
            rng = fill(ws, "A1", data)
            lo = ws.ListObjects.Add(1, rng, None, 1)
            lo.Name = name
            return lo, rng

        def record(case: str, ws, slicers, rng=None, extra=None):
            rec = {"sheet": ws.Name, "slicers": [slicer_info(s) for s in slicers]}
            if rng is not None:
                rec["hiddenRows"] = hidden_rows(ws, rng)
            if extra:
                rec.update(extra)
            out["cases"][case] = rec
            for s in slicers:
                shot(s, f"{case}-{s.Name}")
            print(case, json.dumps(rec, ensure_ascii=False)[:300], flush=True)

        # 1. Defaults: Slicers.Add with only the sheet.
        ws = sheet("Standard")
        lo, rng = table(ws, "Noten")
        sc = wb.SlicerCaches.Add2(lo, "Fach")
        s1 = sc.Slicers.Add(ws)
        record("default", ws, [s1], rng)
        # A second slicer of the same table, also with defaults (where does it go?).
        sc2 = wb.SlicerCaches.Add2(lo, "Klasse")
        s2 = sc2.Slicers.Add(ws)
        record("default2", ws, [s1, s2], rng)

        # 2. One item off.
        ws = sheet("Filter")
        lo, rng = table(ws, "Noten2")
        sc = wb.SlicerCaches.Add2(lo, "Fach")
        a = sc.Slicers.Add(ws, Top=10, Left=300, Width=144, Height=198)
        sc.SlicerItems("Mathe").Selected = False
        sc3 = wb.SlicerCaches.Add2(lo, "Name")
        b = sc3.Slicers.Add(ws, Top=10, Left=460, Width=144, Height=198)
        record("one-off", ws, [a, b], rng)

        # 3. Only Deutsch in Fach and 8c off in Klasse.
        ws = sheet("Zwei")
        lo, rng = table(ws, "Noten3")
        sc = wb.SlicerCaches.Add2(lo, "Fach")
        a = sc.Slicers.Add(ws, Top=10, Left=300, Width=144, Height=198)
        sc.SlicerItems("Mathe").Selected = False
        sc.SlicerItems("Englisch").Selected = False
        scb = wb.SlicerCaches.Add2(lo, "Klasse")
        b = scb.Slicers.Add(ws, Top=10, Left=460, Width=144, Height=198)
        record("two-filters", ws, [a, b], rng)
        # Filter löschen on Fach.
        sc.ClearManualFilter()
        record("cleared", ws, [a, b], rng)

        # 4. The Elemente ohne Daten settings (Name slicer filtered by Fach = Englisch only).
        for k in (1, 2, 3, 4):
            ws = sheet(f"Kreuz{k}")
            lo, rng = table(ws, f"Kreuz{k}")
            sc = wb.SlicerCaches.Add2(lo, "Fach")
            a = sc.Slicers.Add(ws, Top=10, Left=300, Width=144, Height=198)
            sc.SlicerItems("Mathe").Selected = False
            sc.SlicerItems("Deutsch").Selected = False
            scb = wb.SlicerCaches.Add2(lo, "Name")
            b = scb.Slicers.Add(ws, Top=10, Left=460, Width=144, Height=198)
            scb.CrossFilterType = k
            record(f"cross{k}", ws, [a, b], rng)

        # 5. Settings of one slicer.
        ws = sheet("Einstellungen")
        lo, rng = table(ws, "Noten5")
        sc = wb.SlicerCaches.Add2(lo, "Fach")
        a = sc.Slicers.Add(ws, Name="FachA", Caption="Unterrichtsfach", Top=10, Left=300, Width=200, Height=120)
        a.NumberOfColumns = 2
        a.RowHeight = 30
        a.ColumnWidth = 80
        a.Style = "SlicerStyleDark2"
        sc.SortItems = 2  # XlSlicerSort: 1 data source order, 2 ascending (the default), 3 descending (the file shows it)
        scb = wb.SlicerCaches.Add2(lo, "Klasse")
        b = scb.Slicers.Add(ws, Top=150, Left=300, Width=144, Height=120)
        b.DisplayHeader = False
        b.Style = "SlicerStyleLight3"
        scb.SortItems = 3  # descending
        record("settings", ws, [a, b], rng)

        # 6. Numbers, an empty cell, and a date-free numeric sort.
        ws = sheet("Zahlen")
        data = [r[:] for r in DATA]
        data[3][3] = None
        data.append(["Finn", "8c", "Mathe", 10])
        lo, rng = table(ws, "Noten6", data)
        sc = wb.SlicerCaches.Add2(lo, "Note")
        a = sc.Slicers.Add(ws, Top=10, Left=300, Width=144, Height=240)
        sc.SlicerItems(sc.SlicerItems(1).Name).Selected = False
        record("numbers", ws, [a], rng)

        # 7. A PivotTable's slicer.
        ws = sheet("Daten")
        src = fill(ws, "A1")
        pws = sheet("Pivot")
        pc = wb.PivotCaches().Create(1, src)
        pt = pc.CreatePivotTable(pws.Range("A3"), "PivotNoten")
        pt.PivotFields("Klasse").Orientation = 1
        pt.AddDataField(pt.PivotFields("Note"), None, -4157)
        sc = wb.SlicerCaches.Add2(pt, "Fach")
        a = sc.Slicers.Add(pws, Top=10, Left=300, Width=144, Height=198)
        sc.SlicerItems("Mathe").Selected = False
        rng2 = pt.TableRange2
        vals = rng2.Value
        record("pivot", pws, [a], extra={"report": {"at": rng2.Address.replace("$", ""), "values": [list(r) for r in vals]},
                                         "pivotItems": {it.Name: bool(it.Visible) for it in pt.PivotFields("Fach").PivotItems()}})

        out["styles"] = style_elements(wb)
        out["defaultStyle"] = safe(lambda: wb.DefaultSlicerStyle.Name if not isinstance(wb.DefaultSlicerStyle, str) else wb.DefaultSlicerStyle)
        for prop in ("Author", "Last author"):
            safe(lambda: setattr(wb.BuiltinDocumentProperties.Item(prop), "Value", NEUTRAL))
        if path.exists():
            path.unlink()
        wb.SaveAs(str(path), 51)
    finally:
        x.close()

    with zipfile.ZipFile(path) as z:
        out["parts"] = sorted(z.namelist())
    fixture = FIX / "excel-slicers.xlsx"
    fixture.write_bytes(path.read_bytes())
    subprocess.run([sys.executable, str(HERE / "anonymize_office.py"), str(fixture)], check=False)
    clean_fixture(fixture)
    (FIX / "excel-slicers.json").write_text(json.dumps(out, ensure_ascii=False, indent=1, default=str), encoding="utf-8", newline="\n")
    print("written", fixture.name)


def clean_fixture(path: Path) -> None:
    """The repository is public: the save path Excel notes in workbook.xml goes, and the fallback text Excel
    puts into each slicer's drawing anchor (Office's prose, D-058) is replaced by our own words."""
    import io
    import re
    data = path.read_bytes()
    buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as z, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as w:
        for info in z.infolist():
            t = z.read(info.filename)
            if info.filename == "xl/workbook.xml":
                t = re.sub(rb'<mc:AlternateContent[^>]*><mc:Choice Requires="x15"><x15ac:absPath\b.*?</mc:AlternateContent>', b"", t, flags=re.S)
            if info.filename.startswith("xl/drawings/") and info.filename.endswith(".xml"):
                t = re.sub(rb"(<mc:Fallback>.*?<a:t>).*?(</a:t>)", lambda m: m.group(1) + "Hier steht ein Datenschnitt.".encode() + m.group(2), t, flags=re.S)
            w.writestr(info, t)
    path.write_bytes(buf.getvalue())


if __name__ == "__main__":
    if sys.argv[1:] == ["--clean"]:
        clean_fixture(FIX / "excel-slicers.xlsx")
    else:
        main()
