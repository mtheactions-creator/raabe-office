"""Probe: Excel's built-in table styles (Tabellenformatvorlagen: Hell 1-21, Mittel 1-28, Dunkel 1-11) as
data (D-066): for every style and table-style element the fill, the font colour and weight and the
borders, each as theme colour + tint where Excel gives one (so RaabeCalc draws them in any theme) and
as RGB in the Office 2013-2022 theme. Also the gallery's default for a new table. Our own invisible
Excel, COM only. Result: results/<date>-excel-tablestyles.json.
"""
from __future__ import annotations

import datetime
import json
from pathlib import Path

from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
THEME_2013 = r"C:\Program Files\Microsoft Office\root\Document Themes 16\Office 2013 - 2022 Theme.thmx"
# XlTableStyleElementType
ELEMENTS = {"wholeTable": 0, "headerRow": 1, "totalRow": 2, "firstColumn": 3, "lastColumn": 4, "firstRowStripe": 5, "secondRowStripe": 6,
            "firstColumnStripe": 7, "secondColumnStripe": 8, "firstHeaderCell": 9, "lastHeaderCell": 10, "firstTotalCell": 11, "lastTotalCell": 12}
# XlBordersIndex
BORDERS = {"left": 7, "top": 8, "bottom": 9, "right": 10, "insideV": 11, "insideH": 12}


def rgb(n) -> str | None:
    if n is None or n < 0 or n > 0xFFFFFF:
        return None
    n = int(n)
    return f"{n & 0xFF:02X}{(n >> 8) & 0xFF:02X}{(n >> 16) & 0xFF:02X}"


def safe(f, d=None):
    try:
        return f()
    except Exception:  # noqa: BLE001
        return d


def theme_of(obj) -> dict:
    tc = safe(lambda: obj.ThemeColor)
    out = {}
    if isinstance(tc, int) and tc > 0:
        out["theme"] = tc
        t = safe(lambda: obj.TintAndShade, 0)
        if t:
            out["tint"] = round(t, 4)
    return out


def main() -> None:
    x = ExcelTarget(visible=False)
    out: dict = {"styles": {}}
    try:
        x.wb.ApplyTheme(THEME_2013)
        ws = x.sheet
        ws.Range("A1:C4").Value = (("A", "B", "C"), (1, 2, 3), (4, 5, 6), (7, 8, 9))
        lo = ws.ListObjects.Add(1, ws.Range("A1:C4"), None, 1)
        out["default"] = lo.TableStyle.Name
        names = [f"TableStyleLight{i}" for i in range(1, 22)] + [f"TableStyleMedium{i}" for i in range(1, 29)] + [f"TableStyleDark{i}" for i in range(1, 12)]
        for name in names:
            st = x.wb.TableStyles(name)
            rec: dict = {"local": safe(lambda: st.NameLocal)}
            for el, k in ELEMENTS.items():
                e = st.TableStyleElements(k)
                if not safe(lambda: e.HasFormat, False):
                    continue
                d: dict = {}
                pat = safe(lambda: e.Interior.Pattern)
                if pat not in (None, -4142):  # xlNone
                    d["fill"] = {"rgb": rgb(safe(lambda: e.Interior.Color)), **theme_of(e.Interior)}
                fc = safe(lambda: e.Font.Color)
                if fc is not None and safe(lambda: e.Font.ColorIndex) not in (-4105, None):
                    d["font"] = {"rgb": rgb(fc), **theme_of(e.Font)}
                if safe(lambda: e.Font.Bold):
                    d["bold"] = True
                b: dict = {}
                for side, bi in BORDERS.items():
                    bd = e.Borders(bi)
                    ls = safe(lambda: bd.LineStyle)
                    if ls in (None, -4142):
                        continue
                    b[side] = {"style": ls, "weight": safe(lambda: bd.Weight), "rgb": rgb(safe(lambda: bd.Color)), **theme_of(bd)}
                if b:
                    d["borders"] = b
                if el in ("firstRowStripe", "secondRowStripe", "firstColumnStripe", "secondColumnStripe"):
                    size = safe(lambda: e.StripeSize, 1)
                    if size and size != 1:
                        d["stripe"] = size
                if d:
                    rec[el] = d
            out["styles"][name] = rec
            print(name, len(rec) - 1, flush=True)
    finally:
        x.close()
    (HERE / "results" / f"{datetime.date.today().isoformat()}-excel-tablestyles.json").write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
