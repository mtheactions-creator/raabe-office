"""Seitenlayout > Designs, Farben, Schriftarten against Excel (D-083), without real input:
1. Data: our own invisible Excel applies Office's installed designs (.thmx); each must give the colours and
   fonts of the design of the same colours in RaabeCalc's gallery (packages/shell/src/office-designs.gen.ts).
2. RaabeCalc -> Excel: RaabeCalc (off screen) colours cells with design colours from the palette and a
   standard colour, applies the design Ion and the font set Arial, saves; Excel must see the same theme
   colours and fonts, the design colour references (ThemeColor, TintAndShade) and the same RGB values.
3. Excel -> RaabeCalc: Excel colours a cell with Akzent 1, heller 40 % under Ion and saves; RaabeCalc opens it
   (same fill), applies the design Office, and the cell must get the colour Excel gives it under Office.
Result: results/<date>-calc-designs.json.
"""
from __future__ import annotations

import datetime
import json
import re
import tempfile
from pathlib import Path

from feellab.calc_catalogue import Case, CalcSide
from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
ROOT = HERE.parents[1]
TMP = Path(tempfile.gettempdir())
THEMES = Path(r"C:\Program Files\Microsoft Office\root\Document Themes 16")
FILES = ["Office Theme", "Office 2013 - 2022 Theme", "Facet", "Integral", "Ion", "Ion Boardroom", "Organic", "Retrospect", "Slice", "Wisp", "Gallery"]
KEYS = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]


def bgr(n: int) -> str:
    n = int(n)
    return f"{n & 255:02X}{(n >> 8) & 255:02X}{(n >> 16) & 255:02X}"


def designs() -> list:
    text = (ROOT / "packages/shell/src/office-designs.gen.ts").read_text(encoding="utf-8")
    m = re.search(r"export const OFFICE_DESIGNS[^=]*= (\[.*?\]);\n", text, re.S)
    return json.loads(m.group(1))


def theme_of(wb) -> dict:
    cs = wb.Theme.ThemeColorScheme
    fs = wb.Theme.ThemeFontScheme
    return {"colors": {k: bgr(cs.Colors(i + 1).RGB) for i, k in enumerate(KEYS)}, "major": fs.MajorFont.Item(1).Name, "minor": fs.MinorFont.Item(1).Name}


def main() -> None:
    report: dict = {"data": [], "raabe->excel": [], "excel->raabe": []}
    ours = designs()
    x = ExcelTarget(visible=False)
    try:
        wb = x.wb
        for f in FILES:
            path = THEMES / f"{f}.thmx"
            if not path.exists():
                report["data"].append({"file": f, "ok": False, "why": "not installed"})
                continue
            wb.ApplyTheme(str(path))
            t = theme_of(wb)
            match = next((d for d in ours if all(d["colors"].get(k, "").upper() == t["colors"][k] for k in KEYS if k not in ("dk1", "lt1"))
                          and d["major"] == t["major"] and d["minor"] == t["minor"]), None)
            report["data"].append({"file": f, "ok": match is not None, "ours": match["name"] if match else None, **({} if match else {"excel": t})})
            print("data", f, match["name"] if match else "NO MATCH", flush=True)
        # 3. Excel's file for RaabeCalc.
        wb.ApplyTheme(str(THEMES / "Ion.thmx"))
        sh = wb.Worksheets(1)
        sh.Range("A1").Interior.ThemeColor = 5  # xlThemeColorAccent1
        sh.Range("A1").Interior.TintAndShade = 0.3999755851924192
        sh.Range("B1").Font.ThemeColor = 6
        sh.Range("B1").Font.TintAndShade = -0.249977111117893
        sh.Range("B1").Value = "Text"
        ion = {"A1": bgr(sh.Range("A1").Interior.Color), "B1": bgr(sh.Range("B1").Font.Color)}
        excel_file = TMP / "excel-designs.xlsx"
        if excel_file.exists():
            excel_file.unlink()
        wb.SaveCopyAs(str(excel_file))
        wb.ApplyTheme(str(THEMES / "Office Theme.thmx"))
        office = {"A1": bgr(sh.Range("A1").Interior.Color), "B1": bgr(sh.Range("B1").Font.Color)}
    finally:
        x.close()

    c = CalcSide()
    try:
        # 2. RaabeCalc's file for Excel.
        c.reset(Case("designs", {"B1": "Text", "D1": "Text"}, [], "A1", sheets=1))
        c.js("""(() => { const app = __calc.app; const at = (a) => { const r = __calc.parse(a); app.grid.selectCell(r.r0, r.c0); };
          at("A1"); app.setFillColor("accent1,60000,40000"); at("B1"); app.setFontColor("accent2,75000"); at("C1"); app.setFillColor("FF0000");
          app.designs.applyDesign("Ion"); app.designs.applyFonts("Arial"); })()""")
        raabe = c.js("""(() => { const app = __calc.app; const st = (a) => { const r = __calc.parse(a); return app.wb.styles.get(app.sheet.get(r.r0, r.c0)?.s); };
          return { theme: app.wb.theme, fonts: app.wb.themeFonts, A1: st("A1").fill, B1: st("B1").font.color, C1: st("C1").fill, D1font: st("D1").font.name }; })()""")
        raabe_file = TMP / "raabe-designs.xlsx"
        c.js(f"__calc.save({json.dumps(str(raabe_file))})")
        # 3. Excel's file in RaabeCalc.
        c.js(f"__calc.open({json.dumps(str(excel_file))})")
        opened = c.js("""(() => { const app = __calc.app; const st = (a) => { const r = __calc.parse(a); return app.wb.styles.get(app.sheet.get(r.r0, r.c0)?.s); };
          const before = { A1: st("A1").fill, B1: st("B1").font.color, theme: app.wb.theme };
          app.designs.applyDesign("Office");
          return { before, after: { A1: st("A1").fill, B1: st("B1").font.color } }; })()""")
    finally:
        c.close()

    x = ExcelTarget(visible=False)
    try:
        wb = x.app.Workbooks.Open(str(raabe_file))
        sh = wb.Worksheets(1)
        t = theme_of(wb)
        a1, b1, c1 = sh.Range("A1"), sh.Range("B1"), sh.Range("C1")
        got = {"theme": t, "A1": {"color": bgr(a1.Interior.Color), "theme": a1.Interior.ThemeColor, "tint": round(a1.Interior.TintAndShade, 4)},
               "B1": {"color": bgr(b1.Font.Color), "theme": b1.Font.ThemeColor, "tint": round(b1.Font.TintAndShade, 4)},
               "C1": {"color": bgr(c1.Interior.Color)}, "D1font": sh.Range("D1").Font.Name}
        wb.Close(False)
    finally:
        x.close()
    ion_design = next(d for d in ours if d["name"] == "Ion")
    hx = lambda v: (v or "").replace("#", "").upper()  # noqa: E731
    checks = [
        ("theme colours are Ion's", all(t["colors"][k] == ion_design["colors"][k].upper() for k in KEYS if k not in ("dk1", "lt1"))),
        ("fonts are the set Arial", t["major"] == "Arial" and t["minor"] == "Arial"),
        ("A1 fill is Akzent 1", got["A1"]["theme"] == 5),
        ("A1 tint heller 40 %", abs(got["A1"]["tint"] - 0.4) < 0.001),
        ("A1 RGB the same", got["A1"]["color"] == hx(raabe["A1"])),
        ("B1 font is Akzent 2", got["B1"]["theme"] == 6),
        ("B1 tint dunkler 25 %", abs(got["B1"]["tint"] + 0.25) < 0.001),
        ("B1 RGB the same", got["B1"]["color"] == hx(raabe["B1"])),
        ("C1 stays red", got["C1"]["color"] == "FF0000"),
        ("D1 body font follows (Arial)", got["D1font"] == "Arial" and raabe["D1font"] == "Arial"),
    ]
    for name, ok in checks:
        report["raabe->excel"].append({"check": name, "ok": ok})
        print("R->E", name, ok, flush=True)
    checks = [
        ("A1 as Excel under Ion", hx(opened["before"]["A1"]) == ion["A1"]),
        ("B1 as Excel under Ion", hx(opened["before"]["B1"]) == ion["B1"]),
        ("A1 as Excel after Office", hx(opened["after"]["A1"]) == office["A1"]),
        ("B1 as Excel after Office", hx(opened["after"]["B1"]) == office["B1"]),
    ]
    for name, ok in checks:
        report["excel->raabe"].append({"check": name, "ok": ok})
        print("E->R", name, ok, flush=True)
    report["detail"] = {"raabe": raabe, "excel_reads": got, "opened": opened, "excel_ion": ion, "excel_office": office}
    (HERE / "results" / f"{datetime.date.today().isoformat()}-calc-designs.json").write_text(json.dumps(report, ensure_ascii=False, indent=1, default=str), encoding="utf-8", newline="\n")
    for key in ("data", "raabe->excel", "excel->raabe"):
        v = report[key]
        print(key, sum(1 for r in v if r["ok"]), "/", len(v))


if __name__ == "__main__":
    main()
