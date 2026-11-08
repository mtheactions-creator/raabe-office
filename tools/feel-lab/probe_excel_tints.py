"""Records how Excel tints design colours (D-083): for the colours of several designs and the palette's tints,
the RGB value Excel gives a cell with that theme colour and TintAndShade, as the reference for
packages/ooxml/test/tints-excel.test.ts.

    .venv/Scripts/python.exe probe_excel_tints.py

Writes packages/ooxml/test/fixtures/excel-tints.json. Our own invisible Excel; no real input."""

from __future__ import annotations

import json
from pathlib import Path

from feellab.excel import ExcelTarget

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "packages" / "ooxml" / "test" / "fixtures" / "excel-tints.json"
THEMES = Path(r"C:\Program Files\Microsoft Office\root\Document Themes 16")
FILES = ["Office Theme", "Office 2013 - 2022 Theme", "Ion", "Facet", "Retrospect"]
TINTS = [26213 / 32767, 19660 / 32767, 13106 / 32767, -8191 / 32767, -16383 / 32767, 16383 / 32767, 11468 / 32767, 8191 / 32767,
         4915 / 32767, -1638 / 32767, -4915 / 32767, -11468 / 32767, -3276 / 32767, -24575 / 32767, -29490 / 32767]
# xlThemeColor: 1 dark1, 2 light1, 3 dark2, 4 light2, 5-10 accents.
THEME_COLORS = list(range(1, 11))


def bgr(n: int) -> str:
    n = int(n)
    return f"{n & 255:02X}{(n >> 8) & 255:02X}{(n >> 16) & 255:02X}"


def main() -> None:
    xl = ExcelTarget(visible=False)
    out: list = []
    try:
        sh = xl.wb.Worksheets(1)
        for f in FILES:
            xl.wb.ApplyTheme(str(THEMES / f"{f}.thmx"))
            cs = xl.wb.Theme.ThemeColorScheme
            for tc in THEME_COLORS:
                # msoThemeColorSchemeIndex order: 1 dk1, 2 lt1, 3 dk2, 4 lt2, 5.. accents; xlThemeColor 1 dark1, 2 light1 …
                # A cell's ThemeColor 1 and 2 are Hintergrund 1 (lt1) and Text 1 (dk1), 3 and 4 Hintergrund 2 and
                # Text 2 (recorded): the opposite of the colour scheme's order.
                base = bgr(cs.Colors({1: 2, 2: 1, 3: 4, 4: 3}.get(tc, tc)).RGB)
                for i, t in enumerate(TINTS):
                    c = sh.Cells(i + 1, tc)
                    c.Interior.ThemeColor = tc
                    c.Interior.TintAndShade = t
                    out.append({"design": f, "theme": tc, "base": base, "tint": c.Interior.TintAndShade, "rgb": bgr(c.Interior.Color)})
    finally:
        xl.close()
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(out, indent=0) + "\n", encoding="utf-8", newline="\n")
    print(len(out), "cases")


if __name__ == "__main__":
    main()
