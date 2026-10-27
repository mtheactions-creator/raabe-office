"""Probe: number formats in Excel (Step 7, D-061). In an invisible Excel of our own (Calibri 11 as
the Standard style's font, as in RaabeCalc) every format code below is set in the file's form
(NumberFormat read and written with the English LCID, otherwise Excel reads it as German) on
a column, and every value on a row; `Text` is recorded with a wide column. Then the General
format and a few others with columns of many widths (Excel shrinks General numbers to fit and
shows "####" for the others), random numbers of every size, where fractions turn into "####",
the German form of each code (NumberFormatLocal) both ways with the code Excel writes into the
file, the codes of the Zellen formatieren categories, the 1904 date system, and the built-in
format ids as Excel reads and writes them in a .xlsx built here (in TEMP, deleted afterwards).
Runs invisibly, no input.
Result: packages/sheet-engine/test/fixtures/excel-numfmt.json (the tests replay it).

    .venv/Scripts/python.exe probe_excel_numfmt.py
"""
from __future__ import annotations

import ctypes
import datetime
import json
import os
import re
import sys
import tempfile
import zipfile
from ctypes import wintypes
from pathlib import Path

import pythoncom

from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
OUT = HERE.parent.parent / "packages" / "sheet-engine" / "test" / "fixtures" / "excel-numfmt.json"

EN = 0x409
GET = pythoncom.DISPATCH_PROPERTYGET
PUT = pythoncom.DISPATCH_PROPERTYPUT


def put(obj, name: str, value, lcid: int = EN) -> None:
    o = obj._oleobj_
    o.Invoke(o.GetIDsOfNames(name), lcid, PUT, 0, value)


def get(obj, name: str, lcid: int = EN):
    o = obj._oleobj_
    return o.Invoke(o.GetIDsOfNames(name), lcid, GET, 1)


VALUES = [
    0, 1, -1, 0.5, -0.5, 1.5, -1.5, 2.5, 0.125, 12.345, -12.345, 123.456, 1234.5678, -1234.5678,
    12345.6789, 123456.789, 1234567.891, -1234567.891, 12345678.9, 1e9, 1e10, 99999999999, 1e11,
    123456789012, 1e15, 1.5e100, -1.5e100, 1e300, 0.1, 0.01, 0.001, 0.0001, 0.00001, 1.23456e-5,
    1e-10, 1e-300, 0.1 + 0.2, 1 / 3, 2 / 3, -1 / 3, 1 / 7, 3.14159265358979, 0.005, 0.015, 1.005,
    2.675, 0.999999, 9.9999999, 99.995, 0.0005, 0.75, 0.5 / 1440, 1 / 86400, 0.99999999, 59, 60, 61,
    366.25, 45000.5, 45658.75, 45658.999994213, 46022.4375, 2958465, 2958465.9999, 2958466, -0.25,
    1.25, 25.75, 100.123456, 32.5, 1 / 24, 1 + 1 / 24, 0.6, 7, 1000, 1000000, -1000, 0.05, 1.999,
    44927, 45292.123456789, 12.5, 0.0123,
]

CURRENCY = [
    '#,##0.00\\ "€"', '#,##0.00 "€"', '#,##0\\ "€"', '#,##0\\ "€";\\-#,##0\\ "€"', '#,##0\\ "€";[Red]\\-#,##0\\ "€"',
    '#,##0.00\\ "€";\\-#,##0.00\\ "€"', '#,##0.00\\ "€";[Red]\\-#,##0.00\\ "€"', '#,##0.00\\ "€";[Red]#,##0.00\\ "€"',
    '#,##0.00\\ [$€-407]', '[$€-407]\\ #,##0.00', '#,##0.00\\ [$€-1]', '[$€-2]\\ #,##0.00', '#,##0.00\\ [$EUR]',
    '[$$-409]#,##0.00', '[$CHF-807]\\ #,##0.00', '[$£-809]#,##0.00', '$#,##0.00', '\\$0.00', '#,##0.00 €',
    '_-* #,##0.00\\ "€"_-;\\-* #,##0.00\\ "€"_-;_-* "-"??\\ "€"_-;_-@_-',
    '_-* #,##0\\ "€"_-;\\-* #,##0\\ "€"_-;_-* "-"\\ "€"_-;_-@_-',
    '_-* #,##0.00_-;\\-* #,##0.00_-;_-* "-"??_-;_-@_-', '_-* #,##0_-;\\-* #,##0_-;_-* "-"_-;_-@_-',
    '_-[$€-407]\\ * #,##0.00_-;\\-[$€-407]\\ * #,##0.00_-;_-[$€-407]\\ * "-"??_-;_-@_-',
    '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
]

NUMBER = [
    "General", "0", "0.0", "0.00", "0.000", "#,##0", "#,##0.00", "#,##0.000", "#", "#.##", "0.0#", "#,###",
    "#,", "#,##0,", "0.0,,", '#,##0.0,," Mio"', "000", "00000", "000.000", "?", "??.??", "#.???", "0.?",
    "#,##0_);(#,##0)", "#,##0_);[Red](#,##0)", "#,##0.00_);(#,##0.00)", "#,##0.00_);[Red](#,##0.00)",
    "0;-0;;@", '0;(0);"Null"', '0;-0;0;"Text: "@', "General;-General", 'General" kg"', '0" kg"',
    '0 "Stück"', '"Summe: "0.00', "0\\ \\k\\g", "0.00_ ;\\-0.00\\ ", "#,##0.00_ ;[Red]\\-#,##0.00\\ ",
    "0.00;[Red]0.00", "0.00;[Red]\\-0.00", "#,##0.0#", "0 0", "00-00", "000\\-0000", "##-##", '""0',
    '0" "0', "0,0", "(0)", "-0", "+0", "0-", "#\\ ##0", "\\a\\b\\c0", "0\\.0", '0"a"0', "0_)", "0_-", "0_€",
    "#,##0.00_€", "0.0_0", "#0.0", ".00", "0.", "#.", "0.#", "#,##0.", "00.##", "0000.0000", "#,##0.00##",
]

PERCENT = ["0%", "0.0%", "0.00%", "0.000%", "#,##0%", "0%;[Red]-0%", "0.00\\%", "%0", "0%%", "# ?/?%"]

SCIENTIFIC = [
    "0.00E+00", "0.0E+0", "##0.0E+0", "0E+00", "0.00E-00", "0.00e+00", "#E+0", "0.000E+000", "#,##0.00E+00",
    "00.00E+00", "###.##E+00", "0.0##E+00", "?.??E+00", "#0.0E+0", "0E-0",
]

FRACTION = [
    "# ?/?", "# ??/??", "# ???/???", "?/?", "??/??", "# ?/2", "# ?/4", "# ?/8", "# ??/16", "# ?/10",
    "# ??/100", "0 ?/?", "0/0", "# 0/0", "#/#", "# ?/3", "?/8", "0 ??/??", "# #/#", "#,##0 ?/?", "# ?/?;-# ?/?",
]

DATE = [
    "m/d/yyyy", "d/m/yyyy", "dd/mm/yyyy", "dd/mm/yy", "d/m/yy", "d-mmm-yy", "d-mmm", "mmm-yy", "dd. mmm yy",
    "d. mmmm yyyy", "dddd, d. mmmm yyyy", "ddd, dd.mm.yyyy", "dd.mm.yyyy", "d.m.", "yyyy-mm-dd", "yyyy", "yy",
    "mmmm", "mmm", "mmmmm", "mm", "m", "dddd", "ddd", "dd", "d", "mmmm yyyy", "mmm yy", "[$-407]mmmm yyyy",
    "[$-409]mmmm yyyy", "[$-409]dddd, mmmm d, yyyy", "[$-409]d-mmm-yy", "[$-F800]dddd, mmmm dd, yyyy",
    "[$-407]d. mmm. yyyy", "[$-407]dddd, d. mmmm yyyy", "[$-407]mmm yy", "[$-407]d/ mmm/ yyyy", "d/m/yyyy h:mm",
    "dd/mm/yyyy hh:mm:ss", "m/d/yyyy h:mm", "yyyy-mm-dd hh:mm:ss", "e", "ee", "yyyymmdd", "ddddd", "mmmmmm",
    "yyy", "y", "dd/mm/yyyy;@", "[$-407]dd/mm/yyyy;@", "d/m;@", "[$-407]mmmmm;@", "[$-407]mmmmm yy;@", "yy-mm",
    "d\\.m\\.yyyy", "dd\\.mm\\.yyyy", "dd.mm.yy", "mm/yyyy", "General dd/mm",
]

TIME = [
    "h:mm", "hh:mm", "h:mm:ss", "hh:mm:ss", "h:mm AM/PM", "h:mm:ss AM/PM", "hh:mm am/pm", "h:mm A/P", "h:mm a/p",
    "h AM/PM", "mm:ss", "mm:ss.0", "[h]:mm:ss", "[h]:mm", "[mm]:ss", "[ss]", "[m]", "[hh]:mm", "h:mm:ss.000",
    "ss.00", "[h]:mm:ss.0", "h", "s", "m:ss", "h:m", "hh:mm:ss.00", "[$-F400]h:mm:ss AM/PM", "[$-409]h:mm:ss AM/PM",
    "[$-407]h:mm:ss", "yyyy-mm-dd h:mm AM/PM", "[h]", "[s].00", "h:mm;@", "[$-407]h:mm:ss;@", "[$-409]h:mm AM/PM;@",
    "hh\\h mm", 'h" Std. "mm" Min."', "h:mm AM/PM;[Red]h:mm", "[mm]", "[h]:mm:ss;@", "mm:ss.000", "AM/PM h:mm",
]

TEXT = [
    "@", '"Text: "@', '@" (x)"', ";;;", "0;0;0;@*.", "@*-", "@@", '@" "@', "*-@", "[Red]@", "0;0;0;[Blue]@",
]

FILL = ["*-0", "0*-", "* #,##0", "#,##0.00 *€", "**0", "*x0.00", "0*0", "*.@"]

COND = [
    "[Red]0", "[Blue]0.00;[Red]-0.00", "[Green]0;[Magenta]-0;[Cyan]0;[Yellow]@", "[Color10]0", "[Color3]0",
    "[Black]0;[White]-0", "[<100]0;[>=100]0.00", '[<0]"neg";[>0]"pos";"null"', "[Red][<=0]0;[Blue][>0]0",
    "[>=1000]#,##0;0.00", "[<1]0.000;0", '[=0]"-";General', '[<>0]0;"zero"', "[>100][Red]0;[<-100][Blue]0;0",
    "[Red]General", "0;[Red]0", '0.00;-0.00;"0"', "0;;", ";0", "0;-0;", ";;0", "[Color56]0", "[Color 5]0",
    "[<=99999]0000;0", "[>0.5]0.00;0", "[Red][>1]0;[Blue][<-1]0;0.0", "[Rot]0", "[<0]0;0", "[<0]-0;0",
    "[>1]0;[>0]0.0;0.00", "[>=1]0;0.0;\"z\"", "[Magenta]0;[Red]-0", "[RED]0", "[Color 10]0",
]

SPECIAL = ["00000", "0000", '"D-"00000', "000 00 000", "0#########", "#,##0 \"Stk.\"", "[<=9999999]###-####;(###) ###-####",
           "000-00-0000", "00000-0000", "\\(0\\)", "0000\\ 0000\\ 0000\\ 0000"]

FORMATS = []
for group in (NUMBER, CURRENCY, PERCENT, SCIENTIFIC, FRACTION, DATE, TIME, TEXT, FILL, COND, SPECIAL):
    for f in group:
        if f not in FORMATS:
            FORMATS.append(f)

# Rows below the numbers: other kinds of values, as formulas (the English LCID).
OTHER = [("text", '="Text"'), ("empty text", '=""'), ("true", "=TRUE"), ("false", "=FALSE"), ("#N/A", "=NA()"),
         ("#DIV/0!", "=1/0"), ("#VALUE!", '=1+"a"'), ("#NUM!", "=SQRT(-1)"), ("number text", '="12"'),
         ("negative text", '="-5"')]

WIDTH_FORMATS = ["General", "0.00", "#,##0", "0%", "m/d/yyyy", "h:mm:ss", "0.00E+00", "# ?/?", "@", "0", "[h]:mm",
                 '#,##0.00\\ "€"', "General;-General", '0" kg"', "#,##0.00_);(#,##0.00)", "0.00;;", ";;;"]
WIDTH_VALUES = [
    0, 1, -1, 7, 12, -12.5, 0.5, -0.5, 3.14159265358979, 1234567.891, -1234567.891, 123456789012, 0.000123456,
    1 / 3, -1 / 3, 2 / 3, 12.345, 99.995, 1e-10, 1.5e100, 45000.5, 0.75, 99999999999, 100, 1000, 12345, 123456,
    0.001, 0.00001, 1e15, 1e-15, 9.99999, 0.0999999, 1234.5678, 0.12345, 10000000, 98765.4321, 0.95, 0.05,
    2958465, -45000.5, 1.0000001, 55.5555555, 9999999999,
]
WIDTHS = [0.5, 1, 1.5, 2, 2.5, 3, 3.5, 4, 4.5, 5, 5.5, 6, 6.5, 7, 7.5, 8, 8.43, 9, 9.5, 10, 10.5, 11, 11.5, 12,
          13, 14, 15, 16, 18, 20]

LOCAL = [
    "Standard", "0", "0,00", "#.##0", "#.##0,00", "#.##0,00 €", "#.##0,00 [$€-407]", "0%", "0,00%", "0,00E+00",
    "# ?/?", "TT.MM.JJJJ", "T.M.JJ", "TT. MMM JJ", "TTTT, T. MMMM JJJJ", "JJJJ-MM-TT", "MMMM JJJJ", "MMM JJ",
    "hh:mm", "hh:mm:ss", "h:mm AM/PM", "[hh]:mm:ss", "[mm]:ss", "TT.MM.JJJJ hh:mm", "@", '"Text"@', "[Rot]0",
    "[Blau]0;[Rot]-0", "[Grün]0", "[Gelb]0", "[Farbe10]0", "[Schwarz]0", "[Weiß]0", "[Magenta]0", "[Zyan]0",
    '0,0 "kg"', "#.##0.", "0,0..", "#.##0,00_ ;[Rot]-#.##0,00\\ ", '_-* #.##0,00 €_-;-* #.##0,00 €_-;_-* "-"?? €_-;_-@_-',
    "MM:SS,0", "mm:ss,0", "hh:mm:ss,000", "T", "TT", "TTT", "TTTT", "M", "MM", "MMM", "MMMM", "MMMMM", "J", "JJ",
    "JJJJ", "h", "s", "m", "mm", "ss", "[h]", "[m]", "[s]", "TT.MM.", "T. MMM", "MM.JJJJ", "TT/MM/JJJJ", "TT-MM-JJJJ",
    "JJJJMMTT", "hh:mm TT.MM.", "[$-409]TTTT, MMMM T, JJJJ", "[$-F800]TTTT, MMMM TT, JJJJ", "[$-F400]h:mm:ss AM/PM",
    "0,00 \\€", "\\T\\T", '"TT"TT', "[<100]0;[>=100]0,00", "[Rot][<=0]0;[Blau][>0]0", "0,0#", "#.##0,0..\" Mio\"",
    "[$EUR] #.##0,00", "0,00E-00", "TT.MM.JJJJ;@", "e", "E", "jjjj", "tt.mm.jjjj", "HH:MM", "Tt.Mm.Jjjj", "S", "SS",
    "MM:SS", "hh:MM", "[MM]:SS", "[HH]:mm", "[SS]", "AM/PM", "mmm", "MMM", "Standard;-Standard", "STANDARD", "standard",
    "0,00_ ;-0,00\\ ", "#.##0 €;[Rot]-#.##0 €", '#.##0,00 "€"', "# ??/??", "0.0", "#,##0", "0;0;0;@", "G/Standard",
]


# Zellen formatieren > Zahlen: the lists of the categories as German Excel offers them (from the
# dialog as remembered; COM cannot open it invisibly), each checked here: valid, file form, sample.
CATEGORIES = {
    "Zahl": ["0,00", "0,00;[Rot]0,00", r"0,00_ ;-0,00\ ", r"0,00_ ;[Rot]-0,00\ ", "#.##0,00", "#.##0,00;[Rot]#.##0,00",
             r"#.##0,00_ ;-#.##0,00\ ", r"#.##0,00_ ;[Rot]-#.##0,00\ "],
    "Währung": ["#.##0,00 €", "#.##0,00 €;[Rot]#.##0,00 €", "#.##0,00 €;-#.##0,00 €", "#.##0,00 €;[Rot]-#.##0,00 €"],
    "Buchhaltung": ['_-* #.##0,00 €_-;-* #.##0,00 €_-;_-* "-"?? €_-;_-@_-', '_-* #.##0,00_-;-* #.##0,00_-;_-* "-"??_-;_-@_-'],
    "Datum": ["TT.MM.JJJJ", "[$-F800]TTTT, MMMM TT, JJJJ", "TT.MM.JJ;@", "JJJJ-MM-TT;@", "[$-407]TT. MMM JJ;@",
              "[$-407]T. MMMM JJJJ;@", "[$-407]MMMMM;@", "[$-407]MMMMM JJ;@", "T.M.;@", "T.M.JJ;@", "TT.MM.JJJJ;@",
              "[$-407]T. MMM.;@", "[$-407]T. MMM. JJ;@", "[$-407]MMM. JJ;@", "[$-407]MMMM JJ;@", "T.M.JJ h:mm;@",
              "[$-409]T.M.JJ h:mm AM/PM;@"],
    "Uhrzeit": ["[$-F400]h:mm:ss AM/PM", "hh:mm;@", "[$-409]h:mm AM/PM;@", "hh:mm:ss;@", "[$-409]h:mm:ss AM/PM;@",
                "mm:ss,0;@", "[h]:mm:ss;@", "[$-409]T.M.JJ h:mm AM/PM;@", "T.M.JJ h:mm;@"],
    "Prozent": ["0,00%"],
    "Bruch": ["# ?/?", "# ??/??", "# ???/???", "# ?/2", "# ?/4", "# ?/8", "# ??/16", "# ?/10", "# ??/100"],
    "Wissenschaft": ["0,00E+00"],
    "Text": ["@"],
    "Sonderformat": ["00000"],
}
# Excel's sample in the dialog: 14.03.2012 13:30:55.
SAMPLE = 41348.563136574


# Random numbers of every size (fixed seed): General's rounding and switching to scientific,
# decimals and fractions on values nobody picked by hand.
RANDOM_FORMATS = ["General", "0.00", "#,##0.000", "0.00E+00", "##0.00E+0", "# ??/??", "# ???/???", "0.0%", "?/?",
                  "0.###############", "#,##0", "0.0000000000"]
RANDOM_WIDTHS = [2, 3, 4, 5, 6, 7, 8.43, 9, 10, 11, 12]
FRACTION_LIMIT_FORMATS = ["?/8", "?/?", "??/??", "0/0", "# ?/8", "?/16", "?/2", "??/100", "#/#"]
FRACTION_LIMIT_VALUES = [100.5, 1000.25, 4095.875, 4096.125, 8191.5, 16383.75, 32767 / 8, 4096, 8192, 32767, 32768,
                         65535, 65536, 99999, 100000, 262143.5, 1e6, 2147483647, 2147483648, 4294967295, 1e10, -4096.5]


def random_values() -> list:
    import random
    r = random.Random(20261027)
    out = []
    for _ in range(400):
        kind = r.random()
        exp = r.randint(-12, 16)
        if kind < 0.4:
            v = r.random() * 10 ** exp
        elif kind < 0.7:
            v = round(r.random() * 10 ** r.randint(0, 9), r.randint(0, 6))
        elif kind < 0.85:
            v = r.randint(0, 10 ** r.randint(1, 15))
        else:
            v = round(r.random() * 100000, 5)
        if r.random() < 0.3:
            v = -v
        out.append(v)
    return out


def random_widths(sh, values: list) -> dict:
    n = len(values)
    last = col_letter(len(RANDOM_WIDTHS) - 1)
    for j, w in enumerate(RANDOM_WIDTHS):
        sh.Columns(j + 1).ColumnWidth = w
    cols = [{"width": sh.Columns(j + 1).ColumnWidth, "points": sh.Columns(j + 1).Width} for j in range(len(RANDOM_WIDTHS))]
    rng = sh.Range(f"A1:{last}{n}")
    rng.Value2 = tuple(tuple(v for _ in RANDOM_WIDTHS) for v in values)
    texts = [[sh.Cells(i + 1, j + 1).Text for j in range(len(RANDOM_WIDTHS))] for i in range(n)]
    sh.Cells.Clear()
    return {"columns": cols, "texts": texts}


def gdi_widths(chars: str) -> dict:
    """Advance widths in pixels of Calibri 11 at the screen's DPI, as GDI (and Excel) measures."""
    user32, gdi32 = ctypes.windll.user32, ctypes.windll.gdi32
    try:
        user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
    except Exception:  # noqa: BLE001
        user32.SetProcessDPIAware()
    user32.GetDC.restype = wintypes.HDC
    user32.ReleaseDC.argtypes = [wintypes.HWND, wintypes.HDC]
    gdi32.GetDeviceCaps.argtypes = [wintypes.HDC, ctypes.c_int]
    gdi32.SelectObject.argtypes = [wintypes.HDC, wintypes.HGDIOBJ]
    gdi32.DeleteObject.argtypes = [wintypes.HGDIOBJ]
    gdi32.GetTextExtentPoint32W.argtypes = [wintypes.HDC, wintypes.LPCWSTR, ctypes.c_int, ctypes.POINTER(wintypes.SIZE)]
    hdc = user32.GetDC(0)
    dpi = gdi32.GetDeviceCaps(hdc, 88)
    gdi32.CreateFontW.restype = wintypes.HFONT
    hf = gdi32.CreateFontW(-round(11 * dpi / 72), 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 5, 0, "Calibri")
    gdi32.SelectObject(hdc, hf)
    size = wintypes.SIZE()
    out = {}
    for ch in chars:
        gdi32.GetTextExtentPoint32W(hdc, ch, 1, ctypes.byref(size))
        out[ch] = size.cx
    gdi32.DeleteObject(hf)
    user32.ReleaseDC(0, hdc)
    return {"dpi": dpi, "px": round(11 * dpi / 72), "widths": out}


def col_letter(i: int) -> str:
    s = ""
    n = i + 1
    while n:
        n, m = divmod(n - 1, 26)
        s = chr(65 + m) + s
    return s


def table(sh, formats, values, others, width=None):
    """Formats across columns, values down rows; returns per format the format's forms and texts."""
    n = len(values)
    last = col_letter(len(formats) - 1)
    rng = sh.Range(f"A1:{last}{n}")
    rng.Value2 = tuple(tuple(v for _ in formats) for v in values)
    for k, (_, f) in enumerate(others):
        put(sh.Range(f"A{n + 1 + k}:{last}{n + 1 + k}"), "Formula", f)
    if width is not None:
        sh.Range(f"A1:{last}1").EntireColumn.ColumnWidth = width
    out = []
    for j, code in enumerate(formats):
        col = sh.Range(f"{col_letter(j)}1:{col_letter(j)}{n + len(others)}")
        rec: dict = {"code": code}
        try:
            put(col, "NumberFormat", code)
        except Exception as e:  # noqa: BLE001
            rec["invalid"] = True
            rec["error"] = str(e)[:60]
            out.append(rec)
            continue
        rec["en"] = get(col.Cells(1, 1), "NumberFormat")
        rec["local"] = col.Cells(1, 1).NumberFormatLocal
        rec["texts"] = [col.Cells(i + 1, 1).Text for i in range(n)]
        rec["other"] = [col.Cells(n + 1 + k, 1).Text for k in range(len(others))]
        out.append(rec)
    sh.Cells.Clear()
    return out


def widths_section(sh) -> list:
    out = []
    n = len(WIDTH_VALUES)
    last = col_letter(len(WIDTHS) - 1)
    for j, w in enumerate(WIDTHS):
        sh.Columns(j + 1).ColumnWidth = w
    cols = [{"width": sh.Columns(j + 1).ColumnWidth, "points": sh.Columns(j + 1).Width} for j in range(len(WIDTHS))]
    for code in WIDTH_FORMATS:
        rng = sh.Range(f"A1:{last}{n}")
        rng.Value2 = tuple(tuple(v for _ in WIDTHS) for v in WIDTH_VALUES)
        try:
            put(rng, "NumberFormat", code)
        except Exception as e:  # noqa: BLE001
            print("width format failed", code, e)
            continue
        texts = [[sh.Cells(i + 1, j + 1).Text for j in range(len(WIDTHS))] for i in range(n)]
        out.append({"code": code, "texts": texts})
        # text values for "@" and General
        rng.ClearContents()
    strings = ["abc", "Hallo Welt", "12345678901234567890"]
    sh.Cells.Clear()
    rng = sh.Range(f"A1:{last}{len(strings)}")
    rng.Value2 = tuple(tuple(s for _ in WIDTHS) for s in strings)
    stexts = [[sh.Cells(i + 1, j + 1).Text for j in range(len(WIDTHS))] for i in range(len(strings))]
    sh.Cells.Clear()
    return {"columns": cols, "values": WIDTH_VALUES, "formats": out, "strings": strings, "stringTexts": stexts}


def local_section(sh, codes: list, value: float = 45000.5) -> list:
    """Each German code on a cell of its own: the English form Excel reports, the German form it
    shows back, the text, and the code Excel writes into the file (saved to TEMP and read back)."""
    out = []
    for k, loc in enumerate(codes):
        c = sh.Range(f"A{k + 1}")
        c.Value2 = value
        rec: dict = {"local": loc}
        try:
            c.NumberFormatLocal = loc
            rec["en"] = get(c, "NumberFormat")
            rec["back"] = c.NumberFormatLocal
            rec["text"] = c.Text
        except Exception as e:  # noqa: BLE001
            rec["invalid"] = True
            rec["error"] = str(e)[:60]
        out.append(rec)
    written = saved_formats(sh.Parent, [f"A{k + 1}" for k in range(len(codes))])
    for rec, w in zip(out, written):
        if not rec.get("invalid"):
            rec["file"] = w
    return out


def saved_formats(wb, cells: list) -> list:
    """Saves the workbook as .xlsx in TEMP and returns (numFmtId, code) of each cell's format."""
    tmp = Path(tempfile.mkdtemp(prefix="raabe-numfmt-"))
    path = tmp / "saved.xlsx"
    wb.SaveCopyAs(str(path))
    with zipfile.ZipFile(path) as z:
        styles = z.read("xl/styles.xml").decode("utf-8")
        sheet = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
    numfmts = {m.group(1): _unxml(m.group(2)) for m in re.finditer(r'<numFmt numFmtId="(\d+)" formatCode="([^"]*)"', styles)}
    xf_part = re.search(r"<cellXfs[^>]*>(.*?)</cellXfs>", styles, re.S).group(1)
    xf_ids = [int(m.group(1)) for m in re.finditer(r'<xf numFmtId="(\d+)"', xf_part)]
    out = []
    for cell in cells:
        m = re.search(rf'<c r="{cell}" s="(\d+)"', sheet)
        fid = xf_ids[int(m.group(1))] if m else 0
        out.append({"id": fid, "code": numfmts.get(str(fid))})
    path.unlink()
    tmp.rmdir()
    return out


def builtin_section(x) -> dict:
    """A .xlsx whose cells use numFmtId 0-81 without defining them: what Excel reads for each id,
    and what it writes back when it saves the workbook."""
    tmp = Path(tempfile.mkdtemp(prefix="raabe-numfmt-"))
    src = tmp / "builtin.xlsx"
    ids = list(range(0, 82))
    xfs = "".join(f'<xf numFmtId="{i}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' for i in ids)
    rows = "".join(f'<row r="{k + 1}"><c r="A{k + 1}" s="{k}"><v>45000.5</v></c><c r="B{k + 1}" s="{k}"><v>-1234.5</v></c></row>'
                   for k, _ in enumerate(ids))
    files = {
        "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
        "_rels/.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Tabelle1" sheetId="1" r:id="rId1"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
        "xl/styles.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="{len(ids)}">{xfs}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>',
        "xl/worksheets/sheet1.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols><col min="1" max="2" width="60" customWidth="1"/></cols><sheetData>{rows}</sheetData></worksheet>',
    }
    with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data)
    wb = x.app.Workbooks.Open(str(src))
    out: dict = {"read": {}, "written": {}}
    try:
        sh = wb.Worksheets(1)
        for k, i in enumerate(ids):
            a, b = sh.Range(f"A{k + 1}"), sh.Range(f"B{k + 1}")
            out["read"][i] = {"en": get(a, "NumberFormat"), "local": a.NumberFormatLocal, "text": a.Text, "neg": b.Text}
        # Each code set anew on a fresh cell, then saved: which numFmtId and code does Excel write?
        for k, i in enumerate(ids):
            c = sh.Range(f"D{k + 1}")
            c.Value2 = 1
            put(c, "NumberFormat", out["read"][i]["en"])
        saved = tmp / "saved.xlsx"
        wb.SaveAs(str(saved), 51)
    finally:
        wb.Close(False)
    with zipfile.ZipFile(saved) as z:
        styles = z.read("xl/styles.xml").decode("utf-8")
        sheet = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
    numfmts = {m.group(1): _unxml(m.group(2)) for m in re.finditer(r'<numFmt numFmtId="(\d+)" formatCode="([^"]*)"', styles)}
    xf_part = re.search(r"<cellXfs[^>]*>(.*?)</cellXfs>", styles, re.S).group(1)
    xf_ids = [int(m.group(1)) for m in re.finditer(r'<xf numFmtId="(\d+)"', xf_part)]
    for k, i in enumerate(ids):
        m = re.search(rf'<c r="D{k + 1}" s="(\d+)"', sheet)
        if not m:
            continue
        fid = xf_ids[int(m.group(1))]
        out["written"][i] = {"id": fid, "code": numfmts.get(str(fid))}
    for f in tmp.iterdir():
        f.unlink()
    tmp.rmdir()
    return out


def _unxml(s: str) -> str:
    return s.replace("&quot;", '"').replace("&lt;", "<").replace("&gt;", ">").replace("&apos;", "'").replace("&amp;", "&")


class Book:
    """A workbook of our own with Calibri 11 as the Standard style's font. Excel keeps only about
    250 custom number formats per workbook, so each batch of formats gets a new one."""

    def __init__(self, x):
        self.wb = x.app.Workbooks.Add()

    def __enter__(self):
        st = self.wb.Styles("Standard")
        st.Font.Name = "Calibri"
        st.Font.Size = 11
        return self.wb.Worksheets(1)

    def __exit__(self, *exc):
        self.wb.Close(False)


def chunks(items: list, n: int):
    for i in range(0, len(items), n):
        yield items[i:i + n]


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    chars = "".join(chr(c) for c in range(32, 127)) + "ÄÖÜäöüß€£§°µ"
    measure = gdi_widths(chars)
    x = ExcelTarget(visible=False)
    out: dict = {"date": str(datetime.date.today()), "excel": x.app.Version, "measure": measure}
    try:
        print("formats", len(FORMATS), "values", len(VALUES))
        out["values"] = VALUES
        out["other"] = [k for k, _ in OTHER]
        out["formats"] = []
        for part in chunks(FORMATS, 100):
            with Book(x) as sh:
                sh.Cells.ColumnWidth = 120
                out["formats"] += table(sh, part, VALUES, OTHER)
        print("widths")
        with Book(x) as sh:
            out["standardWidth"] = {"width": sh.StandardWidth, "points": sh.Columns(1).Width}
            out["widths"] = widths_section(sh)
        print("local")
        out["local"] = []
        for part in chunks(LOCAL, 100):
            with Book(x) as sh:
                out["local"] += local_section(sh, part)
        print("categories")
        out["categories"] = {}
        for name, codes in CATEGORIES.items():
            with Book(x) as sh:
                out["categories"][name] = local_section(sh, codes, SAMPLE)
        with Book(x) as sh:
            out["categoryNegative"] = local_section(sh, [c for cs in CATEGORIES.values() for c in cs], -1234.1)
        print("1904")
        with Book(x) as sh:
            sh.Parent.Date1904 = True
            sh.Cells.ColumnWidth = 120
            dates = [f for f in DATE + TIME if not f.startswith("General")]
            vals = [0, 1, 1.5, 59, 60, 61, 366, 45000.5, -1, 2957003, 2957004]
            out["date1904"] = {"values": vals, "formats": table(sh, dates, vals, [])}
        print("random")
        rnd = random_values()
        out["random"] = {"values": rnd, "formats": []}
        for part in chunks(RANDOM_FORMATS, 100):
            with Book(x) as sh:
                sh.Cells.ColumnWidth = 120
                out["random"]["formats"] += table(sh, part, rnd, [])
        with Book(x) as sh:
            out["randomWidths"] = random_widths(sh, rnd)
        print("fraction limits")
        with Book(x) as sh:
            sh.Cells.ColumnWidth = 120
            out["fractionLimits"] = {"values": FRACTION_LIMIT_VALUES,
                                     "formats": table(sh, FRACTION_LIMIT_FORMATS, FRACTION_LIMIT_VALUES, [])}
        print("builtin")
        out["builtin"] = builtin_section(x)
    finally:
        x.close()
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))
    print("written", OUT.name)


if __name__ == "__main__":
    main()
