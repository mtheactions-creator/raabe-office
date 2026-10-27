"""Records how Excel computes and writes formulas, as the reference for RaabeCalc's formula
engine (Step 7 package C, D-062).

Starts an invisible Excel of its own (`ExcelTarget(visible=False)`), fills two sheets with a
fixed data block, then enters every case through `Formula2Local` (the German form a user
types) and reads back `Formula2` (the model's file form), `Formula2Local`, `Value2`, `Text`
and, for dynamic arrays, the whole spill range. Further sections copy formulas to other cells
(Range.Copy with a destination, no clipboard), insert and delete rows and columns, cut and
paste cells, insert and delete cells with shifting, and save a workbook to TEMP to read which
functions carry the `_xlfn.` prefix in the file.

    .venv/Scripts/python.exe probe_excel_formulas.py [cases] [shift] [adjust] [move] [cells] [xlfn]

Sections not named keep their earlier results in the fixture.

Writes packages/sheet-engine/test/fixtures/excel-formulas.json. No real input, no visible
window; only our own Excel process is touched."""

from __future__ import annotations

import json
import os
import re
import sys
import tempfile
import time
import zipfile
from pathlib import Path

from feellab.excel import ExcelTarget

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "packages" / "sheet-engine" / "test" / "fixtures" / "excel-formulas.json"

ERRORS = {2000: "#NULL!", 2007: "#DIV/0!", 2015: "#VALUE!", 2023: "#REF!", 2029: "#NAME?", 2036: "#NUM!",
          2042: "#N/A", 2043: "#GETTING_DATA", 2045: "#SPILL!", 2047: "#BLOCKED!", 2050: "#CALC!"}

# ----- the data block ----------------------------------------------------------------------
# (sheet, address, value) with values as Value2 (numbers, text, booleans) or "=..." formulas
# in the German form. Dates and times are serial numbers.
SETUP: list[tuple[str, str, object]] = []


def col(sheet: str, letter: str, values: list, start: int = 1) -> None:
    for i, v in enumerate(values):
        if v is not None:
            SETUP.append((sheet, f"{letter}{start + i}", v))


T1, T2 = "Tabelle1", "Daten 2"
col(T1, "A", [3, 1, 4, 1, 5, 9, 2, 6, 5, 3])
col(T1, "B", ["Apfel", "Birne", "apfel", "Kirsche", "Banane", "Pflaume", "Apfelsaft", "a*b", "?", "~"])
col(T1, "C", [1, "a", True, None, 2.5, '=""', "'10", "=1/0", -3, False])
col(T1, "D", [45292, 45337, 45351, 45291, 61, 59, 45818, 1, 45322, 45473])
col(T1, "E", [10, 20, 30, 40, 50, 60, 70, 80, 90, 100])
col(T1, "F", ["Anton", "Berta", "Cäsar", "Dora", "Emil", "Friedrich", "Gustav", "Heinrich", "Ida", "Julius"])
col(T1, "G", [2.5, -1, 0, 7, 3.25, 100, -20, 0.5, 8, 0.001])
col(T1, "H", ["x", "y", "x", "z", "y", "x", "x", "z", "y", "x"])
col(T1, "I", [0.25, 0.5, 0.75, 1 / 3, 1.5, 0.999988425925926, 0.0000115740740740741, 45292.5, 0.6, 0.041666666666666664])
col(T1, "J", ["1,5", " 7 ", "12.3.2024", "abc", "", "1E3", "50%", "WAHR", "-4", "3 €"])
col(T1, "K", [1, 2, 3], 1)
col(T1, "L", [4, 5, 6], 1)
col(T1, "M", [7, 8, 9], 1)
col(T1, "K", ["Name", "Alter", "Ort"], 5)
col(T1, "L", ["Max", 15, "Celle"], 5)
col(T1, "M", ["Mia", 16, "Peine"], 5)
col(T1, "N", ["Ole", 15, "Celle"], 5)
col(T1, "K", [1, 1, 2, 2, 3], 10)
col(T1, "L", [2, 4, 5, 4, 5], 10)
col(T1, "N", [1, 2, 3, 4, 5, 6, 7, 8, 9, 10], 11)
col(T2, "A", [10, 20, 30, 40, 50])
col(T2, "B", ["eins", "zwei", "drei", "vier", "fünf"])
col(T2, "C", ["=A1*2", "=SUMME(A1:A5)", "=Tabelle1!A1+1"])
NAMES = [("Preise", "=Tabelle1!$E$1:$E$10"), ("Steuer", "=0,19"), ("Namen", "=Tabelle1!$F$1:$F$10")]

# ----- the cases ---------------------------------------------------------------------------
CASES: list[dict] = []


def case(group: str, *formulas: str, addr: str = "P5") -> None:
    for f in formulas:
        CASES.append({"group": group, "local": f, "addr": addr})


def each(group: str, template: str, args: list[str], addr: str = "P5") -> None:
    for a in args:
        case(group, template.replace("@@", a), addr=addr)


NUMS = ["0", "1", "-1", "2,5", "-2,5", "0,5", "-0,5", "1,005", "123,456", "-123,456", "1E+15", "1E-10", "\"3\"", "WAHR", "\"abc\"", "C4", "C8", "#NV"]

# Operators, precedence, coercion.
case("ops", "=1+2*3", "=(1+2)*3", "=2^3^2", "=-2^2", "=2^-2", "=-A1^2", "=1-0,9-0,1", "=0,1+0,2-0,3", "=(0,1+0,2-0,3)",
     "=1*(0,1+0,2-0,3)", "=0,3-0,2-0,1", "=0,1+0,2", "=0,1*3", "=1/3", "=2/3*3", "=10/4", "=5%", "=50%%", "=200%*3", "=-5%",
     "=1/0", "=0/0", "=0^0", "=(-8)^(1/3)", "=(-2)^0,5", "=2^1024", "=1E+308*10", "=-1E+308*10", "=1E-308/1E+10",
     "=\"2\"+3", "=\"2,5\"*2", "=\"abc\"+1", "=WAHR+WAHR", "=FALSCH-1", "=C4+1", "=C4&\"x\"", "=C6+1", "=C6&\"x\"", "=\" 4 \"+1",
     "=\"1.000\"+1", "=\"1.000,5\"+1", "=\"1E3\"+0", "=\"50%\"+0", "=\"(5)\"+0", "=\"5 €\"+0", "=\"€5\"+0", "=\"-5\"*1", "=\"+5\"*1",
     "=\"12:30\"+0", "=\"1.1.2024\"+0", "=\"2024-01-15\"+0", "=\"15.1.24\"+0", "=\"1. Januar 2024\"+0", "=\"Jan 2024\"+0",
     "=\"31.2.2024\"+0", "=\"12:30:45\"*1", "=\"25:00\"*1", "=\"1.1.2024 12:00\"*1", "=\"3/4\"+0", "=\"0 3/4\"+0", "=\"1,5E+2\"+0",
     "=1&2", "=1,5&\"\"", "=WAHR&\"\"", "=\"a\"&#NV", "=#NV&#DIV/0!", "=1/0+#NV", "=#NV+1/0", "=-\"3\"", "=+\"3\"", "=--\"3\"",
     "=1=1", "=1=\"1\"", "=\"a\"=\"A\"", "=\"a\"<\"B\"", "=\"Z\"<\"a\"", "=1<\"a\"", "=\"a\"<WAHR", "=1<WAHR", "=FALSCH<WAHR",
     "=\"\"<\"a\"", "=C4=0", "=C4=\"\"", "=C4=FALSCH", "=C4<1", "=C4<\"a\"", "=C6=C4", "=0,1+0,2=0,3", "=0,3=0,1+0,2",
     "=(0,1+0,2-0,3)=0", "=1E-16=0", "=1+1E-15=1", "=1+1E-16=1", "=\"ä\"=\"a\"", "=\"ä\">\"a\"", "=\"ä\"<\"b\"", "=\"ß\"=\"ss\"",
     "=\"a-b\"<\"ab\"", "=\"a b\"<\"ab\"", "=\"10\"<\"9\"", "=\"abc\"<\"abd\"", "=\"Ab\"=\"aB\"", "=\"a\"<>\"A\"",
     "=1<>1", "=2>=2", "=2<=1", "=\"x\"&1/0", "=1=1=1", "=1+\"\"", "=\"\"&\"\"", "=A1:A3+1", "=A1:A3*{1;2;3}", "={1.2.3}+{10;20}",
     "={1.2}+{1.2.3}", "={1;2}*{3;4;5}", "=A1:B2", "=-A1:A2", "=A1:A3=1", "=A1:A3&\"x\"", "=(A1:A3)", "=A1:A3%",
     "=1E+15+1", "=123456789012345678", "=0,1+0,7", "=4,35*100", "=1,1*1,1", "=3*1,1", "=100*1,1", "=0,57*100", "=1-0,99",
     "=5,1-5", "=1E+20+1-1E+20", "=2^53+1", "=10^15+0,3", "=\"1,5\"=1,5", "=WAHR=1", "=1/3*3=1")
case("ops-ws", "= 1 + 2", "=SUMME( A1 ; A2 )", "=SUMME(A1:A3 ; 2)", "=WENN( A1 >1 ; 1 ; 2 ) ", "=A1 :A2", "=- -1", "=1 %")
case("refs", "=A1", "=$A$1", "=A$1+$A1", "=C4", "=C6", "=C2", "=C3", "=C8", "='Daten 2'!A1", "='Daten 2'!C2", "='Daten 2'!B5",
     "=SUMME('Daten 2'!A1:A5)", "=SUMME('Daten 2'!A:A)", "=SUMME(A:A)", "=SUMME(1:1)", "=SUMME($A:$A)", "=SUMME(A1:A3 A2:B2)",
     "=SUMME(A1:B3 C1:D3)", "=SUMME((A1:A2;A3))", "=ANZAHL((A1:A2;A3;B1))", "=A1:A3 A2:C2", "=SUMME(A1:B2:C3)",
     "=Preise", "=SUMME(Preise)", "=Steuer*100", "=INDEX(Namen;3)", "=Unbekannt", "=XFD1", "=XFE1", "=A1048577", "=@A1:A3",
     "=@A2:A4", "=@B1:D1", "=Tabelle1!A1:Tabelle1!B2", "=SUMME(Tabelle1!A1:Tabelle1!B2)", "=ZEILE()", "=SPALTE()")
case("refs-at", "=A1:A10", addr="P3")
case("refs-at", "=@A1:A10", "=@E:E", "=@3:3", addr="P3")
case("refs-at", "=@A1:A10", addr="Q20")

# Math.
each("SUMME", "=SUMME(@@)", ["A1:A10", "B1:B10", "C1:C10", "C1:C7", "A1:A10;5", "1;2;3", "\"3\";4", "WAHR;2", "\"abc\";1", "{1.\"2\".WAHR}", "", "A1:A10;C8",
                              "G1:G10", "D1:D3", "0,1;0,2;-0,3", "1E+308;1E+308", "A1;;A2", "N1:N20", "A:A", "Preise", "-A1:A3", "A1:A3*2"])
each("PRODUKT", "=PRODUKT(@@)", ["A1:A5", "A1:A5;2", "C1:C7", "\"2\";3", "B1:B3", "1E+200;1E+200"])
each("RUNDEN", "=RUNDEN(@@)", ["2,5;0", "-2,5;0", "2,675;2", "1,005;2", "-1,005;2", "1234,5678;-2", "1234,5678;1,9", "0,285;2", "1,15;1", "1,25;1",
                              "1,35;1", "8,345;2", "0,1+0,2;15", "1E+20;2", "123456789,123456789;5", "5,5;-1", "-5,5;-1", "\"2,5\";0", "WAHR;0", "\"x\";0",
                              "2,5;\"x\"", "C4;0", "0,5;", "A1:A3;0", "1,23456789;300", "1,5;-300", "4,0000000000000005;0", "2,4999999999999996;0"])
each("AUFRUNDEN", "=AUFRUNDEN(@@)", ["2,1;0", "-2,1;0", "0,1+0,2;1", "1,0000000001;0", "3,2;-1", "-3,2;-1", "1,005;2", "0;0", "2,99999999999999;0", "4,35*100;0"])
each("ABRUNDEN", "=ABRUNDEN(@@)", ["2,9;0", "-2,9;0", "0,3*3;1", "1,005;2", "3,2;-1", "4,35*100;0", "0,57*100;0", "1-0,99;2"])
each("GANZZAHL", "=GANZZAHL(@@)", NUMS + ["4,35*100", "-0,0000001", "1E+300", "0,57*100"])
each("KÜRZEN", "=KÜRZEN(@@)", ["2,7", "-2,7", "2,789;2", "-2,789;2", "1234;-2", "4,35*100", "0,57*100;0", "\"x\""])
each("ABS", "=ABS(@@)", NUMS)
each("VORZEICHEN", "=VORZEICHEN(@@)", NUMS)
each("WURZEL", "=WURZEL(@@)", NUMS + ["2", "1E+300"])
each("POTENZ", "=POTENZ(@@)", ["2;10", "2;0,5", "-8;1/3", "0;0", "0;-1", "10;-2", "-2;3", "\"2\";2", "2;1024", "4;0,5"])
each("REST", "=REST(@@)", ["10;3", "-10;3", "10;-3", "-10;-3", "5,5;2", "1;0", "0;5", "10;0,3", "1E+15;7", "123456789012;3", "2^53;3", "3;2,5", "-1;1E+300", "0,3;0,1", "1E+20;3"])
each("QUOTIENT", "=QUOTIENT(@@)", ["10;3", "-10;3", "5,5;2", "1;0", "-7,5;2"])
each("PI", "=PI()@@", ["", "*2"])
each("EXP", "=EXP(@@)", ["0", "1", "-1", "710", "709", "\"x\""])
each("LN", "=LN(@@)", ["1", "2,718281828459045", "0", "-1", "10"])
each("LOG", "=LOG(@@)", ["100", "8;2", "1000", "1;10", "0", "10;1", "10;0", "100;-10", "0,001"])
each("LOG10", "=LOG10(@@)", ["1000", "0,1", "-1", "2"])
for fn, args in [("SIN", ["0", "PI()", "PI()/2", "PI()/6", "1E+10", "2^27", "2^28"]), ("COS", ["0", "PI()", "PI()/3", "1"]),
                 ("TAN", ["0", "PI()/4", "PI()/2", "1"]), ("ARCSIN", ["0", "1", "0,5", "2", "-1"]), ("ARCCOS", ["0", "1", "-1", "1,5"]),
                 ("ARCTAN", ["0", "1", "-1", "1E+300"]), ("ARCTAN2", ["1;1", "-1;1", "0;0", "0;1", "-1;0"]), ("BOGENMASS", ["180", "90", "-45"]),
                 ("GRAD", ["PI()", "1", "PI()/2"]), ("SINHYP", ["1", "0"]), ("COSHYP", ["1"]), ("TANHYP", ["1", "100"]), ("COT", ["1", "0"])]:
    each(fn, f"={fn}(@@)", args)
each("FAKULTÄT", "=FAKULTÄT(@@)", ["0", "1", "5", "5,9", "-1", "170", "171", "20", "\"3\""])
each("KOMBINATIONEN", "=KOMBINATIONEN(@@)", ["5;2", "10;0", "10;10", "5;6", "5,9;2,9", "-1;0", "49;6", "1000;500", "0;0"])
each("VARIATIONEN", "=VARIATIONEN(@@)", ["5;2", "10;0", "5;6", "5,9;2,9", "49;6"])
each("GGT", "=GGT(@@)", ["12;18", "0;5", "7;13", "12,7;18", "-4;6", "A1:A10", "48;36;24", "0", "2^53;6"])
each("KGV", "=KGV(@@)", ["4;6", "0;5", "3;5;7", "A1:A4", "-4;6", "1,5;2"])
each("GERADE", "=GERADE(@@)", ["1,5", "3", "2", "-1", "-2,5", "0", "\"x\""])
each("UNGERADE", "=UNGERADE(@@)", ["1,5", "3", "2", "-1", "-2,5", "0"])
each("OBERGRENZE.MATHEMATIK", "=OBERGRENZE.MATHEMATIK(@@)", ["4,3", "-4,3", "4,3;2", "-4,3;2", "-4,3;2;1", "24,3;5", "6,7;0", "4,3;-2", "0,1+0,2;0,1", "4,35*100"])
each("UNTERGRENZE.MATHEMATIK", "=UNTERGRENZE.MATHEMATIK(@@)", ["4,7", "-4,7", "4,7;2", "-4,7;2", "-4,7;2;1", "24,3;5", "4,7;-2", "0,7;0,1"])
each("OBERGRENZE", "=OBERGRENZE(@@)", ["4,3;1", "-4,3;-1", "-4,3;1", "4,3;-1", "2,5;0,5", "0;0", "4,3;0", "1,5;0,1", "-2,5;2"])
each("UNTERGRENZE", "=UNTERGRENZE(@@)", ["4,7;1", "-4,7;-1", "-4,7;1", "4,7;-1", "2,5;0,5", "4,7;0", "0;0", "-2,5;2", "0,7;0,1"])
each("VRUNDEN", "=VRUNDEN(@@)", ["10;3", "-10;-3", "1,3;0,2", "5;-2", "7,5;5", "-7,5;-5", "0,15;0,1", "1,25;0,5", "6,05;0,1", "0;0", "5;0"])
each("QUADRATESUMME", "=QUADRATESUMME(@@)", ["3;4", "A1:A10", "C1:C10", "\"2\";1", "{1.2.3}"])
each("SUMMENPRODUKT", "=SUMMENPRODUKT(@@)", ["A1:A10;E1:E10", "A1:A3;A1:A4", "A1:A10", "C1:C10;A1:A10", "(H1:H10=\"x\")*A1:A10", "--(A1:A10>3)",
                                             "(A1:A10>3)*(E1:E10)", "B1:B3;A1:A3", "{1.2.3};{4.5.6}", "A1:A3*2", "K1:M3;K1:M3", "1;2", "(A1:A10>3)+0", "A1:A10>3"])
each("ZUFALLSZAHL", "=@@", ["UND(ZUFALLSZAHL()>=0;ZUFALLSZAHL()<1)", "ISTZAHL(ZUFALLSBEREICH(1;6))", "UND(ZUFALLSBEREICH(3;3)=3)", "ZUFALLSBEREICH(5;1)",
                            "ZUFALLSBEREICH(1,5;1,6)", "ANZAHL(ZUFALLSMATRIX(2;3))"])
each("TEILERGEBNIS", "=TEILERGEBNIS(@@)", ["9;A1:A10", "1;A1:A10", "2;C1:C10", "3;C1:C10", "4;A1:A10", "5;A1:A10", "109;A1:A10", "12;A1:A10", "0;A1:A10", "6;A1:A3"])

# Statistics.
for fn in ["MITTELWERT", "MIN", "MAX", "ANZAHL", "ANZAHL2", "MEDIAN", "MODUS.EINF", "STABW.S", "STABW.N", "VAR.S", "VAR.P", "MITTELABW", "GEOMITTEL",
           "MAXA", "MINA", "MITTELWERTA", "STABW", "VARIANZ", "HARMITTEL", "KURT", "SCHIEFE"]:
    each(fn, f"={fn}(@@)", ["A1:A10", "C1:C10", "B1:B10", "G1:G10", "A1:A3;\"4\"", "WAHR;2", "D1:D10", "N1:N20", "A1;;A2", "\"x\";1", "C4", "E1:E10;A1:A10", "{1.2.\"3\"}"])
each("ANZAHLLEEREZELLEN", "=ANZAHLLEEREZELLEN(@@)", ["C1:C10", "A1:A10", "N1:N20", "J1:J10", "C4"])
each("RANG.GLEICH", "=RANG.GLEICH(@@)", ["5;A1:A10", "5;A1:A10;1", "1;A1:A10", "7;A1:A10", "9;A1:A10;0", "3;A1:A10;1", "2,5;G1:G10", "\"5\";A1:A10"])
each("RANG", "=RANG(@@)", ["5;A1:A10", "1;A1:A10;1"])
each("RANG.MITTELW", "=RANG.MITTELW(@@)", ["5;A1:A10", "1;A1:A10;1", "3;A1:A10"])
each("KGRÖSSTE", "=KGRÖSSTE(@@)", ["A1:A10;1", "A1:A10;3", "A1:A10;10", "A1:A10;11", "A1:A10;0", "C1:C10;1", "A1:A10;2,5", "{3.1.2};2", "A1:A10;{1;2}"])
each("KKLEINSTE", "=KKLEINSTE(@@)", ["A1:A10;1", "A1:A10;3", "A1:A10;10", "A1:A10;11", "G1:G10;2", "B1:B10;1"])
for fn in ["KORREL", "STEIGUNG", "ACHSENABSCHNITT", "BESTIMMTHEITSMASS", "KOVARIANZ.P", "KOVARIANZ.S", "STFEHLERYX"]:
    each(fn, f"={fn}(@@)", ["A1:A10;E1:E10", "E1:E10;A1:A10", "G1:G10;A1:A10", "A1:A10;A1:A9", "C1:C10;A1:A10", "K1:K3;L1:L3", "{1.2.3};{2.4.7}", "A1:A2;{1;1}", "{1;1;1};{2;3;4}"])
each("SCHÄTZER", "=SCHÄTZER(@@)", ["11;E1:E10;A1:A10", "4;A1:A10;E1:E10"])
each("PROGNOSE.LINEAR", "=PROGNOSE.LINEAR(@@)", ["11;E1:E10;A1:A10"])
for fn in ["QUANTIL.INKL", "QUANTIL", "QUANTIL.EXKL"]:
    each(fn, f"={fn}(@@)", ["A1:A10;0", "A1:A10;0,25", "A1:A10;0,5", "A1:A10;0,9", "A1:A10;1", "A1:A10;1,1", "A1:A10;-0,1", "G1:G10;0,33", "A1:A10;0,05"])
for fn in ["QUARTILE.INKL", "QUARTILE", "QUARTILE.EXKL"]:
    each(fn, f"={fn}(@@)", ["A1:A10;0", "A1:A10;1", "A1:A10;2", "A1:A10;3", "A1:A10;4", "A1:A10;5", "A1:A10;1,7", "G1:G10;3"])
each("QUANTILSRANG.INKL", "=QUANTILSRANG.INKL(@@)", ["A1:A10;4", "A1:A10;3,5", "A1:A10;10", "A1:A10;1;2"])
each("NORM.S.VERT", "=NORM.S.VERT(@@)", ["0;WAHR", "1;WAHR", "-1,96;WAHR", "1;FALSCH", "0;FALSCH", "5;WAHR", "-8;WAHR", "-40;WAHR", "1,5;1"])
each("NORM.VERT", "=NORM.VERT(@@)", ["42;40;1,5;WAHR", "42;40;1,5;FALSCH", "40;40;0;WAHR", "1;0;1;WAHR", "1;0;-1;WAHR", "100;50;10;WAHR"])
each("NORM.INV", "=NORM.INV(@@)", ["0,908789;40;1,5", "0,5;0;1", "0,025;0;1", "0;0;1", "1;0;1", "0,000001;0;1", "0,99;10;2", "0,5;0;0"])
each("NORM.S.INV", "=NORM.S.INV(@@)", ["0,975", "0,5", "0,001", "0,9999999", "1E-10"])
each("BINOM.VERT", "=BINOM.VERT(@@)", ["6;10;0,5;FALSCH", "6;10;0,5;WAHR", "0;10;0,3;FALSCH", "10;10;0,3;WAHR", "11;10;0,3;FALSCH", "3,7;10;0,5;FALSCH",
                                       "50;100;0,5;FALSCH", "500;1000;0,5;WAHR", "2;5;1,5;WAHR", "-1;5;0,5;WAHR", "2;5;0;FALSCH"])
each("BINOM.INV", "=BINOM.INV(@@)", ["10;0,5;0,5", "100;0,3;0,95"])
each("POISSON.VERT", "=POISSON.VERT(@@)", ["2;5;FALSCH", "2;5;WAHR", "0;0;WAHR", "10;2,5;FALSCH", "-1;2;WAHR", "2,9;5;FALSCH", "100;100;WAHR"])
each("EXPON.VERT", "=EXPON.VERT(@@)", ["0,2;10;WAHR", "0,2;10;FALSCH"])
each("T.VERT", "=T.VERT(@@)", ["1,96;60;WAHR", "1;5;FALSCH", "-2;3;WAHR"])
each("T.INV", "=T.INV(@@)", ["0,95;10", "0,05;3"])
each("T.INV.2S", "=T.INV.2S(@@)", ["0,05;10"])
each("T.TEST", "=T.TEST(@@)", ["A1:A10;E1:E10;2;1", "A1:A10;E1:E10;2;2", "A1:A10;E1:E10;1;3", "A1:A10;G1:G10;2;3", "A1:A10;G1:G10;1;1", "A1:A5;G1:G5;2;2", "A1:A10;E1:E9;2;1"])
each("STANDARDISIERUNG", "=STANDARDISIERUNG(@@)", ["42;40;1,5", "1;0;0"])
each("HÄUFIGKEIT", "=HÄUFIGKEIT(@@)", ["A1:A10;{2.4.6}", "G1:G10;{0;5}"])

# Logic and information.
each("WENN", "=WENN(@@)", ["WAHR;1;2", "FALSCH;1;2", "1;\"a\";\"b\"", "0;1", "FALSCH;1", "WAHR;;2", "\"WAHR\";1;2", "\"TRUE\";1;2", "\"x\";1;2", "C4;1;2",
                           "#NV;1;2", "WAHR;1;1/0", "A1>2;\"groß\";\"klein\"", "A1:A3>2;1;0", "{WAHR.FALSCH};{1.2};3", "C6;1;2", "\"\";1;2", "2,5;1;2"])
each("WENNS", "=WENNS(@@)", ["A1>5;1;A1>2;2", "FALSCH;1", "FALSCH;1;WAHR;2", "1;\"a\"", "\"x\";1", "A1>5;1;WAHR;\"sonst\"", "FALSCH;1;FALSCH;2"])
each("WENNFEHLER", "=WENNFEHLER(@@)", ["1/0;\"Fehler\"", "1;2", "#NV;0", "C8;5", "C4;5", "A1:A3/0;7"])
each("WENNNV", "=WENNNV(@@)", ["#NV;1", "1/0;1", "SVERWEIS(99;E1:F10;2;FALSCH);\"fehlt\""])
each("UND", "=UND(@@)", ["WAHR;WAHR", "WAHR;FALSCH", "1;2", "0;1", "\"WAHR\";WAHR", "\"x\";WAHR", "B1:B3", "C1:C10", "A1:A10", "A1:A10>0", "C4", "#NV;FALSCH", "FALSCH;#NV", "{1.0}", "C1:C3;C9"])
each("ODER", "=ODER(@@)", ["FALSCH;FALSCH", "WAHR;FALSCH", "0;0", "B1:B3", "C1:C10", "A1:A10>8", "\"x\";FALSCH", "#NV;WAHR"])
each("XODER", "=XODER(@@)", ["WAHR;WAHR", "WAHR;FALSCH", "WAHR;WAHR;WAHR", "1;0;1;1", "A1:A3>2", "B1:B3"])
each("NICHT", "=NICHT(@@)", ["WAHR", "0", "5", "\"x\"", "C4", "\"FALSCH\"", "A1:A3>2"])
each("ERSTERWERT", "=ERSTERWERT(@@)", ["2;1;\"eins\";2;\"zwei\"", "3;1;\"eins\";2;\"zwei\";\"sonst\"", "3;1;\"eins\";2;\"zwei\"", "\"a\";\"A\";1;2",
                                       "A1;3;\"drei\";\"?\"", "C4;0;\"null\";\"leer\""])
each("WAHL", "=WAHL(@@)", ["2;\"a\";\"b\";\"c\"", "1,9;\"a\";\"b\"", "4;\"a\";\"b\"", "0;1", "\"2\";1;2", "{1.2};\"a\";\"b\"", "2;A1;A2:A3"])
for fn in ["ISTZAHL", "ISTTEXT", "ISTLEER", "ISTFEHLER", "ISTFEHL", "ISTNV", "ISTLOG", "ISTKTEXT", "ISTBEZUG", "TYP", "N", "T"]:
    each(fn, f"={fn}(@@)", ["1", "\"1\"", "WAHR", "C4", "C6", "C8", "#NV", "A1", "B1", "\"\"", "A1:A3", "{1.\"a\"}", "1/0", "C3"])
each("ISTGERADE", "=ISTGERADE(@@)", ["2", "3", "2,9", "-3", "0", "\"4\"", "WAHR", "C4"])
each("ISTUNGERADE", "=ISTUNGERADE(@@)", ["2", "3", "-3,5"])
each("ISTFORMEL", "=ISTFORMEL(@@)", ["A1", "C6", "C8", "'Daten 2'!C1"])
each("FEHLER.TYP", "=FEHLER.TYP(@@)", ["#NULL!", "1/0", "#WERT!", "#BEZUG!", "#NAME?", "#ZAHL!", "#NV", "1", "\"x\""])
each("NV", "=NV()@@", [""])
each("ZEILE", "=ZEILE(@@)", ["", "A5", "A5:B9", "C:C", "3:5", "'Daten 2'!B3"])
each("SPALTE", "=SPALTE(@@)", ["", "C5", "C5:E9", "C:E", "3:5"])
each("ZEILEN", "=ZEILEN(@@)", ["A1:A10", "A:A", "1:3", "{1.2;3.4;5.6}", "5", "K1:M3"])
each("SPALTEN", "=SPALTEN(@@)", ["A1:C10", "A:C", "1:3", "{1.2;3.4;5.6}", "5"])
each("BEREICHE", "=BEREICHE(@@)", ["A1:B2", "(A1:B2;C3)"])
each("INDIREKT", "=INDIREKT(@@)", ["\"A2\"", "\"Tabelle1!A3\"", "\"'Daten 2'!A2\"", "\"A\"&2", "\"Z1S1\";FALSCH", "\"Z2S1\";FALSCH", "\"x\"", "\"Preise\"", "\"A1:A3\""])
case("INDIREKT", "=SUMME(INDIREKT(\"A1:A10\"))", "=SUMME(INDIREKT(\"E\"&1&\":E\"&3))", "=ZEILE(INDIREKT(\"B7\"))")
each("BEREICH.VERSCHIEBEN", "=BEREICH.VERSCHIEBEN(@@)", ["A1;2;0", "A1;0;4", "A1;-1;0", "A1;0;0;1;1", "A1:A3;1;0", "A1;0;0;3;1", "A1;1;1;0;1", "A1;0;0;-2;1"])
case("BEREICH.VERSCHIEBEN", "=SUMME(BEREICH.VERSCHIEBEN(A1;0;0;5;1))", "=SUMME(BEREICH.VERSCHIEBEN(E1;2;0;3))", "=ZEILEN(BEREICH.VERSCHIEBEN(A1;0;0;4;2))",
     "=SUMME(BEREICH.VERSCHIEBEN(A1:B2;1;1))", "=SUMME(BEREICH.VERSCHIEBEN(A5;-2;0;5;1))")
each("ADRESSE", "=ADRESSE(@@)", ["1;1", "5;3;4", "5;3;2", "5;3;3", "5;3;1;FALSCH", "5;3;4;FALSCH", "1;1;1;WAHR;\"Daten 2\"", "1;1;1;WAHR;\"Tab\"", "0;1", "1;16385"])

# Lookup.
each("SVERWEIS", "=SVERWEIS(@@)", ["30;E1:F10;2;FALSCH", "35;E1:F10;2", "35;E1:F10;2;WAHR", "5;E1:F10;2", "200;E1:F10;2", "30;E1:F10;3;FALSCH", "30;E1:F10;0;FALSCH",
                                   "\"dora\";F1:F10;1;FALSCH", "\"D*\";F1:F10;1;FALSCH", "\"?r?\";F1:G10;1;FALSCH", "\"Xaver\";F1:F10;1;FALSCH", "\"Bz\";F1:F10;1",
                                   "\"apfel\";B1:C10;2;FALSCH", "\"a~*b\";B1:C10;2;FALSCH", "\"~?\";B1:C10;2;FALSCH", "\"~~\";B1:B10;1;FALSCH", "30;E1:F10;2;0", "\"30\";E1:F10;2;FALSCH",
                                   "1;A1:B10;2;FALSCH", "{30;50};E1:F10;2;FALSCH", "30;E1:F10;{1.2};FALSCH", "C4;C1:C10;1;FALSCH", "WAHR;C1:C10;1;FALSCH", "2,5;C1:C10;1;FALSCH",
                                   "4;A1:A10;1", "8;A1:B10;2", "\"Kirsche\";B1:B10;1", "25;E1:E10;1;WAHR", "30;E1:F10;2,9;FALSCH", "30;E:F;2;FALSCH", "\"Max\";L5:N7;2;FALSCH"])
each("WVERWEIS", "=WVERWEIS(@@)", ["\"Mia\";L5:N7;2;FALSCH", "\"Ole\";L5:N7;3;FALSCH", "\"Ole\";L5:N7;4;FALSCH", "2;K1:M3;3", "5;K1:M3;2", "\"x\";K5:N7;2;FALSCH"])
each("VERGLEICH", "=VERGLEICH(@@)", ["30;E1:E10;0", "35;E1:E10", "35;E1:E10;1", "5;E1:E10;1", "35;E1:E10;-1", "\"dora\";F1:F10;0", "\"C*\";F1:F10;0", "4;A1:A10;0",
                                     "1;A1:A10;0", "100;E1:E10;1", "1000;E1:E10", "\"Emil\";F1:F10", "\"Zz\";F1:F10", "35;{10.20.30.40};1", "9;A1:A10;1", "2;A1:A10",
                                     "70;E1:E10;-1", "\"a?fel\";B1:B10;0", "WAHR;C1:C10;0", "\"10\";C1:C10;0", "10;C1:C10;0", "35;E1:E10;2", "\"\";C1:C10;0"])
each("XVERWEIS", "=XVERWEIS(@@)", ["30;E1:E10;F1:F10", "35;E1:E10;F1:F10", "35;E1:E10;F1:F10;\"fehlt\"", "35;E1:E10;F1:F10;;-1", "35;E1:E10;F1:F10;;1",
                                   "\"D*\";F1:F10;E1:E10;;2", "\"d?ra\";F1:F10;E1:E10;;2", "\"D*\";F1:F10;E1:E10", "5;A1:A10;B1:B10;;0;-1", "5;A1:A10;B1:B10",
                                   "30;E1:E10;F1:G10", "\"Mia\";L5:N5;L6:N7", "70;E1:E10;F1:F10;;0;2", "75;E1:E10;F1:F10;;-1;2", "75;E1:E10;F1:F10;;1;-2",
                                   "30;E1:E10;F1:F9", "\"apfel\";B1:B10;A1:A10", "{30;40};E1:E10;F1:F10", "200;E1:E10;F1:F10;;-1", "0;E1:E10;F1:F10;;1"])
each("XVERGLEICH", "=XVERGLEICH(@@)", ["30;E1:E10", "35;E1:E10;-1", "35;E1:E10;1", "\"D*\";F1:F10;2", "5;A1:A10;0;-1", "45;E1:E10;1;2"])
each("INDEX", "=INDEX(@@)", ["K1:M3;2;3", "K1:M3;2", "K1:M3;0;2", "K1:M3;2;0", "K1:M3;4;1", "A1:A10;5", "A1:A10;5;1", "A1:A10;1;2", "K1:K3;2;1", "{1.2.3;4.5.6};2;3",
                               "{1.2.3};2", "K1:M1;3", "(K1:M3;A1:A10);3;1;2", "(K1:M3;A1:A10);3;1;3", "K1:M3;1,9;1,9", "K1:M3;-1;1", "A1:A10;\"2\"", "K1:M3;{1;2};1"])
case("INDEX", "=SUMME(INDEX(K1:M3;0;2))", "=SUMME(INDEX(K1:M3;2;0))", "=SUMME(A1:INDEX(A1:A10;4))", "=ZEILE(INDEX(A1:A10;7))")
each("VERWEIS", "=VERWEIS(@@)", ["35;E1:E10;F1:F10", "35;E1:F10", "5;E1:E10;F1:F10", "\"Dora\";F1:F10;E1:E10", "100;E1:E10", "2;K1:M3", "7;{1.5.10};{\"a\".\"b\".\"c\"}"])

# Criteria functions.
CRIT = ["\">3\"", "\"<3\"", "\">=3\"", "\"<=3\"", "\"<>3\"", "\"=3\"", "3", "\"3\"", "\"apfel\"", "\"Apfel*\"", "\"*e\"", "\"?????\"", "\"a~*b\"", "\"~?\"",
        "\"~~\"", "\"*\"", "\"?*\"", "\"\"", "\"=\"", "\"<>\"", "WAHR", "\"WAHR\"", "\">b\"", "\"<Birne\"", "\">=apfel\"", "\"<>apfel\"", "\"10\"", "10",
        "\">1,5\"", "\">1.5\"", "\"#DIV/0!\"", "\">\"&A3", "A3", "\"<\"&2", "\"*a*\"", "\">=1.1.2024\"", "\"=1\"", "\"=a\"", "\"<=\"", "\">-5\"", "2,5", "\"x\"",
        "\"=*\"", "\"<>*\"", "\"a*b\"", "C4"]
for rng in ["A1:A10", "B1:B10", "C1:C10", "D1:D10", "J1:J10"]:
    each("ZÄHLENWENN", f"=ZÄHLENWENN({rng};@@)", CRIT)
each("ZÄHLENWENN", "=ZÄHLENWENN(@@)", ["A:A;\">3\"", "A1:A10;{3.5}", "A1:B10;\">3\"", "N1:N20;\"\"", "N1:N20;\"<>\"", "N1:N20;\"=\""])
each("SUMMEWENN", "=SUMMEWENN(@@)", ["A1:A10;\">3\"", "H1:H10;\"x\";A1:A10", "H1:H10;\"x\";E1:E10", "B1:B10;\"apfel*\";A1:A10", "A1:A10;3;E1:E10",
                                     "H1:H10;\"x\";A1:A3", "H1:H10;\"x\";A5", "A1:A10;\"<>3\"", "C1:C10;\">0\"", "H1:H10;\"<>x\";G1:G10", "A1:A10;\">\"&5;G1:G10"])
each("SUMMEWENNS", "=SUMMEWENNS(@@)", ["A1:A10;H1:H10;\"x\"", "A1:A10;H1:H10;\"x\";A1:A10;\">2\"", "E1:E10;A1:A10;\">=3\";A1:A10;\"<=5\"", "A1:A10;H1:H9;\"x\"",
                                       "A1:A10;B1:B10;\"*e*\"", "G1:G10;H1:H10;\"<>z\""])
each("ZÄHLENWENNS", "=ZÄHLENWENNS(@@)", ["H1:H10;\"x\";A1:A10;\">2\"", "A1:A10;\">1\";A1:A10;\"<5\"", "H1:H10;\"x\"", "A1:A10;\">1\";A1:A9;\"<5\"", "B1:B10;\"<>apfel\""])
each("MITTELWERTWENN", "=MITTELWERTWENN(@@)", ["A1:A10;\">3\"", "H1:H10;\"x\";A1:A10", "H1:H10;\"q\";A1:A10", "B1:B10;\"apfel\";A1:A10", "C1:C10;\"<>0\""])
each("MITTELWERTWENNS", "=MITTELWERTWENNS(@@)", ["A1:A10;H1:H10;\"x\"", "A1:A10;H1:H10;\"x\";A1:A10;\">2\"", "A1:A10;H1:H10;\"q\""])
each("MAXWENNS", "=MAXWENNS(@@)", ["A1:A10;H1:H10;\"x\"", "G1:G10;H1:H10;\"y\";A1:A10;\">2\"", "A1:A10;H1:H10;\"q\"", "A1:A10;H1:H9;\"x\""])
each("MINWENNS", "=MINWENNS(@@)", ["A1:A10;H1:H10;\"x\"", "G1:G10;H1:H10;\"z\"", "A1:A10;H1:H10;\"q\""])

# Text.
each("LINKS", "=LINKS(@@)", ["\"Hallo\";2", "\"Hallo\"", "\"Hallo\";0", "\"Hallo\";10", "\"Hallo\";-1", "123,45;3", "WAHR;2", "C4;2", "\"Hallo\";2,9", "B1:B3;2"])
each("RECHTS", "=RECHTS(@@)", ["\"Hallo\";2", "\"Hallo\"", "\"Hallo\";0", "\"Hallo\";10", "1/3;4", "\"Hallo\";-1"])
each("TEIL", "=TEIL(@@)", ["\"Hallo Welt\";7;4", "\"Hallo\";2;100", "\"Hallo\";10;2", "\"Hallo\";0;2", "\"Hallo\";2;-1", "\"Hallo\";2;0", "12345;2;3"])
each("LÄNGE", "=LÄNGE(@@)", ["\"Hallo\"", "\"\"", "C4", "1/3", "WAHR", "123,456", "1E+20", "-0,5", "\"Grüße\"", "A1:A3", "2^0,5", "1E-10", "0,1+0,2"])
each("GROSS", "=GROSS(@@)", ["\"hallo straße\"", "\"äöü\"", "1,5"])
each("KLEIN", "=KLEIN(@@)", ["\"HALLO ÄÖÜ\"", "WAHR"])
each("GROSS2", "=GROSS2(@@)", ["\"hallo welt\"", "\"o'neil mc-donald 2te\"", "\"ÄRGER über\"", "\"x1y\""])
each("GLÄTTEN", "=GLÄTTEN(@@)", ["\"  a   b  \"", "\"a\"&ZEICHEN(160)&\" b\"", "\" \"", "J2"])
each("FINDEN", "=FINDEN(@@)", ["\"l\";\"Hallo\"", "\"l\";\"Hallo\";4", "\"L\";\"Hallo\"", "\"\";\"Hallo\"", "\"x\";\"Hallo\"", "\"l\";\"Hallo\";0", "\"l\";\"Hallo\";6", "\"\";\"Hallo\";6", "\"\";\"Hallo\";7", "\"*\";\"a*b\""])
each("SUCHEN", "=SUCHEN(@@)", ["\"L\";\"Hallo\"", "\"l?o\";\"Hallo\"", "\"*o\";\"Hallo\"", "\"~*\";\"a*b\"", "\"x\";\"Hallo\"", "\"\";\"Hallo\"", "\"a\";\"Hallo\";3", "\"l*\";\"Hallo\""])
each("WECHSELN", "=WECHSELN(@@)", ["\"aaa\";\"a\";\"b\"", "\"aaa\";\"a\";\"b\";2", "\"aaa\";\"a\";\"b\";4", "\"aaa\";\"A\";\"b\"", "\"aaa\";\"\";\"b\"", "\"aaa\";\"a\";\"b\";0",
                                   "\"abab\";\"ab\";\"x\"", "1,5;\",\";\".\""])
each("ERSETZEN", "=ERSETZEN(@@)", ["\"Hallo\";2;3;\"xy\"", "\"Hallo\";10;1;\"!\"", "\"Hallo\";0;1;\"x\"", "\"Hallo\";1;0;\"x\"", "12345;2;2;\"-\""])
each("WIEDERHOLEN", "=WIEDERHOLEN(@@)", ["\"ab\";3", "\"ab\";0", "\"ab\";2,9", "\"ab\";-1", "\"ab\";20000", "1;3"])
each("IDENTISCH", "=IDENTISCH(@@)", ["\"a\";\"A\"", "\"a\";\"a\"", "1;\"1\"", "WAHR;\"WAHR\"", "C4;\"\""])
each("SÄUBERN", "=SÄUBERN(@@)", ["\"a\"&ZEICHEN(10)&\"b\"", "\"a\"&ZEICHEN(160)&\"b\"&ZEICHEN(7)"])
each("ZEICHEN", "=ZEICHEN(@@)", ["65", "97", "128", "196", "223", "256", "0", "1", "10", "65,9", "138", "150"])
each("CODE", "=CODE(@@)", ["\"A\"", "\"Ä\"", "\"€\"", "\"\"", "\"abc\"", "1", "\"ß\"", "\"–\""])
each("UNIZEICHEN", "=UNIZEICHEN(@@)", ["8364", "65", "0"])
each("UNICODE", "=UNICODE(@@)", ["\"€\"", "\"A\""])
each("VERKETTEN", "=VERKETTEN(@@)", ["\"a\";\"b\";1", "A1:A3", "1,5;WAHR", "C4;\"x\""])
each("TEXTKETTE", "=TEXTKETTE(@@)", ["A1:A3", "\"a\";B1:B3", "C1:C7", "1,5;WAHR;\"x\""])
each("TEXTVERKETTEN", "=TEXTVERKETTEN(@@)", ["\", \";WAHR;B1:B3", "\";\";WAHR;C1:C7", "\";\";FALSCH;C1:C7", "\"-\";WAHR;1;\"\";2", "{\"-\".\"+\"};WAHR;1;2;3", "\"\";WAHR;A1:A3"])
each("WERT", "=WERT(@@)", ["\"1,5\"", "\"1.000\"", "\"1.000,25\"", "\" 12 \"", "\"12%\"", "\"1E3\"", "\"abc\"", "\"\"", "\"12:30\"", "\"1.1.2024\"", "\"-3\"",
                           "\"(4)\"", "\"4 €\"", "\"1,2,3\"", "\"1.5\"", "\"0,5E-2\"", "1", "WAHR", "C4", "\"1 000\"", "\"$5\"", "\"5-\"", "\"1/2\"", "\"0 1/2\"", "\"10.000.000\""])
each("ZAHLENWERT", "=ZAHLENWERT(@@)", ["\"1.234,5\"", "\"1,234.5\";\".\";\",\"", "\"12%\"", "\"\"", "\"1 234\""])
each("FEST", "=FEST(@@)", ["1234,567", "1234,567;1", "1234,567;-2", "1234,567;2;WAHR", "-0,5;0", "1234567,891;3"])
each("TEXT", "=TEXT(@@)", ["1234,5;\"0,00\"", "1234,5;\"#.##0,00\"", "0,25;\"0%\"", "45292;\"TT.MM.JJJJ\"", "45292;\"TTTT\"", "45292;\"MMMM JJJJ\"",
                           "0,5;\"hh:mm\"", "3,14159;\"0,0\"", "-1;\"0;(0)\"", "1;\"\"", "\"abc\";\"0\"", "1234;\"00000\"", "45292,75;\"TT.MM.JJ hh:mm\""])
each("T", "=T(@@)", ["\"a\"", "1"])

# Dates and times.
each("DATUM", "=DATUM(@@)", ["2024;1;15", "2024;13;1", "2024;0;1", "2024;2;30", "2024;1;0", "1900;1;1", "1900;2;29", "1900;3;1", "0;1;1", "99;1;1", "1899;12;31",
                             "10000;1;1", "-1;1;1", "2024;-1;1", "2024,9;1,9;1,9", "1900;1;-1", "9999;12;31", "\"2024\";1;1", "2024;1;1+0,5"])
each("JAHR", "=JAHR(@@)", ["45292", "0", "1", "60", "61", "-1", "2958465", "2958466", "\"1.1.2024\"", "D1:D3", "45292,9"])
each("MONAT", "=MONAT(@@)", ["45292", "0", "60", "59", "61", "\"15.3.2024\""])
each("TAG", "=TAG(@@)", ["45292", "0", "60", "59", "61", "1", "45351"])
each("WOCHENTAG", "=WOCHENTAG(@@)", ["45292", "45292;2", "45292;3", "45292;11", "45292;12", "45292;13", "45292;14", "45292;15", "45292;16", "45292;17", "45292;4",
                                     "1", "0", "60", "59", "61", "-1", "45292;1,5"])
each("KALENDERWOCHE", "=KALENDERWOCHE(@@)", ["45292", "45292;2", "45291", "45291;2", "45291;21", "45658;21", "45657;21", "45293;11", "45296;16", "1", "45291;3"])
each("ISOKALENDERWOCHE", "=ISOKALENDERWOCHE(@@)", ["45292", "45291", "45658", "45657", "44197", "1", "60"])
each("EDATUM", "=EDATUM(@@)", ["45322;1", "45351;12", "45322;-1", "45322;1,9", "45292;0", "45322;-1500", "D1;3"])
each("MONATSENDE", "=MONATSENDE(@@)", ["45292;0", "45292;1", "45292;-1", "45337;0", "45292;0,9", "45292;-1500"])
each("DATEDIF", "=DATEDIF(@@)", ["45292;45658;\"Y\"", "45292;45658;\"M\"", "45292;45658;\"D\"", "45337;45818;\"YM\"", "45337;45818;\"MD\"", "45337;45818;\"YD\"",
                                 "45658;45292;\"D\"", "45292;45658;\"x\"", "45322;45351;\"M\"", "45351;45716;\"Y\"", "45322;45382;\"MD\"", "45292;45292;\"D\"",
                                 "36556;45337;\"YM\"", "36556;45337;\"md\""])
each("TAGE", "=TAGE(@@)", ["45658;45292", "45292;45658", "\"1.1.2025\";\"1.1.2024\""])
each("TAGE360", "=TAGE360(@@)", ["45292;45658", "45351;45382", "45351;45382;WAHR", "45322;45351"])
each("NETTOARBEITSTAGE", "=NETTOARBEITSTAGE(@@)", ["45292;45322", "45322;45292", "45292;45322;45293", "45292;45322;D1:D10", "45297;45298", "45292;45292"])
each("ARBEITSTAG", "=ARBEITSTAG(@@)", ["45292;10", "45292;-10", "45297;1", "45292;10;45293", "45292;0", "45292;5,9"])
each("NETTOARBEITSTAGE.INTL", "=NETTOARBEITSTAGE.INTL(@@)", ["45292;45322;11", "45292;45322;\"0000011\"", "45292;45322;7"])
each("ARBEITSTAG.INTL", "=ARBEITSTAG.INTL(@@)", ["45292;10;11", "45292;10;\"1111100\""])
each("DATWERT", "=DATWERT(@@)", ["\"1.1.2024\"", "\"01.01.24\"", "\"2024-01-15\"", "\"15. März 2024\"", "\"15.Mär.2024\"", "\"29.2.1900\"", "\"1.1.1900\"", "\"31.12.9999\"",
                                 "\"1.1.10000\"", "\"x\"", "\"1.1.2024 12:00\"", "\"15/3/2024\"", "\"März 2024\"", "\"15-Mär-2024\"", "\"1.13.2024\"", "45292", "\"1.1.29\"", "\"1.1.30\""])
each("ZEITWERT", "=ZEITWERT(@@)", ["\"12:00\"", "\"6:30:15\"", "\"18:45\"", "\"25:00\"", "\"12:00 PM\"", "\"1:00 AM\"", "\"1.1.2024 6:00\"", "\"x\"", "\"0:90\"", "\"12:30:15,5\""])
each("ZEIT", "=ZEIT(@@)", ["12;30;0", "25;0;0", "0;90;0", "0;0;-1", "-1;0;0", "12,9;30,9;0", "0;0;86400", "23;59;59", "32767;0;0", "32768;0;0"])
for fn in ["STUNDE", "MINUTE", "SEKUNDE"]:
    each(fn, f"={fn}(@@)", ["0,5", "0,75", "1,5", "I4", "I6", "I7", "45292,123456", "\"12:34:56\"", "-0,5", "0,999999999", "0,99999999999", "I1:I3", "2958466"])
each("HEUTE", "=@@", ["HEUTE()=GANZZAHL(JETZT())", "JETZT()>=HEUTE()", "JETZT()-HEUTE()<1", "JAHR(HEUTE())>=2026"])
each("BRTEILJAHRE", "=BRTEILJAHRE(@@)", ["45292;45658", "45292;45658;1", "45292;45658;2", "45292;45658;3", "45292;45658;4", "45351;45716;1"])

# Finance basics.
each("RMZ", "=RMZ(@@)", ["0,05/12;60;10000", "0;12;1200", "0,05;10;-1000;0;1", "0,05/12;60;10000;5000"])
each("ZW", "=ZW(@@)", ["0,05;10;-100", "0,05/12;12;-100;-1000;1", "0;10;-100"])
each("BW", "=BW(@@)", ["0,05;10;-100", "0;10;-100"])
each("ZZR", "=ZZR(@@)", ["0,05/12;-200;10000", "0;-100;1000"])
each("ZINS", "=ZINS(@@)", ["60;-200;10000", "10;-100;800"])

# Dynamic arrays.
each("SORTIEREN", "=SORTIEREN(@@)", ["A1:A10", "A1:A10;;-1", "B1:B10", "F1:F5;1;-1", "K10:L14;2", "K10:L14;2;-1", "K1:M3;1;-1;WAHR", "C1:C10", "{3.1.2};1;1;WAHR", "H1:H10"])
each("SORTIERENNACH", "=SORTIERENNACH(@@)", ["F1:F5;A1:A5", "F1:F5;A1:A5;-1", "K10:K14;L10:L14;1;K10:K14;-1"])
each("FILTER", "=FILTER(@@)", ["A1:A10;A1:A10>4", "B1:B10;H1:H10=\"x\"", "A1:A10;A1:A10>100", "A1:A10;A1:A10>100;\"keine\"", "K10:L14;L10:L14=5", "A1:A10;E1:E9>30",
                               "K1:M3;{1.0.1}", "A1:A10;(A1:A10>2)*(H1:H10=\"x\")"])
each("EINDEUTIG", "=EINDEUTIG(@@)", ["A1:A10", "H1:H10", "A1:A10;;WAHR", "K10:L14", "B1:B3", "{1.1.2};WAHR", "C1:C10"])
each("SEQUENZ", "=SEQUENZ(@@)", ["5", "2;3", "3;1;10;5", "2;2;0;0,5", "0", "1;1;-1"])
each("MTRANS", "=MTRANS(@@)", ["K1:M2", "A1:A3", "{1.2;3.4}"])
each("arrays", "=@@", ["A1:A3*E1:E3", "SUMME(A1:A10*E1:E10)", "MAX(WENN(H1:H10=\"x\";A1:A10))", "ZEILE(A1:A4)", "SPALTE(A1:D1)", "LÄNGE(B1:B4)",
                        "WENN(A1:A5>3;\"hoch\";\"tief\")", "RUNDEN(G1:G4;0)", "INDEX(K1:M3;{1;3};2)", "SVERWEIS(E1:E3;E1:F10;2;FALSCH)", "A1:A3&B1:B3",
                        "MITTELWERT(WENN(H1:H10=\"y\";A1:A10))", "SUMME((A1:A10>3)*1)", "ANZAHL(EINDEUTIG(A1:A10))", "{1.2.3}*{1;2}", "K1:M3*{1;10;100}",
                        "MTRANS(SEQUENZ(3))", "SUMME(SEQUENZ(100))", "MAX(LÄNGE(F1:F10))", "ISTZAHL(C1:C5)", "XVERWEIS(E1:E2;E1:E10;F1:F10)"])

# Volatile and calculation properties.
each("calc", "=@@", ["'Daten 2'!C1", "'Daten 2'!C2", "'Daten 2'!C3", "C6", "C8"])

# Numbers turned into text (concatenation): where Excel switches to E notation.
for k in range(-25, 26):
    case("numtext", f"=\"#\"&10^{k}", f"=\"#\"&-1,23456789012345*10^{k}", f"=\"#\"&1/3*10^{k}")
case("numtext", "=\"#\"&0,1+0,2", "=\"#\"&2^0,5", "=\"#\"&1/7", "=\"#\"&123456789012345678", "=\"#\"&0,000123456789012345", "=\"#\"&1E+308",
     "=\"#\"&-0,5", "=\"#\"&99999999999999,9", "=\"#\"&999999999999999", "=\"#\"&0,00001", "=\"#\"&0,000001", "=\"#\"&1,5E-9", "=\"#\"&1,234E-10")
# Number literals in formulas: where Excel writes E notation.
each("literals", "=@@", ["1E+21", "1E+22", "1E+100", "1E+300", "1,5E+25", "1E-20", "1E-21", "1E-100", "1,234E-15", "0,000000000000000123", "1E+16", "1,7976931348623E+308"])
for k in [25, 27, 30, 35, 38, 39, 40, 41, 42, 45, 48, 50, 52]:
    case("REST", f"=REST(2^{k}+0,5;1)", f"=REST(2^{k};3)", f"=REST(2^{k};7)")
case("A-arrays", "=MAXA({1.2.\"3\"})", "=MINA({1.2.\"3\";-1.WAHR.\"x\"})", "=MITTELWERTA({1.2.\"3\"})", "=MITTELWERTA({1.WAHR})",
     "=STABWA({1.2.\"3\"})", "=MAXA(C1:C10)", "=MINA(C1:C10)", "=MITTELWERTA(C1:C10)", "=ANZAHL2({1.\"\"})")
case("compare", "=\"a-b\"=\"ab\"", "=\"a'b\"=\"ab\"", "=\"a b\"=\"ab\"", "=\"ab\"<\"a-c\"", "=\"Straße\"=\"STRASSE\"", "=\"Æ\"=\"AE\"",
     "=\"a\"<\"ä\"", "=\"ä\"<\"b\"", "=\"z\"<\"ä\"", "=\"_\"<\"a\"", "=\"1\"<\"a\"", "=\"-\"<\"a\"", "=\" \"<\"a\"", "=\"\"<\" \"",
     "=\"a1\"<\"a10\"", "=\"a2\"<\"a10\"", "=\"é\"=\"e\"", "=\"ab\"<\"aB\"", "=\"Ab\"<\"ab\"", "=VERGLEICH(\"b\";{\"a\".\"B\".\"c\"};0)")
case("convert", "='Tabelle1'!A1", "=tabelle1!a1", "='Daten 2'!a1:b2", "=SUMME('Daten 2':Tabelle1!A1)", "=A1 A2", "=(A1)", "=((1))", "=-(-1)", "={1;2}",
     "=WENN(A1=1;\"x\")", "=Preise", "=preise", "=SUMME(preise)", "=1+ 2", "=1 +2", "=SUMME(A1;" + chr(10) + "2)","=a1:b2 b1", "=$a1:b$2", "=A1:A1", "=B2:A1",
     "=A$1:$A1", "=SUMME(Z:Z)", "=ZÄHLENWENN(A1:A10;\">\"&A1)", "=\"Ä\"&\"ö\"", "=1%%", "=-1%", "=+A1", "=--A1", "=WAHR()", "=FALSCH()",
     "={\"a\";\"b\"}", "={1.-2}", "={-1.+2}", "={#DIV/0!.#NV}", "=#BEZUG!", "=Tabelle1!#BEZUG!", "=SUMME(#BEZUG!)", "=#NULL!", "=#ZAHL!", "=#NAME?")

# Generated cases: many inputs through the same functions (dates, rounding, number text).
DATES = [1, 2, 59, 60, 61, 366, 367, 1462, 36526, 36556, 36585, 36891, 44196, 44197, 44198, 44561, 44562, 44926, 44927, 45291, 45292,
         45293, 45296, 45297, 45351, 45352, 45657, 45658, 45659, 45664, 46022, 46023, 46387, 2958465]
for d in DATES:
    case("gen-date", f"=WOCHENTAG({d})", f"=WOCHENTAG({d};2)", f"=WOCHENTAG({d};3)", f"=KALENDERWOCHE({d})", f"=KALENDERWOCHE({d};2)",
         f"=KALENDERWOCHE({d};21)", f"=KALENDERWOCHE({d};15)", f"=ISOKALENDERWOCHE({d})", f"=JAHR({d})&\"-\"&MONAT({d})&\"-\"&TAG({d})",
         f"=EDATUM({d};1)", f"=EDATUM({d};-13)", f"=MONATSENDE({d};0)", f"=MONATSENDE({d};-1)", f"=DATEDIF({d};{d}+400;\"md\")",
         f"=DATEDIF({d};{d}+400;\"ym\")", f"=DATEDIF({d};{d}+400;\"yd\")", f"=DATEDIF({d};46000;\"m\")", f"=TAGE360({d};{d}+100)",
         f"=TAGE360({d};{d}+100;WAHR)", f"=BRTEILJAHRE({d};{d}+400)", f"=BRTEILJAHRE({d};{d}+400;1)", f"=NETTOARBEITSTAGE({d};{d}+30)",
         f"=ARBEITSTAG({d};7)", f"=ARBEITSTAG({d};-3)")
VALS = ["0,5", "1,5", "2,5", "-2,5", "0,125", "0,375", "1,005", "1,015", "1,025", "2,345", "2,355", "0,045", "1234,5", "-0,05", "0,0049999", "99,995",
        "1/3", "2/3", "-1/3", "10/7", "1E+15/7", "4,35*100", "1,1*1,1", "0,1*3", "1-0,9", "0,7+0,1", "123,456789"]
for v in VALS:
    for k in ["0", "1", "2", "-1"]:
        case("gen-round", f"=RUNDEN({v};{k})", f"=AUFRUNDEN({v};{k})", f"=ABRUNDEN({v};{k})")
    case("gen-round", f"=GANZZAHL({v})", f"=KÜRZEN({v})", f"=OBERGRENZE.MATHEMATIK({v};0,05)", f"=UNTERGRENZE.MATHEMATIK({v};0,05)", f"=VRUNDEN({v};0,05)",
         f"=REST({v};0,1)", f"=\"#\"&{v}", f"=LÄNGE({v})", f"=WERT(\"\"&{v})", f"=({v})*3=({v})+({v})+({v})")
for t in ["A1:A10", "G1:G10", "E1:E10", "{1.2.2.3.3.3.4}", "{-5.0.5.10.15}"]:
    for f in ["MEDIAN", "MODUS.EINF", "STABW.S", "VAR.P", "MITTELABW", "KURT", "SCHIEFE"]:
        case("gen-stat", f"={f}({t})")
    for k in ["0,1", "0,2", "0,3", "0,75", "0,95"]:
        case("gen-stat", f"=QUANTIL.INKL({t};{k})", f"=QUANTIL.EXKL({t};{k})", f"=QUANTILSRANG.INKL({t};2)")
for x in ["-3", "-1,5", "-0,5", "0,1", "0,7", "1,2", "2", "3,5"]:
    case("gen-dist", f"=NORM.S.VERT({x};WAHR)", f"=NORM.S.VERT({x};FALSCH)", f"=NORM.VERT({x};1;2;WAHR)", f"=T.VERT({x};4;WAHR)", f"=T.VERT({x};30;FALSCH)",
         f"=EXPON.VERT(ABS({x});0,5;WAHR)", f"=POISSON.VERT(ABS({x});3;WAHR)", f"=BINOM.VERT(ABS({x});12;0,3;WAHR)")
for p in ["0,001", "0,01", "0,1", "0,3", "0,5", "0,7", "0,9", "0,99", "0,999"]:
    case("gen-dist", f"=NORM.S.INV({p})", f"=NORM.INV({p};100;15)", f"=T.INV({p};5)", f"=T.INV.2S({p};10)", f"=BINOM.INV(20;0,4;{p})")
case("more", "=TEXTVOR(\"a-b-c\";\"-\")", "=TEXTVOR(\"a-b-c\";\"-\";2)", "=TEXTVOR(\"a-b-c\";\"-\";-1)", "=TEXTNACH(\"a-b-c\";\"-\")",
     "=TEXTNACH(\"a-b-c\";\"x\")", "=TEXTNACH(\"a-b-c\";\"x\";;;;\"?\")", "=TEXTNACH(\"aXbxc\";\"x\";1;1)", "=DEZINBIN(10)", "=DEZINBIN(10;8)",
     "=DEZINBIN(-1)", "=DEZINHEX(255)", "=DEZINHEX(-1)", "=BININDEZ(\"1010\")", "=BININDEZ(\"1111111111\")", "=HEXINDEZ(\"FF\")", "=HEXINDEZ(\"fffffffff0\")",
     "=ZINSZ(0,05/12;1;60;10000)", "=KAPZ(0,05/12;1;60;10000)", "=ZINSZ(0,05;2;10;1000;0;1)", "=NBW(0,1;-100;50;60)", "=SUMME(MMULT(K1:M3;K1:M3))",
     "=MMULT({1.2};{3;4})", "=INDEX(MMULT(K1:M3;K1:M3);2;3)", "=WAHL({1.2.3};\"a\";\"b\";\"c\")", "=XVERGLEICH(40;E1:E10;0;2)", "=XVERGLEICH(45;E1:E10;-1;2)",
     "=XVERGLEICH(45;E1:E10;1;2)", "=VERWEIS(2;{1.2.3};{\"a\".\"b\".\"c\"})", "=VERWEIS(\"x\";B1:B10)", "=SUMMEWENN(D1:D10;\">=1.1.2024\";A1:A10)",
     "=ZÄHLENWENN(D1:D10;\"<\"&DATUM(2024;3;1))", "=ZÄHLENWENN(A1:A10;\">=\"&3)", "=ZÄHLENWENN(B1:B10;\"<>*e*\")", "=ZÄHLENWENN(B1:B10;\"A*\")",
     "=ZÄHLENWENNS(A1:A10;\">2\";H1:H10;\"<>x\";B1:B10;\"<>\")", "=MITTELWERTWENNS(G1:G10;A1:A10;\">3\";H1:H10;\"x\")",
     "=ADRESSE(3;4;2;FALSCH)", "=ADRESSE(3;4;1;FALSCH;\"Daten 2\")", "=ZEILE(INDIREKT(\"Z(2)S\";FALSCH))", "=INDIREKT(\"Z(-1)S(-2)\";FALSCH)",
     "=FEHLER.TYP(#ÜBERLAUF!)", "=ISTFEHLER(EINDEUTIG(A1:A10))", "=ANZAHL(A1:A10;\"5\";WAHR;\"x\")", "=ANZAHL2(A1;\"\";C4)", "=SUMME(\"1.1.2024\";1)",
     "=MAX(\"12:00\";0,1)", "=PRODUKT(C1:C3)", "=QUADRATESUMME(\"3\";WAHR)", "=GGT(0;0)", "=KGV(0;0)", "=MITTELWERT(\"\")", "=SUMME(\"\")")

SHIFT = [  # (formula as file form, from, to)
    ("A1+B2", "P5", "P6"), ("A1+B2", "P5", "Q5"), ("$A$1+A$1+$A1", "P5", "R8"), ("A1", "P5", "P1"), ("A1", "P5", "O5"), ("A2", "P5", "P3"),
    ("SUM(A1:A3)", "P5", "P10"), ("SUM(A1:A3)", "P5", "P4"), ("SUM(A$1:A3)", "P5", "P3"), ("SUM(A:A)", "P5", "Q9"), ("SUM(1:1)", "P5", "Q9"),
    ("SUM($A:A)", "P5", "R5"), ("SUM(A:A)", "P5", "O5"), ("'Daten 2'!A1*2", "P5", "P7"), ("A1048576", "P5", "P6"), ("XFD1", "P5", "Q5"),
    ("A1:B2 B1:C3", "P5", "Q6"), ("INDIRECT(\"A1\")+A1", "P5", "P6"), ("\"A1\"&A1", "P5", "P6"), ("Preise", "P5", "P6"), ("SUM(A1:A3,C1)", "P5", "O4"),
    ("A1#", "P5", "P6"), ("#REF!+A1", "P5", "P6"), ("A1048576", "P5", "P4"), ("A1:A1048576", "P5", "P6"), ("B5:A1", "P5", "P6"),
    ("SUM(A1:A3)", "P5", "P3"), ("A1:B2", "P5", "O5"), ("$A1", "P5", "P1"), ("A$1", "P5", "O1"), ("'Daten 2'!A1", "P5", "P4"),
    ("SUM(A:A)", "P5", "P1"), ("SUM(2:3)", "P5", "P3"), ("A1#", "P5", "P4"), ("A2+A1", "P5", "P4"), ("SUM(A2:B3 A1)", "P5", "P4"),
]

ADJUST_FORMULAS = [
    "A5", "$A$5", "A1:A10", "A5:A10", "A6:A10", "A1:A4", "A1:A5", "A4:A5", "A10:A20", "A:A", "1:10", "5:5", "B5:C6", "SUM(A5:A10)", "A3:A4",
    "SUM(A5,A6,A7)", "A4", "A11", "C:D", "B:B", "A:C", "B2:D3", "$B$2:$D$3", "SUM(B:B)", "XFD1", "A1048576", "A5#",
]
ADJUST_OPS = [("rows", 5, 1), ("rows", 5, 3), ("rows", 5, -1), ("rows", 5, -2), ("rows", 4, -2), ("rows", 1, 2), ("rows", 11, 1), ("rows", 10, -1),
              ("rows", 6, -5), ("rows", 2, -9), ("cols", 2, 1), ("cols", 2, -1), ("cols", 3, -1), ("cols", 2, -2), ("cols", 4, 2), ("cols", 1, 1)]


def encode(v):
    """Value2 as JSON: numbers, text, booleans, null, errors as {"e": code}."""
    if isinstance(v, bool) or v is None or isinstance(v, str):
        return v
    if isinstance(v, int):
        return {"e": ERRORS.get(v + 2146828288, f"#{v}")}
    if isinstance(v, float):
        return v
    if isinstance(v, tuple):
        return [[encode(c) for c in row] for row in v]
    return str(v)


def fill(x, wb) -> None:
    for sheet, addr, v in SETUP:
        ws = wb.Worksheets(sheet)
        r = ws.Range(addr)
        if isinstance(v, str) and (v.startswith("=") or v.startswith("'")):
            r.Formula2Local = v
        else:
            r.Value2 = v
    for name, ref in NAMES:
        wb.Names.Add(name, RefersToLocal=ref)


def record_setup(wb) -> list[dict]:
    out = []
    for sheet, addr, v in SETUP:
        r = wb.Worksheets(sheet).Range(addr)
        e = {"sheet": sheet, "addr": addr, "value": encode(r.Value2)}
        if r.HasFormula:
            e["formula"] = r.Formula2[1:]
        out.append(e)
    return out


def run_cases(x, ws) -> None:
    n = len(CASES)
    t0 = time.time()
    for i, c in enumerate(CASES):
        cell = ws.Range(c["addr"])
        try:
            cell.Formula2Local = c["local"]
        except Exception as e:  # noqa: BLE001  Excel refuses the formula
            c["refused"] = True
            continue
        c["file"] = cell.Formula2[1:]
        c["back"] = cell.Formula2Local
        c["value"] = encode(cell.Value2)
        c["text"] = cell.Text
        if cell.HasSpill:
            c["spill"] = encode(cell.SpillingToRange.Value2)
        cell.ClearContents()
        if i % 250 == 0:
            print(f"  {i}/{n} cases, {time.time() - t0:.0f} s", flush=True)


def run_shift(ws) -> list[dict]:
    out = []
    for f, src, dst in SHIFT:
        print(f"  {f} {src} {dst}", flush=True)
        a = ws.Range(src)
        a.Formula2 = "=" + f
        b = ws.Range(dst)
        # Range.Copy with a destination copies like Strg+C/Strg+V without the clipboard.
        a.Copy(b)
        out.append({"formula": f, "from": src, "to": dst, "result": b.Formula2[1:]})
        a.ClearContents()
        b.ClearContents()
    return out


REF_RE = re.compile(r"(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?#?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?\d+:\$?\d+)")


def qualify(f: str, sheet: str) -> str:
    """The same formula on another sheet: every reference with the sheet's name."""
    return REF_RE.sub(lambda m: f"{sheet}!{m.group(1)}", f)


def moved(row: int, col: int, kind: str, at: int, count: int) -> tuple[int, int]:
    if kind == "rows" and row >= at:
        return row + count, col
    if kind == "cols" and col >= at:
        return row, col + count
    return row, col


def run_adjust(wb) -> list[dict]:
    """Every formula on the changed sheet (in cells the operation does not delete) and on
    another sheet, then one insert or delete; Excel's formulas afterwards. Entered through the
    legacy Formula property: as dynamic arrays, A:A and 1:10 would spill (and recalculate)
    arrays of a million cells on every change."""
    ws = wb.Worksheets.Add()
    ws.Name = "Adj"
    other = wb.Worksheets.Add()
    other.Name = "Andere"
    out = []
    for kind, at, count in ADJUST_OPS:
        print(f"  {kind} {at} {count}", flush=True)
        ws.Cells.Clear()
        other.Cells.Clear()
        homes = []
        for i, f in enumerate(ADJUST_FORMULAS):
            # Same sheet: row 1 from column 200 for row operations, column 200 from row 2 for column ones.
            r, c = (1, 200 + i) if kind == "rows" else (2 + i, 200)
            if kind == "rows" and at <= 1 and count < 0:
                r = 1000
            ws.Cells(r, c).Formula = "=" + f
            other.Cells(1 + i, 1).Formula = "=" + qualify(f, "Adj")
            homes.append((r, c))
        rng = ws.Rows(f"{at}:{at + abs(count) - 1}") if kind == "rows" else ws.Range(ws.Cells(1, at), ws.Cells(1, at + abs(count) - 1)).EntireColumn
        if count > 0:
            rng.Insert()
        else:
            rng.Delete()
        for i, f in enumerate(ADJUST_FORMULAS):
            r, c = moved(*homes[i], kind, at, count)
            out.append({"kind": kind, "at": at, "count": count, "formula": f, "same": ws.Cells(r, c).Formula[1:],
                        "otherBefore": qualify(f, "Adj"), "other": other.Cells(1 + i, 1).Formula[1:]})
    return out


MOVE_FORMULAS = ["B5", "C6", "B5:C6", "B5:B6", "B6", "A5:C6", "B4:C6", "B5:D7", "A1:A10", "B5:C7", "B5:B100", "A5:C5", "E10", "E10:F11",
                 "SUM(B5:C6)", "$B$5", "B:B", "5:5", "D5", "B5:C6 B6"]
MOVES = [("B5:C6", 5, 3), ("B5:C6", 0, 1), ("B5:C6", 1, 0), ("B5:B6", 0, 10), ("B5:C6", -2, 0)]


def run_move(wb) -> list[dict]:
    """Cut and paste (Range.Cut with a destination, no clipboard): formulas pointing at the
    moved block from the same and from another sheet, and a formula inside the block."""
    ws = wb.Worksheets("Adj")
    other = wb.Worksheets("Andere")
    out = []
    for src, dr, dc in MOVES:
        print(f"  move {src} {dr} {dc}", flush=True)
        ws.Cells.Clear()
        other.Cells.Clear()
        for i, f in enumerate(MOVE_FORMULAS):
            ws.Cells(300 + i, 300).Formula = "=" + f
            other.Cells(1 + i, 1).Formula = "=" + qualify(f, "Adj")
        block = ws.Range(src)
        r0, c0 = block.Row, block.Column
        # A formula inside the block: its own references do not shift when it moves.
        ws.Cells(r0, c0).Formula = "=A1+C6+B5"
        dest = ws.Cells(r0 + dr, c0 + dc)
        block.Cut(dest)
        inner = ws.Cells(r0 + dr, c0 + dc).Formula[1:]
        for i, f in enumerate(MOVE_FORMULAS):
            out.append({"from": src, "dRow": dr, "dCol": dc, "formula": f, "same": ws.Cells(300 + i, 300).Formula[1:],
                        "otherBefore": qualify(f, "Adj"), "other": other.Cells(1 + i, 1).Formula[1:]})
        out.append({"from": src, "dRow": dr, "dCol": dc, "formula": "A1+C6+B5", "inside": True, "same": inner})
    return out


CELL_FORMULAS = ["B5", "C6", "B5:C6", "B4", "B7", "B8:C9", "B4:C9", "A5:C6", "B5:D6", "B1:B10", "B6:B20", "A1:D10", "D5", "B:B", "5:5",
                 "SUM(B7:C8)", "$B$7", "B5:B6", "C7", "B3:C5", "B6:C7", "A6", "D6"]
CELL_OPS = [("B5:C6", "down"), ("B5:C6", "right"), ("B5:C6", "up"), ("B5:C6", "left"), ("B5:B5", "down"), ("A5:C6", "up")]


def run_cells(wb) -> list[dict]:
    """Zellen einfügen/löschen with shifting (Range.Insert/Delete with a direction): the formulas
    on another sheet that point at the shifted area, entered as legacy formulas (see run_adjust)."""
    ws = wb.Worksheets("Adj")
    other = wb.Worksheets("Andere")
    out = []
    shift = {"down": -4121, "right": -4161, "up": -4162, "left": -4159}
    for rng, d in CELL_OPS:
        print(f"  cells {rng} {d}", flush=True)
        ws.Cells.Clear()
        other.Cells.Clear()
        for i, f in enumerate(CELL_FORMULAS):
            other.Cells(1 + i, 1).Formula = "=" + qualify(f, "Adj")
        r = ws.Range(rng)
        if d in ("down", "right"):
            r.Insert(Shift=shift[d])
        else:
            r.Delete(Shift=shift[d])
        for i, f in enumerate(CELL_FORMULAS):
            out.append({"range": rng, "dir": d, "formula": qualify(f, "Adj"), "other": other.Cells(1 + i, 1).Formula[1:]})
    return out


def run_xlfn(x) -> dict:
    """Saves a workbook with one formula per function to TEMP and reads the prefixes Excel writes."""
    wb = x.app.Workbooks.Add()
    ws = wb.Worksheets(1)
    seen = {}
    row = 1
    for c in CASES_FOR_XLFN:
        if c.get("refused") or "file" not in c:
            continue
        for name in re.findall(r"([A-Z][A-Z0-9.]*)\(", c["file"]):
            if name not in seen:
                seen[name] = c["file"]
    for name, f in seen.items():
        try:
            ws.Cells(row, 1).Formula2 = "=" + f
            row += 1
        except Exception:  # noqa: BLE001
            pass
    path = os.path.join(tempfile.gettempdir(), "raabe-probe-formulas.xlsx")
    if os.path.exists(path):
        os.remove(path)
    wb.SaveAs(path, 51)
    wb.Close(False)
    with zipfile.ZipFile(path) as z:
        xml = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
    os.remove(path)
    prefixed = sorted(set(re.findall(r"(_xlfn\.(?:_xlws\.)?[A-Z][A-Z0-9.]*)\(", xml)))
    return {p.split(".", 1)[1].replace("_xlws.", ""): p for p in prefixed}


SECTIONS = ["cases", "shift", "adjust", "move", "cells", "xlfn"]


def main() -> None:
    """Runs the sections named on the command line (all by default) and merges them into the
    fixture, which is written after every section."""
    wanted = [a for a in sys.argv[1:] if a in SECTIONS] or SECTIONS
    data = json.loads(OUT.read_text(encoding="utf-8")) if OUT.exists() else {}

    def save() -> None:
        OUT.parent.mkdir(parents=True, exist_ok=True)
        OUT.write_text(json.dumps(data, ensure_ascii=False, indent=0), encoding="utf-8", newline=chr(10))

    x = ExcelTarget(visible=False)
    try:
        wb = x.wb
        ws = wb.Worksheets(1)
        ws.Name = T1
        ws2 = wb.Worksheets.Add(After=ws)
        ws2.Name = T2
        ws.Activate()
        fill(x, wb)
        data.update({
            "excel": f"{x.app.Version} build {int(x.app.Build)}",
            "recorded": time.strftime("%Y-%m-%d"),
            "sheets": [T1, T2],
            "names": [{"name": n, "ref": wb.Names(n).RefersTo[1:]} for n, _ in NAMES],
            "setup": record_setup(wb),
        })
        if "cases" in wanted:
            print(f"{len(CASES)} cases", flush=True)
            run_cases(x, ws)
            data["cases"] = CASES
            save()
        if "shift" in wanted:
            print("shift", flush=True)
            data["shift"] = run_shift(ws)
            save()
        if "adjust" in wanted:
            print("adjust", flush=True)
            data["adjust"] = run_adjust(wb)
            save()
        if "move" in wanted:
            print("move", flush=True)
            data["move"] = run_move(wb)
            save()
        if "cells" in wanted:
            print("cells", flush=True)
            for name in ("Adj", "Andere"):
                if not any(s.Name == name for s in wb.Worksheets):
                    wb.Worksheets.Add().Name = name
            data["cells"] = run_cells(wb)
            save()
        if "xlfn" in wanted and "cases" in data:
            print("xlfn", flush=True)
            global CASES_FOR_XLFN
            CASES_FOR_XLFN = data["cases"]
            data["xlfn"] = run_xlfn(x)
            save()
    finally:
        try:
            x.close()
        except TypeError:  # the process is gone already (tasklist's message is not cp1252)
            pass
    print(f"done -> {OUT.relative_to(ROOT)}")


CASES_FOR_XLFN: list[dict] = []


if __name__ == "__main__":
    sys.exit(main())
