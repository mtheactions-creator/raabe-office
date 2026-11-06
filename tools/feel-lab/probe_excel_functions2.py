"""Records Excel's results for the functions added in the second round of RaabeCalc's formula
engine (statistics and distributions, maths, finance, engineering, text, arrays), as the
reference for packages/sheet-engine/test/formula-functions2.test.ts.

Starts an invisible Excel of its own (`ExcelTarget(visible=False)`), fills one sheet with a
fixed data block, then enters every case in the file form (`Formula2`, English names) and
reads back `Formula2Local` (the German form, which also gives Excel's German function names),
`Value2` and, for dynamic arrays, the whole spill range. A last section saves a workbook to
TEMP to read which of the functions carry the `_xlfn.` prefix in the file.

    .venv/Scripts/python.exe probe_excel_functions2.py

Writes packages/sheet-engine/test/fixtures/excel-functions2.json. No real input, no visible
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
OUT = ROOT / "packages" / "sheet-engine" / "test" / "fixtures" / "excel-functions2.json"

ERRORS = {2000: "#NULL!", 2007: "#DIV/0!", 2015: "#VALUE!", 2023: "#REF!", 2029: "#NAME?", 2036: "#NUM!",
          2042: "#N/A", 2043: "#GETTING_DATA", 2045: "#SPILL!", 2047: "#BLOCKED!", 2050: "#CALC!"}

SHEET = "Tabelle1"
AT = "AB2"

# ----- the data block (address, value; "=..." in the file form) ----------------------------
SETUP: list[tuple[str, object]] = []


def col(letter: str, values: list, start: int = 1) -> None:
    for i, v in enumerate(values):
        if v is not None:
            SETUP.append((f"{letter}{start + i}", v))


col("A", [3, 1, 4, 1, 5, 9, 2, 6, 5, 3])
col("B", [2.5, -1, 0, 7, 3.25, 100, -20, 0.5, 8, 0.001])
col("C", [1, "a", True, None, 2.5, '=""', "'10", 4, -3, False])
col("D", [1, 2, 3, 4, 5, 6, 7, 8])
col("E", [2.1, 3.9, 8.2, 15.8, 32.5, 63.0, 129.0, 255.0])
col("F", [-10000, 2750, 4250, 3250, 2750])
col("G", [39448, 39508, 39751, 39859, 39904])
col("F", [-120000, 39000, 30000, 21000, 37000, 46000], 7)
col("H", [1, 0, 5]); col("I", [2, 1, 6]); col("J", [3, 4, 0])
col("H", [1, 2], 5); col("I", [2, 4], 5)
col("H", [4, 2], 8); col("I", [7, 6], 8)
col("K", [3.1, 4.9, 7.2, 8.8, 11.1, 13.0])
col("L", [1, 2, 3, 4, 5, 6])
col("M", [2, 1, 4, 3, 6, 5])
col("N", ["Nord", "Süd", "Nord", "Ost", "Süd", "Nord", "Ost", "Süd"])
col("O", ["A", "A", "B", "B", "A", "B", "A", "B"])
col("P", [10, 20, 30, 40, 50, 60, 70, 80])
col("Q", [58, 11, 10]); col("R", [35, 25, 23])
col("S", [45.35, 17.56, 16.09]); col("T", [47.65, 18.44, 16.91])
col("U", ["a,b;c,d", "Hallo Welt", "abc123def45", "2024-01-15"])
col("V", [6, 7, 9, 15, 21])
col("W", [20, 28, 31, 38, 40])
col("X", [1, 2, 2, 3, 3, 4])
col("Y", [5, "=1/0", 3, "=NA()", 8, 1])

# ----- the cases (file form) ---------------------------------------------------------------
CASES: list[dict] = []


def case(group: str, *formulas: str) -> None:
    for f in formulas:
        CASES.append({"group": group, "file": f})


def each(group: str, names: list[str], *arglists: str) -> None:
    for nm in names:
        for a in arglists:
            case(group, f"{nm}({a})")


# Statistics and distributions.
case("stats",
     "BETA.DIST(2,8,10,TRUE,1,3)", "BETA.DIST(2,8,10,FALSE,1,3)", "BETA.DIST(0.4,2,3,TRUE)", "BETA.DIST(0.4,2,3,FALSE)",
     "BETA.DIST(0.5,0.5,0.5,FALSE)", "BETA.DIST(0,2,3,TRUE)", "BETA.DIST(1.5,2,3,TRUE)", "BETA.DIST(0.4,0,3,TRUE)",
     "BETA.DIST(\"a\",2,3,TRUE)", "BETA.DIST(0.999,30,2,TRUE)", "BETA.DIST(0,2,3,FALSE)", "BETA.DIST(1,2,3,FALSE)",
     "BETA.INV(0.685470581,8,10,1,3)", "BETA.INV(0.5,2,3)", "BETA.INV(0.01,0.5,0.5)", "BETA.INV(0.99,5,1)",
     "BETA.INV(0,2,3)", "BETA.INV(1,2,3)", "BETA.INV(1.5,2,3)", "BETA.INV(0.2,2,3,1,1)",
     "BETADIST(2,8,10,1,3)", "BETADIST(0.4,2,3)", "BETADIST(1,2,3)", "BETAINV(0.685470581,8,10,1,3)", "BETAINV(0.3,2,3)",
     "BINOM.DIST.RANGE(60,0.75,48)", "BINOM.DIST.RANGE(60,0.75,45,50)", "BINOM.DIST.RANGE(10,0.5,11)",
     "BINOM.DIST.RANGE(10,0.5,6,4)", "BINOM.DIST.RANGE(10,1.5,3)", "BINOM.DIST.RANGE(10.7,0.3,2.9)",
     "BINOMDIST(6,10,0.5,FALSE)", "BINOMDIST(6,10,0.5,TRUE)", "BINOMDIST(11,10,0.5,TRUE)",
     "CHISQ.DIST(0.5,1,TRUE)", "CHISQ.DIST(2,3,FALSE)", "CHISQ.DIST(10,4,TRUE)", "CHISQ.DIST(0,2,FALSE)",
     "CHISQ.DIST(0,1,FALSE)", "CHISQ.DIST(-1,2,TRUE)", "CHISQ.DIST(1,0,TRUE)", "CHISQ.DIST(3,2.7,TRUE)",
     "CHISQ.DIST.RT(18.307,10)", "CHISQ.DIST.RT(0.5,1)", "CHISQ.DIST.RT(100,3)", "CHISQ.DIST.RT(-1,2)", "CHISQ.DIST.RT(0,2)",
     "CHISQ.INV(0.93,1)", "CHISQ.INV(0.6,2)", "CHISQ.INV(0.001,10)", "CHISQ.INV(1,2)", "CHISQ.INV(0,2)", "CHISQ.INV(0.5,300)",
     "CHISQ.INV.RT(0.050001,10)", "CHISQ.INV.RT(0.5,3)", "CHISQ.INV.RT(1E-10,5)", "CHISQ.INV.RT(0,2)", "CHISQ.INV.RT(1,2)",
     "CHISQ.TEST(Q1:R3,S1:T3)", "CHISQ.TEST(Q1:Q3,S1:S3)", "CHISQ.TEST(Q1:R3,S1:S3)", "CHISQ.TEST({1,2,3},{1,2,3})",
     "CHISQ.TEST(Q1:R1,S1:T1)",
     "CHIDIST(18.307,10)", "CHIDIST(2,1)", "CHIINV(0.050001,10)", "CHIINV(0.5,2)", "CHITEST(Q1:R3,S1:T3)",
     "F.DIST(15.2068649,6,4,TRUE)", "F.DIST(15.2068649,6,4,FALSE)", "F.DIST(1,1,1,TRUE)", "F.DIST(0.5,2,2,FALSE)",
     "F.DIST(-1,2,3,TRUE)", "F.DIST(1,0,3,TRUE)", "F.DIST(0,2,3,FALSE)", "F.DIST(0,1,3,TRUE)",
     "F.DIST.RT(15.2068649,6,4)", "F.DIST.RT(2,10,20)", "F.DIST.RT(0,1,1)",
     "F.INV(0.01,6,4)", "F.INV(0.95,3,20)", "F.INV(0,2,3)", "F.INV(1,2,3)",
     "F.INV.RT(0.01,6,4)", "F.INV.RT(0.5,10,10)", "F.INV.RT(1,2,3)",
     "F.TEST(V1:V5,W1:W5)", "F.TEST(A1:A10,B1:B10)", "F.TEST({1},{2,3})",
     "FDIST(15.2068649,6,4)", "FINV(0.01,6,4)", "FTEST(V1:V5,W1:W5)",
     "FISHER(0.75)", "FISHER(0)", "FISHER(1)", "FISHER(-0.5)", "FISHERINV(0.972955)", "FISHERINV(-2)", "FISHERINV(0)",
     "GAMMA(2.5)", "GAMMA(-3.75)", "GAMMA(0)", "GAMMA(-2)", "GAMMA(5)", "GAMMA(0.5)", "GAMMA(171)", "GAMMA(172)",
     "GAMMA(1E-5)", "GAMMA(30.3)", "GAMMA(-0.5)",
     "GAMMA.DIST(10.00001131,9,2,FALSE)", "GAMMA.DIST(10.00001131,9,2,TRUE)", "GAMMA.DIST(0,1,1,FALSE)",
     "GAMMA.DIST(2,0.5,1,TRUE)", "GAMMA.DIST(-1,2,2,TRUE)", "GAMMA.DIST(50,3,2,TRUE)", "GAMMA.DIST(0,0.5,1,FALSE)",
     "GAMMA.DIST(1,0,1,TRUE)",
     "GAMMA.INV(0.068094,9,2)", "GAMMA.INV(0.5,0.5,1)", "GAMMA.INV(0.999,3,2)", "GAMMA.INV(0,2,2)", "GAMMA.INV(1,2,2)",
     "GAMMADIST(10.00001131,9,2,TRUE)", "GAMMAINV(0.068094,9,2)",
     "GAMMALN(4)", "GAMMALN(0.5)", "GAMMALN(100)", "GAMMALN(0)", "GAMMALN(-1)", "GAMMALN(1)", "GAMMALN(2)", "GAMMALN(1E+300)",
     "GAMMALN.PRECISE(4)", "GAMMALN.PRECISE(1E-10)", "GAMMALN.PRECISE(2.5)",
     "GAUSS(2)", "GAUSS(-1)", "GAUSS(0)", "GAUSS(10)",
     "TRIMMEAN(A1:A10,0.2)", "TRIMMEAN(B1:B10,0.25)", "TRIMMEAN(A1:A10,0)", "TRIMMEAN(A1:A10,1)", "TRIMMEAN(A1:A10,0.1)",
     "TRIMMEAN({1,2,3},0.5)", "TRIMMEAN(A1:A10,0.39)", "TRIMMEAN(C1:C10,0.2)",
     "HYPGEOM.DIST(1,4,8,20,TRUE)", "HYPGEOM.DIST(1,4,8,20,FALSE)", "HYPGEOM.DIST(5,4,8,20,FALSE)",
     "HYPGEOM.DIST(0,4,8,20,TRUE)", "HYPGEOM.DIST(3,4,8,20,TRUE)", "HYPGEOMDIST(1,4,8,20)",
     "CONFIDENCE.NORM(0.05,2.5,50)", "CONFIDENCE.NORM(0,2.5,50)", "CONFIDENCE.NORM(0.05,2.5,0.5)",
     "CONFIDENCE.T(0.05,1,50)", "CONFIDENCE.T(0.05,1,1)", "CONFIDENCE.T(0.1,2,5.9)", "CONFIDENCE(0.05,2.5,50)",
     "LOGNORM.DIST(4,3.5,1.2,TRUE)", "LOGNORM.DIST(4,3.5,1.2,FALSE)", "LOGNORM.DIST(0,3.5,1.2,TRUE)",
     "LOGNORM.INV(0.039084,3.5,1.2)", "LOGNORM.INV(1,3.5,1.2)", "LOGNORMDIST(4,3.5,1.2)", "LOGINV(0.039084,3.5,1.2)",
     "MODE.MULT(X1:X6)", "MODE.MULT(A1:A10)", "MODE.MULT({1,2,3})", "MODE.MULT(1,2,2,3,3)", "MODE.MULT(B1:B10)",
     "NEGBINOM.DIST(10,5,0.25,FALSE)", "NEGBINOM.DIST(10,5,0.25,TRUE)", "NEGBINOM.DIST(-1,5,0.25,TRUE)",
     "NEGBINOM.DIST(0,1,1,FALSE)", "NEGBINOMDIST(10,5,0.25)",
     "PEARSON(A1:A10,B1:B10)", "PEARSON(D1:D8,E1:E8)", "PEARSON({1,2},{3})", "PEARSON(C1:C10,A1:A10)",
     "PHI(0.75)", "PHI(0)", "PHI(-3)",
     "PERCENTRANK.EXC(A1:A10,4)", "PERCENTRANK.EXC(A1:A10,4.5,5)", "PERCENTRANK.EXC(A1:A10,1)",
     "PERCENTRANK.EXC(A1:A10,10)", "PERCENTRANK.EXC(A1:A10,3,0)", "PERCENTRANK.EXC(A1:A10,9)", "PERCENTRANK.EXC(B1:B10,2)",
     "PERCENTRANK(A1:A10,4)", "PERCENTRANK(A1:A10,4.5,5)", "PERCENTRANK(B1:B10,2)",
     "LINEST(K1:K6,L1:L6)", "LINEST(K1:K6,L1:L6,TRUE,TRUE)", "LINEST(K1:K6,L1:M6,TRUE,TRUE)", "LINEST(K1:K6,L1:L6,FALSE,TRUE)",
     "LINEST(K1:K6)", "LINEST(K1:K6,L1:M6,FALSE)", "LINEST(K1:K6,L1:M6,FALSE,TRUE)", "LINEST({1,2,3},{1,1,1},TRUE,TRUE)",
     "LINEST(K1:K6,L1:L5)",
     "LOGEST(E1:E8,D1:D8)", "LOGEST(E1:E8,D1:D8,TRUE,TRUE)", "LOGEST(E1:E8,D1:D8,FALSE,FALSE)", "LOGEST(E1:E8)",
     "LOGEST(B1:B10,A1:A10)",
     "SKEW.P(A1:A10)", "SKEW.P(B1:B10)", "SKEW.P(1,1)",
     "STDEVPA(C1:C10)", "STDEVPA(A1:A10)", "STDEVPA(1,TRUE)", "VARPA(C1:C10)", "VARPA(A1:A10)", "VARPA({1,TRUE,\"a\"})",
     "DEVSQ(A1:A10)", "DEVSQ(B1:B10,1)", "DEVSQ(C1:C10)", "DEVSQ(\"a\")",
     "T.DIST.2T(1.96,60)", "T.DIST.2T(0,5)", "T.DIST.2T(-1,5)", "T.DIST.2T(1,0.5)", "T.DIST.2T(2,1.9)",
     "T.DIST.RT(1.96,60)", "T.DIST.RT(-1,5)", "T.DIST.RT(0,1)",
     "TREND(K1:K6,L1:L6)", "TREND(K1:K6,L1:L6,{7;8})", "TREND(K1:K6,L1:M6,L1:M2)", "TREND(K1:K6,L1:L6,7,FALSE)",
     "TREND(K1:K6)", "TREND(K1:K6,L1:L6,{7,8})",
     "GROWTH(E1:E8,D1:D8)", "GROWTH(E1:E8,D1:D8,{9;10})", "GROWTH(E1:E8,D1:D8,9,FALSE)", "GROWTH(B1:B10,A1:A10)",
     "PERMUTATIONA(3,2)", "PERMUTATIONA(2,2)", "PERMUTATIONA(0,0)", "PERMUTATIONA(-1,2)", "PERMUTATIONA(2.7,1.9)",
     "PROB({0,1,2,3},{0.2,0.3,0.1,0.4},2)", "PROB({0,1,2,3},{0.2,0.3,0.1,0.4},1,3)", "PROB({0,1},{0.5,0.6},0)",
     "PROB({0,1,2,3},{0.2,0.3,0.1,0.4},5)", "PROB({0,1},{0.5,0.5},1,0)", "PROB({0,1},{0.5},1)",
     "WEIBULL.DIST(105,20,100,TRUE)", "WEIBULL.DIST(105,20,100,FALSE)", "WEIBULL.DIST(-1,20,100,TRUE)",
     "WEIBULL.DIST(0,1,1,FALSE)", "WEIBULL(105,20,100,TRUE)",
     "COVAR(A1:A10,B1:B10)", "CRITBINOM(6,0.5,0.75)", "EXPONDIST(0.2,10,TRUE)", "NORMDIST(42,40,1.5,TRUE)",
     "NORMINV(0.908789,40,1.5)", "NORMSDIST(1.333333)", "NORMSINV(0.908789)", "POISSON(2,5,TRUE)",
     "TDIST(1.96,60,2)", "TDIST(1.96,60,1)", "TDIST(-1,60,1)", "TINV(0.05,60)", "TTEST(A1:A10,B1:B10,2,1)",
     "Z.TEST(A1:A10,4)", "Z.TEST(A1:A10,4,2)", "Z.TEST(A1:A10,5)", "ZTEST(A1:A10,4)", "Z.TEST({1},1)")

# Maths.
case("math",
     "AGGREGATE(9,6,Y1:Y6)", "AGGREGATE(9,0,Y1:Y6)", "AGGREGATE(4,6,Y1:Y6)", "AGGREGATE(12,6,Y1:Y6)",
     "AGGREGATE(16,6,Y1:Y6,0.3)", "AGGREGATE(19,6,Y1:Y6,1)", "AGGREGATE(15,6,Y1:Y6,5)", "AGGREGATE(20,6,Y1:Y6)",
     "AGGREGATE(3,0,A1:A10)", "AGGREGATE(14,6,A1:A10/(A1:A10>3),1)", "AGGREGATE(1,5,A1:A10)", "AGGREGATE(7,3,B1:B10)",
     "AGGREGATE(13,6,A1:A10)", "AGGREGATE(17,6,A1:A10,3)", "AGGREGATE(18,6,A1:A10,0.25)", "AGGREGATE(9,4,Y1:Y6)",
     "AGGREGATE(2,6,C1:C10)", "AGGREGATE(6,6,Y1:Y6)", "AGGREGATE(5,7,Y1:Y6)", "AGGREGATE(8,6,A1:A10)",
     "AGGREGATE(10,6,A1:A10)", "AGGREGATE(11,6,A1:A10)", "AGGREGATE(9,6,Y1:Y6,A1:A2)", "AGGREGATE(14,6,Y1:Y6)",
     "AGGREGATE(9,8,Y1:Y6)",
     "ARABIC(\"MCMXCIV\")", "ARABIC(\"mmxxvi\")", "ARABIC(\"\")", "ARABIC(\"-XIV\")", "ARABIC(\"ABC\")",
     "ARABIC(\"IIII\")", "ARABIC(\"IM\")", "ARABIC(\" XL \")", "ARABIC(12)",
     "ACOT(2)", "ACOT(0)", "ACOT(-1)", "ACOTH(6)", "ACOTH(1)", "ACOTH(-2)", "ACOTH(0.5)",
     "BASE(7,2)", "BASE(100,16,4)", "BASE(255,36)", "BASE(-1,2)", "BASE(10,1)", "BASE(10,37)", "BASE(2^53,2)",
     "BASE(15.9,2)", "BASE(5,2,300)", "BASE(5,2,256)", "BASE(2^53,36)",
     "CSC(15)", "CSC(0)", "CSCH(1.5)", "CSCH(0)", "COTH(2)", "COTH(0)", "COTH(-0.5)", "SEC(45)", "SEC(0)",
     "SECH(45)", "SECH(0)", "SECH(1000)", "CSC(2^27)", "SEC(2^27)",
     "DECIMAL(\"FF\",16)", "DECIMAL(\"111\",2)", "DECIMAL(\"zap\",36)", "DECIMAL(\"12\",2)", "DECIMAL(\"\",10)",
     "DECIMAL(\"ff\",16)", "DECIMAL(\"10\",1)", "DECIMAL(101,2)", "DECIMAL(\"7FFFFFFFFFFFFFFF\",16)",
     "MDETERM(H1:J3)", "MDETERM(H5:I6)", "MDETERM({1,2;3,4})", "MDETERM(H1:I3)", "MDETERM({3,6,1;1,1,0;3,10,2})",
     "MDETERM(C1:D2)", "MDETERM(5)",
     "MINVERSE(H1:J3)", "MINVERSE(H5:I6)", "MINVERSE(H8:I9)", "MINVERSE({4})", "MINVERSE(H1:I3)",
     "MUNIT(3)", "MUNIT(1)", "MUNIT(0)", "MUNIT(2.9)",
     "MULTINOMIAL(2,3,4)", "MULTINOMIAL(1)", "MULTINOMIAL(-1,2)", "MULTINOMIAL(2.5,1)", "MULTINOMIAL(A1:A3)",
     "SERIESSUM(0.5,0,2,{1,-0.5,0.0416666667,-0.0013888889})", "SERIESSUM(2,1,1,{1,2,3})", "SERIESSUM(2,1,1,A1:A3)",
     "SERIESSUM(\"a\",1,1,{1})",
     "ROMAN(499)", "ROMAN(499,1)", "ROMAN(499,2)", "ROMAN(499,3)", "ROMAN(499,4)", "ROMAN(1994,TRUE)",
     "ROMAN(1994,FALSE)", "ROMAN(0)", "ROMAN(4000)", "ROMAN(-1)", "ROMAN(3999)", "ROMAN(1999,1)", "ROMAN(1999,2)",
     "ROMAN(1999,3)", "ROMAN(1999,4)", "ROMAN(45,1)", "ROMAN(490,1)", "ROMAN(95,2)", "ROMAN(49,3)", "ROMAN(999,1)",
     "ROMAN(3999,4)", "ROMAN(1490,2)", "ROMAN(4,5)", "ROMAN(12.9)",
     "SUMX2MY2(A1:A10,B1:B10)", "SUMX2PY2(A1:A10,B1:B10)", "SUMXMY2(A1:A10,B1:B10)", "SUMXMY2({1,2},{1})",
     "SUMX2MY2(C1:C10,A1:A10)", "SUMX2PY2({1,2;3,4},{1,2;3,4})", "SUMXMY2(Y1:Y6,A1:A6)")

# Finance.
COUP = ["COUPDAYBS", "COUPDAYS", "COUPDAYSNC", "COUPNCD", "COUPNUM", "COUPPCD"]
each("finance", COUP,
     "DATE(2011,1,25),DATE(2011,11,15),2,1", "DATE(2011,1,25),DATE(2011,11,15),4,0",
     "DATE(2011,2,28),DATE(2016,8,31),2,0", "DATE(2011,1,25),DATE(2011,11,15),4,3", "DATE(2011,1,25),DATE(2011,11,15),2,4",
     "DATE(2011,11,15),DATE(2011,1,25),2", "DATE(2011,1,25),DATE(2011,11,15),3", "DATE(2012,2,29),DATE(2015,8,31),2,1",
     "DATE(2011,1,31),DATE(2011,11,30),1,2", "DATE(2011,1,25),DATE(2011,11,15),2,5")
case("finance",
     "ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0)",
     "ACCRINT(DATE(2008,3,5),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0,FALSE)",
     "ACCRINT(DATE(2008,4,5),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0,TRUE)",
     "ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,1)",
     "ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,2)",
     "ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,3)",
     "ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,4)",
     "ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,3,0)",
     "ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2008,2,1),0.1,1000,2,0)",
     "ACCRINT(DATE(2007,1,10),DATE(2007,7,10),DATE(2009,3,15),0.05,1000,4,1)",
     "ACCRINT(DATE(2007,1,10),DATE(2007,7,10),DATE(2009,3,15),0.05,1000,1,0,FALSE)",
     "ACCRINT(DATE(2007,1,10),DATE(2007,7,10),DATE(2009,3,15),0.05,1000,2,3,FALSE)",
     "ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0,1000,2,0)",
     "ACCRINTM(DATE(2008,4,1),DATE(2008,6,15),0.1,1000,3)", "ACCRINTM(DATE(2008,4,1),DATE(2008,6,15),0.1,1000)",
     "ACCRINTM(DATE(2008,4,1),DATE(2008,6,15),0.1,1000,1)", "ACCRINTM(DATE(2008,4,1),DATE(2008,6,15),0.1,1000,2)",
     "ACCRINTM(DATE(2008,4,1),DATE(2008,6,15),0.1,1000,4)", "ACCRINTM(DATE(2008,6,15),DATE(2008,4,1),0.1,1000)",
     "ACCRINTM(DATE(2008,4,1),DATE(2008,6,15),0.1,-1)",
     "AMORLINC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15,1)",
     "AMORLINC(2400,DATE(2008,8,19),DATE(2008,12,31),300,0,0.15,1)",
     "AMORLINC(2400,DATE(2008,8,19),DATE(2008,12,31),300,2,0.15,1)",
     "AMORLINC(2400,DATE(2008,8,19),DATE(2008,12,31),300,5,0.15,1)",
     "AMORLINC(2400,DATE(2008,8,19),DATE(2008,12,31),300,6,0.15,1)",
     "AMORLINC(2400,DATE(2008,8,19),DATE(2008,12,31),300,7,0.15,1)",
     "AMORLINC(2400,DATE(2008,8,19),DATE(2008,12,31),300,0,0.15,0)",
     "AMORLINC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15,3)",
     "AMORLINC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15,2)",
     "AMORLINC(2400,DATE(2008,8,19),DATE(2008,12,31),300,0,0.15,4)",
     "AMORLINC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15,5)",
     "AMORLINC(1000,DATE(2008,1,1),DATE(2008,12,31),0,0,0.3)",
     "AMORLINC(1000,DATE(2008,1,1),DATE(2008,12,31),0,3,0.3)",
     "AMORLINC(1000,DATE(2008,1,1),DATE(2008,12,31),0,4,0.3)",
     "RECEIVED(DATE(2008,2,15),DATE(2008,5,15),1000000,0.0575,2)", "RECEIVED(DATE(2008,2,15),DATE(2008,5,15),1000000,0.0575)",
     "RECEIVED(DATE(2008,2,15),DATE(2008,5,15),1000000,0.0575,1)", "RECEIVED(DATE(2008,2,15),DATE(2008,5,15),1000000,5)",
     "DB(1000000,100000,6,1,7)", "DB(1000000,100000,6,2,7)", "DB(1000000,100000,6,7,7)", "DB(1000000,100000,6,1)",
     "DB(1000000,100000,6,7)", "DB(1000,0,5,1)", "DB(1000,100,5,2.5)", "DB(1000,2000,5,1)", "DB(1000,100,5,1,13)",
     "DDB(2400,300,10*365,1)", "DDB(2400,300,120,1,2)", "DDB(2400,300,10,1,2)", "DDB(2400,300,10,10)",
     "DDB(2400,300,10,8,1.5)", "DDB(2400,300,10,11)", "DDB(2400,300,10,0.5)", "DDB(2400,300,10,2.5)", "DDB(2400,3000,10,1)",
     "DISC(DATE(2018,7,1),DATE(2048,1,1),97.975,100,1)", "DISC(DATE(2018,7,1),DATE(2048,1,1),97.975,100)",
     "DISC(DATE(2018,7,1),DATE(2018,1,1),97.975,100)", "DISC(DATE(2018,7,1),DATE(2019,1,1),97.975,100,3)",
     "DOLLARDE(1.02,16)", "DOLLARDE(1.1,32)", "DOLLARDE(1.1,0)", "DOLLARDE(-1.02,16)", "DOLLARDE(1.02,16.9)",
     "DOLLARDE(1.1,-1)", "DOLLARDE(1.5,10)",
     "DOLLARFR(1.125,16)", "DOLLARFR(1.125,32)", "DOLLARFR(-1.125,32)", "DOLLARFR(1.125,0)", "DOLLARFR(1.3,3)",
     "DURATION(DATE(2018,7,1),DATE(2048,1,1),0.08,0.09,2,1)", "DURATION(DATE(2018,7,1),DATE(2048,1,1),0.08,0.09,2)",
     "DURATION(DATE(2018,7,1),DATE(2048,1,1),0.08,0.09,1,0)", "DURATION(DATE(2018,7,1),DATE(2048,1,1),0.08,0.09,4,3)",
     "DURATION(DATE(2018,7,1),DATE(2048,1,1),-0.08,0.09,2)", "DURATION(DATE(2018,7,1),DATE(2018,12,1),0.08,0.09,2)",
     "MDURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1)", "MDURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,1)",
     "EFFECT(0.0525,4)", "EFFECT(0.05,1)", "EFFECT(0.05,0.5)", "EFFECT(0.05,4.9)", "EFFECT(-0.05,4)",
     "NOMINAL(0.053543,4)", "NOMINAL(0.05,1)", "NOMINAL(0.05,0)", "NOMINAL(0.05,12.5)",
     "IRR(F7:F12)", "IRR(F7:F10)", "IRR(F7:F12,-0.1)", "IRR({1,2,3})", "IRR(F7:F9,-0.1)", "IRR(F1:F5)", "IRR({-100,0,0,0,150})",
     "ISPMT(0.1/12,1,36,8000000)", "ISPMT(0.1,1,3,8000000)", "ISPMT(0.1,0,0,1)", "ISPMT(0.1,5,3,1000)",
     "CUMIPMT(0.09/12,30*12,125000,13,24,0)", "CUMIPMT(0.09/12,30*12,125000,1,1,0)", "CUMIPMT(0.09/12,30*12,125000,1,1,1)",
     "CUMIPMT(0.09/12,30*12,125000,13,24,2)", "CUMIPMT(0.09/12,30*12,-125000,13,24,0)", "CUMIPMT(0.09/12,30*12,125000,24,13,0)",
     "CUMIPMT(0.09/12,30*12,125000,2,5,1)", "CUMIPMT(0.09/12,30*12.5,125000,1.5,5.5,0)",
     "CUMPRINC(0.09/12,30*12,125000,13,24,0)", "CUMPRINC(0.09/12,30*12,125000,1,1,0)", "CUMPRINC(0.09/12,30*12,125000,1,1,1)",
     "CUMPRINC(0.09/12,30*12,125000,2,5,1)", "CUMPRINC(0,12,1200,1,3,0)",
     "PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0)", "PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,1)",
     "PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,2)", "PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,3)",
     "PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,4)", "PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,1)",
     "PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,4)", "PRICE(DATE(2008,2,15),DATE(2008,4,15),0.0575,0.065,100,2,0)",
     "PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,-0.065,100,2,0)", "PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,0,2)",
     "PRICE(DATE(2008,2,29),DATE(2017,8,31),0.0575,0.065,100,2,0)",
     "PRICEDISC(DATE(2008,2,16),DATE(2008,3,1),0.0525,100,2)", "PRICEDISC(DATE(2008,2,16),DATE(2008,3,1),0.0525,100)",
     "PRICEDISC(DATE(2008,2,16),DATE(2009,3,1),0.0525,100,1)",
     "PRICEMAT(DATE(2008,2,15),DATE(2008,4,13),DATE(2007,11,11),0.061,0.061,0)",
     "PRICEMAT(DATE(2008,2,15),DATE(2008,4,13),DATE(2007,11,11),0.061,0.061,1)",
     "PRICEMAT(DATE(2008,2,15),DATE(2009,4,13),DATE(2007,11,11),0.061,0.05,3)",
     "MIRR(F7:F12,0.1,0.12)", "MIRR(F7:F10,0.1,0.12)", "MIRR({1,2},0.1,0.1)", "MIRR({-1,2},-1,0.1)",
     "SLN(30000,7500,10)", "SLN(1,0,0)", "SYD(30000,7500,10,1)", "SYD(30000,7500,10,10)", "SYD(30000,7500,10,11)",
     "SYD(30000,7500,10,2.5)",
     "PDURATION(0.025,2000,2200)", "PDURATION(0,1,2)", "PDURATION(0.1,2,1)",
     "RRI(96,10000,11000)", "RRI(0,1,2)", "RRI(10,0,100)", "RRI(10,100,0)", "RRI(10,-100,-200)",
     "FVSCHEDULE(1,{0.09,0.11,0.1})", "FVSCHEDULE(100,A1:A3)", "FVSCHEDULE(100,C1:C3)",
     "TBILLEQ(DATE(2008,3,31),DATE(2008,6,1),0.0914)", "TBILLEQ(DATE(2008,3,31),DATE(2009,2,1),0.0914)",
     "TBILLEQ(DATE(2008,3,31),DATE(2009,4,1),0.0914)", "TBILLEQ(DATE(2008,3,31),DATE(2008,6,1),0)",
     "TBILLPRICE(DATE(2008,3,31),DATE(2008,6,1),0.09)", "TBILLPRICE(DATE(2008,3,31),DATE(2009,2,1),0.09)",
     "TBILLPRICE(DATE(2008,3,31),DATE(2008,6,1),5)",
     "TBILLYIELD(DATE(2008,3,31),DATE(2008,6,1),98.45)", "TBILLYIELD(DATE(2008,3,31),DATE(2009,6,1),98.45)",
     "TBILLYIELD(DATE(2008,3,31),DATE(2008,6,1),0)",
     "ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1)",
     "ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,1,15),DATE(2009,3,1),0.0785,0.0625,100,2,1)",
     "ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0)",
     "ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,4,3)",
     "ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,1,15),DATE(2009,3,1),0.0785,0.0625,100,1,0)",
     "ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,12,15),DATE(2009,3,1),0.0785,0.0625,100,2,1)",
     "ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0575,84.5,100,2,0)",
     "ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,1,15),DATE(2009,3,1),0.0575,84.5,100,2,1)",
     "ODDLPRICE(DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,2,0)",
     "ODDLPRICE(DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,2,1)",
     "ODDLPRICE(DATE(2008,2,7),DATE(2008,6,15),DATE(2007,6,15),0.0375,0.0405,100,4,0)",
     "ODDLPRICE(DATE(2007,8,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,2,0)",
     "ODDLYIELD(DATE(2008,4,20),DATE(2008,6,15),DATE(2007,12,24),0.0375,99.875,100,2,0)",
     "ODDLYIELD(DATE(2008,4,20),DATE(2008,6,15),DATE(2007,12,24),0.0375,99.875,100,2,1)",
     "ODDLYIELD(DATE(2008,4,20),DATE(2008,6,15),DATE(2007,8,24),0.0375,99.875,100,4,3)",
     "VDB(2400,300,3650,0,1)", "VDB(2400,300,120,0,1)", "VDB(2400,300,10,0,0.875,1.5)", "VDB(2400,300,10,6,10)",
     "VDB(2400,300,10,6,10,2,TRUE)", "VDB(2400,300,10,0,10)", "VDB(2400,300,10,3.5,7.25)", "VDB(2400,300,10,5,4)",
     "VDB(2400,300,10,0,11)", "VDB(2400,300,10,2,8,3)", "VDB(2400,300,10,2,8,3,TRUE)",
     "XIRR(F1:F5,G1:G5)", "XIRR(F1:F5,G1:G5,0.5)", "XIRR({1,2},{39448,39508})", "XIRR(F1:F5,G1:G4)",
     "XNPV(0.09,F1:F5,G1:G5)", "XNPV(0.09,F1:F3,G1:G5)", "XNPV(-1,F1:F5,G1:G5)", "XNPV(0,F1:F5,G1:G5)",
     "INTRATE(DATE(2008,2,15),DATE(2008,5,15),1000000,1014420,2)", "INTRATE(DATE(2008,2,15),DATE(2008,5,15),1000000,1014420)",
     "INTRATE(DATE(2008,2,15),DATE(2008,5,15),0,1014420)",
     "YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,95.04287,100,2,0)",
     "YIELD(DATE(2008,2,15),DATE(2008,4,15),0.0575,99.5,100,2,0)",
     "YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,95.04287,100,2,1)",
     "YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,95.04287,100,1,3)",
     "YIELD(DATE(2008,2,15),DATE(2016,11,15),0,95.04287,100,4)",
     "YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,0,100,2)",
     "YIELDDISC(DATE(2008,2,16),DATE(2008,3,1),99.795,100,2)", "YIELDDISC(DATE(2008,2,16),DATE(2008,3,1),99.795,100)",
     "YIELDMAT(DATE(2008,3,15),DATE(2008,11,3),DATE(2007,11,8),0.0625,100.0123,0)",
     "YIELDMAT(DATE(2008,3,15),DATE(2008,11,3),DATE(2007,11,8),0.0625,100.0123,1)",
     "YIELDMAT(DATE(2008,3,15),DATE(2009,11,3),DATE(2007,11,8),0.0625,100.0123,2)")

# Engineering.
case("eng",
     "BESSELI(1.5,1)", "BESSELI(1.5,0)", "BESSELI(10,3)", "BESSELI(-2,1)", "BESSELI(1,-1)", "BESSELI(2.5,2.9)", "BESSELI(0,0)",
     "BESSELI(0.01,5)", "BESSELI(50,2)",
     "BESSELJ(1.9,2)", "BESSELJ(0,0)", "BESSELJ(10,1)", "BESSELJ(50,3)", "BESSELJ(-2,1)", "BESSELJ(0.5,10)", "BESSELJ(3,0)",
     "BESSELJ(100,0)", "BESSELJ(5,5)",
     "BESSELK(1.5,1)", "BESSELK(0.1,0)", "BESSELK(5,4)", "BESSELK(0,1)", "BESSELK(3,0)", "BESSELK(0.5,3)", "BESSELK(-1,0)",
     "BESSELY(2.5,1)", "BESSELY(0.5,0)", "BESSELY(10,3)", "BESSELY(0,1)", "BESSELY(5,0)", "BESSELY(1,5)", "BESSELY(30,2)",
     "BIN2HEX(11111011,4)", "BIN2HEX(1110)", "BIN2HEX(1111111111)", "BIN2HEX(2)", "BIN2HEX(1,0)", "BIN2HEX(\"\")",
     "BIN2OCT(1001,3)", "BIN2OCT(1100100)", "BIN2OCT(1111111111)", "BIN2OCT(111,2)",
     "BITAND(1,5)", "BITAND(13,25)", "BITAND(-1,1)", "BITAND(2^48,1)", "BITAND(1.5,1)", "BITAND(2^48-1,7)",
     "BITOR(23,10)", "BITOR(0,0)", "BITXOR(5,3)", "BITXOR(5,\"a\")",
     "BITLSHIFT(4,2)", "BITLSHIFT(4,-1)", "BITLSHIFT(1,48)", "BITLSHIFT(1,47)", "BITLSHIFT(1,54)", "BITLSHIFT(5,-3)",
     "BITLSHIFT(1,-54)",
     "BITRSHIFT(13,2)", "BITRSHIFT(13,-2)", "BITRSHIFT(13,60)", "BITRSHIFT(2^47,47)",
     "DELTA(5,4)", "DELTA(5,5)", "DELTA(0.5)", "DELTA(\"a\",1)", "DELTA(0)",
     "DEC2OCT(58,3)", "DEC2OCT(-100)", "DEC2OCT(536870912)", "DEC2OCT(-536870912)", "DEC2OCT(8,1)",
     "ERF.PRECISE(0.745)", "ERF.PRECISE(1)", "ERF.PRECISE(-1)", "ERF(0.745)", "ERF(0,1.5)", "ERF(1,2)", "ERF(-1)",
     "ERF(2,1)", "ERF(10)",
     "ERFC(1)", "ERFC(-1)", "ERFC(5)", "ERFC(0)", "ERFC.PRECISE(1)", "ERFC.PRECISE(-2)",
     "GESTEP(5,4)", "GESTEP(5,5)", "GESTEP(-4,-5)", "GESTEP(-1)", "GESTEP(0)",
     "HEX2BIN(\"F\",8)", "HEX2BIN(\"B7\")", "HEX2BIN(\"FFFFFFFFFF\")", "HEX2BIN(\"200\")", "HEX2BIN(\"FFFFFFFE00\")",
     "HEX2OCT(\"F\",3)", "HEX2OCT(\"3B4E\")", "HEX2OCT(\"FFFFFFFF00\")", "HEX2OCT(\"FFE0000000\")", "HEX2OCT(\"20000000\")",
     "OCT2BIN(3,3)", "OCT2BIN(7777777000)", "OCT2BIN(1000)", "OCT2BIN(\"7777777777\")",
     "OCT2DEC(54)", "OCT2DEC(7777777533)", "OCT2DEC(8)", "OCT2HEX(100,4)", "OCT2HEX(7777777533)", "OCT2HEX(3777777777)",
     "COMPLEX(3,4)", "COMPLEX(3,4,\"j\")", "COMPLEX(0,1)", "COMPLEX(0,-1)", "COMPLEX(1,0)", "COMPLEX(0,0)",
     "COMPLEX(1.5,-2.25)", "COMPLEX(3,4,\"k\")", "COMPLEX(1E+20,1)", "COMPLEX(1/3,2/3)", "COMPLEX(3,4,\"I\")",
     "COMPLEX(-0.000001,1E-20)",
     "IMABS(\"5+12i\")", "IMABS(\"i\")", "IMABS(\"-j\")", "IMABS(\"abc\")", "IMABS(3)", "IMABS(\"1e2+1i\")",
     "IMABS(\"1+2ii\")", "IMABS(\"+i\")", "IMABS(\"3.5\")", "IMABS(\"1,5+2i\")",
     "IMAGINARY(\"3+4i\")", "IMAGINARY(\"0-j\")", "IMAGINARY(4)", "IMAGINARY(\"4\")", "IMAGINARY(\"-2.5i\")",
     "IMREAL(\"6-9i\")", "IMREAL(\"i\")",
     "IMARGUMENT(\"3+4i\")", "IMARGUMENT(0)", "IMARGUMENT(\"-1\")", "IMARGUMENT(\"-i\")",
     "IMCONJUGATE(\"3+4i\")", "IMCONJUGATE(\"-2j\")", "IMCONJUGATE(5)",
     "IMDIV(\"-238+240i\",\"10+24i\")", "IMDIV(\"1\",\"0\")", "IMDIV(\"1+i\",\"1-i\")", "IMDIV(\"1+i\",\"j\")",
     "IMPOWER(\"2+3i\",3)", "IMPOWER(\"2+3i\",0.5)", "IMPOWER(\"i\",2)", "IMPOWER(0,0)", "IMPOWER(0,2)", "IMPOWER(\"-4\",0.5)",
     "IMPOWER(\"1+i\",-1)",
     "IMPRODUCT(\"3+4i\",\"5-3i\")", "IMPRODUCT(\"1+2i\",30)", "IMPRODUCT({\"1+i\",\"2\"})", "IMPRODUCT(\"i\",\"i\")",
     "IMSUB(\"13+4i\",\"5+3i\")", "IMSUB(\"1+i\",\"1+i\")",
     "IMSUM(\"3+4i\",\"5-3i\")", "IMSUM(\"1+j\",\"i\")", "IMSUM(\"1+j\",\"2\")", "IMSUM(1,2,3)", "IMSUM(\"\",1)")
each("eng", ["IMCOS", "IMCOSH", "IMCOT", "IMCSC", "IMCSCH", "IMEXP", "IMLN", "IMLOG10", "IMLOG2", "IMSEC", "IMSECH",
             "IMSIN", "IMSINH", "IMSQRT", "IMTAN"],
     "\"1+i\"", "\"4-3j\"", "\"-2\"", "0", "\"0.5i\"")
case("eng",
     "CONVERT(1,\"lbm\",\"kg\")", "CONVERT(68,\"F\",\"C\")", "CONVERT(2.5,\"ft\",\"sec\")", "CONVERT(100,\"km\",\"mi\")",
     "CONVERT(1,\"kWh\",\"J\")", "CONVERT(1,\"Gbyte\",\"Mbyte\")", "CONVERT(1,\"Gibyte\",\"Mibyte\")", "CONVERT(1,\"m2\",\"ft2\")",
     "CONVERT(1,\"l\",\"gal\")", "CONVERT(1,\"hp\",\"W\")", "CONVERT(0,\"C\",\"K\")", "CONVERT(1,\"yr\",\"day\")",
     "CONVERT(1,\"cm\",\"in\")", "CONVERT(1,\"mmHg\",\"Pa\")", "CONVERT(1,\"ang\",\"m\")", "CONVERT(1,\"T\",\"ga\")",
     "CONVERT(1,\"cal\",\"J\")", "CONVERT(1,\"m/s\",\"mph\")", "CONVERT(1,\"kn\",\"m/h\")", "CONVERT(1,\"mi\",\"Nmi\")",
     "CONVERT(1,\"ly\",\"km\")", "CONVERT(1,\"m3\",\"l\")", "CONVERT(1,\"uk_pt\",\"l\")", "CONVERT(1,\"tsp\",\"ml\")",
     "CONVERT(1,\"atm\",\"psi\")", "CONVERT(1,\"Wh\",\"BTU\")", "CONVERT(1,\"eV\",\"J\")", "CONVERT(1,\"lbf\",\"N\")",
     "CONVERT(1,\"dyn\",\"N\")", "CONVERT(1,\"ozm\",\"g\")", "CONVERT(1,\"stone\",\"kg\")", "CONVERT(1,\"ton\",\"kg\")",
     "CONVERT(1,\"uk_ton\",\"kg\")", "CONVERT(1,\"u\",\"kg\")", "CONVERT(1,\"grain\",\"g\")", "CONVERT(1,\"Rank\",\"K\")",
     "CONVERT(1,\"Reau\",\"C\")", "CONVERT(100,\"C\",\"F\")", "CONVERT(1,\"hh\",\"m\")", "CONVERT(1,\"Pica\",\"pt\")",
     "CONVERT(1,\"survey_mi\",\"m\")", "CONVERT(1,\"ell\",\"m\")", "CONVERT(1,\"parsec\",\"ly\")", "CONVERT(1,\"yd\",\"m\")",
     "CONVERT(1,\"ha\",\"m2\")", "CONVERT(1,\"acre\",\"m2\")", "CONVERT(1,\"uk_acre\",\"m2\")", "CONVERT(1,\"Morgen\",\"m2\")",
     "CONVERT(1,\"cup\",\"l\")", "CONVERT(1,\"oz\",\"ml\")", "CONVERT(1,\"qt\",\"l\")", "CONVERT(1,\"uk_qt\",\"l\")",
     "CONVERT(1,\"uk_gal\",\"l\")", "CONVERT(1,\"barrel\",\"l\")", "CONVERT(1,\"bushel\",\"l\")", "CONVERT(1,\"GRT\",\"m3\")",
     "CONVERT(1,\"MTON\",\"m3\")", "CONVERT(1,\"in3\",\"l\")", "CONVERT(1,\"ft3\",\"l\")", "CONVERT(1,\"tbs\",\"ml\")",
     "CONVERT(1,\"hr\",\"s\")", "CONVERT(1,\"mn\",\"s\")", "CONVERT(1,\"min\",\"s\")", "CONVERT(1,\"Pa\",\"Torr\")",
     "CONVERT(1,\"bar\",\"Pa\")", "CONVERT(1,\"PS\",\"W\")", "CONVERT(1,\"HP\",\"kW\")", "CONVERT(1,\"BTU\",\"J\")",
     "CONVERT(1,\"flb\",\"J\")", "CONVERT(1,\"HPh\",\"J\")", "CONVERT(1,\"e\",\"J\")", "CONVERT(1,\"c\",\"J\")",
     "CONVERT(1,\"pond\",\"N\")", "CONVERT(1,\"brton\",\"kg\")", "CONVERT(1,\"cwt\",\"kg\")", "CONVERT(1,\"uk_cwt\",\"kg\")",
     "CONVERT(1,\"sg\",\"kg\")", "CONVERT(1,\"bit\",\"byte\")", "CONVERT(1,\"kbyte\",\"byte\")", "CONVERT(1,\"kibyte\",\"byte\")",
     "CONVERT(1,\"Ybyte\",\"Zbyte\")", "CONVERT(1,\"km/h\",\"m/s\")", "CONVERT(1,\"admkn\",\"kn\")", "CONVERT(1,\"ang2\",\"m2\")",
     "CONVERT(1,\"Nmi2\",\"m2\")", "CONVERT(1,\"mi3\",\"km3\")", "CONVERT(1,\"lt\",\"l\")", "CONVERT(1,\"L\",\"ml\")",
     "CONVERT(1,\"mi\",\"km\")", "CONVERT(1,\"uk_pt\",\"pt\")", "CONVERT(1,\"pt\",\"l\")", "CONVERT(1,\"us_pt\",\"l\")",
     "CONVERT(1,\"ft\",\"m\")", "CONVERT(1,\"Picapt\",\"mm\")", "CONVERT(1,\"pica\",\"mm\")", "CONVERT(1,\"lbm\",\"ozm\")",
     "CONVERT(1,\"m\",\"m\")", "CONVERT(1,\"kg\",\"m\")", "CONVERT(1,\"xyz\",\"m\")", "CONVERT(1,\"km\",\"Mm\")",
     "CONVERT(1,\"dam\",\"m\")", "CONVERT(1,\"da\",\"m\")", "CONVERT(1,\"Kibyte\",\"byte\")", "CONVERT(1,\"kft\",\"m\")",
     "CONVERT(1,\"mg\",\"kg\")", "CONVERT(1,\"ug\",\"g\")", "CONVERT(1,\"Gm\",\"m\")", "CONVERT(1,\"mil\",\"mm\")",
     "CONVERT(1,\"N\",\"dyn\")", "CONVERT(1,\"J\",\"erg\")", "CONVERT(1,\"W\",\"w\")", "CONVERT(1,\"sec\",\"day\")",
     "CONVERT(1,\"Wb\",\"Mx\")", "CONVERT(1,\"Mx\",\"Wb\")", "CONVERT(1,\"Wh\",\"kJ\")", "CONVERT(1,\"cel\",\"fah\")",
     "CONVERT(1,\"kel\",\"C\")", "CONVERT(\"a\",\"m\",\"km\")", "CONVERT(1,\"ly\",\"pc\")", "CONVERT(1,\"Nmi\",\"m\")",
     "CONVERT(1,\"in\",\"m\")", "CONVERT(1,\"ft2\",\"in2\")", "CONVERT(1,\"m3\",\"ft3\")", "CONVERT(1,\"gal\",\"in3\")",
     "CONVERT(1,\"ang3\",\"l\")", "CONVERT(1,\"Pica3\",\"l\")", "CONVERT(1,\"yd3\",\"m3\")", "CONVERT(1,\"m/hr\",\"m/s\")",
     "CONVERT(1,\"mmHg\",\"atm\")", "CONVERT(1,\"p\",\"Pa\")", "CONVERT(1,\"at\",\"Pa\")", "CONVERT(1,\"psi\",\"Pa\")",
     "CONVERT(1,\"ozm\",\"lbm\")", "CONVERT(1,\"cal\",\"BTU\")", "CONVERT(1,\"pt\",\"in\")", "CONVERT(1,\"Pica\",\"in\")",
     "CONVERT(1,\"uk_gal\",\"gal\")", "CONVERT(1,\"gal\",\"oz\")", "CONVERT(1,\"cup\",\"oz\")", "CONVERT(1,\"m\",\"km2\")")

# Text.
case("text",
     "BAHTTEXT(1234)", "BAHTTEXT(0)", "BAHTTEXT(21.25)", "BAHTTEXT(-5)", "BAHTTEXT(1000000)", "BAHTTEXT(11)", "BAHTTEXT(101)",
     "BAHTTEXT(21)", "BAHTTEXT(0.5)", "BAHTTEXT(1234567.89)", "BAHTTEXT(20000000)", "BAHTTEXT(0.999)",
     "DOLLAR(1234.567)", "DOLLAR(1234.567,-2)", "DOLLAR(-1234.567,4)", "DOLLAR(0.123,2)", "DOLLAR(99.5,0)", "DOLLAR(-0.004)",
     "DOLLAR(1234567)", "DOLLAR(\"a\")",
     "ARRAYTOTEXT(A1:B2)", "ARRAYTOTEXT(A1:B2,1)", "ARRAYTOTEXT(C1:C10)", "ARRAYTOTEXT(C1:C10,1)",
     "ARRAYTOTEXT({1,\"a\";TRUE,2.5},1)", "ARRAYTOTEXT(Y1:Y6)", "ARRAYTOTEXT(Y1:Y6,1)", "ARRAYTOTEXT(1/3)",
     "ARRAYTOTEXT(A1:A2,2)",
     "VALUETOTEXT(2.5)", "VALUETOTEXT(\"a\",1)", "VALUETOTEXT(TRUE)", "VALUETOTEXT(1/0)", "VALUETOTEXT(\"a\")",
     "VALUETOTEXT(C1:C3)", "VALUETOTEXT(C4,1)", "VALUETOTEXT(1/3,1)", "VALUETOTEXT(TRUE,1)",
     "TEXTSPLIT(\"a,b,c\",\",\")", "TEXTSPLIT(U1,\",\",\";\")", "TEXTSPLIT(\"a,,b\",\",\",,TRUE)", "TEXTSPLIT(\"a,,b\",\",\")",
     "TEXTSPLIT(\"a,b;c\",\",\",\";\",FALSE,0,\"-\")", "TEXTSPLIT(\"aXbxc\",\"x\",,,1)", "TEXTSPLIT(\"a-b+c\",{\"-\",\"+\"})",
     "TEXTSPLIT(\"abc\",\"\")", "TEXTSPLIT(\"a,b;c\",,\";\")", "TEXTSPLIT(\"a,b;c\",\",\",\";\")", "TEXTSPLIT(\"\",\",\")",
     "TEXTSPLIT(12.5,\",\")", "TEXTSPLIT(\"abc\",\"x\")",
     "REGEXTEST(\"abc123\",\"[0-9]+\")", "REGEXTEST(\"ABC\",\"abc\")", "REGEXTEST(\"ABC\",\"abc\",1)", "REGEXTEST(\"abc\",\"(\")",
     "REGEXTEST(\"abc\",\"^b\")",
     "REGEXEXTRACT(U3,\"[0-9]+\")", "REGEXEXTRACT(U3,\"[0-9]+\",1)", "REGEXEXTRACT(U4,\"(\\d+)-(\\d+)\",2)",
     "REGEXEXTRACT(\"abc\",\"x\")", "REGEXEXTRACT(\"ABC\",\"b\",0,1)", "REGEXEXTRACT(\"abc\",\"x\",1)",
     "REGEXREPLACE(U3,\"[0-9]+\",\"#\")", "REGEXREPLACE(U3,\"[0-9]+\",\"#\",1)", "REGEXREPLACE(U3,\"[0-9]+\",\"#\",-1)",
     "REGEXREPLACE(\"Hallo Welt\",\"(\\w+) (\\w+)\",\"$2 $1\")", "REGEXREPLACE(\"aAa\",\"a\",\"x\",0,1)",
     "REGEXREPLACE(\"abc\",\"x\",\"y\")",
     "ENCODEURL(\"a b&c=ä\")", "ENCODEURL(\"https://x.de/?q=1\")", "ENCODEURL(\"-_.~!*'()\")", "ENCODEURL(12.5)",
     "ENCODEURL(\"€\")")

# Arrays.
case("array",
     "EXPAND({1,2;3,4},3,3)", "EXPAND({1,2;3,4},3,3,0)", "EXPAND({1,2;3,4},1,3)", "EXPAND(A1:A3,,2,\"x\")",
     "EXPAND({1,2;3,4},2)", "EXPAND(5,2,2)",
     "HSTACK({1,2},{3;4})", "HSTACK(A1:A3,B1:B3)", "HSTACK(1,\"a\")", "HSTACK(C1:C2,Y1:Y3)",
     "VSTACK({1,2},{3})", "VSTACK(A1:B2,A3:B3)", "VSTACK(1,{2;3})",
     "WRAPCOLS({1,2,3,4,5},2)", "WRAPCOLS({1,2,3,4,5},2,0)", "WRAPCOLS({1,2;3,4},2)", "WRAPCOLS(A1:A5,0)",
     "WRAPCOLS(A1:A5,10)",
     "WRAPROWS({1,2,3,4,5},2)", "WRAPROWS({1,2,3,4,5},2,0)", "WRAPROWS(A1:A5,3,\"\")", "WRAPROWS({1;2;3},2)",
     "CHOOSECOLS(H1:J3,1,3)", "CHOOSECOLS(H1:J3,-1)", "CHOOSECOLS(H1:J3,4)", "CHOOSECOLS(H1:J3,{2,1})", "CHOOSECOLS(H1:J3,0)",
     "CHOOSEROWS(H1:J3,1,3)", "CHOOSEROWS(H1:J3,-1)", "CHOOSEROWS(H1:J3,4)", "CHOOSEROWS(H1:J3,{2;1})",
     "TAKE(H1:J3,2)", "TAKE(H1:J3,-1)", "TAKE(H1:J3,,2)", "TAKE(H1:J3,5,-2)", "TAKE(H1:J3,0)", "TAKE(H1:J3,-2,-2)",
     "DROP(H1:J3,1)", "DROP(H1:J3,-1,1)", "DROP(H1:J3,3)", "DROP(H1:J3,,-2)", "DROP(H1:J3,0,0)",
     "TOCOL(H1:J3)", "TOCOL(H1:J3,0,TRUE)", "TOCOL(Y1:Y6,2)", "TOCOL(C1:C10,1)", "TOCOL(C1:C10,3)", "TOCOL(C1:C10)",
     "TOROW(H1:J3)", "TOROW(H1:J3,0,TRUE)", "TOROW(Y1:Y6,2)", "TOROW(C1:C4,1)", "TOROW(H1:J3,4)",
     "TRIMRANGE(A1:A20)", "TRIMRANGE(C1:D20)", "TRIMRANGE(C1:C20,2)", "TRIMRANGE(C1:C20,0,2)", "TRIMRANGE(Z1:Z5)",
     "GROUPBY(N1:N8,P1:P8,SUM)", "GROUPBY(N1:N8,P1:P8,AVERAGE)", "GROUPBY(N1:N8,P1:P8,COUNT,0,0)",
     "GROUPBY(N1:O8,P1:P8,SUM)", "GROUPBY(N1:N8,P1:P8,SUM,,0)", "GROUPBY(N1:N8,P1:P8,SUM,,,-2)",
     "GROUPBY(N1:N8,P1:P8,MAX,0,1,-1)", "GROUPBY(N1:N8,P1:P8,SUM,,,,N1:N8<>\"Ost\")", "GROUPBY(N1:O8,P1:P8,SUM,,2)",
     "GROUPBY(N1:N8,P1:P8,MIN,,,2)", "GROUPBY(O1:O8,P1:P8,SUM,0,-1)",
     "PIVOTBY(N1:N8,O1:O8,P1:P8,SUM)", "PIVOTBY(N1:N8,O1:O8,P1:P8,SUM,,0,,0)", "PIVOTBY(N1:N8,O1:O8,P1:P8,COUNT)",
     "PIVOTBY(N1:N8,O1:O8,P1:P8,SUM,0,1,-1,1,1)")


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


def fill(ws) -> list[dict]:
    out = []
    for addr, v in SETUP:
        r = ws.Range(addr)
        if isinstance(v, str) and (v.startswith("=") or v.startswith("'")):
            r.Formula2 = v
        else:
            r.Value2 = v
    for addr, v in SETUP:
        r = ws.Range(addr)
        e = {"addr": addr, "value": encode(r.Value2)}
        if r.HasFormula:
            e["formula"] = r.Formula2[1:]
        out.append(e)
    return out


NAME_RE = re.compile(r"([A-Za-zÄÖÜäöüß_][A-Za-z0-9ÄÖÜäöüß_.]*)\(")


def run_cases(ws) -> dict[str, str]:
    names: dict[str, str] = {}
    cell = ws.Range(AT)
    t0 = time.time()
    for i, c in enumerate(CASES):
        try:
            cell.Formula2 = "=" + c["file"]
        except Exception:  # noqa: BLE001  Excel refuses the formula
            c["refused"] = True
            continue
        c["file"] = cell.Formula2[1:]
        c["local"] = cell.Formula2Local
        c["value"] = encode(cell.Value2)
        if cell.HasSpill:
            c["spill"] = encode(cell.SpillingToRange.Value2)
        en = NAME_RE.findall(c["file"])
        de = NAME_RE.findall(c["local"])
        if len(en) == len(de):
            for a, b in zip(en, de):
                names.setdefault(a.upper(), b.upper())
        cell.ClearContents()
        if i % 200 == 0:
            print(f"  {i}/{len(CASES)} cases, {time.time() - t0:.0f} s", flush=True)
    return names


def run_xlfn(app, cases: list[dict]) -> dict[str, str]:
    """Saves a workbook with one formula per function to TEMP and reads the prefixes Excel writes."""
    wb = app.Workbooks.Add()
    ws = wb.Worksheets(1)
    seen: dict[str, str] = {}
    for c in cases:
        if c.get("refused"):
            continue
        for nm in re.findall(r"([A-Z][A-Z0-9.]*)\(", c["file"]):
            seen.setdefault(nm, c["file"])
    row = 1
    for f in seen.values():
        try:
            ws.Cells(row, 1).Formula2 = "=" + f
            row += 1
        except Exception:  # noqa: BLE001
            pass
    path = os.path.join(tempfile.gettempdir(), "raabe-probe-functions2.xlsx")
    if os.path.exists(path):
        os.remove(path)
    wb.SaveAs(path, 51)
    wb.Close(False)
    with zipfile.ZipFile(path) as z:
        xml = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
    os.remove(path)
    prefixed = sorted(set(re.findall(r"(_xlfn\.(?:_xlws\.)?[A-Z][A-Z0-9.]*)\(", xml)))
    return {p.split(".", 1)[1].replace("_xlws.", ""): p for p in prefixed}


def main() -> None:
    x = ExcelTarget(visible=False)
    try:
        ws = x.wb.Worksheets(1)
        ws.Name = SHEET
        setup = fill(ws)
        print(f"{len(CASES)} cases", flush=True)
        names = run_cases(ws)
        xlfn = run_xlfn(x.app, CASES)
        data = {
            "excel": f"{x.app.Version} build {int(x.app.Build)}",
            "recorded": time.strftime("%Y-%m-%d"),
            "sheet": SHEET,
            "at": AT,
            "setup": setup,
            "names": dict(sorted(names.items())),
            "xlfn": xlfn,
            "cases": CASES,
        }
        OUT.parent.mkdir(parents=True, exist_ok=True)
        OUT.write_text(json.dumps(data, ensure_ascii=False, indent=0), encoding="utf-8", newline=chr(10))
    finally:
        try:
            x.close()
        except TypeError:  # the process is gone already (tasklist's message is not cp1252)
            pass
    print(f"done -> {OUT.relative_to(ROOT)}")


if __name__ == "__main__":
    sys.exit(main())
