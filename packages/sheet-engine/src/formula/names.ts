// Excel's German and English function names (recorded from Excel 365 through Formula2 and
// Formula2Local, 2026-10-27) and the argument hints shown while a formula is typed. The hints
// are written in Excel's style ("Zahl1; [Zahl2]; ..."): Excel offers no way to read its own
// through COM, so they follow its conventions rather than copying its help texts.

/** English name | German name | argument hint. */
const TABLE = `ABS|ABS|Zahl
ACOS|ARCCOS|Zahl
ACOSH|ARCCOSHYP|Zahl
ADDRESS|ADRESSE|Zeile; Spalte; [Abs]; [A1]; [Tabellenname]
AGGREGATE|AGGREGAT|Funktion; Optionen; Bezug1; ...
AND|UND|Wahrheitswert1; [Wahrheitswert2]; ...
ARABIC|ARABISCH|Text
AREAS|BEREICHE|Bezug
ASIN|ARCSIN|Zahl
ASINH|ARCSINHYP|Zahl
ATAN|ARCTAN|Zahl
ATAN2|ARCTAN2|x_Koordinate; y_Koordinate
ATANH|ARCTANHYP|Zahl
AVEDEV|MITTELABW|Zahl1; [Zahl2]; ...
AVERAGE|MITTELWERT|Zahl1; [Zahl2]; ...
AVERAGEA|MITTELWERTA|Wert1; [Wert2]; ...
AVERAGEIF|MITTELWERTWENN|Bereich; Kriterien; [Mittelwert_Bereich]
AVERAGEIFS|MITTELWERTWENNS|Mittelwert_Bereich; Kriterien_Bereich1; Kriterien1; ...
BASE|BASIS|Zahl; Basis; [Mindestlänge]
BIN2DEC|BININDEZ|Zahl
BINOM.DIST|BINOM.VERT|Zahl_Erfolge; Versuche; Erfolgswahrsch; Kumuliert
BINOM.INV|BINOM.INV|Versuche; Erfolgswahrsch; Alpha
BYCOL|NACHSPALTE|Matrix; Lambda
BYROW|NACHZEILE|Matrix; Lambda
CEILING|OBERGRENZE|Zahl; Schritt
CEILING.MATH|OBERGRENZE.MATHEMATIK|Zahl; [Schritt]; [Modus]
CELL|ZELLE|Infotyp; [Bezug]
CHAR|ZEICHEN|Zahl
CHISQ.TEST|CHIQU.TEST|Beob_Messwerte; Erwart_Werte
CHOOSE|WAHL|Index; Wert1; [Wert2]; ...
CHOOSECOLS|SPALTENWAHL|Matrix; Spaltenzahl1; [Spaltenzahl2]; ...
CHOOSEROWS|ZEILENWAHL|Matrix; Zeilenzahl1; [Zeilenzahl2]; ...
CLEAN|SÄUBERN|Text
CODE|CODE|Text
COLUMN|SPALTE|[Bezug]
COLUMNS|SPALTEN|Matrix
COMBIN|KOMBINATIONEN|Zahl; gewählte_Zahl
COMBINA|KOMBINATIONEN2|Zahl; gewählte_Zahl
COMPLEX|KOMPLEXE|Realteil; Imaginärteil; [Suffix]
CONCAT|TEXTKETTE|Text1; [Text2]; ...
CONCATENATE|VERKETTEN|Text1; [Text2]; ...
CONFIDENCE.NORM|KONFIDENZ.NORM|Alpha; Standabwn; Umfang
CONVERT|UMWANDELN|Zahl; Von_Maßeinheit; In_Maßeinheit
CORREL|KORREL|Matrix1; Matrix2
COS|COS|Zahl
COSH|COSHYP|Zahl
COT|COT|Zahl
COUNT|ANZAHL|Wert1; [Wert2]; ...
COUNTA|ANZAHL2|Wert1; [Wert2]; ...
COUNTBLANK|ANZAHLLEEREZELLEN|Bereich
COUNTIF|ZÄHLENWENN|Bereich; Suchkriterien
COUNTIFS|ZÄHLENWENNS|Kriterienbereich1; Kriterien1; ...
COVARIANCE.P|KOVARIANZ.P|Matrix1; Matrix2
COVARIANCE.S|KOVARIANZ.S|Matrix1; Matrix2
DATE|DATUM|Jahr; Monat; Tag
DATEDIF|DATEDIF|Ausgangsdatum; Enddatum; Einheit
DATEVALUE|DATWERT|Datumstext
DAY|TAG|Zahl
DAYS|TAGE|Zieldatum; Ausgangsdatum
DAYS360|TAGE360|Ausgangsdatum; Enddatum; [Methode]
DEC2BIN|DEZINBIN|Zahl; [Stellen]
DEC2HEX|DEZINHEX|Zahl; [Stellen]
DECIMAL|DEZIMAL|Text; Basis
DEGREES|GRAD|Winkel
DELTA|DELTA|Zahl1; [Zahl2]
DOLLAR|DM|Zahl; [Dezimalstellen]
EDATE|EDATUM|Ausgangsdatum; Monate
DROP|WEGLASSEN|Matrix; Zeilen; [Spalten]
EFFECT|EFFEKTIV|Nominalzins; Perioden
EOMONTH|MONATSENDE|Ausgangsdatum; Monate
ERF|GAUSSFEHLER|Untere_Grenze; [Obere_Grenze]
ERFC|GAUSSFKOMPL|Untere_Grenze
ERROR.TYPE|FEHLER.TYP|Fehlerwert
EVEN|GERADE|Zahl
EXACT|IDENTISCH|Text1; Text2
EXP|EXP|Zahl
EXPAND|ERWEITERN|Matrix; Zeilen; [Spalten]; [auffüllen_mit]
EXPON.DIST|EXPON.VERT|x; Lambda; Kumuliert
FACT|FAKULTÄT|Zahl
FACTDOUBLE|ZWEIFAKULTÄT|Zahl
FALSE|FALSCH|
FILTER|FILTER|Matrix; einschließen; [wenn_leer]
FIND|FINDEN|Suchtext; Text; [Erstes_Zeichen]
FIXED|FEST|Zahl; [Dezimalstellen]; [Keine_Punkte]
FLOOR|UNTERGRENZE|Zahl; Schritt
FLOOR.MATH|UNTERGRENZE.MATHEMATIK|Zahl; [Schritt]; [Modus]
FORECAST|SCHÄTZER|x; Y_Werte; X_Werte
FORECAST.LINEAR|PROGNOSE.LINEAR|x; Y_Werte; X_Werte
FORMULATEXT|FORMELTEXT|Bezug
FREQUENCY|HÄUFIGKEIT|Daten; Klassen
FV|ZW|Zins; Zzr; Rmz; [Bw]; [F]
GAMMA|GAMMA|x
GCD|GGT|Zahl1; [Zahl2]; ...
GEOMEAN|GEOMITTEL|Zahl1; [Zahl2]; ...
GESTEP|GGANZZAHL|Zahl; [Schritt]
HARMEAN|HARMITTEL|Zahl1; [Zahl2]; ...
HEX2DEC|HEXINDEZ|Zahl
HLOOKUP|WVERWEIS|Suchkriterium; Matrix; Zeilenindex; [Bereich_Verweis]
HOUR|STUNDE|Zahl
HSTACK|HSTAPELN|Matrix1; [Matrix2]; ...
HYPERLINK|HYPERLINK|Hyperlink_Adresse; [Freundlicher_Name]
IF|WENN|Wahrheitstest; [Wert_wenn_wahr]; [Wert_wenn_falsch]
IFERROR|WENNFEHLER|Wert; Wert_falls_Fehler
IFNA|WENNNV|Wert; Wert_bei_NV
IFS|WENNS|Wahrheitstest1; Wert_wenn_wahr1; ...
IMAGE|BILD|Quelle; [Alternativtext]; [Größe]; [Höhe]; [Breite]
IMSUM|IMSUMME|Komplexe_Zahl1; [Komplexe_Zahl2]; ...
INDEX|INDEX|Matrix; Zeile; [Spalte]; [Bereich]
INDIRECT|INDIREKT|Bezug; [A1]
INFO|INFO|Typ
INT|GANZZAHL|Zahl
INTERCEPT|ACHSENABSCHNITT|Y_Werte; X_Werte
IPMT|ZINSZ|Zins; Zr; Zzr; Bw; [Zw]; [F]
ISBLANK|ISTLEER|Wert
ISERR|ISTFEHL|Wert
ISERROR|ISTFEHLER|Wert
ISEVEN|ISTGERADE|Zahl
ISFORMULA|ISTFORMEL|Bezug
ISLOGICAL|ISTLOG|Wert
ISNA|ISTNV|Wert
ISNONTEXT|ISTKTEXT|Wert
ISNUMBER|ISTZAHL|Wert
ISODD|ISTUNGERADE|Zahl
ISOMITTED|WURDEAUSGELASSEN|Argument
ISOWEEKNUM|ISOKALENDERWOCHE|Datum
ISREF|ISTBEZUG|Wert
ISTEXT|ISTTEXT|Wert
KURT|KURT|Zahl1; [Zahl2]; ...
LAMBDA|LAMBDA|[Parameter_oder_Berechnung]; ...
LARGE|KGRÖSSTE|Matrix; k
LCM|KGV|Zahl1; [Zahl2]; ...
LEFT|LINKS|Text; [Anzahl_Zeichen]
LEN|LÄNGE|Text
LET|LET|Name1; Namenswert1; Berechnung_oder_Name2; ...
LINEST|RGP|Y_Werte; [X_Werte]; [Konstante]; [Stats]
LN|LN|Zahl
LOG|LOG|Zahl; [Basis]
LOG10|LOG10|Zahl
LOOKUP|VERWEIS|Suchkriterium; Suchvektor; [Ergebnisvektor]
LOWER|KLEIN|Text
MAKEARRAY|MATRIXERSTELLEN|Zeilen; Spalten; Lambda
MAP|MAP|Matrix1; Lambda_oder_Matrix2; ...
MATCH|VERGLEICH|Suchkriterium; Suchmatrix; [Vergleichstyp]
MAX|MAX|Zahl1; [Zahl2]; ...
MAXA|MAXA|Wert1; [Wert2]; ...
MAXIFS|MAXWENNS|Max_Bereich; Kriterien_Bereich1; Kriterien1; ...
MDETERM|MDET|Matrix
MEDIAN|MEDIAN|Zahl1; [Zahl2]; ...
MID|TEIL|Text; Erstes_Zeichen; Anzahl_Zeichen
MIN|MIN|Zahl1; [Zahl2]; ...
MINA|MINA|Wert1; [Wert2]; ...
MINIFS|MINWENNS|Min_Bereich; Kriterien_Bereich1; Kriterien1; ...
MINUTE|MINUTE|Zahl
MINVERSE|MINV|Matrix
MMULT|MMULT|Array1; Array2
MOD|REST|Zahl; Divisor
MODE|MODALWERT|Zahl1; [Zahl2]; ...
MODE.MULT|MODUS.VIELF|Zahl1; [Zahl2]; ...
MODE.SNGL|MODUS.EINF|Zahl1; [Zahl2]; ...
MONTH|MONAT|Zahl
MROUND|VRUNDEN|Zahl; Vielfaches
N|N|Wert
NA|NV|
NETWORKDAYS|NETTOARBEITSTAGE|Ausgangsdatum; Enddatum; [Freie_Tage]
NETWORKDAYS.INTL|NETTOARBEITSTAGE.INTL|Ausgangsdatum; Enddatum; [Wochenende]; [Freie_Tage]
NOMINAL|NOMINAL|Effektiver_Zins; Perioden
NORM.DIST|NORM.VERT|x; Mittelwert; Standabwn; Kumuliert
NORM.INV|NORM.INV|Wahrsch; Mittelwert; Standabwn
NORM.S.DIST|NORM.S.VERT|z; Kumuliert
NORM.S.INV|NORM.S.INV|Wahrsch
NOT|NICHT|Wahrheitswert
NOW|JETZT|
NPER|ZZR|Zins; Rmz; Bw; [Zw]; [F]
NPV|NBW|Zins; Wert1; [Wert2]; ...
NUMBERVALUE|ZAHLENWERT|Text; [Dezimaltrennzeichen]; [Gruppentrennzeichen]
ODD|UNGERADE|Zahl
OFFSET|BEREICH.VERSCHIEBEN|Bezug; Zeilen; Spalten; [Höhe]; [Breite]
OR|ODER|Wahrheitswert1; [Wahrheitswert2]; ...
PERCENTILE|QUANTIL|Matrix; k
PERCENTILE.EXC|QUANTIL.EXKL|Matrix; k
PERCENTILE.INC|QUANTIL.INKL|Matrix; k
PERCENTRANK.INC|QUANTILSRANG.INKL|Matrix; x; [Genauigkeit]
PERMUT|VARIATIONEN|Zahl; gewählte_Zahl
PI|PI|
PMT|RMZ|Zins; Zzr; Bw; [Zw]; [F]
POISSON.DIST|POISSON.VERT|x; Mittelwert; Kumuliert
POWER|POTENZ|Zahl; Potenz
PPMT|KAPZ|Zins; Zr; Zzr; Bw; [Zw]; [F]
PRODUCT|PRODUKT|Zahl1; [Zahl2]; ...
PROPER|GROSS2|Text
PV|BW|Zins; Zzr; Rmz; [Zw]; [F]
QUARTILE|QUARTILE|Matrix; Quartile
QUARTILE.EXC|QUARTILE.EXKL|Matrix; Quartile
QUARTILE.INC|QUARTILE.INKL|Matrix; Quartile
QUOTIENT|QUOTIENT|Zähler; Nenner
RADIANS|BOGENMASS|Winkel
RAND|ZUFALLSZAHL|
RANDARRAY|ZUFALLSMATRIX|[Zeilen]; [Spalten]; [Min]; [Max]; [ganze_Zahl]
RANDBETWEEN|ZUFALLSBEREICH|Untere_Zahl; Obere_Zahl
RANK|RANG|Zahl; Bezug; [Reihenfolge]
RANK.AVG|RANG.MITTELW|Zahl; Bezug; [Reihenfolge]
RANK.EQ|RANG.GLEICH|Zahl; Bezug; [Reihenfolge]
RATE|ZINS|Zzr; Rmz; Bw; [Zw]; [F]; [Schätzwert]
REDUCE|REDUCE|[Anfangswert]; Matrix; Lambda
REPLACE|ERSETZEN|Alter_Text; Erstes_Zeichen; Anzahl_Zeichen; Neuer_Text
REPT|WIEDERHOLEN|Text; Multiplikator
RIGHT|RECHTS|Text; [Anzahl_Zeichen]
ROMAN|RÖMISCH|Zahl; [Typ]
ROUND|RUNDEN|Zahl; Anzahl_Stellen
ROUNDDOWN|ABRUNDEN|Zahl; Anzahl_Stellen
ROUNDUP|AUFRUNDEN|Zahl; Anzahl_Stellen
ROW|ZEILE|[Bezug]
ROWS|ZEILEN|Matrix
RSQ|BESTIMMTHEITSMASS|Y_Werte; X_Werte
SCAN|SCAN|[Anfangswert]; Matrix; Lambda
SEARCH|SUCHEN|Suchtext; Text; [Erstes_Zeichen]
SECOND|SEKUNDE|Zahl
SEQUENCE|SEQUENZ|Zeilen; [Spalten]; [Anfang]; [Schritt]
SIGN|VORZEICHEN|Zahl
SIN|SIN|Zahl
SINH|SINHYP|Zahl
SKEW|SCHIEFE|Zahl1; [Zahl2]; ...
SLOPE|STEIGUNG|Y_Werte; X_Werte
SMALL|KKLEINSTE|Matrix; k
SORT|SORTIEREN|Matrix; [Sortierindex]; [Sortierreihenfolge]; [Nach_Spalte]
SORTBY|SORTIERENNACH|Matrix; Nach_Matrix1; [Sortierreihenfolge1]; ...
SQRT|WURZEL|Zahl
SQRTPI|WURZELPI|Zahl
STANDARDIZE|STANDARDISIERUNG|x; Mittelwert; Standabwn
STDEV|STABW|Zahl1; [Zahl2]; ...
STDEV.P|STABW.N|Zahl1; [Zahl2]; ...
STDEV.S|STABW.S|Zahl1; [Zahl2]; ...
STDEVA|STABWA|Wert1; [Wert2]; ...
STDEVP|STABWN|Zahl1; [Zahl2]; ...
STEYX|STFEHLERYX|Y_Werte; X_Werte
SUBSTITUTE|WECHSELN|Text; Alter_Text; Neuer_Text; [ntes_Auftreten]
SUBTOTAL|TEILERGEBNIS|Funktion; Bezug1; ...
SUM|SUMME|Zahl1; [Zahl2]; ...
SUMIF|SUMMEWENN|Bereich; Suchkriterien; [Summe_Bereich]
SUMIFS|SUMMEWENNS|Summe_Bereich; Kriterien_Bereich1; Kriterien1; ...
SUMPRODUCT|SUMMENPRODUKT|Array1; [Array2]; [Array3]; ...
SUMSQ|QUADRATESUMME|Zahl1; [Zahl2]; ...
SUMX2MY2|SUMMEX2MY2|Matrix_x; Matrix_y
SUMX2PY2|SUMMEX2PY2|Matrix_x; Matrix_y
SUMXMY2|SUMMEXMY2|Matrix_x; Matrix_y
SWITCH|ERSTERWERT|Ausdruck; Wert1; Ergebnis1; [Standard_oder_Wert2]; ...
T|T|Wert
T.DIST|T.VERT|x; Freiheitsgrade; Kumuliert
T.INV|T.INV|Wahrsch; Freiheitsgrade
T.INV.2T|T.INV.2S|Wahrsch; Freiheitsgrade
T.TEST|T.TEST|Matrix1; Matrix2; Seiten; Typ
TAKE|ÜBERNEHMEN|Matrix; Zeilen; [Spalten]
TAN|TAN|Zahl
TANH|TANHYP|Zahl
TEXT|TEXT|Wert; Textformat
TEXTAFTER|TEXTNACH|Text; Trennzeichen; [Instanznummer]; [Vergleichsmodus]; [Übereinstimmungsende]; [Wenn_nicht_gefunden]
TEXTBEFORE|TEXTVOR|Text; Trennzeichen; [Instanznummer]; [Vergleichsmodus]; [Übereinstimmungsende]; [Wenn_nicht_gefunden]
TEXTJOIN|TEXTVERKETTEN|Trennzeichen; Leer_ignorieren; Text1; [Text2]; ...
TEXTSPLIT|TEXTTEILEN|Text; Spaltentrennzeichen; [Zeilentrennzeichen]; [Leer_ignorieren]; [Vergleichsmodus]; [Auffüllen_mit]
TIME|ZEIT|Stunde; Minute; Sekunde
TIMEVALUE|ZEITWERT|Zeit
TOCOL|ZUSPALTE|Matrix; [ignorieren]; [nach_Spalte_scannen]
TODAY|HEUTE|
TOROW|ZUZEILE|Matrix; [ignorieren]; [nach_Spalte_scannen]
TRANSPOSE|MTRANS|Matrix
TREND|TREND|Y_Werte; [X_Werte]; [Neue_x_Werte]; [Konstante]
TRIM|GLÄTTEN|Text
TRUE|WAHR|
TRUNC|KÜRZEN|Zahl; [Anzahl_Stellen]
TYPE|TYP|Wert
UNICHAR|UNIZEICHEN|Zahl
UNICODE|UNICODE|Text
UNIQUE|EINDEUTIG|Matrix; [nach_Spalte]; [genau_einmal]
UPPER|GROSS|Text
VALUE|WERT|Text
VALUETOTEXT|WERTZUTEXT|Wert; [Format]
ARRAYTOTEXT|MATRIXZUTEXT|Matrix; [Format]
VAR|VARIANZ|Zahl1; [Zahl2]; ...
VAR.P|VAR.P|Zahl1; [Zahl2]; ...
VAR.S|VAR.S|Zahl1; [Zahl2]; ...
VARA|VARIANZA|Wert1; [Wert2]; ...
VARP|VARIANZEN|Zahl1; [Zahl2]; ...
VLOOKUP|SVERWEIS|Suchkriterium; Matrix; Spaltenindex; [Bereich_Verweis]
VSTACK|VSTAPELN|Matrix1; [Matrix2]; ...
WEEKDAY|WOCHENTAG|Zahl; [Typ]
WEEKNUM|KALENDERWOCHE|Fortlaufende_Zahl; [Zahl_Typ]
WORKDAY|ARBEITSTAG|Ausgangsdatum; Tage; [Freie_Tage]
WORKDAY.INTL|ARBEITSTAG.INTL|Ausgangsdatum; Tage; [Wochenende]; [Freie_Tage]
WRAPCOLS|SPALTENUMBRUCH|Vektor; Umbruch_Anzahl; [auffüllen_mit]
WRAPROWS|ZEILENUMBRUCH|Vektor; Umbruch_Anzahl; [auffüllen_mit]
XLOOKUP|XVERWEIS|Suchkriterium; Suchmatrix; Rückgabematrix; [Wenn_nicht_gefunden]; [Vergleichsmodus]; [Suchmodus]
XMATCH|XVERGLEICH|Suchkriterium; Suchmatrix; [Vergleichsmodus]; [Suchmodus]
XOR|XODER|Wahrheitswert1; [Wahrheitswert2]; ...
YEAR|JAHR|Zahl
YEARFRAC|BRTEILJAHRE|Ausgangsdatum; Enddatum; [Basis]`;

const EN_DE = new Map<string, string>();
const DE_EN = new Map<string, string>();
const HINTS: { name: string; args: string }[] = [];
for (const line of TABLE.split("\n")) {
  const [en, de, args] = line.split("|");
  EN_DE.set(en, de);
  DE_EN.set(de, en);
  HINTS.push({ name: de, args });
}

/** German function name for an English one (both upper case); undefined when unknown. */
export function germanName(en: string): string | undefined {
  return EN_DE.get(en.toUpperCase());
}

/** Whether Excel knows an English function name (upper case). */
export function isEnglishName(en: string): boolean {
  return EN_DE.has(en) || EXTRA_EN.has(en);
}

/** Functions known in the file form only (Excel shows them under the same name). */
const EXTRA_EN = new Set(["SINGLE"]);

/** English function name for a German one, any case; undefined when unknown. */
export function englishName(de: string): string | undefined {
  return DE_EN.get(de.toUpperCase());
}

/** Every German function name Excel knows, with its argument hint, sorted. */
export const ALL_FUNCTIONS_DE: { name: string; args: string }[] = HINTS.sort((a, b) => a.name.localeCompare(b.name, "de"));

/**
 * Functions that Excel writes with a prefix in the .xlsx file (added after Excel 2007): the
 * recorded ones from a workbook Excel saved (probe_excel_formulas.py), the rest from the
 * same naming rule. The model keeps formulas without the prefixes, as Excel's Range.Formula2
 * shows them (D-062); toXlsxFormula and fromXlsxFormula convert.
 */
export const XLFN: Record<string, string> = {
  "ACOT": "_xlfn.ACOT",
  "ACOTH": "_xlfn.ACOTH",
  "AGGREGATE": "_xlfn.AGGREGATE",
  "ARABIC": "_xlfn.ARABIC",
  "ARRAYTOTEXT": "_xlfn.ARRAYTOTEXT",
  "BASE": "_xlfn.BASE",
  "BINOM.DIST": "_xlfn.BINOM.DIST",
  "BINOM.INV": "_xlfn.BINOM.INV",
  "BYCOL": "_xlfn.BYCOL",
  "BYROW": "_xlfn.BYROW",
  "CEILING.MATH": "_xlfn.CEILING.MATH",
  "CHISQ.TEST": "_xlfn.CHISQ.TEST",
  "CHOOSECOLS": "_xlfn.CHOOSECOLS",
  "CHOOSEROWS": "_xlfn.CHOOSEROWS",
  "COMBINA": "_xlfn.COMBINA",
  "CONCAT": "_xlfn.CONCAT",
  "CONFIDENCE.NORM": "_xlfn.CONFIDENCE.NORM",
  "CONFIDENCE.T": "_xlfn.CONFIDENCE.T",
  "COT": "_xlfn.COT",
  "COTH": "_xlfn.COTH",
  "COVARIANCE.P": "_xlfn.COVARIANCE.P",
  "COVARIANCE.S": "_xlfn.COVARIANCE.S",
  "CSC": "_xlfn.CSC",
  "CSCH": "_xlfn.CSCH",
  "DAYS": "_xlfn.DAYS",
  "DECIMAL": "_xlfn.DECIMAL",
  "DROP": "_xlfn.DROP",
  "EXPAND": "_xlfn.EXPAND",
  "EXPON.DIST": "_xlfn.EXPON.DIST",
  "F.DIST": "_xlfn.F.DIST",
  "FILTER": "_xlfn._xlws.FILTER",
  "FLOOR.MATH": "_xlfn.FLOOR.MATH",
  "FORECAST.LINEAR": "_xlfn.FORECAST.LINEAR",
  "FORMULATEXT": "_xlfn.FORMULATEXT",
  "GAMMA": "_xlfn.GAMMA",
  "GAMMA.DIST": "_xlfn.GAMMA.DIST",
  "HSTACK": "_xlfn.HSTACK",
  "IFNA": "_xlfn.IFNA",
  "IFS": "_xlfn.IFS",
  "IMAGE": "_xlfn.IMAGE",
  "ISFORMULA": "_xlfn.ISFORMULA",
  "ISOMITTED": "_xlfn.ISOMITTED",
  "ISOWEEKNUM": "_xlfn.ISOWEEKNUM",
  "LAMBDA": "_xlfn.LAMBDA",
  "LET": "_xlfn.LET",
  "MAKEARRAY": "_xlfn.MAKEARRAY",
  "MAP": "_xlfn.MAP",
  "MAXIFS": "_xlfn.MAXIFS",
  "MINIFS": "_xlfn.MINIFS",
  "MODE.MULT": "_xlfn.MODE.MULT",
  "MODE.SNGL": "_xlfn.MODE.SNGL",
  "NETWORKDAYS.INTL": "_xlfn.NETWORKDAYS.INTL",
  "NORM.DIST": "_xlfn.NORM.DIST",
  "NORM.INV": "_xlfn.NORM.INV",
  "NORM.S.DIST": "_xlfn.NORM.S.DIST",
  "NORM.S.INV": "_xlfn.NORM.S.INV",
  "NUMBERVALUE": "_xlfn.NUMBERVALUE",
  "PERCENTILE.EXC": "_xlfn.PERCENTILE.EXC",
  "PERCENTILE.INC": "_xlfn.PERCENTILE.INC",
  "PERCENTRANK.EXC": "_xlfn.PERCENTRANK.EXC",
  "PERCENTRANK.INC": "_xlfn.PERCENTRANK.INC",
  "POISSON.DIST": "_xlfn.POISSON.DIST",
  "QUARTILE.EXC": "_xlfn.QUARTILE.EXC",
  "QUARTILE.INC": "_xlfn.QUARTILE.INC",
  "RANDARRAY": "_xlfn.RANDARRAY",
  "RANK.AVG": "_xlfn.RANK.AVG",
  "RANK.EQ": "_xlfn.RANK.EQ",
  "REDUCE": "_xlfn.REDUCE",
  "SCAN": "_xlfn.SCAN",
  "SEC": "_xlfn.SEC",
  "SECH": "_xlfn.SECH",
  "SEQUENCE": "_xlfn.SEQUENCE",
  "SHEET": "_xlfn.SHEET",
  "SHEETS": "_xlfn.SHEETS",
  "SORT": "_xlfn._xlws.SORT",
  "SORTBY": "_xlfn.SORTBY",
  "STDEV.P": "_xlfn.STDEV.P",
  "STDEV.S": "_xlfn.STDEV.S",
  "SWITCH": "_xlfn.SWITCH",
  "T.DIST": "_xlfn.T.DIST",
  "T.INV": "_xlfn.T.INV",
  "T.INV.2T": "_xlfn.T.INV.2T",
  "T.TEST": "_xlfn.T.TEST",
  "TAKE": "_xlfn.TAKE",
  "TEXTAFTER": "_xlfn.TEXTAFTER",
  "TEXTBEFORE": "_xlfn.TEXTBEFORE",
  "TEXTJOIN": "_xlfn.TEXTJOIN",
  "TEXTSPLIT": "_xlfn.TEXTSPLIT",
  "TOCOL": "_xlfn.TOCOL",
  "TOROW": "_xlfn.TOROW",
  "UNICHAR": "_xlfn.UNICHAR",
  "UNICODE": "_xlfn.UNICODE",
  "UNIQUE": "_xlfn.UNIQUE",
  "VALUETOTEXT": "_xlfn.VALUETOTEXT",
  "VAR.P": "_xlfn.VAR.P",
  "VAR.S": "_xlfn.VAR.S",
  "VSTACK": "_xlfn.VSTACK",
  "WORKDAY.INTL": "_xlfn.WORKDAY.INTL",
  "WRAPCOLS": "_xlfn.WRAPCOLS",
  "WRAPROWS": "_xlfn.WRAPROWS",
  "XLOOKUP": "_xlfn.XLOOKUP",
  "XMATCH": "_xlfn.XMATCH",
  "XOR": "_xlfn.XOR",
};
