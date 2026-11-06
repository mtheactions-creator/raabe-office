// Excel's German and English function names (recorded from Excel 365 through Formula2 and
// Formula2Local, 2026-10-27) and the argument hints shown while a formula is typed. The hints
// are written in Excel's style ("Zahl1; [Zahl2]; ..."): Excel offers no way to read its own
// through COM, so they follow its conventions rather than copying its help texts.

/** English name | German name | argument hint. */
const TABLE = `ABS|ABS|Zahl
ACCRINT|AUFGELZINS|Emission; Erster_Zinstermin; Abrechnung; Nominalzins; Nennwert; Häufigkeit; [Basis]; [Berechnungsmethode]
ACCRINTM|AUFGELZINSF|Emission; Abrechnung; Nominalzins; Nennwert; [Basis]
ACOS|ARCCOS|Zahl
ACOSH|ARCCOSHYP|Zahl
ACOT|ARCCOT|Zahl
ACOTH|ARCCOTHYP|Zahl
ADDRESS|ADRESSE|Zeile; Spalte; [Abs]; [A1]; [Tabellenname]
AGGREGATE|AGGREGAT|Funktion; Optionen; Bezug1; ...
AMORLINC|AMORLINEARK|Ansch_Wert; Kaufdatum; Erster_Zinstermin; Restwert; Periode; Rate; [Basis]
AND|UND|Wahrheitswert1; [Wahrheitswert2]; ...
ARABIC|ARABISCH|Text
AREAS|BEREICHE|Bezug
ARRAYTOTEXT|MATRIXZUTEXT|Matrix; [Format]
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
BAHTTEXT|BAHTTEXT|Zahl
BASE|BASIS|Zahl; Basis; [Mindestlänge]
BESSELI|BESSELI|x; n
BESSELJ|BESSELJ|x; n
BESSELK|BESSELK|x; n
BESSELY|BESSELY|x; n
BETA.DIST|BETA.VERT|x; Alpha; Beta; Kumuliert; [A]; [B]
BETA.INV|BETA.INV|Wahrsch; Alpha; Beta; [A]; [B]
BETADIST|BETAVERT|x; Alpha; Beta; [A]; [B]
BETAINV|BETAINV|Wahrsch; Alpha; Beta; [A]; [B]
BIN2DEC|BININDEZ|Zahl
BIN2HEX|BININHEX|Zahl; [Stellen]
BIN2OCT|BININOKT|Zahl; [Stellen]
BINOM.DIST|BINOM.VERT|Zahl_Erfolge; Versuche; Erfolgswahrsch; Kumuliert
BINOM.DIST.RANGE|BINOM.VERT.BEREICH|Versuche; Erfolgswahrsch; Zahl_Erfolge; [Zahl2_Erfolge]
BINOM.INV|BINOM.INV|Versuche; Erfolgswahrsch; Alpha
BINOMDIST|BINOMVERT|Zahl_Erfolge; Versuche; Erfolgswahrsch; Kumuliert
BITAND|BITUND|Zahl1; Zahl2
BITLSHIFT|BITLVERSCHIEB|Zahl; Verschiebebetrag
BITOR|BITODER|Zahl1; Zahl2
BITRSHIFT|BITRVERSCHIEB|Zahl; Verschiebebetrag
BITXOR|BITXODER|Zahl1; Zahl2
BYCOL|NACHSPALTE|Matrix; Lambda
BYROW|NACHZEILE|Matrix; Lambda
CEILING|OBERGRENZE|Zahl; Schritt
CEILING.MATH|OBERGRENZE.MATHEMATIK|Zahl; [Schritt]; [Modus]
CELL|ZELLE|Infotyp; [Bezug]
CHAR|ZEICHEN|Zahl
CHIDIST|CHIVERT|x; Freiheitsgrade
CHIINV|CHIINV|Wahrsch; Freiheitsgrade
CHISQ.DIST|CHIQU.VERT|x; Freiheitsgrade; Kumuliert
CHISQ.DIST.RT|CHIQU.VERT.RE|x; Freiheitsgrade
CHISQ.INV|CHIQU.INV|Wahrsch; Freiheitsgrade
CHISQ.INV.RT|CHIQU.INV.RE|Wahrsch; Freiheitsgrade
CHISQ.TEST|CHIQU.TEST|Beob_Messwerte; Erwart_Werte
CHITEST|CHITEST|Beob_Messwerte; Erwart_Werte
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
CONFIDENCE|KONFIDENZ|Alpha; Standabwn; Umfang
CONFIDENCE.NORM|KONFIDENZ.NORM|Alpha; Standabwn; Umfang
CONFIDENCE.T|KONFIDENZ.T|Alpha; Standabwn; Umfang
CONVERT|UMWANDELN|Zahl; Von_Maßeinheit; In_Maßeinheit
CORREL|KORREL|Matrix1; Matrix2
COS|COS|Zahl
COSH|COSHYP|Zahl
COT|COT|Zahl
COTH|COTHYP|Zahl
COUNT|ANZAHL|Wert1; [Wert2]; ...
COUNTA|ANZAHL2|Wert1; [Wert2]; ...
COUNTBLANK|ANZAHLLEEREZELLEN|Bereich
COUNTIF|ZÄHLENWENN|Bereich; Suchkriterien
COUNTIFS|ZÄHLENWENNS|Kriterienbereich1; Kriterien1; ...
COUPDAYBS|ZINSTERMTAGVA|Abrechnung; Fälligkeit; Häufigkeit; [Basis]
COUPDAYS|ZINSTERMTAGE|Abrechnung; Fälligkeit; Häufigkeit; [Basis]
COUPDAYSNC|ZINSTERMTAGNZ|Abrechnung; Fälligkeit; Häufigkeit; [Basis]
COUPNCD|ZINSTERMNZ|Abrechnung; Fälligkeit; Häufigkeit; [Basis]
COUPNUM|ZINSTERMZAHL|Abrechnung; Fälligkeit; Häufigkeit; [Basis]
COUPPCD|ZINSTERMVZ|Abrechnung; Fälligkeit; Häufigkeit; [Basis]
COVAR|KOVAR|Matrix1; Matrix2
COVARIANCE.P|KOVARIANZ.P|Matrix1; Matrix2
COVARIANCE.S|KOVARIANZ.S|Matrix1; Matrix2
CRITBINOM|KRITBINOM|Versuche; Erfolgswahrsch; Alpha
CSC|COSEC|Zahl
CSCH|COSECHYP|Zahl
CUMIPMT|KUMZINSZ|Zins; Zzr; Bw; Zeitraum_Anfang; Zeitraum_Ende; F
CUMPRINC|KUMKAPITAL|Zins; Zzr; Bw; Zeitraum_Anfang; Zeitraum_Ende; F
DATE|DATUM|Jahr; Monat; Tag
DATEDIF|DATEDIF|Ausgangsdatum; Enddatum; Einheit
DATEVALUE|DATWERT|Datumstext
DAY|TAG|Zahl
DAYS|TAGE|Zieldatum; Ausgangsdatum
DAYS360|TAGE360|Ausgangsdatum; Enddatum; [Methode]
DB|GDA2|Ansch_Wert; Restwert; Nutzungsdauer; Periode; [Monate]
DDB|GDA|Ansch_Wert; Restwert; Nutzungsdauer; Periode; [Faktor]
DEC2BIN|DEZINBIN|Zahl; [Stellen]
DEC2HEX|DEZINHEX|Zahl; [Stellen]
DEC2OCT|DEZINOKT|Zahl; [Stellen]
DECIMAL|DEZIMAL|Text; Basis
DEGREES|GRAD|Winkel
DELTA|DELTA|Zahl1; [Zahl2]
DEVSQ|SUMQUADABW|Zahl1; [Zahl2]; ...
DISC|DISAGIO|Abrechnung; Fälligkeit; Kurs; Rückzahlung; [Basis]
DOLLAR|DM|Zahl; [Dezimalstellen]
DOLLARDE|NOTIERUNGDEZ|Zahl; Teiler
DOLLARFR|NOTIERUNGBRU|Zahl; Teiler
DROP|WEGLASSEN|Matrix; Zeilen; [Spalten]
DURATION|DURATION|Abrechnung; Fälligkeit; Nominalzins; Rendite; Häufigkeit; [Basis]
EDATE|EDATUM|Ausgangsdatum; Monate
EFFECT|EFFEKTIV|Nominalzins; Perioden
ENCODEURL|URLCODIEREN|Text
EOMONTH|MONATSENDE|Ausgangsdatum; Monate
ERF|GAUSSFEHLER|Untere_Grenze; [Obere_Grenze]
ERF.PRECISE|GAUSSF.GENAU|X
ERFC|GAUSSFKOMPL|Untere_Grenze
ERFC.PRECISE|GAUSSFKOMPL.GENAU|X
ERROR.TYPE|FEHLER.TYP|Fehlerwert
EVEN|GERADE|Zahl
EXACT|IDENTISCH|Text1; Text2
EXP|EXP|Zahl
EXPAND|ERWEITERN|Matrix; Zeilen; [Spalten]; [auffüllen_mit]
EXPON.DIST|EXPON.VERT|x; Lambda; Kumuliert
EXPONDIST|EXPONVERT|x; Lambda; Kumuliert
F.DIST|F.VERT|x; Freiheitsgrade1; Freiheitsgrade2; Kumuliert
F.DIST.RT|F.VERT.RE|x; Freiheitsgrade1; Freiheitsgrade2
F.INV|F.INV|Wahrsch; Freiheitsgrade1; Freiheitsgrade2
F.INV.RT|F.INV.RE|Wahrsch; Freiheitsgrade1; Freiheitsgrade2
F.TEST|F.TEST|Matrix1; Matrix2
FACT|FAKULTÄT|Zahl
FACTDOUBLE|ZWEIFAKULTÄT|Zahl
FALSE|FALSCH|
FDIST|FVERT|x; Freiheitsgrade1; Freiheitsgrade2
FILTER|FILTER|Matrix; einschließen; [wenn_leer]
FIND|FINDEN|Suchtext; Text; [Erstes_Zeichen]
FINV|FINV|Wahrsch; Freiheitsgrade1; Freiheitsgrade2
FISHER|FISHER|x
FISHERINV|FISHERINV|y
FIXED|FEST|Zahl; [Dezimalstellen]; [Keine_Punkte]
FLOOR|UNTERGRENZE|Zahl; Schritt
FLOOR.MATH|UNTERGRENZE.MATHEMATIK|Zahl; [Schritt]; [Modus]
FORECAST|SCHÄTZER|x; Y_Werte; X_Werte
FORECAST.LINEAR|PROGNOSE.LINEAR|x; Y_Werte; X_Werte
FORMULATEXT|FORMELTEXT|Bezug
FREQUENCY|HÄUFIGKEIT|Daten; Klassen
FTEST|FTEST|Matrix1; Matrix2
FV|ZW|Zins; Zzr; Rmz; [Bw]; [F]
FVSCHEDULE|ZW2|Kapital; Zinsen
GAMMA|GAMMA|x
GAMMA.DIST|GAMMA.VERT|x; Alpha; Beta; Kumuliert
GAMMA.INV|GAMMA.INV|Wahrsch; Alpha; Beta
GAMMADIST|GAMMAVERT|x; Alpha; Beta; Kumuliert
GAMMAINV|GAMMAINV|Wahrsch; Alpha; Beta
GAMMALN|GAMMALN|x
GAMMALN.PRECISE|GAMMALN.GENAU|x
GAUSS|GAUSS|x
GCD|GGT|Zahl1; [Zahl2]; ...
GEOMEAN|GEOMITTEL|Zahl1; [Zahl2]; ...
GESTEP|GGANZZAHL|Zahl; [Schritt]
GROUPBY|GRUPPIERENNACH|Zeilenfelder; Werte; Funktion; [Feldüberschriften]; [Gesamttiefe]; [Sortierreihenfolge]; [Filtermatrix]; [Feldbeziehung]
GROWTH|VARIATION|Y_Werte; [X_Werte]; [Neue_x_Werte]; [Konstante]
HARMEAN|HARMITTEL|Zahl1; [Zahl2]; ...
HEX2BIN|HEXINBIN|Zahl; [Stellen]
HEX2DEC|HEXINDEZ|Zahl
HEX2OCT|HEXINOKT|Zahl; [Stellen]
HLOOKUP|WVERWEIS|Suchkriterium; Matrix; Zeilenindex; [Bereich_Verweis]
HOUR|STUNDE|Zahl
HSTACK|HSTAPELN|Matrix1; [Matrix2]; ...
HYPERLINK|HYPERLINK|Hyperlink_Adresse; [Freundlicher_Name]
HYPGEOM.DIST|HYPGEOM.VERT|Erfolge_S; Umfang_S; Erfolge_G; Umfang_G; Kumuliert
HYPGEOMDIST|HYPGEOMVERT|Erfolge_S; Umfang_S; Erfolge_G; Umfang_G
IF|WENN|Wahrheitstest; [Wert_wenn_wahr]; [Wert_wenn_falsch]
IFERROR|WENNFEHLER|Wert; Wert_falls_Fehler
IFNA|WENNNV|Wert; Wert_bei_NV
IFS|WENNS|Wahrheitstest1; Wert_wenn_wahr1; ...
IMABS|IMABS|Komplexe_Zahl
IMAGE|BILD|Quelle; [Alternativtext]; [Größe]; [Höhe]; [Breite]
IMAGINARY|IMAGINÄRTEIL|Komplexe_Zahl
IMARGUMENT|IMARGUMENT|Komplexe_Zahl
IMCONJUGATE|IMKONJUGIERTE|Komplexe_Zahl
IMCOS|IMCOS|Komplexe_Zahl
IMCOSH|IMCOSHYP|Komplexe_Zahl
IMCOT|IMCOT|Komplexe_Zahl
IMCSC|IMCOSEC|Komplexe_Zahl
IMCSCH|IMCOSECHYP|Komplexe_Zahl
IMDIV|IMDIV|Komplexe_Zahl1; Komplexe_Zahl2
IMEXP|IMEXP|Komplexe_Zahl
IMLN|IMLN|Komplexe_Zahl
IMLOG10|IMLOG10|Komplexe_Zahl
IMLOG2|IMLOG2|Komplexe_Zahl
IMPOWER|IMAPOTENZ|Komplexe_Zahl; Potenz
IMPRODUCT|IMPRODUKT|Komplexe_Zahl1; [Komplexe_Zahl2]; ...
IMREAL|IMREALTEIL|Komplexe_Zahl
IMSEC|IMSEC|Komplexe_Zahl
IMSECH|IMSECHYP|Komplexe_Zahl
IMSIN|IMSIN|Komplexe_Zahl
IMSINH|IMSINHYP|Komplexe_Zahl
IMSQRT|IMWURZEL|Komplexe_Zahl
IMSUB|IMSUB|Komplexe_Zahl1; Komplexe_Zahl2
IMSUM|IMSUMME|Komplexe_Zahl1; [Komplexe_Zahl2]; ...
IMTAN|IMTAN|Komplexe_Zahl
INDEX|INDEX|Matrix; Zeile; [Spalte]; [Bereich]
INDIRECT|INDIREKT|Bezug; [A1]
INFO|INFO|Typ
INT|GANZZAHL|Zahl
INTERCEPT|ACHSENABSCHNITT|Y_Werte; X_Werte
INTRATE|ZINSSATZ|Abrechnung; Fälligkeit; Anlage; Rückzahlung; [Basis]
IPMT|ZINSZ|Zins; Zr; Zzr; Bw; [Zw]; [F]
IRR|IKV|Werte; [Schätzwert]
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
ISPMT|ISPMT|Zins; Zr; Zzr; Bw
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
LOGEST|RKP|Y_Werte; [X_Werte]; [Konstante]; [Stats]
LOGINV|LOGINV|Wahrsch; Mittelwert; Standabwn
LOGNORM.DIST|LOGNORM.VERT|x; Mittelwert; Standabwn; Kumuliert
LOGNORM.INV|LOGNORM.INV|Wahrsch; Mittelwert; Standabwn
LOGNORMDIST|LOGNORMVERT|x; Mittelwert; Standabwn
LOOKUP|VERWEIS|Suchkriterium; Suchvektor; [Ergebnisvektor]
LOWER|KLEIN|Text
MAKEARRAY|MATRIXERSTELLEN|Zeilen; Spalten; Lambda
MAP|MAP|Matrix1; Lambda_oder_Matrix2; ...
MATCH|VERGLEICH|Suchkriterium; Suchmatrix; [Vergleichstyp]
MAX|MAX|Zahl1; [Zahl2]; ...
MAXA|MAXA|Wert1; [Wert2]; ...
MAXIFS|MAXWENNS|Max_Bereich; Kriterien_Bereich1; Kriterien1; ...
MDETERM|MDET|Matrix
MDURATION|MDURATION|Abrechnung; Fälligkeit; Nominalzins; Rendite; Häufigkeit; [Basis]
MEDIAN|MEDIAN|Zahl1; [Zahl2]; ...
MID|TEIL|Text; Erstes_Zeichen; Anzahl_Zeichen
MIN|MIN|Zahl1; [Zahl2]; ...
MINA|MINA|Wert1; [Wert2]; ...
MINIFS|MINWENNS|Min_Bereich; Kriterien_Bereich1; Kriterien1; ...
MINUTE|MINUTE|Zahl
MINVERSE|MINV|Matrix
MIRR|QIKV|Werte; Investition; Reinvestition
MMULT|MMULT|Array1; Array2
MOD|REST|Zahl; Divisor
MODE|MODALWERT|Zahl1; [Zahl2]; ...
MODE.MULT|MODUS.VIELF|Zahl1; [Zahl2]; ...
MODE.SNGL|MODUS.EINF|Zahl1; [Zahl2]; ...
MONTH|MONAT|Zahl
MROUND|VRUNDEN|Zahl; Vielfaches
MULTINOMIAL|POLYNOMIAL|Zahl1; [Zahl2]; ...
MUNIT|MEINHEIT|Größe
N|N|Wert
NA|NV|
NEGBINOM.DIST|NEGBINOM.VERT|Zahl_Mißerfolge; Zahl_Erfolge; Erfolgswahrsch; Kumuliert
NEGBINOMDIST|NEGBINOMVERT|Zahl_Mißerfolge; Zahl_Erfolge; Erfolgswahrsch
NETWORKDAYS|NETTOARBEITSTAGE|Ausgangsdatum; Enddatum; [Freie_Tage]
NETWORKDAYS.INTL|NETTOARBEITSTAGE.INTL|Ausgangsdatum; Enddatum; [Wochenende]; [Freie_Tage]
NOMINAL|NOMINAL|Effektiver_Zins; Perioden
NORM.DIST|NORM.VERT|x; Mittelwert; Standabwn; Kumuliert
NORM.INV|NORM.INV|Wahrsch; Mittelwert; Standabwn
NORM.S.DIST|NORM.S.VERT|z; Kumuliert
NORM.S.INV|NORM.S.INV|Wahrsch
NORMDIST|NORMVERT|x; Mittelwert; Standabwn; Kumuliert
NORMINV|NORMINV|Wahrsch; Mittelwert; Standabwn
NORMSDIST|STANDNORMVERT|z
NORMSINV|STANDNORMINV|Wahrsch
NOT|NICHT|Wahrheitswert
NOW|JETZT|
NPER|ZZR|Zins; Rmz; Bw; [Zw]; [F]
NPV|NBW|Zins; Wert1; [Wert2]; ...
NUMBERVALUE|ZAHLENWERT|Text; [Dezimaltrennzeichen]; [Gruppentrennzeichen]
OCT2BIN|OKTINBIN|Zahl; [Stellen]
OCT2DEC|OKTINDEZ|Zahl
OCT2HEX|OKTINHEX|Zahl; [Stellen]
ODD|UNGERADE|Zahl
ODDFPRICE|UNREGER.KURS|Abrechnung; Fälligkeit; Emission; Erster_Zinstermin; Zins; Rendite; Rückzahlung; Häufigkeit; [Basis]
ODDFYIELD|UNREGER.REND|Abrechnung; Fälligkeit; Emission; Erster_Zinstermin; Zins; Kurs; Rückzahlung; Häufigkeit; [Basis]
ODDLPRICE|UNREGLE.KURS|Abrechnung; Fälligkeit; Letzter_Zinstermin; Zins; Rendite; Rückzahlung; Häufigkeit; [Basis]
ODDLYIELD|UNREGLE.REND|Abrechnung; Fälligkeit; Letzter_Zinstermin; Zins; Kurs; Rückzahlung; Häufigkeit; [Basis]
OFFSET|BEREICH.VERSCHIEBEN|Bezug; Zeilen; Spalten; [Höhe]; [Breite]
OR|ODER|Wahrheitswert1; [Wahrheitswert2]; ...
PDURATION|PDURATION|Zins; Bw; Zw
PEARSON|PEARSON|Matrix1; Matrix2
PERCENTILE|QUANTIL|Matrix; k
PERCENTILE.EXC|QUANTIL.EXKL|Matrix; k
PERCENTILE.INC|QUANTIL.INKL|Matrix; k
PERCENTRANK|QUANTILSRANG|Matrix; x; [Genauigkeit]
PERCENTRANK.EXC|QUANTILSRANG.EXKL|Matrix; x; [Genauigkeit]
PERCENTRANK.INC|QUANTILSRANG.INKL|Matrix; x; [Genauigkeit]
PERMUT|VARIATIONEN|Zahl; gewählte_Zahl
PERMUTATIONA|VARIATIONEN2|Zahl; gewählte_Zahl
PHI|PHI|x
PI|PI|
PIVOTBY|PIVOTMIT|Zeilenfelder; Spaltenfelder; Werte; Funktion; [Feldüberschriften]; [Zeilensummentiefe]; [Zeilensortierreihenfolge]; [Spaltensummentiefe]; [Spaltensortierreihenfolge]; [Filtermatrix]; [Relative_zu]
PMT|RMZ|Zins; Zzr; Bw; [Zw]; [F]
POISSON|POISSON|x; Mittelwert; Kumuliert
POISSON.DIST|POISSON.VERT|x; Mittelwert; Kumuliert
POWER|POTENZ|Zahl; Potenz
PPMT|KAPZ|Zins; Zr; Zzr; Bw; [Zw]; [F]
PRICE|KURS|Abrechnung; Fälligkeit; Zins; Rendite; Rückzahlung; Häufigkeit; [Basis]
PRICEDISC|KURSDISAGIO|Abrechnung; Fälligkeit; Disagio; Rückzahlung; [Basis]
PRICEMAT|KURSFÄLLIG|Abrechnung; Fälligkeit; Emission; Zins; Rendite; [Basis]
PROB|WAHRSCHBEREICH|Beob_Werte; Beob_Wahrsch; Untergrenze; [Obergrenze]
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
RECEIVED|AUSZAHLUNG|Abrechnung; Fälligkeit; Anlage; Disagio; [Basis]
REDUCE|REDUCE|[Anfangswert]; Matrix; Lambda
REGEXEXTRACT|REGEXEXTRAHIEREN|Text; Muster; [Rückgabemodus]; [Groß-/Kleinschreibung]
REGEXREPLACE|REGEXERSETZEN|Text; Muster; Ersatz; [Vorkommen]; [Groß-/Kleinschreibung]
REGEXTEST|REGEXTESTEN|Text; Muster; [Groß-/Kleinschreibung]
REPLACE|ERSETZEN|Alter_Text; Erstes_Zeichen; Anzahl_Zeichen; Neuer_Text
REPT|WIEDERHOLEN|Text; Multiplikator
RIGHT|RECHTS|Text; [Anzahl_Zeichen]
ROMAN|RÖMISCH|Zahl; [Typ]
ROUND|RUNDEN|Zahl; Anzahl_Stellen
ROUNDDOWN|ABRUNDEN|Zahl; Anzahl_Stellen
ROUNDUP|AUFRUNDEN|Zahl; Anzahl_Stellen
ROW|ZEILE|[Bezug]
ROWS|ZEILEN|Matrix
RRI|ZSATZINVEST|Zzr; Bw; Zw
RSQ|BESTIMMTHEITSMASS|Y_Werte; X_Werte
SCAN|SCAN|[Anfangswert]; Matrix; Lambda
SEARCH|SUCHEN|Suchtext; Text; [Erstes_Zeichen]
SEC|SEC|Zahl
SECH|SECHYP|Zahl
SECOND|SEKUNDE|Zahl
SEQUENCE|SEQUENZ|Zeilen; [Spalten]; [Anfang]; [Schritt]
SERIESSUM|POTENZREIHE|x; n; m; Koeffizienten
SIGN|VORZEICHEN|Zahl
SIN|SIN|Zahl
SINH|SINHYP|Zahl
SKEW|SCHIEFE|Zahl1; [Zahl2]; ...
SKEW.P|SCHIEFE.P|Zahl1; [Zahl2]; ...
SLN|LIA|Ansch_Wert; Restwert; Nutzungsdauer
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
STDEVPA|STABWNA|Wert1; [Wert2]; ...
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
SYD|DIA|Ansch_Wert; Restwert; Nutzungsdauer; Zr
T|T|Wert
T.DIST|T.VERT|x; Freiheitsgrade; Kumuliert
T.DIST.2T|T.VERT.2S|x; Freiheitsgrade
T.DIST.RT|T.VERT.RE|x; Freiheitsgrade
T.INV|T.INV|Wahrsch; Freiheitsgrade
T.INV.2T|T.INV.2S|Wahrsch; Freiheitsgrade
T.TEST|T.TEST|Matrix1; Matrix2; Seiten; Typ
TAKE|ÜBERNEHMEN|Matrix; Zeilen; [Spalten]
TAN|TAN|Zahl
TANH|TANHYP|Zahl
TBILLEQ|TBILLÄQUIV|Abrechnung; Fälligkeit; Abzinsungssatz
TBILLPRICE|TBILLKURS|Abrechnung; Fälligkeit; Abzinsungssatz
TBILLYIELD|TBILLRENDITE|Abrechnung; Fälligkeit; Kurs
TDIST|TVERT|x; Freiheitsgrade; Seiten
TEXT|TEXT|Wert; Textformat
TEXTAFTER|TEXTNACH|Text; Trennzeichen; [Instanznummer]; [Vergleichsmodus]; [Übereinstimmungsende]; [Wenn_nicht_gefunden]
TEXTBEFORE|TEXTVOR|Text; Trennzeichen; [Instanznummer]; [Vergleichsmodus]; [Übereinstimmungsende]; [Wenn_nicht_gefunden]
TEXTJOIN|TEXTVERKETTEN|Trennzeichen; Leer_ignorieren; Text1; [Text2]; ...
TEXTSPLIT|TEXTTEILEN|Text; Spaltentrennzeichen; [Zeilentrennzeichen]; [Leer_ignorieren]; [Vergleichsmodus]; [Auffüllen_mit]
TIME|ZEIT|Stunde; Minute; Sekunde
TIMEVALUE|ZEITWERT|Zeit
TINV|TINV|Wahrsch; Freiheitsgrade
TOCOL|ZUSPALTE|Matrix; [ignorieren]; [nach_Spalte_scannen]
TODAY|HEUTE|
TOROW|ZUZEILE|Matrix; [ignorieren]; [nach_Spalte_scannen]
TRANSPOSE|MTRANS|Matrix
TREND|TREND|Y_Werte; [X_Werte]; [Neue_x_Werte]; [Konstante]
TRIM|GLÄTTEN|Text
TRIMMEAN|GESTUTZTMITTEL|Matrix; Prozent
TRIMRANGE|ABSCHNBEREICH|Bereich; [Zeilen_abschneiden]; [Spalten_abschneiden]
TRUE|WAHR|
TRUNC|KÜRZEN|Zahl; [Anzahl_Stellen]
TTEST|TTEST|Matrix1; Matrix2; Seiten; Typ
TYPE|TYP|Wert
UNICHAR|UNIZEICHEN|Zahl
UNICODE|UNICODE|Text
UNIQUE|EINDEUTIG|Matrix; [nach_Spalte]; [genau_einmal]
UPPER|GROSS|Text
VALUE|WERT|Text
VALUETOTEXT|WERTZUTEXT|Wert; [Format]
VAR|VARIANZ|Zahl1; [Zahl2]; ...
VAR.P|VAR.P|Zahl1; [Zahl2]; ...
VAR.S|VAR.S|Zahl1; [Zahl2]; ...
VARA|VARIANZA|Wert1; [Wert2]; ...
VARP|VARIANZEN|Zahl1; [Zahl2]; ...
VARPA|VARIANZENA|Wert1; [Wert2]; ...
VDB|VDB|Ansch_Wert; Restwert; Nutzungsdauer; Anfang; Ende; [Faktor]; [Nicht_wechseln]
VLOOKUP|SVERWEIS|Suchkriterium; Matrix; Spaltenindex; [Bereich_Verweis]
VSTACK|VSTAPELN|Matrix1; [Matrix2]; ...
WEEKDAY|WOCHENTAG|Zahl; [Typ]
WEEKNUM|KALENDERWOCHE|Fortlaufende_Zahl; [Zahl_Typ]
WEIBULL|WEIBULL|x; Alpha; Beta; Kumuliert
WEIBULL.DIST|WEIBULL.VERT|x; Alpha; Beta; Kumuliert
WORKDAY|ARBEITSTAG|Ausgangsdatum; Tage; [Freie_Tage]
WORKDAY.INTL|ARBEITSTAG.INTL|Ausgangsdatum; Tage; [Wochenende]; [Freie_Tage]
WRAPCOLS|SPALTENUMBRUCH|Vektor; Umbruch_Anzahl; [auffüllen_mit]
WRAPROWS|ZEILENUMBRUCH|Vektor; Umbruch_Anzahl; [auffüllen_mit]
XIRR|XINTZINSFUSS|Werte; Zeitpkte; [Schätzwert]
XLOOKUP|XVERWEIS|Suchkriterium; Suchmatrix; Rückgabematrix; [Wenn_nicht_gefunden]; [Vergleichsmodus]; [Suchmodus]
XMATCH|XVERGLEICH|Suchkriterium; Suchmatrix; [Vergleichsmodus]; [Suchmodus]
XNPV|XKAPITALWERT|Zins; Werte; Zeitpkte
XOR|XODER|Wahrheitswert1; [Wahrheitswert2]; ...
YEAR|JAHR|Zahl
YEARFRAC|BRTEILJAHRE|Ausgangsdatum; Enddatum; [Basis]
YIELD|RENDITE|Abrechnung; Fälligkeit; Zins; Kurs; Rückzahlung; Häufigkeit; [Basis]
YIELDDISC|RENDITEDIS|Abrechnung; Fälligkeit; Kurs; Rückzahlung; [Basis]
YIELDMAT|RENDITEFÄLL|Abrechnung; Fälligkeit; Emission; Zins; Kurs; [Basis]
Z.TEST|G.TEST|Matrix; x; [Sigma]
ZTEST|GTEST|Matrix; x; [Sigma]`;

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
  "BETA.DIST": "_xlfn.BETA.DIST",
  "BETA.INV": "_xlfn.BETA.INV",
  "BINOM.DIST": "_xlfn.BINOM.DIST",
  "BINOM.DIST.RANGE": "_xlfn.BINOM.DIST.RANGE",
  "BINOM.INV": "_xlfn.BINOM.INV",
  "BITAND": "_xlfn.BITAND",
  "BITLSHIFT": "_xlfn.BITLSHIFT",
  "BITOR": "_xlfn.BITOR",
  "BITRSHIFT": "_xlfn.BITRSHIFT",
  "BITXOR": "_xlfn.BITXOR",
  "BYCOL": "_xlfn.BYCOL",
  "BYROW": "_xlfn.BYROW",
  "CEILING.MATH": "_xlfn.CEILING.MATH",
  "CHISQ.DIST": "_xlfn.CHISQ.DIST",
  "CHISQ.DIST.RT": "_xlfn.CHISQ.DIST.RT",
  "CHISQ.INV": "_xlfn.CHISQ.INV",
  "CHISQ.INV.RT": "_xlfn.CHISQ.INV.RT",
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
  "ENCODEURL": "_xlfn.ENCODEURL",
  "ERF.PRECISE": "_xlfn.ERF.PRECISE",
  "ERFC.PRECISE": "_xlfn.ERFC.PRECISE",
  "EXPAND": "_xlfn.EXPAND",
  "EXPON.DIST": "_xlfn.EXPON.DIST",
  "F.DIST": "_xlfn.F.DIST",
  "F.DIST.RT": "_xlfn.F.DIST.RT",
  "F.INV": "_xlfn.F.INV",
  "F.INV.RT": "_xlfn.F.INV.RT",
  "F.TEST": "_xlfn.F.TEST",
  "FILTER": "_xlfn._xlws.FILTER",
  "FLOOR.MATH": "_xlfn.FLOOR.MATH",
  "FORECAST.LINEAR": "_xlfn.FORECAST.LINEAR",
  "FORMULATEXT": "_xlfn.FORMULATEXT",
  "GAMMA": "_xlfn.GAMMA",
  "GAMMA.DIST": "_xlfn.GAMMA.DIST",
  "GAMMA.INV": "_xlfn.GAMMA.INV",
  "GAMMALN.PRECISE": "_xlfn.GAMMALN.PRECISE",
  "GAUSS": "_xlfn.GAUSS",
  "GROUPBY": "_xlfn.GROUPBY",
  "HSTACK": "_xlfn.HSTACK",
  "HYPGEOM.DIST": "_xlfn.HYPGEOM.DIST",
  "IFNA": "_xlfn.IFNA",
  "IFS": "_xlfn.IFS",
  "IMAGE": "_xlfn.IMAGE",
  "IMCOSH": "_xlfn.IMCOSH",
  "IMCOT": "_xlfn.IMCOT",
  "IMCSC": "_xlfn.IMCSC",
  "IMCSCH": "_xlfn.IMCSCH",
  "IMSEC": "_xlfn.IMSEC",
  "IMSECH": "_xlfn.IMSECH",
  "IMSINH": "_xlfn.IMSINH",
  "IMTAN": "_xlfn.IMTAN",
  "ISFORMULA": "_xlfn.ISFORMULA",
  "ISOMITTED": "_xlfn.ISOMITTED",
  "ISOWEEKNUM": "_xlfn.ISOWEEKNUM",
  "LAMBDA": "_xlfn.LAMBDA",
  "LET": "_xlfn.LET",
  "LOGNORM.DIST": "_xlfn.LOGNORM.DIST",
  "LOGNORM.INV": "_xlfn.LOGNORM.INV",
  "MAKEARRAY": "_xlfn.MAKEARRAY",
  "MAP": "_xlfn.MAP",
  "MAXIFS": "_xlfn.MAXIFS",
  "MINIFS": "_xlfn.MINIFS",
  "MODE.MULT": "_xlfn.MODE.MULT",
  "MODE.SNGL": "_xlfn.MODE.SNGL",
  "MUNIT": "_xlfn.MUNIT",
  "NEGBINOM.DIST": "_xlfn.NEGBINOM.DIST",
  "NETWORKDAYS.INTL": "_xlfn.NETWORKDAYS.INTL",
  "NORM.DIST": "_xlfn.NORM.DIST",
  "NORM.INV": "_xlfn.NORM.INV",
  "NORM.S.DIST": "_xlfn.NORM.S.DIST",
  "NORM.S.INV": "_xlfn.NORM.S.INV",
  "NUMBERVALUE": "_xlfn.NUMBERVALUE",
  "PDURATION": "_xlfn.PDURATION",
  "PERCENTILE.EXC": "_xlfn.PERCENTILE.EXC",
  "PERCENTILE.INC": "_xlfn.PERCENTILE.INC",
  "PERCENTRANK.EXC": "_xlfn.PERCENTRANK.EXC",
  "PERCENTRANK.INC": "_xlfn.PERCENTRANK.INC",
  "PERMUTATIONA": "_xlfn.PERMUTATIONA",
  "PHI": "_xlfn.PHI",
  "PIVOTBY": "_xlfn.PIVOTBY",
  "POISSON.DIST": "_xlfn.POISSON.DIST",
  "QUARTILE.EXC": "_xlfn.QUARTILE.EXC",
  "QUARTILE.INC": "_xlfn.QUARTILE.INC",
  "RANDARRAY": "_xlfn.RANDARRAY",
  "RANK.AVG": "_xlfn.RANK.AVG",
  "RANK.EQ": "_xlfn.RANK.EQ",
  "REDUCE": "_xlfn.REDUCE",
  "REGEXEXTRACT": "_xlfn.REGEXEXTRACT",
  "REGEXREPLACE": "_xlfn.REGEXREPLACE",
  "REGEXTEST": "_xlfn.REGEXTEST",
  "RRI": "_xlfn.RRI",
  "SCAN": "_xlfn.SCAN",
  "SEC": "_xlfn.SEC",
  "SECH": "_xlfn.SECH",
  "SEQUENCE": "_xlfn.SEQUENCE",
  "SHEET": "_xlfn.SHEET",
  "SHEETS": "_xlfn.SHEETS",
  "SKEW.P": "_xlfn.SKEW.P",
  "SORT": "_xlfn._xlws.SORT",
  "SORTBY": "_xlfn.SORTBY",
  "STDEV.P": "_xlfn.STDEV.P",
  "STDEV.S": "_xlfn.STDEV.S",
  "SWITCH": "_xlfn.SWITCH",
  "T.DIST": "_xlfn.T.DIST",
  "T.DIST.2T": "_xlfn.T.DIST.2T",
  "T.DIST.RT": "_xlfn.T.DIST.RT",
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
  "TRIMRANGE": "_xlfn.TRIMRANGE",
  "UNICHAR": "_xlfn.UNICHAR",
  "UNICODE": "_xlfn.UNICODE",
  "UNIQUE": "_xlfn.UNIQUE",
  "VALUETOTEXT": "_xlfn.VALUETOTEXT",
  "VAR.P": "_xlfn.VAR.P",
  "VAR.S": "_xlfn.VAR.S",
  "VSTACK": "_xlfn.VSTACK",
  "WEIBULL.DIST": "_xlfn.WEIBULL.DIST",
  "WORKDAY.INTL": "_xlfn.WORKDAY.INTL",
  "WRAPCOLS": "_xlfn.WRAPCOLS",
  "WRAPROWS": "_xlfn.WRAPROWS",
  "XLOOKUP": "_xlfn.XLOOKUP",
  "XMATCH": "_xlfn.XMATCH",
  "XOR": "_xlfn.XOR",
  "Z.TEST": "_xlfn.Z.TEST",
};
