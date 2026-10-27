// Number formats as German Excel shows them. Codes are kept in the file's (English) form, as
// Excel writes them into styles.xml ("#,##0.00", "dd/mm/yyyy", "General"); the display follows
// a German installation of Excel 365: decimal comma, "." for thousands, "." as the date
// separator, German month and day names, WAHR/FALSCH and German error names.
// Measured against Excel itself (tools/feel-lab/probe_excel_numfmt.py, the tests replay it).

import { CellError, ERROR_TEXT_DE, type Value } from "./model";
import {
  DAYS_DE, DAYS_EN, DAYS_SHORT_DE, MONTHS_DE, MONTHS_EN, MONTHS_SHORT_DE, maxSerial, serialToDate,
} from "./dates";

export interface FormatOptions {
  /** Width of the cell's text area in pixels (the column's width); with `measure`, General
   * numbers are shortened to fit and other numbers and dates turn into "####" as in Excel. */
  widthPx?: number;
  /** Width in pixels of a string in the cell's font. */
  measure?: (s: string) => number;
  /** Margin Excel keeps inside a cell; default 5/7 of a digit's width (5 px for Calibri 11 at 96 dpi). */
  padPx?: number;
  date1904?: boolean;
}

export interface Formatted {
  text: string;
  /** Colour of the chosen section ([Rot], [Farbe10], …) as "#RRGGBB". */
  color?: string;
  /** Excel's default alignment of the value (numbers right, text left, booleans and errors centred). */
  align?: "left" | "right" | "center";
  /** "*x": the character repeated to fill the cell, inserted at `fillAt` in `text`. */
  fill?: string;
  fillAt?: number;
  /** The value does not fit: `text` is the row of "#" Excel shows instead. */
  hashes?: boolean;
}

// ---------------------------------------------------------------------------------------------
// Built-in formats

/** numFmtId → format code as German Excel uses it: what it reads for an id without a <numFmt>
 * (ids 0-81, recorded) and, for the currency and accounting ids it writes with an explicit code
 * (5-8, 41-44), that code. Ids 14-22 and 37-40 are shown the German way (see LOCALE_DISPLAY). */
export const BUILTIN_FORMATS: Record<number, string> = {
  0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00",
  5: '#,##0\\ "€";\\-#,##0\\ "€"', 6: '#,##0\\ "€";[Red]\\-#,##0\\ "€"',
  7: '#,##0.00\\ "€";\\-#,##0.00\\ "€"', 8: '#,##0.00\\ "€";[Red]\\-#,##0.00\\ "€"',
  9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??",
  14: "m/d/yyyy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy", 18: "h:mm AM/PM", 19: "h:mm:ss AM/PM",
  20: "h:mm", 21: "h:mm:ss", 22: "m/d/yyyy h:mm",
  // 23-36 and 50-81 are East Asian ids; German Excel reads them as General or its short date / time.
  23: "General", 24: "General", 25: "General", 26: "General",
  27: "m/d/yyyy", 28: "m/d/yyyy", 29: "m/d/yyyy", 30: "m/d/yyyy", 31: "m/d/yyyy",
  32: "h:mm:ss", 33: "h:mm:ss", 34: "h:mm:ss", 35: "h:mm:ss", 36: "m/d/yyyy",
  37: "#,##0_);(#,##0)", 38: "#,##0_);[Red](#,##0)", 39: "#,##0.00_);(#,##0.00)", 40: "#,##0.00_);[Red](#,##0.00)",
  41: '_-* #,##0_-;\\-* #,##0_-;_-* "-"_-;_-@_-',
  42: '_-* #,##0\\ "€"_-;\\-* #,##0\\ "€"_-;_-* "-"\\ "€"_-;_-@_-',
  43: '_-* #,##0.00_-;\\-* #,##0.00_-;_-* "-"??_-;_-@_-',
  44: '_-* #,##0.00\\ "€"_-;\\-* #,##0.00\\ "€"_-;_-* "-"??\\ "€"_-;_-@_-',
  45: "mm:ss", 46: "[h]:mm:ss", 47: "mm:ss.0", 48: "##0.0E+0", 49: "@",
  50: "m/d/yyyy", 51: "m/d/yyyy", 52: "m/d/yyyy", 53: "m/d/yyyy", 54: "m/d/yyyy", 55: "m/d/yyyy",
  56: "m/d/yyyy", 57: "m/d/yyyy", 58: "m/d/yyyy", 59: "0", 60: "0.00", 61: "#,##0", 62: "#,##0.00",
  63: '#,##0\\ "€";\\-#,##0\\ "€"', 64: '#,##0\\ "€";[Red]\\-#,##0\\ "€"',
  65: '#,##0.00\\ "€";\\-#,##0.00\\ "€"', 66: '#,##0.00\\ "€";[Red]\\-#,##0.00\\ "€"',
  67: "0%", 68: "0.00%", 69: "# ?/?", 70: "# ??/??", 71: "m/d/yyyy", 72: "m/d/yyyy", 73: "d-mmm-yy",
  74: "d-mmm", 75: "mmm-yy", 76: "h:mm", 77: "h:mm:ss", 78: "m/d/yyyy h:mm", 79: "mm:ss", 80: "[h]:mm:ss",
  81: "mm:ss.0",
};

/** The id Excel writes for a code (ids 0-22 and 37-49; 5-8 and 41-44 with the code itself).
 * Excel maps the East Asian ids to these when it saves. undefined: a custom format. */
export function builtinId(code: string): number | undefined {
  return BUILTIN_IDS.get(code);
}

const BUILTIN_IDS = new Map<string, number>();
for (let i = 49; i >= 0; i--) if (i < 23 || i > 36) BUILTIN_IDS.set(BUILTIN_FORMATS[i], i);

/** Built-in codes that depend on the installation's locale, as German Excel shows them. Excel
 * does this for the code itself too: a cell set to "h:mm" shows "08:16" ("hh:mm"). */
export const LOCALE_DISPLAY: Record<string, string> = {
  "m/d/yyyy": "dd/mm/yyyy",
  "d-mmm-yy": "dd/\\ mmm\\ yy",
  "d-mmm": "dd/\\ mmm",
  "mmm-yy": "mmm\\ yy",
  "h:mm": "hh:mm",
  "h:mm:ss": "hh:mm:ss",
  "m/d/yyyy h:mm": "dd/mm/yyyy\\ hh:mm",
  "#,##0_);(#,##0)": "#,##0;\\-#,##0",
  "#,##0_);[Red](#,##0)": "#,##0;[Red]\\-#,##0",
  "#,##0.00_);(#,##0.00)": "#,##0.00;\\-#,##0.00",
  "#,##0.00_);[Red](#,##0.00)": "#,##0.00;[Red]\\-#,##0.00",
  "$#,##0_);($#,##0)": BUILTIN_FORMATS[5],
  "$#,##0_);[Red]($#,##0)": BUILTIN_FORMATS[6],
  "$#,##0.00_);($#,##0.00)": BUILTIN_FORMATS[7],
  "$#,##0.00_);[Red]($#,##0.00)": BUILTIN_FORMATS[8],
  '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)': BUILTIN_FORMATS[41],
  '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)': BUILTIN_FORMATS[42],
  '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)': BUILTIN_FORMATS[43],
  '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)': BUILTIN_FORMATS[44],
};

/** Codes that differ only in escaped spaces and minus signs are the same format. */
function canonical(code: string): string {
  return code.replace(/\\([ -])/g, "$1");
}

// German codes Excel recognises as built-in formats when they are typed (fileFormat).
const TYPED_BUILTIN = new Map<string, string>();
for (let i = 49; i >= 0; i--) if (i < 23 || i > 36) TYPED_BUILTIN.set(canonical(BUILTIN_FORMATS[i]), BUILTIN_FORMATS[i]);
for (const [k, v] of Object.entries(LOCALE_DISPLAY)) if (!k.startsWith("$") && !k.startsWith("_(")) TYPED_BUILTIN.set(canonical(v), k);

// ---------------------------------------------------------------------------------------------
// Parsing

type Tok =
  | { t: "lit"; s: string }
  | { t: "pad"; s: string }
  | { t: "fill"; s: string }
  | { t: "text" }
  | { t: "general" }
  | { t: "digit"; c: string }
  | { t: "point" }
  | { t: "comma" }
  | { t: "percent" }
  | { t: "exp"; s: string }
  | { t: "slash" }
  | { t: "denom"; n: string }
  | { t: "date"; k: string; n: number }
  | { t: "elapsed"; k: string; n: number }
  | { t: "sub"; n: number }
  | { t: "ampm"; s: string };

interface Cond { op: string; v: number }

interface Section {
  toks: Tok[];
  color?: string;
  cond?: Cond;
  lcid?: number;
  kind: "general" | "number" | "date" | "text" | "literal";
  num?: NumPlan;
}

interface NumPlan {
  intIdx: number[];
  decIdx: number[];
  expIdx: number[];
  pointAt: number;
  expAt: number;
  grouping: boolean;
  /** Literals between the integer placeholders ("0 0", "000-0000"). */
  intLits: boolean;
  scale: number;
  percent: number;
  frac?: { intIdx: number[]; numIdx: number[]; denIdx: number[]; fixed?: number; slashAt: number };
}

interface Parsed {
  sections: Section[];
  /** Numeric sections (the first up to three; a trailing section with "@" is the text section). */
  nums: Section[];
  text?: Section;
  hasCond: boolean;
}

const COLORS: Record<string, string> = {
  black: "#000000", blue: "#0000FF", cyan: "#00FFFF", green: "#00FF00", magenta: "#FF00FF", red: "#FF0000",
  white: "#FFFFFF", yellow: "#FFFF00", schwarz: "#000000", blau: "#0000FF", zyan: "#00FFFF", "grün": "#00FF00",
  rot: "#FF0000", "weiß": "#FFFFFF", gelb: "#FFFF00",
};

/** Excel's 56-colour palette for [ColorN] / [FarbeN]. */
export const PALETTE = [
  "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF", "#800000", "#008000",
  "#000080", "#808000", "#800080", "#008080", "#C0C0C0", "#808080", "#9999FF", "#993366", "#FFFFCC", "#CCFFFF",
  "#660066", "#FF8080", "#0066CC", "#CCCCFF", "#000080", "#FF00FF", "#FFFF00", "#00FFFF", "#800080", "#800000",
  "#008080", "#0000FF", "#00CCFF", "#CCFFFF", "#CCFFCC", "#FFFF99", "#99CCFF", "#FF99CC", "#CC99FF", "#FFCC99",
  "#3366FF", "#33CCCC", "#99CC00", "#FFCC00", "#FF9900", "#FF6600", "#666699", "#969696", "#003366", "#339966",
  "#003300", "#333300", "#993300", "#993366", "#333399", "#333333",
];

/** Splits a code into its sections at ";" outside quotes, brackets and escapes. */
function splitSections(code: string): string[] {
  const out: string[] = [];
  let cur = "";
  for (let i = 0; i < code.length; i++) {
    const c = code[i];
    if (c === '"') {
      const j = code.indexOf('"', i + 1);
      const end = j < 0 ? code.length : j + 1;
      cur += code.slice(i, end);
      i = end - 1;
    } else if (c === "\\" || c === "_" || c === "*") {
      cur += code.slice(i, i + 2);
      i++;
    } else if (c === "[") {
      const j = code.indexOf("]", i);
      const end = j < 0 ? code.length : j + 1;
      cur += code.slice(i, end);
      i = end - 1;
    } else if (c === ";") {
      out.push(cur);
      cur = "";
    } else cur += c;
  }
  out.push(cur);
  return out;
}

const SECTION_CACHE = new Map<string, Parsed>();

function parse(code: string): Parsed {
  let p = SECTION_CACHE.get(code);
  if (!p) {
    p = parseUncached(code);
    if (SECTION_CACHE.size > 2000) SECTION_CACHE.clear();
    SECTION_CACHE.set(code, p);
  }
  return p;
}

function parseUncached(code: string): Parsed {
  const sections = splitSections(code).map(parseSection);
  let text: Section | undefined;
  let nums = sections.slice(0, 3);
  if (sections.length >= 4) text = sections[3];
  else {
    const last = sections[sections.length - 1];
    if (last.toks.some((t) => t.t === "text")) {
      text = last;
      nums = sections.slice(0, sections.length - 1);
    }
  }
  return { sections, nums, text, hasCond: nums.some((s) => s.cond) };
}

function parseSection(src: string): Section {
  const sec: Section = { toks: [], kind: "literal" };
  const toks = sec.toks;
  const lit = (s: string) => toks.push({ t: "lit", s });
  let digitsSeen = false;
  for (let i = 0; i < src.length;) {
    const c = src[i];
    const rest = src.slice(i);
    if (c === '"') {
      const j = src.indexOf('"', i + 1);
      lit(j < 0 ? src.slice(i + 1) : src.slice(i + 1, j));
      i = j < 0 ? src.length : j + 1;
    } else if (c === "\\") {
      if (i + 1 < src.length) lit(src[i + 1]);
      i += 2;
    } else if (c === "_") {
      toks.push({ t: "pad", s: src[i + 1] ?? " " });
      i += 2;
    } else if (c === "*") {
      toks.push({ t: "fill", s: src[i + 1] ?? " " });
      i += 2;
    } else if (c === "[") {
      const j = src.indexOf("]", i);
      const b = src.slice(i + 1, j < 0 ? src.length : j);
      i = j < 0 ? src.length : j + 1;
      bracket(sec, b);
    } else if (/^general/i.test(rest)) {
      toks.push({ t: "general" });
      i += 7;
    } else if (c === "@") {
      toks.push({ t: "text" });
      i++;
    } else if (c === "0" || c === "#" || c === "?") {
      toks.push({ t: "digit", c });
      digitsSeen = true;
      i++;
    } else if (c === ".") {
      toks.push({ t: "point" });
      i++;
    } else if (c === ",") {
      toks.push({ t: "comma" });
      i++;
    } else if (c === "%") {
      toks.push({ t: "percent" });
      i++;
    } else if ((c === "E" || c === "e") && (src[i + 1] === "+" || src[i + 1] === "-") && digitsSeen) {
      toks.push({ t: "exp", s: c + src[i + 1] });
      i += 2;
    } else if (c === "/") {
      const m = /^\/([1-9][0-9]*)/.exec(rest);
      if (m) {
        toks.push({ t: "slash" }, { t: "denom", n: m[1] });
        i += m[0].length;
      } else {
        toks.push({ t: "slash" });
        i++;
      }
    } else if (/^am\/pm/i.test(rest)) {
      toks.push({ t: "ampm", s: "AM/PM" });
      i += 5;
    } else if (/^a\/p/i.test(rest)) {
      toks.push({ t: "ampm", s: rest.slice(0, 3) });
      i += 3;
    } else if (/[yYmMdDhHsSeE]/.test(c)) {
      const k = c.toLowerCase();
      let n = 1;
      while (i + n < src.length && src[i + n].toLowerCase() === k) n++;
      if (k === "e" && /[eE]/.test(c)) toks.push({ t: "date", k: "y", n: 4 });
      else toks.push({ t: "date", k, n });
      i += n;
    } else {
      // An unquoted "$" is the installation's currency symbol: "€" in German Excel.
      lit(c === "$" ? "€" : c);
      i++;
    }
  }
  classify(sec);
  return sec;
}

function bracket(sec: Section, b: string): void {
  const low = b.toLowerCase().trim();
  if (/^(h+|m+|s+)$/.test(low)) {
    sec.toks.push({ t: "elapsed", k: low[0], n: low.length });
    return;
  }
  const cm = /^(<=|>=|<>|<|>|=)\s*(-?[0-9.]+(?:[eE][+-]?[0-9]+)?)$/.exec(b.trim());
  if (cm) {
    sec.cond = { op: cm[1], v: Number(cm[2]) };
    return;
  }
  if (COLORS[low]) {
    sec.color = COLORS[low];
    return;
  }
  const pm = /^(?:color|farbe)\s*(\d+)$/.exec(low);
  if (pm) {
    sec.color = PALETTE[Number(pm[1]) - 1] ?? "#000000";
    return;
  }
  if (b.startsWith("$")) {
    const m = /^\$([^-]*)(?:-([0-9A-Fa-f]+))?/.exec(b);
    if (m) {
      if (m[1]) sec.toks.push({ t: "lit", s: m[1] });
      if (m[2]) sec.lcid = parseInt(m[2].slice(-8), 16);
    }
  }
  // Other brackets ([DBNum1], calendars) do not change the display here.
}

const LONG_DATE = "dddd, d. mmmm yyyy";
const LONG_TIME = "hh:mm:ss";

function classify(sec: Section): void {
  if (sec.lcid !== undefined && (sec.lcid & 0xffff) === 0xf800) {
    sec.toks = parseSection(LONG_DATE).toks;
  } else if (sec.lcid !== undefined && (sec.lcid & 0xffff) === 0xf400) {
    sec.toks = parseSection(LONG_TIME).toks;
  }
  const toks = sec.toks;
  const isDate = toks.some((t) => t.t === "date" || t.t === "elapsed" || t.t === "ampm");
  if (isDate) {
    sec.kind = "date";
    sec.toks = dateTokens(toks);
    return;
  }
  // Letters that are not codes in a number section stay literal.
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.t === "date") toks[i] = { t: "lit", s: t.k.repeat(t.n) };
  }
  if (toks.some((t) => t.t === "general")) sec.kind = "general";
  else if (toks.some((t) => t.t === "digit")) {
    sec.kind = "number";
    sec.num = numPlan(toks);
  } else if (toks.some((t) => t.t === "text")) sec.kind = "text";
  else sec.kind = "literal";
}

/** Date sections: m after h or before s is minutes; "." before 0s after seconds starts the
 * fraction of a second; "/" is the date separator ("." in German); other number signs are text. */
function dateTokens(toks: Tok[]): Tok[] {
  const out: Tok[] = [];
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.t === "point") {
      const prev = lastDate(out);
      let n = 0;
      while (toks[i + 1 + n]?.t === "digit" && (toks[i + 1 + n] as { c: string }).c === "0") n++;
      if (prev && (prev.k === "s") && n > 0) {
        out.push({ t: "sub", n: Math.min(n, 3) });
        i += n;
      } else out.push({ t: "lit", s: "." });
    } else if (t.t === "slash") out.push({ t: "lit", s: "." });
    else if (t.t === "denom") out.push({ t: "lit", s: t.n });
    else if (t.t === "comma") out.push({ t: "lit", s: "," });
    else if (t.t === "percent") out.push({ t: "lit", s: "%" });
    else if (t.t === "digit") out.push({ t: "lit", s: t.c });
    else if (t.t === "exp") out.push({ t: "lit", s: t.s });
    else if (t.t === "general" || t.t === "text") continue;
    else out.push(t);
  }
  // Month or minute: a minute follows an hour or precedes a second.
  for (let i = 0; i < out.length; i++) {
    const t = out[i];
    if (t.t !== "date" || t.k !== "m" || t.n > 2) continue;
    const prev = lastDate(out.slice(0, i));
    let next: { k: string } | undefined;
    for (let j = i + 1; j < out.length; j++) {
      const u = out[j];
      if (u.t === "date" || u.t === "elapsed") { next = u; break; }
    }
    if ((prev && prev.k === "h") || (next && next.k === "s")) out[i] = { t: "date", k: "M", n: t.n };
  }
  return out;
}

function lastDate(toks: Tok[]): { k: string } | undefined {
  for (let i = toks.length - 1; i >= 0; i--) {
    const t = toks[i];
    if (t.t === "date") return { k: t.k === "M" ? "m" : t.k };
    if (t.t === "elapsed") return { k: t.k };
  }
  return undefined;
}

function numPlan(toks: Tok[]): NumPlan {
  const plan: NumPlan = { intIdx: [], decIdx: [], expIdx: [], pointAt: -1, expAt: -1, grouping: false, intLits: false, scale: 0, percent: 0 };
  const slashAt = toks.findIndex((t) => t.t === "slash");
  plan.expAt = toks.findIndex((t) => t.t === "exp");
  if (slashAt >= 0 && plan.expAt < 0) {
    // Fraction: [integer] [numerator]/[denominator].
    const denIdx: number[] = [];
    let fixed: number | undefined;
    for (let j = slashAt + 1; j < toks.length; j++) {
      const t = toks[j];
      if (t.t === "denom") { fixed = Number(t.n); continue; }
      if (t.t === "digit") denIdx.push(j);
      else if (denIdx.length || fixed !== undefined) break;
      else break;
    }
    const numIdx: number[] = [];
    let j = slashAt - 1;
    while (j >= 0 && toks[j].t === "digit") numIdx.unshift(j--);
    const intIdx: number[] = [];
    for (let k = 0; k <= j; k++) if (toks[k].t === "digit") intIdx.push(k);
    plan.frac = { intIdx, numIdx, denIdx, fixed, slashAt };
    plan.grouping = toks.slice(0, j + 1).some((t, k) => t.t === "comma" && isBetweenDigits(toks, k, j + 1));
    for (const t of toks) if (t.t === "percent") plan.percent++;
    return plan;
  }
  const end = plan.expAt >= 0 ? plan.expAt : toks.length;
  for (let i = 0; i < end; i++) if (toks[i].t === "point") { plan.pointAt = i; break; }
  const intEnd = plan.pointAt >= 0 ? plan.pointAt : end;
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.t === "digit") {
      if (i < intEnd) plan.intIdx.push(i);
      else if (i < end) plan.decIdx.push(i);
      else plan.expIdx.push(i);
    } else if (t.t === "percent") plan.percent++;
    else if (t.t === "comma") {
      if (i < intEnd && isBetweenDigits(toks, i, intEnd)) plan.grouping = true;
      else if (precededByDigit(toks, i)) plan.scale++;
      else toks[i] = { t: "lit", s: "," };
    }
  }
  if (plan.intIdx.length > 1) {
    const a = plan.intIdx[0], b = plan.intIdx[plan.intIdx.length - 1];
    for (let i = a; i < b; i++) {
      const t = toks[i];
      if (t.t === "lit" || t.t === "pad" || t.t === "fill") plan.intLits = true;
    }
  }
  return plan;
}

function isBetweenDigits(toks: Tok[], i: number, end: number): boolean {
  let before = false, after = false;
  for (let k = i - 1; k >= 0; k--) if (toks[k].t === "digit") { before = true; break; }
  for (let k = i + 1; k < end; k++) if (toks[k].t === "digit") { after = true; break; }
  return before && after;
}

function precededByDigit(toks: Tok[], i: number): boolean {
  for (let k = i - 1; k >= 0; k--) {
    const t = toks[k];
    if (t.t === "digit") return true;
    if (t.t !== "comma" && t.t !== "point") return false;
  }
  return false;
}

// ---------------------------------------------------------------------------------------------
// Decimal arithmetic on Excel's 15 significant digits

interface Dec { d: string; p: number }

/** x > 0 rounded to 15 significant digits: x = 0.d × 10^p, d without leading or trailing zeros. */
function decimal(x: number): Dec {
  const s = x.toExponential(14);
  const e = s.indexOf("e");
  const digits = (s[0] + s.slice(2, e)).replace(/0+$/, "");
  return { d: digits || "0", p: Number(s.slice(e + 1)) + 1 };
}

/** Rounds 0.d × 10^p to n decimals (half away from zero): integer digits ("" for zero) and
 * exactly n decimals. */
function roundDec(x: Dec, n: number): { int: string; frac: string } {
  const keep = x.p + n;
  let s = keep > 0 ? x.d.slice(0, keep).padEnd(keep, "0") : "";
  const next = keep >= 0 ? (x.d[keep] ?? "0") : "0";
  if (next >= "5") s = increment(s);
  s = s.replace(/^0+/, "");
  if (n <= 0) return { int: s, frac: "" };
  const padded = s.padStart(n, "0");
  return { int: padded.slice(0, padded.length - n), frac: padded.slice(padded.length - n) };
}

function increment(s: string): string {
  const a = s.split("");
  let i = a.length - 1;
  while (i >= 0) {
    if (a[i] === "9") { a[i] = "0"; i--; }
    else { a[i] = String.fromCharCode(a[i].charCodeAt(0) + 1); return a.join(""); }
  }
  return "1" + a.join("");
}

function round15(x: number): number {
  return x === 0 || !isFinite(x) ? x : Number(x.toPrecision(15));
}

// ---------------------------------------------------------------------------------------------
// General ("Standard")

/** Excel's General display of a number in at most `budget` characters (sign included); the
 * column rules: decimals are dropped to fit, then scientific notation, then "0"; null when
 * nothing fits ("#"). Without a column Excel allows 11 characters besides the sign. */
export function generalNumber(v: number, budget?: number): string | null {
  if (!isFinite(v)) return null;
  const neg = v < 0;
  const sign = neg ? "-" : "";
  const x = Math.abs(v);
  const cap = 11 + sign.length;
  let b = budget === undefined ? cap : Math.min(budget, cap);
  if (x === 0) return "0";
  const dx = decimal(x);
  const e0 = dx.p - 1;
  // Numbers shown in decimals (from 0,0001 on) still show a single digit in the narrowest column.
  if (e0 >= -4) b = Math.max(b, 1);
  const avail = b - sign.length;
  // Decimal candidate: as many decimals as fit, not rounded to zero.
  let dec: string | null = null, posOk = false, exact = false;
  if (avail >= 1) {
    const intLen = Math.max(dx.p, 1);
    const nd = Math.max(avail - intLen - 1, 0);
    const r = roundDec(dx, nd);
    const frac = r.frac.replace(/0+$/, "");
    const s = (r.int || "0") + (frac ? "," + frac : "");
    if (s.length <= avail && /[1-9]/.test(r.int + frac)) {
      dec = s;
      // Shows a digit where the value's first digit is (0,001 for 0,0005 does not).
      posOk = nd >= -e0;
      exact = dx.d.length - dx.p <= nd;
    }
  }
  // Scientific candidate: as many mantissa decimals as fit (at most 5).
  let sci: string | null = null;
  for (let md = Math.min(5, avail - 2); md >= 0; md--) {
    const s = sciString(dx, md);
    if (s.length <= avail) { sci = s; break; }
  }
  if (dec !== null) {
    const num = (t: string) => Number(t.replace(",", "."));
    let useDec: boolean;
    if (sci === null) useDec = true;
    // Below 0,0001 decimals only when they show every digit (0,00001).
    else if (e0 <= -5) useDec = exact;
    else if (!posOk) useDec = false;
    // Otherwise the closer of the two, decimals when they are equal (0,1 rather than 1E-01).
    else useDec = Math.abs(x - num(dec)) <= Math.abs(x - num(sci)) * (1 + 1e-12);
    if (useDec) return sign + dec;
  }
  if (sci !== null) return sign + sci;
  if (e0 < 0 && avail >= 1) return sign + "0";
  return null;
}

/** d,dddddE+xx with `md` mantissa decimals (trailing zeros dropped), Excel's General style. */
function sciString(x: Dec, md: number): string {
  let e = x.p - 1;
  let r = roundDec({ d: x.d, p: 1 }, md);
  if (r.int.length > 1) {
    // 9,99995 rounds up to 10: one more in the exponent.
    e += 1;
    r = { int: "1", frac: "0".repeat(md) };
  }
  const frac = r.frac.replace(/0+$/, "");
  const ae = Math.abs(e);
  return (r.int || "0") + (frac ? "," + frac : "") + "E" + (e < 0 ? "-" : "+") + String(ae).padStart(2, "0");
}

// ---------------------------------------------------------------------------------------------
// Formatting

const BOOL_TEXT = { true: "WAHR", false: "FALSCH" };

/** The text German Excel shows for a value in a cell with the given format code (file form). */
export function formatValue(value: Value, code: string, opts: FormatOptions = {}): Formatted {
  if (value === null || value === undefined) return { text: "" };
  if (value instanceof CellError) return { text: ERROR_TEXT_DE[value.code], align: "center" };
  const p = parse(LOCALE_DISPLAY[code] ?? (code || "General"));
  if (typeof value === "boolean" || typeof value === "string") {
    const s = typeof value === "boolean" ? BOOL_TEXT[`${value}`] : value;
    const align = typeof value === "boolean" ? "center" : "left";
    if (!p.text) return { text: s, align };
    const out = render(p.text, () => s);
    return { ...out, color: p.text.color, align };
  }
  if (typeof value !== "number" || !isFinite(value)) return { text: String(value) };
  const pick = pickSection(p, value);
  if (!pick) return hashes(opts, false);
  const { sec, abs } = pick;
  const color = sec.color;
  const v = abs ? Math.abs(value) : value;
  let out: Out | null;
  let general = false;
  const plainGeneral = p.sections.length === 1 && sec.kind === "general" && sec.toks.length === 1;
  if (plainGeneral) {
    general = true;
    const s = generalNumber(v, columnChars(opts));
    if (s === null) return { ...hashes(opts, true), color };
    out = { text: s };
  } else if (sec.kind === "general") {
    // General with text or in a format of several sections: shortened like General in the
    // room the text leaves; nothing in the narrowest column, "0" where no digit fits.
    general = true;
    const n = columnChars(opts);
    if (n === undefined) {
      const s = generalNumber(v);
      out = s === null ? null : render(sec, () => s);
    } else if (n === 0) return { text: "", align: "right", color };
    else {
      const room = n - render(sec, () => "").text.length;
      const s = room >= 1 ? generalNumber(v, room) : null;
      if (s === null && room >= 1) return { ...hashes(opts, true), color };
      out = s === null || (s === "0" && v !== 0) ? { text: "0" } : render(sec, () => s);
    }
  } else if (sec.kind === "number") out = formatNumber(sec, v);
  // Dates and times of negative numbers are "####" whatever section shows them (1900 system).
  else if (sec.kind === "date") out = formatDate(sec, opts.date1904 ? v : value, opts.date1904 ?? false);
  else if (sec.kind === "text" && p.nums.length === 0) {
    // A text-only format shows numbers as General, negative ones like "General;-General".
    return formatValue(value, "General;-General", opts);
  } else out = render(sec, () => "");
  if (out === null) return { ...hashes(opts, false), color };
  if (!general && opts.widthPx !== undefined && opts.measure && sec.kind !== "literal") {
    const room = opts.widthPx - padding(opts);
    if (opts.measure(out.text) > room) return { ...hashes(opts, false), color };
  }
  const res: Formatted = { text: out.text, align: "right" };
  if (color) res.color = color;
  if (out.fill !== undefined) { res.fill = out.fill; res.fillAt = out.fillAt; }
  return res;
}

interface Out { text: string; fill?: string; fillAt?: number }

function padding(opts: FormatOptions): number {
  if (opts.padPx !== undefined) return opts.padPx;
  const digit = opts.measure ? opts.measure("0") : 7;
  return Math.floor((digit * 5) / 7);
}

/** How many digits fit in the column (Excel's General counts characters of digit width). */
function columnChars(opts: FormatOptions): number | undefined {
  if (opts.widthPx === undefined || !opts.measure) return undefined;
  const digit = opts.measure("0") || 1;
  return Math.max(0, Math.floor((opts.widthPx - padding(opts)) / digit));
}

/** The row of "#" for a value that does not fit: as many as digits fit (General at least one). */
function hashes(opts: FormatOptions, general: boolean): Formatted {
  const n = columnChars(opts);
  const count = n === undefined ? 8 : general ? Math.max(1, n) : n;
  return { text: "#".repeat(count), hashes: true, align: "right" };
}

function matches(c: Cond, v: number): boolean {
  switch (c.op) {
    case "<": return v < c.v;
    case ">": return v > c.v;
    case "=": return v === c.v;
    case "<=": return v <= c.v;
    case ">=": return v >= c.v;
    case "<>": return v !== c.v;
  }
  return false;
}

function pickSection(p: Parsed, v: number): { sec: Section; abs: boolean } | null {
  const n = p.nums;
  if (n.length === 0) return { sec: p.sections[0], abs: false };
  if (!p.hasCond) {
    if (n.length === 1) return { sec: n[0], abs: false };
    if (n.length === 2) return v < 0 ? { sec: n[1], abs: true } : { sec: n[0], abs: false };
    return v > 0 ? { sec: n[0], abs: false } : v < 0 ? { sec: n[1], abs: true } : { sec: n[2], abs: false };
  }
  const c0: Cond = n[0].cond ?? (n.length >= 3 ? { op: ">", v: 0 } : { op: ">=", v: 0 });
  if (matches(c0, v)) return { sec: n[0], abs: isNegCond(c0) };
  if (n.length === 1) return null;
  const c1: Cond | undefined = n[1].cond ?? (n.length >= 3 ? { op: "<", v: 0 } : undefined);
  if (!c1 || matches(c1, v)) return { sec: n[1], abs: c1 ? isNegCond(c1) : false };
  if (n.length >= 3) return { sec: n[2], abs: false };
  return null;
}

/** A section chosen for "less than 0" (or less than a negative number) shows the number without
 * its minus sign, as the standard negative section does. */
function isNegCond(c: Cond): boolean {
  return c.op === "<" && c.v <= 0;
}

/** Literal parts of a section around a value string (text sections, General with text). */
function render(sec: Section, value: () => string): Out {
  let text = "";
  let fill: string | undefined, fillAt: number | undefined;
  for (const t of sec.toks) {
    if (t.t === "lit") text += t.s;
    else if (t.t === "pad") text += " ";
    else if (t.t === "fill") { if (fill === undefined) { fill = t.s; fillAt = text.length; } }
    else if (t.t === "text" || t.t === "general") text += value();
    else if (t.t === "digit") text += t.c === "0" ? "0" : "";
    else if (t.t === "point") text += ",";
    else if (t.t === "percent") text += "%";
    else if (t.t === "slash") text += "/";
  }
  return fill === undefined ? { text } : { text, fill, fillAt };
}

// ---- numbers

function formatNumber(sec: Section, v: number): Out | null {
  const plan = sec.num!;
  const toks = sec.toks;
  const neg = v < 0;
  // Excel takes the value's 15 digits, then scales for "%" and "," in floating point; fractions
  // work on the binary value itself.
  let x = round15(Math.abs(v)), raw = Math.abs(v);
  for (let i = 0; i < plan.percent; i++) { x *= 100; raw *= 100; }
  for (let i = 0; i < plan.scale; i++) { x /= 1000; raw /= 1000; }
  x = round15(x);
  const out = new Map<number, string>();
  let zero: boolean;
  let expText = "";
  if (plan.frac) {
    const r = fraction(plan, toks, raw, out);
    if (r === null) return null;
    zero = r;
  } else if (plan.expAt >= 0) {
    zero = scientific(plan, toks, x, out);
    expText = out.get(-1) ?? "";
  } else {
    const r = x === 0 ? { int: "", frac: "0".repeat(plan.decIdx.length) } : roundDec(decimal(x), plan.decIdx.length);
    zero = r.int === "" && /^0*$/.test(r.frac);
    placeInt(plan.intIdx, toks, r.int, plan.grouping && !plan.intLits, out);
    // Without integer placeholders (".00") the integer digits still show before the comma.
    if (!plan.intIdx.length && plan.pointAt >= 0 && r.int) out.set(plan.pointAt, r.int + ",");
    placeDec(plan.decIdx, toks, r.frac, out);
  }
  let text = neg && !zero ? "-" : "";
  let fill: string | undefined, fillAt: number | undefined;
  let pointDone = false;
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    const o = out.get(i);
    if (o !== undefined) { text += o; continue; }
    switch (t.t) {
      case "lit": text += t.s; break;
      case "pad": text += " "; break;
      case "fill": if (fill === undefined) { fill = t.s; fillAt = text.length; } break;
      case "point": text += pointDone ? "." : ","; pointDone = true; break;
      case "percent": text += "%"; break;
      case "exp": text += expText; break;
      case "slash": text += "/"; break;
      case "denom": text += t.n; break;
      case "general": text += generalNumber(x) ?? ""; break;
      case "text": break;
      default: break;
    }
  }
  return fill === undefined ? { text } : { text, fill, fillAt };
}

/** Integer digits into their placeholders: from the right, the leftmost one takes the rest. */
function placeInt(idx: number[], toks: Tok[], digits: string, grouping: boolean, out: Map<number, string>): void {
  if (idx.length === 0) return;
  const chars = idx.map((i) => (toks[i] as { c: string }).c);
  if (grouping) {
    // Pad to the placeholders' minimum ('0' zeros, '?' spaces), then group the digits by three.
    let s = digits;
    const firstZero = chars.indexOf("0");
    const minZeros = firstZero < 0 ? 0 : chars.length - firstZero;
    if (s.length < minZeros) s = s.padStart(minZeros, "0");
    let grouped = "";
    for (let k = 0; k < s.length; k++) {
      if (k > 0 && (s.length - k) % 3 === 0) grouped += ".";
      grouped += s[k];
    }
    const firstQ = chars.indexOf("?");
    if (firstQ >= 0 && s.length < chars.length - firstQ) grouped = " ".repeat(chars.length - firstQ - s.length) + grouped;
    out.set(idx[0], grouped);
    for (let k = 1; k < idx.length; k++) out.set(idx[k], "");
    return;
  }
  let rest = digits;
  for (let k = idx.length - 1; k >= 0; k--) {
    if (k === 0) {
      out.set(idx[k], rest.length ? rest : chars[k] === "0" ? "0" : chars[k] === "?" ? " " : "");
    } else if (rest.length) {
      out.set(idx[k], rest[rest.length - 1]);
      rest = rest.slice(0, -1);
    } else out.set(idx[k], chars[k] === "0" ? "0" : chars[k] === "?" ? " " : "");
  }
}

function placeDec(idx: number[], toks: Tok[], frac: string, out: Map<number, string>): void {
  let trimming = true;
  for (let k = idx.length - 1; k >= 0; k--) {
    const c = (toks[idx[k]] as { c: string }).c;
    const d = frac[k] ?? "0";
    if (trimming && d === "0" && c !== "0") out.set(idx[k], c === "?" ? " " : "");
    else {
      trimming = false;
      out.set(idx[k], d);
    }
  }
}

function scientific(plan: NumPlan, toks: Tok[], x: number, out: Map<number, string>): boolean {
  const ni = Math.max(plan.intIdx.length, 1);
  const nd = plan.decIdx.length;
  // With more than one integer placeholder ("##0.0E+0", "00.00E+00") the exponent is a multiple
  // of their number (engineering notation).
  const eng = plan.intIdx.length > 1;
  let e = 0;
  // Zero fills every mantissa placeholder with 0 ("000,0E+0" for "##0.0E+0").
  let r = { int: "0".repeat(ni), frac: "0".repeat(nd) };
  if (x > 0) {
    const dx = decimal(x);
    e = dx.p - 1;
    if (eng) e = Math.floor(e / ni) * ni;
    r = roundDec({ d: dx.d, p: dx.p - e }, nd);
    if (r.int.length > (eng ? ni : 1)) {
      e += eng ? ni : 1;
      r = roundDec({ d: dx.d, p: dx.p - e }, nd);
    }
  }
  placeInt(plan.intIdx, toks, r.int, plan.grouping, out);
  placeDec(plan.decIdx, toks, r.frac, out);
  const et = toks[plan.expAt] as { s: string };
  const expDigits = plan.expIdx.map((i) => (toks[i] as { c: string }).c);
  const minDigits = expDigits.filter((c) => c === "0").length;
  const ae = String(Math.abs(e)).padStart(Math.max(minDigits, 1), "0");
  const sign = e < 0 ? "-" : et.s[1] === "+" ? "+" : "";
  out.set(-1, et.s[0] + sign + ae);
  for (const i of plan.expIdx) out.set(i, "");
  return /^0*$/.test(r.int) && /^0*$/.test(r.frac);
}

/** Fractions ("# ?/?", "?/8"): returns whether the minus sign goes (a numerator of 0 without a
 * whole part; with one, "-0      " keeps it); null for "####". */
function fraction(plan: NumPlan, toks: Tok[], x: number, out: Map<number, string>): boolean | null {
  const f = plan.frac!;
  const hasInt = f.intIdx.length > 0;
  // The fraction comes from the binary value; above 1E14 the whole part is Excel's 15 digits
  // (which may already be rounded up).
  let whole = hasInt ? Math.floor(x >= 1e14 ? round15(x) : x) : 0;
  const part = hasInt ? x - Math.floor(x) : x;
  let num: number, den: number;
  if (f.fixed !== undefined) {
    den = f.fixed;
    num = Math.round(round15(part) * den);
    if (hasInt && num >= den) { whole += Math.floor(num / den); num %= den; }
    if (!hasInt && num > 32767) return null;
  } else {
    const maxDen = Math.max(1, 10 ** Math.max(f.denIdx.length, 1) - 1);
    const best = bestFraction(part, maxDen);
    if (!best) return null;
    [num, den] = best;
    if (hasInt && num === den) { whole += 1; num = 0; den = 1; }
  }
  const numChars = f.numIdx.map((i) => (toks[i] as { c: string }).c);
  const denChars = f.denIdx.map((i) => (toks[i] as { c: string }).c);
  const hide = hasInt && num === 0 && !numChars.includes("0");
  // Integer part.
  if (hasInt) {
    const intDigits = whole === 0 ? (hide ? "0" : "") : roundDec(decimal(whole), 0).int;
    placeInt(f.intIdx, toks, intDigits, plan.grouping, out);
  }
  const lastInt = hasInt ? f.intIdx[f.intIdx.length - 1] : -1;
  const endFrac = Math.max(f.slashAt, ...f.denIdx, toks.findIndex((t, i) => i > f.slashAt && t.t === "denom"));
  const qMark = numChars.includes("?") || denChars.includes("?");
  if (hide) {
    // The fraction's place is kept blank for "?", dropped for "#".
    for (let i = lastInt + 1; i <= endFrac; i++) {
      const t = toks[i];
      const w = t.t === "lit" ? t.s.length : t.t === "denom" ? t.n.length : t.t === "pad" ? 1 : t.t === "fill" ? 0 : 1;
      out.set(i, qMark ? " ".repeat(w) : "");
    }
    return false;
  }
  if (hasInt && whole === 0) {
    // "0 1/2" keeps the space; with "#" the integer and, unless "?" pads it, the space go.
    const intChars = f.intIdx.map((i) => (toks[i] as { c: string }).c);
    if (!intChars.includes("0") && !numChars.includes("?")) {
      for (let i = lastInt + 1; i < f.numIdx[0]; i++) out.set(i, "");
    }
  }
  placeAligned(f.numIdx, toks, String(num), out, "right");
  if (f.fixed === undefined) placeAligned(f.denIdx, toks, String(den), out, "left");
  return !hasInt && num === 0;
}

function placeAligned(idx: number[], toks: Tok[], digits: string, out: Map<number, string>, side: "left" | "right"): void {
  if (!idx.length) return;
  const chars = idx.map((i) => (toks[i] as { c: string }).c);
  if (side === "right") {
    placeInt(idx, toks, digits, false, out);
    return;
  }
  // Denominator: digits from the left, padding after them.
  for (let k = 0; k < idx.length; k++) {
    if (k === idx.length - 1) out.set(idx[k], digits.slice(k) || (chars[k] === "0" ? "0" : chars[k] === "?" ? " " : ""));
    else if (k < digits.length) out.set(idx[k], digits[k]);
    else out.set(idx[k], chars[k] === "0" ? "0" : chars[k] === "?" ? " " : "");
  }
}

/** Excel's fraction: the last convergent of the continued fraction of x whose denominator is at
 * most maxDen, computed in floating point as Excel does (its rounding noise included, which is
 * why 0,4375 becomes 4/9 rather than 3/7 with one digit). Numerators stay below 2^31 - 1;
 * null when even the whole number is too large ("####"). */
function bestFraction(x: number, maxDen: number): [number, number] | null {
  let p0 = 0, q0 = 1, p1 = 1, q1 = 0;
  let b = x;
  for (let i = 0; i < 64; i++) {
    const a = Math.floor(b);
    const p = a * p1 + p0, q = a * q1 + q0;
    if (p >= 2147483647) return q1 === 0 ? null : [p1, q1];
    if (q > maxDen) return [p1, q1];
    p0 = p1; q0 = q1; p1 = p; q1 = q;
    if (b - a < 5e-8) return [p, q];
    b = 1 / (b - a);
  }
  return [p1, q1];
}

// ---- dates and times

function formatDate(sec: Section, v: number, date1904: boolean): Out | null {
  const toks = sec.toks;
  const dateToks = toks.filter((t) => t.t === "date" || t.t === "elapsed" || t.t === "sub" || t.t === "ampm");
  const lone = dateToks.length === 1 && dateToks[0].t === "elapsed";
  if (lone) {
    // A single elapsed field ([h], [mm], [ss]) also shows negative and very large values.
    const t = dateToks[0] as { k: string; n: number };
    const per = t.k === "h" ? 3600 : t.k === "m" ? 60 : 1;
    const x = Math.abs(v) * 86400;
    const digits = x < 1e15 ? String(Math.floor(Math.round(round15(x)) / per)) : roundDec(decimal(x / per), 0).int;
    const s = (v < 0 && digits !== "0" ? "-" : "") + digits.padStart(t.n, "0");
    return renderDate(sec, { elapsed: s });
  }
  if (v < 0 && date1904) {
    // The 1904 system shows negative dates: the date of the absolute value with a minus.
    const r = formatDate(sec, -v, true);
    return r && { ...r, text: "-" + r.text, fillAt: r.fillAt === undefined ? undefined : r.fillAt + 1 };
  }
  if (v < 0 || v >= maxSerial(date1904) + 1) return null;
  const subN = Math.max(0, ...toks.map((t) => (t.t === "sub" ? t.n : 0)));
  const unit = 10 ** subN;
  let days = Math.floor(v);
  let ticks = Math.round((v - days) * 86400 * unit + 1e-6);
  if (ticks >= 86400 * unit) { days += 1; ticks -= 86400 * unit; }
  if (days > maxSerial(date1904)) return null;
  const totalTicks = days * 86400 * unit + ticks;
  const secs = Math.floor(ticks / unit);
  const parts = {
    ymd: serialToDate(days, date1904),
    dow: (((days + (date1904 ? 5 : 6)) % 7) + 7) % 7,
    h: Math.floor(secs / 3600),
    mi: Math.floor((secs % 3600) / 60),
    s: secs % 60,
    sub: String(ticks % unit).padStart(subN, "0"),
    total: totalTicks / unit,
  };
  return renderDate(sec, { parts });
}

interface DateParts {
  ymd: { y: number; m: number; d: number };
  dow: number;
  h: number;
  mi: number;
  s: number;
  sub: string;
  total: number;
}

function renderDate(sec: Section, d: { parts?: DateParts; elapsed?: string }): Out {
  const toks = sec.toks;
  const english = sec.lcid !== undefined && (sec.lcid & 0x3ff) === 0x09;
  const ampm = toks.some((t) => t.t === "ampm");
  let text = "";
  let fill: string | undefined, fillAt: number | undefined;
  const p = d.parts;
  for (const t of toks) {
    switch (t.t) {
      case "lit": text += t.s; break;
      case "pad": text += " "; break;
      case "fill": if (fill === undefined) { fill = t.s; fillAt = text.length; } break;
      case "elapsed": {
        if (d.elapsed !== undefined) { text += d.elapsed; break; }
        const unit = t.k === "h" ? 3600 : t.k === "m" ? 60 : 1;
        text += String(Math.floor(p!.total / unit + 1e-9)).padStart(t.n, "0");
        break;
      }
      case "sub": text += "," + p!.sub.slice(0, t.n); break;
      case "ampm": {
        const pm = p!.h >= 12;
        if (t.s === "AM/PM") text += pm ? "PM" : "AM";
        else text += pm ? t.s[2] : t.s[0];
        break;
      }
      case "date": {
        if (!p) break;
        const { y, m, d: day } = p.ymd;
        switch (t.k) {
          case "y": text += t.n <= 2 ? String(y % 100).padStart(2, "0") : String(y); break;
          case "m":
            if (t.n === 1) text += m;
            else if (t.n === 2) text += String(m).padStart(2, "0");
            else if (t.n === 3) text += english ? MONTHS_EN[m - 1].slice(0, 3) : MONTHS_SHORT_DE[m - 1];
            else if (t.n === 5) text += (english ? MONTHS_EN : MONTHS_DE)[m - 1][0];
            else text += (english ? MONTHS_EN : MONTHS_DE)[m - 1];
            break;
          case "d":
            if (t.n === 1) text += day;
            else if (t.n === 2) text += String(day).padStart(2, "0");
            else if (t.n === 3) text += english ? DAYS_EN[p.dow].slice(0, 3) : DAYS_SHORT_DE[p.dow];
            else text += (english ? DAYS_EN : DAYS_DE)[p.dow];
            break;
          case "h": {
            const h = ampm ? (p.h % 12 || 12) : p.h;
            text += t.n >= 2 ? String(h).padStart(2, "0") : h;
            break;
          }
          case "M": text += t.n >= 2 ? String(p.mi).padStart(2, "0") : p.mi; break;
          case "s": text += t.n >= 2 ? String(p.s).padStart(2, "0") : p.s; break;
          default: text += t.k.repeat(t.n);
        }
        break;
      }
      default: break;
    }
  }
  return fill === undefined ? { text } : { text, fill, fillAt };
}

// ---------------------------------------------------------------------------------------------
// Kinds of formats

/** What a format shows a number as: used by typed input and the ribbon's number format box. */
export function formatKind(code: string): "general" | "number" | "percent" | "scientific" | "fraction" | "currency" | "date" | "time" | "datetime" | "text" | "other" {
  const p = parse(LOCALE_DISPLAY[code] ?? code);
  const s = p.nums[0] ?? p.sections[0];
  if (!s) return "general";
  if (s.kind === "general") return "general";
  if (s.kind === "text" || (p.nums.length === 0 && p.text)) return "text";
  if (s.kind === "date") {
    const hasDate = s.toks.some((t) => t.t === "date" && /[ymd]/.test(t.k));
    const hasTime = s.toks.some((t) => (t.t === "date" && /[hMs]/.test(t.k)) || t.t === "elapsed" || t.t === "ampm" || t.t === "sub");
    return hasDate && hasTime ? "datetime" : hasDate ? "date" : "time";
  }
  if (s.kind === "number") {
    const plan = s.num!;
    if (plan.frac) return "fraction";
    if (plan.expAt >= 0) return "scientific";
    if (plan.percent) return "percent";
    if (s.toks.some((t) => t.t === "lit" && /[€$£¥]/.test(t.s))) return "currency";
    return "number";
  }
  return "other";
}

export function isDateFormat(code: string): boolean {
  const k = formatKind(code);
  return k === "date" || k === "time" || k === "datetime";
}

// ---------------------------------------------------------------------------------------------
// German form of a code (Zellen formatieren > Benutzerdefiniert)

const COLOR_DE: Record<string, string> = {
  black: "Schwarz", blue: "Blau", cyan: "Zyan", green: "Grün", magenta: "Magenta", red: "Rot", white: "Weiß", yellow: "Gelb",
};
const COLOR_EN: Record<string, string> = {
  schwarz: "Black", blau: "Blue", zyan: "Cyan", "grün": "Green", magenta: "Magenta", rot: "Red", "weiß": "White", gelb: "Yellow",
  black: "Black", blue: "Blue", cyan: "Cyan", green: "Green", red: "Red", white: "White", yellow: "Yellow",
};

/** The code as German Excel shows it (NumberFormatLocal): "#.##0,00", "TT.MM.JJJJ", "Standard".
 * The built-in currency and accounting formats show "€" as the currency sign; a custom code
 * keeps its quotes ('#.##0,00 "€"'), as in Excel. */
export function localFormat(code: string): string {
  const src = LOCALE_DISPLAY[code] ?? code;
  const builtinCurrency = CURRENCY_BUILTINS.has(src);
  return splitSections(src).map((s) => localSection(s, builtinCurrency)).join(";");
}

const CURRENCY_BUILTINS = new Set([5, 6, 7, 8, 41, 42, 43, 44].map((i) => BUILTIN_FORMATS[i]));

function localSection(src: string, builtinCurrency: boolean): string {
  // Excel puts a colour before a condition.
  const lead = /^(\[[^\]]*\])+/.exec(src);
  if (lead) {
    const parts = lead[0].match(/\[[^\]]*\]/g) ?? [];
    const isColor = (b: string) => /^\[(black|blue|cyan|green|magenta|red|white|yellow|color\s*\d+)\]$/i.test(b);
    src = [...parts.filter(isColor), ...parts.filter((b) => !isColor(b))].join("") + src.slice(lead[0].length);
  }
  const sec = parseSection(src);
  const isDate = sec.kind === "date";
  // Which "m" runs are minutes: follow the parsed tokens in order.
  const minuteRuns: boolean[] = [];
  for (const t of sec.toks) if (t.t === "date" && (t.k === "m" || t.k === "M")) minuteRuns.push(t.k === "M");
  let run = 0;
  let out = "";
  let afterSeconds = false;
  for (let i = 0; i < src.length;) {
    const c = src[i];
    const rest = src.slice(i);
    if (c === '"') {
      const j = src.indexOf('"', i + 1);
      const end = j < 0 ? src.length : j + 1;
      const s = src.slice(i, end);
      out += s === '"€"' && builtinCurrency ? "€" : s;
      i = end;
    } else if (c === "\\") {
      const x = src[i + 1] ?? "";
      out += keepEscapeLocal(x, isDate) ? "\\" + x : x;
      i += 2;
    } else if (c === "_" || c === "*") {
      out += src.slice(i, i + 2);
      i += 2;
    } else if (c === "[") {
      const j = src.indexOf("]", i);
      const b = src.slice(i + 1, j < 0 ? src.length : j);
      i = j < 0 ? src.length : j + 1;
      const low = b.toLowerCase();
      const cm = /^(color|farbe)\s*(\d+)$/.exec(low);
      if (COLOR_DE[low]) out += `[${COLOR_DE[low]}]`;
      else if (cm) out += `[Farbe${cm[2]}]`;
      else if (/^(<=|>=|<>|<|>|=)/.test(b)) out += "[" + b.replace(/\./g, ",") + "]";
      else {
        out += "[" + b + "]";
        afterSeconds = /^s+$/i.test(b);
      }
    } else if (/^general/i.test(rest)) {
      out += "Standard";
      i += 7;
    } else if (isDate && /^am\/pm/i.test(rest)) {
      out += "AM/PM";
      i += 5;
    } else if (isDate && /^a\/p/i.test(rest)) {
      out += rest.slice(0, 3);
      i += 3;
    } else if (isDate && /[yYmMdDhHsSeE]/.test(c)) {
      const k = c.toLowerCase();
      let n = 1;
      while (i + n < src.length && src[i + n].toLowerCase() === k) n++;
      if (k === "y") out += n <= 2 ? "JJ" : "JJJJ";
      else if (k === "d") out += "T".repeat(Math.min(n, 4));
      else if (k === "m") out += minuteRuns[run++] ? "m".repeat(Math.min(n, 2)) : "M".repeat(n === 5 ? 5 : Math.min(n, 4));
      else if (k === "h") out += "h".repeat(Math.min(n, 2));
      else if (k === "s") out += "s".repeat(Math.min(n, 2));
      else out += "e".repeat(n);
      afterSeconds = k === "s";
      i += n;
    } else if (c === "$" || c === "€") {
      // Unquoted "$" is the currency sign ("€"); a literal "€" needs its escape in German.
      out += c === "$" ? "€" : "\\€";
      i++;
    } else if (isDate) {
      if (c === "/") out += ".";
      else if (c === "." && afterSeconds && src[i + 1] === "0") out += ",";
      else out += c;
      if (c !== "." && c !== "0") afterSeconds = false;
      i++;
    } else {
      out += c === "." ? "," : c === "," ? "." : c;
      i++;
    }
  }
  return out;
}

/** Whether an escaped character keeps its backslash in the German form. */
function keepEscapeLocal(x: string, isDate: boolean): boolean {
  if (" -()+:$!^&'~{}<>=".includes(x)) return false;
  if (isDate && ".,/".includes(x)) return false;
  if (/[a-zA-Z]/.test(x)) return /[aAbBeEgGhHjJmMsStTdDyY]/.test(x);
  return true;
}

/** The file form of a code typed in German (Excel's NumberFormatLocal → NumberFormat), written
 * the way Excel writes it into the file ("#.##0,00 €" → '#,##0.00\ "€"', "TT.MM.JJJJ" →
 * "m/d/yyyy", the built-in short date). */
export function fileFormat(local: string): string {
  const code = splitSections(local).map(fileSection).join(";");
  return TYPED_BUILTIN.get(canonical(code)) ?? code;
}

function fileSection(src: string): string {
  const isDate = localIsDate(src);
  // Month or minute for each M/m run, decided like Excel: after h or before s it is minutes.
  const runs: { k: string; i: number; n: number }[] = [];
  if (isDate) {
    for (let i = 0; i < src.length;) {
      const c = src[i];
      if (c === '"') { const j = src.indexOf('"', i + 1); i = j < 0 ? src.length : j + 1; continue; }
      if (c === "\\" || c === "_" || c === "*") { i += 2; continue; }
      if (c === "[") { const j = src.indexOf("]", i); i = j < 0 ? src.length : j + 1; continue; }
      if (/^standard/i.test(src.slice(i))) { i += 8; continue; }
      if (/^am\/pm/i.test(src.slice(i))) { i += 5; continue; }
      if (/^a\/p/i.test(src.slice(i))) { i += 3; continue; }
      const k = c.toLowerCase();
      if ("tjmhsd".includes(k) || c === "e") {
        let n = 1;
        while (i + n < src.length && src[i + n].toLowerCase() === k) n++;
        runs.push({ k: k === "m" ? c : k, i, n });
        i += n;
        continue;
      }
      i++;
    }
  }
  const minuteAt = new Set<number>();
  runs.forEach((r, j) => {
    if (r.k.toLowerCase() !== "m") return;
    const prev = runs[j - 1], next = runs[j + 1];
    if (r.k === "m" || (prev && prev.k === "h") || (next && next.k === "s")) minuteAt.add(r.i);
  });
  let out = "";
  let afterSeconds = false;
  for (let i = 0; i < src.length;) {
    const c = src[i];
    const rest = src.slice(i);
    if (c === '"') {
      const j = src.indexOf('"', i + 1);
      const end = j < 0 ? src.length : j + 1;
      out += src.slice(i, end);
      i = end;
    } else if (c === "\\") {
      out += src.slice(i, i + 2);
      i += 2;
    } else if (c === "_" || c === "*") {
      out += src.slice(i, i + 2);
      i += 2;
    } else if (c === "[") {
      const j = src.indexOf("]", i);
      const b = src.slice(i + 1, j < 0 ? src.length : j);
      i = j < 0 ? src.length : j + 1;
      const low = b.toLowerCase();
      const cm = /^(color|farbe)\s*(\d+)$/.exec(low);
      if (COLOR_EN[low]) out += `[${COLOR_EN[low]}]`;
      else if (cm) out += `[Color${cm[2]}]`;
      else if (/^(<=|>=|<>|<|>|=)/.test(b)) out += "[" + b.replace(/,/g, ".") + "]";
      else if (/^(h+|m+|s+)$/i.test(b)) out += "[" + b.toLowerCase() + "]";
      else out += "[" + b + "]";
    } else if (/^standard/i.test(rest)) {
      out += "General";
      i += 8;
    } else if (isDate && /^am\/pm/i.test(rest)) {
      out += "AM/PM";
      i += 5;
    } else if (isDate && /^a\/p/i.test(rest)) {
      out += rest.slice(0, 3);
      i += 3;
    } else if (isDate && /[tTjJmMhHsSdD]/.test(c)) {
      const k = c.toLowerCase();
      let n = 1;
      while (i + n < src.length && src[i + n].toLowerCase() === k) n++;
      if (k === "t" || k === "d") out += "d".repeat(Math.min(n, 4));
      else if (k === "j") out += n <= 2 ? "yy" : "yyyy";
      else if (k === "m") out += minuteAt.has(i) ? "m".repeat(Math.min(n, 2)) : "m".repeat(Math.min(n, 5));
      else if (k === "h") out += "h".repeat(Math.min(n, 2));
      else out += "s".repeat(Math.min(n, 2));
      afterSeconds = k === "s";
      i += n;
    } else if (isDate) {
      if (c === ".") out += "/";
      else if (c === "," && afterSeconds && src[i + 1] === "0") out += ".";
      else if (c === "," || c === " " || c === "-") out += "\\" + c;
      else if (c === "/") out += "\\/";
      else if (c === "€") out += '"€"';
      else out += c;
      if (c !== "0" && c !== ",") afterSeconds = false;
      i++;
    } else {
      if (c === ",") out += ".";
      else if (c === ".") out += ",";
      else if (c === " " || (c === "-" && !/[eE]$/.test(out))) out += "\\" + c;
      else if (c === "€") out += '"€"';
      else out += c;
      i++;
    }
  }
  return out;
}

/** A German code is a date or time format when it has T, J, M, h, s or AM/PM outside quotes. */
function localIsDate(src: string): boolean {
  for (let i = 0; i < src.length; i++) {
    const c = src[i];
    if (c === '"') { const j = src.indexOf('"', i + 1); i = j < 0 ? src.length : j; continue; }
    if (c === "\\" || c === "_" || c === "*") { i++; continue; }
    if (c === "[") {
      const j = src.indexOf("]", i);
      const b = src.slice(i + 1, j < 0 ? src.length : j);
      if (/^(h+|m+|s+)$/i.test(b)) return true;
      i = j < 0 ? src.length : j;
      continue;
    }
    if (/^standard/i.test(src.slice(i))) { i += 7; continue; }
    if (/^am\/pm/i.test(src.slice(i)) || /^a\/p/i.test(src.slice(i))) return true;
    if (/[tTjJmMhHsSdD]/.test(c) || c === "e") return true;
  }
  return false;
}

// ---------------------------------------------------------------------------------------------
// Zellen formatieren > Zahlen and the ribbon's number format box

export interface FormatCategory {
  /** Category name as in German Excel. */
  name: string;
  /** The formats the category offers, in the file form (show them with localFormat / formatValue). */
  codes: string[];
}

/** The categories of Zellen formatieren > Zahlen. Standard, Prozent and Wissenschaft offer one
 * code each (their decimals are a spin box); Zahl and Währung list the negative-number variants
 * with two decimals; the lists of Datum, Uhrzeit, Bruch and Sonderformat are German Excel's;
 * Benutzerdefiniert lists the built-in codes in Excel's order. Codes as Excel writes them. */
export const FORMAT_CATEGORIES: FormatCategory[] = [
  { name: "Standard", codes: ["General"] },
  {
    name: "Zahl",
    codes: ["0.00", "0.00;[Red]0.00", "0.00_ ;\\-0.00\\ ", "0.00_ ;[Red]\\-0.00\\ ",
      "#,##0.00", "#,##0.00;[Red]#,##0.00", "#,##0.00_ ;\\-#,##0.00\\ ", "#,##0.00_ ;[Red]\\-#,##0.00\\ "],
  },
  {
    name: "Währung",
    codes: ['#,##0.00\\ "€"', '#,##0.00\\ "€";[Red]#,##0.00\\ "€"', BUILTIN_FORMATS[7], BUILTIN_FORMATS[8]],
  },
  { name: "Buchhaltung", codes: [BUILTIN_FORMATS[44], BUILTIN_FORMATS[43]] },
  {
    name: "Datum",
    codes: ["m/d/yyyy", "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy", "dd/mm/yy;@", "yyyy\\-mm\\-dd;@", "[$-407]dd/\\ mmm\\ yy;@",
      "[$-407]d/\\ mmmm\\ yyyy;@", "[$-407]mmmmm;@", "[$-407]mmmmm\\ yy;@", "d/m/;@", "d/m/yy;@", "dd/mm/yyyy;@",
      "[$-407]d/\\ mmm/;@", "[$-407]d/\\ mmm/\\ yy;@", "[$-407]mmm/\\ yy;@", "[$-407]mmmm\\ yy;@", "d/m/yy\\ h:mm;@",
      "[$-409]d/m/yy\\ h:mm\\ AM/PM;@"],
  },
  {
    name: "Uhrzeit",
    codes: ["[$-F400]h:mm:ss\\ AM/PM", "hh:mm;@", "[$-409]h:mm\\ AM/PM;@", "hh:mm:ss;@", "[$-409]h:mm:ss\\ AM/PM;@",
      "mm:ss.0;@", "[h]:mm:ss;@", "[$-409]d/m/yy\\ h:mm\\ AM/PM;@", "d/m/yy\\ h:mm;@"],
  },
  { name: "Prozent", codes: ["0.00%"] },
  { name: "Bruch", codes: ["# ?/?", "# ??/??", "#\\ ???/???", "#\\ ?/2", "#\\ ?/4", "#\\ ?/8", "#\\ ??/16", "#\\ ?/10", "#\\ ??/100"] },
  { name: "Wissenschaft", codes: ["0.00E+00"] },
  { name: "Text", codes: ["@"] },
  { name: "Sonderformat", codes: ["00000"] },
  {
    name: "Benutzerdefiniert",
    codes: [0, 1, 2, 3, 4, 37, 38, 39, 40, 5, 6, 7, 8, 9, 10, 11, 48, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 47, 49, 46, 42, 41, 44, 43]
      .map((id) => BUILTIN_FORMATS[id]),
  },
];

/** Start > Zahl > Zahlenformat: the entries and the codes Excel applies. */
export const RIBBON_FORMATS: { label: string; code: string }[] = [
  { label: "Standard", code: "General" },
  { label: "Zahl", code: "0.00" },
  { label: "Währung", code: '#,##0.00\\ "€"' },
  { label: "Buchhaltung", code: BUILTIN_FORMATS[44] },
  { label: "Datum, kurz", code: "m/d/yyyy" },
  { label: "Datum, lang", code: "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy" },
  { label: "Uhrzeit", code: "[$-F400]h:mm:ss\\ AM/PM" },
  { label: "Prozent", code: "0.00%" },
  { label: "Bruch", code: "# ?/?" },
  { label: "Wissenschaft", code: "0.00E+00" },
  { label: "Text", code: "@" },
];
