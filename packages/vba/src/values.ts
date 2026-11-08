// The values of our VBA interpreter (D-079) and the conversions between them, as VBA converts: numbers, strings
// (German conventions: "," as the decimal separator, Wahr/Falsch), Booleans (True is -1), dates, arrays and
// objects. Shared by the parser, the runtime, the built-in objects and the forms.

export type VbaValue = string | number | boolean | null | undefined | Date | VbaArray | VbaObject;
/** An array; a multi-dimensional array is an array of arrays (the first index outermost). */
export interface VbaArray { kind: "array"; lb: number; items: VbaValue[] }
/** An object of the app's model: properties, methods and (for collections) items. */
export interface VbaObject {
  kind: "object";
  /** `named`: the named arguments of the call (Key:=…), by lower-case name, when the call had any (D-100). Without
   *  argNames the positional `args` then hold only the unnamed arguments, in order. */
  get(name: string, args: VbaValue[], named?: Record<string, VbaValue>): VbaValue;
  set?(name: string, value: VbaValue, args: VbaValue[]): void;
  call?(name: string, args: VbaValue[], named?: Record<string, VbaValue>): VbaValue;
  /** For Each over a collection. */
  items?(): VbaValue[];
  /** The object's default member (Item for collections, Value, Text …). */
  default?(args: VbaValue[]): VbaValue;
  /** What TypeName shows for the object (Shape, Slide, Range …); "Object" when absent. */
  typeName?(): string;
  /** The parameter names of a method or parameterised property, so that named arguments (Key:=…) reach the right
   *  position (they come in `named` as well); without it they come only in `named`. */
  argNames?(name: string): string[] | undefined;
}

export class VbaError extends Error {
  /** The module and procedure the error was raised in (set when it leaves the procedure). */
  module?: string;
  proc?: string;
  constructor(message: string, readonly number = 5, readonly line = 0) { super(message); }
}

/** Returned by an object's call() when the name is a property, not a method. */
export const NOT_A_METHOD = Symbol("not a method") as unknown as VbaValue;

/** VBA's day number of a date (days since 30.12.1899, the time as the fraction). */
// Dates are local wall-clock times (DateSerial, Now build them so), the serial counts their calendar fields.
export const dateToSerial = (d: Date): number =>
  Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds(), d.getMilliseconds()) / 86400000 + 25569;
export const serialToDate = (n: number): Date => {
  const u = new Date(Math.round((n - 25569) * 86400000));
  return new Date(u.getUTCFullYear(), u.getUTCMonth(), u.getUTCDate(), u.getUTCHours(), u.getUTCMinutes(), u.getUTCSeconds(), u.getUTCMilliseconds());
};

export const toNum = (v: VbaValue): number => {
  if (typeof v === "number") return v;
  if (typeof v === "boolean") return v ? -1 : 0;
  if (v === undefined || v === null) return 0;
  if (v instanceof Date) return dateToSerial(v);
  if (typeof v === "string") { const n = Number(v.trim().replace(",", ".")); if (v.trim() === "" || Number.isNaN(n)) throw new VbaError("Typen unverträglich", 13); return n; }
  if (isObj(v) && v.default) return toNum(v.default([]));
  throw new VbaError("Typen unverträglich", 13);
};
export const toStr = (v: VbaValue): string => {
  if (typeof v === "string") return v;
  if (typeof v === "number") return String(v).replace(".", ",");
  if (typeof v === "boolean") return v ? "Wahr" : "Falsch";
  if (v === undefined || v === null) return "";
  if (v instanceof Date) return v.toLocaleString("de-DE");
  if ((v as VbaObject).kind === "object") { const d = (v as VbaObject).default?.([]); return d === undefined ? "" : toStr(d); }
  throw new VbaError("Typen unverträglich", 13);
};
export const toBool = (v: VbaValue): boolean => (typeof v === "boolean" ? v : typeof v === "string" && /^(wahr|true)$/i.test(v.trim()) ? true : typeof v === "string" && /^(falsch|false)$/i.test(v.trim()) ? false : toNum(v) !== 0);
export const toDate = (v: VbaValue): Date => (v instanceof Date ? v : typeof v === "number" ? serialToDate(v) : parseDate(toStr(v)));
export const isObj = (v: VbaValue): v is VbaObject => !!v && typeof v === "object" && (v as VbaObject).kind === "object";
export const isArr = (v: VbaValue): v is VbaArray => !!v && typeof v === "object" && (v as VbaArray).kind === "array";
/** Rounds as VBA does (banker's rounding) for integer conversions. */
export const bankers = (n: number): number => { const f = Math.floor(n), d = n - f; return d > 0.5 ? f + 1 : d < 0.5 ? f : f % 2 === 0 ? f : f + 1; };

/** A date written the German way (31.12.2024 13:45) or as JavaScript reads it. */
export function parseDate(s: string): Date {
  const m = /^\s*(\d{1,2})\.(\d{1,2})\.(\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?\s*$/.exec(s);
  if (m) { const y = Number(m[3]) < 100 ? 2000 + Number(m[3]) : Number(m[3]); return new Date(y, Number(m[2]) - 1, Number(m[1]), Number(m[4] ?? 0), Number(m[5] ?? 0), Number(m[6] ?? 0)); }
  const d = new Date(s);
  if (Number.isNaN(d.getTime())) throw new VbaError("Typen unverträglich", 13);
  return d;
}

/** The declared types the runtime coerces to, and their initial values. */
const NUMERIC = new Set(["integer", "long", "byte", "longlong", "longptr", "single", "double", "currency", "decimal"]);
export const SCALAR_TYPES = new Set([...NUMERIC, "string", "boolean", "date", "variant"]);
export function defaultOf(type: string | undefined): VbaValue {
  if (!type || type === "variant") return undefined;
  if (NUMERIC.has(type)) return 0;
  if (type === "string") return "";
  if (type === "boolean") return false;
  if (type === "date") return serialToDate(0);
  return null; // an object type: Nothing
}
/** A value assigned to a variable of a declared type, converted as VBA converts it (Overflow for whole numbers out of range). */
export function coerce(type: string | undefined, v: VbaValue): VbaValue {
  if (!type || type === "variant" || !SCALAR_TYPES.has(type)) return v;
  if (isArr(v)) return v;
  const whole = (lo: number, hi: number) => { const n = bankers(toNum(v)); if (n < lo || n > hi) throw new VbaError("Überlauf", 6); return n; };
  switch (type) {
    case "integer": return whole(-32768, 32767);
    case "long": case "longptr": return whole(-2147483648, 2147483647);
    case "longlong": return whole(-Number.MAX_SAFE_INTEGER, Number.MAX_SAFE_INTEGER);
    case "byte": return whole(0, 255);
    case "single": case "double": case "decimal": return toNum(v);
    case "currency": return Math.round(toNum(v) * 10000) / 10000;
    case "string": return toStr(v);
    case "boolean": return toBool(v);
    case "date": return toDate(v);
  }
  return v;
}

export function makeArray(dims: [number, number][], fill: () => VbaValue = () => undefined): VbaArray {
  const [lb, ub] = dims[0];
  return { kind: "array", lb, items: Array.from({ length: Math.max(0, ub - lb + 1) }, () => (dims.length > 1 ? makeArray(dims.slice(1), fill) : fill())) };
}
/** Arrays are values in VBA: an assignment copies them (objects inside stay shared). */
export function copyArray(a: VbaArray): VbaArray { return { kind: "array", lb: a.lb, items: a.items.map((x) => (isArr(x) ? copyArray(x) : x)) }; }
export function flat(a: VbaArray): VbaValue[] { return a.items.flatMap((x) => (isArr(x) ? flat(x) : [x])); }
export function getIn(a: VbaArray, idx: number[]): VbaValue {
  let cur: VbaValue = a;
  for (const i of idx) { if (!isArr(cur)) throw new VbaError("Index außerhalb des gültigen Bereichs", 9); const arr: VbaArray = cur; const k: number = Math.round(i) - arr.lb; if (k < 0 || k >= arr.items.length) throw new VbaError("Index außerhalb des gültigen Bereichs", 9); cur = arr.items[k]; }
  return cur;
}
/** Sets an element and returns the value it replaced. */
export function setIn(a: VbaArray, idx: number[], v: VbaValue): VbaValue {
  let cur: VbaArray = a;
  let old: VbaValue;
  idx.forEach((i, n) => {
    const k = Math.round(i) - cur.lb;
    if (k < 0 || k >= cur.items.length) throw new VbaError("Index außerhalb des gültigen Bereichs", 9);
    if (n === idx.length - 1) { old = cur.items[k]; cur.items[k] = v; } else { const next = cur.items[k]; if (!isArr(next)) throw new VbaError("Index außerhalb des gültigen Bereichs", 9); cur = next; }
  });
  return old;
}

export function like(s: string, p: string, text = false): boolean {
  let re = "";
  for (let i = 0; i < p.length; i++) {
    const c = p[i];
    if (c === "*") re += ".*"; else if (c === "?") re += "."; else if (c === "#") re += "\\d";
    else if (c === "[") { const j = p.indexOf("]", i); if (j < 0) { re += "\\["; continue; } re += "[" + p.slice(i + 1, j).replace(/^!/, "^").replace(/\\/g, "\\\\") + "]"; i = j; }
    else re += c.replace(/[.+^${}()|\\]/g, "\\$&");
  }
  return new RegExp(`^${re}$`, text ? "si" : "s").test(s);
}

export const CONSTANTS: Record<string, VbaValue> = {
  vbcrlf: "\r\n", vbcr: "\r", vblf: "\n", vbnewline: "\r\n", vbtab: "\t", vbnullstring: "", vbnullchar: "\0", vbback: "\b", vbformfeed: "\f", vbverticaltab: "\v",
  vbokonly: 0, vbokcancel: 1, vbabortretryignore: 2, vbyesnocancel: 3, vbyesno: 4, vbretrycancel: 5, vbcritical: 16, vbquestion: 32, vbexclamation: 48, vbinformation: 64,
  vbdefaultbutton1: 0, vbdefaultbutton2: 256, vbdefaultbutton3: 512, vbapplicationmodal: 0, vbsystemmodal: 4096,
  vbok: 1, vbcancel: 2, vbabort: 3, vbretry: 4, vbignore: 5, vbyes: 6, vbno: 7, vbtextcompare: 1, vbbinarycompare: 0,
  vbblack: 0, vbred: 255, vbgreen: 65280, vbyellow: 65535, vbblue: 16711680, vbmagenta: 16711935, vbcyan: 16776960, vbwhite: 16777215,
  vbobjecterror: -2147221504, vbmodal: 1, vbmodeless: 0, vbformcontrolmenu: 0, vbformcode: 1, vbappwindows: 2, vbapptaskmanager: 3,
  vbempty: 0, vbnull: 1, vbinteger: 2, vblong: 3, vbsingle: 4, vbdouble: 5, vbcurrency: 6, vbdate: 7, vbstring: 8, vbobject: 9, vberror: 10, vbboolean: 11, vbvariant: 12, vbbyte: 17, vbarray: 8192,
  vbnormal: 0, vbreadonly: 1, vbhidden: 2, vbsystem: 4, vbdirectory: 16, vbarchive: 32,
  vbuppercase: 1, vblowercase: 2, vbpropercase: 3, vbsunday: 1, vbmonday: 2, vbusesystemdayofweek: 0, vbgeneraldate: 0, vblongdate: 1, vbshortdate: 2, vblongtime: 3, vbshorttime: 4,
  vbtrue: -1, vbfalse: 0, vbusedefault: -2, vbget: 2, vblet: 4, vbset: 8, vbmethod: 1,
  // MS Forms constants (UserForms).
  fmmultiselectsingle: 0, fmmultiselectmulti: 1, fmmultiselectextended: 2, fmstyledropdowncombo: 0, fmstyledropdownlist: 2,
  fmorientationauto: -1, fmorientationvertical: 0, fmorientationhorizontal: 1, fmtextalignleft: 1, fmtextaligncenter: 2, fmtextalignright: 3,
  fmbackstyletransparent: 0, fmbackstyleopaque: 1, fmborderstylenone: 0, fmborderstylesingle: 1,
};

/** Format with the common named and numeric formats (German conventions). */
export function format(v: VbaValue, f: string): string {
  const lf = f.toLowerCase();
  if (v instanceof Date || ["short date", "long date", "short time", "long time", "general date", "medium date", "medium time"].includes(lf)) {
    const d = v instanceof Date ? v : toDate(v);
    if (lf === "short date" || !f) return d.toLocaleDateString("de-DE", { day: "2-digit", month: "2-digit", year: "numeric" });
    if (lf === "long date") return d.toLocaleDateString("de-DE", { weekday: "long", day: "numeric", month: "long", year: "numeric" });
    if (lf === "medium date") return d.toLocaleDateString("de-DE", { day: "2-digit", month: "short", year: "2-digit" });
    if (lf === "short time") return d.toLocaleTimeString("de-DE", { hour: "2-digit", minute: "2-digit" });
    if (lf === "long time" || lf === "medium time") return d.toLocaleTimeString("de-DE");
    if (lf === "general date") return d.toLocaleString("de-DE");
    const p = (x: number, l = 2) => String(x).padStart(l, "0");
    return f.replace(/yyyy|yy|mmmm|mmm|mm|dddd|ddd|dd|hh|nn|ss|m|d|h|n|s/gi, (t) => {
      switch (t.toLowerCase()) {
        case "yyyy": return String(d.getFullYear()); case "yy": return p(d.getFullYear() % 100);
        case "mmmm": return d.toLocaleDateString("de-DE", { month: "long" }); case "mmm": return d.toLocaleDateString("de-DE", { month: "short" });
        case "mm": return p(d.getMonth() + 1); case "m": return String(d.getMonth() + 1);
        case "dddd": return d.toLocaleDateString("de-DE", { weekday: "long" }); case "ddd": return d.toLocaleDateString("de-DE", { weekday: "short" });
        case "dd": return p(d.getDate()); case "d": return String(d.getDate());
        case "hh": return p(d.getHours()); case "h": return String(d.getHours()); case "nn": return p(d.getMinutes()); case "n": return String(d.getMinutes());
        case "ss": return p(d.getSeconds()); case "s": return String(d.getSeconds());
      }
      return t;
    });
  }
  if (v === undefined && !f) return "";
  const n = toNum(v);
  if (lf === "percent") return (n * 100).toLocaleString("de-DE", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " %";
  if (lf === "currency") return n.toLocaleString("de-DE", { style: "currency", currency: "EUR" });
  if (lf === "fixed") return n.toLocaleString("de-DE", { minimumFractionDigits: 2, maximumFractionDigits: 2, useGrouping: false });
  if (lf === "standard") return n.toLocaleString("de-DE", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  if (lf === "yes/no") return n ? "Ja" : "Nein";
  if (lf === "true/false") return n ? "Wahr" : "Falsch";
  if (lf === "on/off") return n ? "Ein" : "Aus";
  const m = /^([#0,]*)(?:\.([#0]+))?(%?)$/.exec(f);
  if (m && f) {
    const dec = (m[2] ?? "").length, min = (m[2] ?? "").replace(/#/g, "").length;
    const x = m[3] ? n * 100 : n;
    const intMin = (m[1].replace(/[^0]/g, "")).length;
    let s = x.toLocaleString("de-DE", { minimumFractionDigits: min, maximumFractionDigits: dec, useGrouping: m[1].includes(","), minimumIntegerDigits: Math.max(1, intMin) });
    if (!intMin && /^-?0(,|$)/.test(s) && m[1]) s = s.replace(/^(-?)0/, "$1");
    return s + (m[3] ? "%" : "");
  }
  return toStr(v);
}

/** Windows-1252 text (VBA's files and the project's MBCS streams, German letters included). */
const CP1252_HIGH: Record<number, string> = { 0x80: "€", 0x82: "‚", 0x83: "ƒ", 0x84: "„", 0x85: "…", 0x86: "†", 0x87: "‡", 0x88: "ˆ", 0x89: "‰", 0x8a: "Š", 0x8b: "‹", 0x8c: "Œ", 0x8e: "Ž", 0x91: "‘", 0x92: "’", 0x93: "“", 0x94: "”", 0x95: "•", 0x96: "–", 0x97: "—", 0x98: "˜", 0x99: "™", 0x9a: "š", 0x9b: "›", 0x9c: "œ", 0x9e: "ž", 0x9f: "Ÿ" };
const CP1252_BACK: Record<string, number> = Object.fromEntries(Object.entries(CP1252_HIGH).map(([k, v]) => [v, Number(k)]));
export function encodeCp1252(s: string): Uint8Array { return Uint8Array.from(Array.from(s), (c) => CP1252_BACK[c] ?? (c.charCodeAt(0) < 256 ? c.charCodeAt(0) : 0x3f)); }
export function decodeCp1252(b: Uint8Array): string { let s = ""; for (const c of b) s += CP1252_HIGH[c] ?? String.fromCharCode(c); return s; }
