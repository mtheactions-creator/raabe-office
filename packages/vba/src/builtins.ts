// VBA's built-in functions that need nothing but their arguments (strings, numbers, dates, arrays, conversions).
// The ones that reach the runtime or the host (MsgBox, files, settings, CreateObject …) are in interp.ts.

import { VbaError, bankers, dateToSerial, format, isArr, isObj, serialToDate, toBool, toDate, toNum, toStr, type VbaArray, type VbaValue } from "./values";
import { ERROR_TEXT } from "./objects";

export const NOT_BUILTIN = Symbol("not builtin");

/** Parameter names of built-ins that are called with named arguments (MsgBox Prompt:=…, Title:=…). */
export const BUILTIN_PARAMS: Record<string, string[]> = {
  msgbox: ["prompt", "buttons", "title", "helpfile", "context"],
  inputbox: ["prompt", "title", "default", "xpos", "ypos", "helpfile", "context"],
  instr: ["start", "string1", "string2", "compare"],
  replace: ["expression", "find", "replace", "start", "count", "compare"],
  split: ["expression", "delimiter", "limit", "compare"],
  format: ["expression", "format", "firstdayofweek", "firstweekofyear"],
  mid: ["string", "start", "length"],
  dateadd: ["interval", "number", "date"],
  datediff: ["interval", "date1", "date2", "firstdayofweek", "firstweekofyear"],
  getsetting: ["appname", "section", "key", "default"],
  savesetting: ["appname", "section", "key", "setting"],
  dir: ["pathname", "attributes"],
};

const days = (d: Date) => Math.floor(dateToSerial(d));
const wholeDiff = (a: Date, b: Date, unit: number) => Math.floor((b.getTime() - a.getTime()) / unit);

function dateAdd(interval: string, n: number, d: Date): Date {
  const r = new Date(d.getTime());
  switch (interval.toLowerCase()) {
    case "yyyy": r.setFullYear(r.getFullYear() + n); break;
    case "q": r.setMonth(r.getMonth() + 3 * n); break;
    case "m": r.setMonth(r.getMonth() + n); break;
    case "y": case "d": case "w": r.setDate(r.getDate() + n); break;
    case "ww": r.setDate(r.getDate() + 7 * n); break;
    case "h": r.setTime(r.getTime() + n * 3600000); break;
    case "n": r.setTime(r.getTime() + n * 60000); break;
    case "s": r.setTime(r.getTime() + n * 1000); break;
    default: throw new VbaError("Ungültiger Prozeduraufruf oder ungültiges Argument", 5);
  }
  return r;
}
function dateDiff(interval: string, a: Date, b: Date): number {
  switch (interval.toLowerCase()) {
    case "yyyy": return b.getFullYear() - a.getFullYear();
    case "q": return (b.getFullYear() - a.getFullYear()) * 4 + Math.floor(b.getMonth() / 3) - Math.floor(a.getMonth() / 3);
    case "m": return (b.getFullYear() - a.getFullYear()) * 12 + b.getMonth() - a.getMonth();
    case "y": case "d": return days(b) - days(a);
    case "w": return Math.trunc((days(b) - days(a)) / 7);
    case "ww": { const sa = days(a) - ((a.getDay() + 7) % 7), sb = days(b) - ((b.getDay() + 7) % 7); return Math.round((sb - sa) / 7); }
    case "h": return wholeDiff(new Date(a.getFullYear(), a.getMonth(), a.getDate(), a.getHours()), new Date(b.getFullYear(), b.getMonth(), b.getDate(), b.getHours()), 3600000);
    case "n": return wholeDiff(new Date(a.getFullYear(), a.getMonth(), a.getDate(), a.getHours(), a.getMinutes()), new Date(b.getFullYear(), b.getMonth(), b.getDate(), b.getHours(), b.getMinutes()), 60000);
    case "s": return Math.round((b.getTime() - a.getTime()) / 1000);
  }
  throw new VbaError("Ungültiger Prozeduraufruf oder ungültiges Argument", 5);
}

const MONTHS = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"];
const WEEKDAYS = ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"];

/** The value of a pure built-in, or NOT_BUILTIN. `text` is Option Compare Text of the calling module. */
// ----- Rnd and Randomize: VBA's own generator (D-105) -----------------------------------------------------
// VBA's random numbers come from a 24-bit linear congruential generator, so a macro that seeds it (Rnd -1, then
// Randomize n) draws the same numbers in Office and here (measured with PowerPoint's macro test NamenZiehen).
let rndSeed = 0x50000;
const scratch = new DataView(new ArrayBuffer(8));
const timerNow = () => { const d = new Date(); return d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds() + d.getMilliseconds() / 1000; };

/** Rnd(n): n < 0 seeds from the number's single-precision bits and steps once; 0 repeats the last; else steps. */
function rnd(x?: number): number {
  if (x !== undefined && x < 0) {
    scratch.setFloat32(0, x, true);
    const b = scratch.getUint32(0, true);
    rndSeed = (b + (b >>> 24)) & 0xffffff;
  } else if (x === 0) return rndSeed / 16777216;
  rndSeed = Number((BigInt(rndSeed) * 0x43fd43fdn + 0xc39ec3n) & 0xffffffn);
  return rndSeed / 16777216;
}

/** Randomize n: the high word of the number's double bits, its halves mixed, become the seed's middle bytes. */
function randomize(x: number): void {
  scratch.setFloat64(0, x, true);
  const hi = scratch.getUint32(4, true);
  const t = (((hi & 0xffff) ^ (hi >>> 16)) << 8) & 0xffff00;
  rndSeed = (rndSeed & 0xff0000ff) | t;
}

export function pureBuiltin(k: string, a: VbaValue[], text: boolean, base: number): VbaValue | typeof NOT_BUILTIN {
  const s = (i: number) => toStr(a[i]), n = (i: number) => toNum(a[i]);
  const has = (i: number) => a[i] !== undefined;
  const cmpText = (i: number) => (has(i) ? n(i) === 1 : text);
  switch (k) {
    case "len": return isArr(a[0]) ? a[0].items.length : a[0] === null ? null : s(0).length;
    case "left": return s(0).slice(0, Math.max(0, n(1)));
    case "right": return n(1) <= 0 ? "" : s(0).slice(-n(1));
    case "mid": { if (n(1) < 1) throw new VbaError("Ungültiger Prozeduraufruf oder ungültiges Argument", 5); return !has(2) ? s(0).slice(n(1) - 1) : s(0).substr(n(1) - 1, Math.max(0, n(2))); }
    case "ucase": return s(0).toUpperCase();
    case "lcase": return s(0).toLowerCase();
    case "trim": return s(0).replace(/^ +| +$/g, "");
    case "ltrim": return s(0).replace(/^ +/, "");
    case "rtrim": return s(0).replace(/ +$/, "");
    case "instr": {
      const withStart = a.length >= 3 && typeof a[0] === "number";
      const start = withStart ? n(0) : 1, hay = withStart ? s(1) : s(0), needle = withStart ? s(2) : s(1);
      const t = withStart ? cmpText(3) : text;
      if (start < 1) throw new VbaError("Ungültiger Prozeduraufruf oder ungültiges Argument", 5);
      return (t ? hay.toLowerCase().indexOf(needle.toLowerCase(), start - 1) : hay.indexOf(needle, start - 1)) + 1;
    }
    case "instrrev": { const t = cmpText(3); const hay = t ? s(0).toLowerCase() : s(0), needle = t ? s(1).toLowerCase() : s(1); const from = has(2) && n(2) > 0 ? n(2) - 1 : hay.length; return hay.lastIndexOf(needle, from - needle.length + 1 < 0 ? -1 : from) + 1; }
    case "replace": {
      const src = s(0), find = s(1), rep = s(2), start = has(3) ? n(3) : 1, count = has(4) ? n(4) : -1;
      if (!find) return src.slice(start - 1);
      const t = cmpText(5);
      let out = "", i = start - 1, done = 0;
      const hay = t ? src.toLowerCase() : src, f = t ? find.toLowerCase() : find;
      for (;;) { const j = count >= 0 && done >= count ? -1 : hay.indexOf(f, i); if (j < 0) { out += src.slice(i); break; } out += src.slice(i, j) + rep; i = j + find.length; done++; }
      return out;
    }
    case "space": return " ".repeat(Math.max(0, n(0)));
    case "string": return (typeof a[1] === "number" ? String.fromCharCode(a[1]) : s(1).charAt(0)).repeat(Math.max(0, n(0)));
    case "strreverse": return Array.from(s(0)).reverse().join("");
    case "strcomp": { const t = cmpText(2); const x = t ? s(0).toLowerCase() : s(0), y = t ? s(1).toLowerCase() : s(1); return x < y ? -1 : x > y ? 1 : 0; }
    case "strconv": { const m = n(1) & 3; if (m === 3) return s(0).toLowerCase().replace(/(^|[\s-])(\S)/g, (_, p, c) => p + c.toUpperCase()); if (m === 1) return s(0).toUpperCase(); if (m === 2) return s(0).toLowerCase(); return s(0); }
    case "chr": case "chrw": return String.fromCharCode(n(0));
    case "asc": case "ascw": { const t = s(0); if (!t) throw new VbaError("Ungültiger Prozeduraufruf oder ungültiges Argument", 5); return t.charCodeAt(0); }
    case "cstr": return toStr(a[0]);
    case "str": return toNum(a[0]) >= 0 ? " " + String(toNum(a[0])) : String(toNum(a[0]));
    case "val": { const m = /^\s*[-+]?(\d[\d ]*)?\.?\d*(e[-+]?\d+)?/i.exec(s(0)); const t = m ? m[0].replace(/ /g, "") : ""; const v = Number(t); return t && !Number.isNaN(v) ? v : /^\s*&h([0-9a-f]+)/i.test(s(0)) ? parseInt(/^\s*&h([0-9a-f]+)/i.exec(s(0))![1], 16) : 0; }
    case "hex": return (n(0) < 0 ? (n(0) >>> 0) : Math.round(n(0))).toString(16).toUpperCase();
    case "oct": return (n(0) < 0 ? (n(0) >>> 0) : Math.round(n(0))).toString(8);
    case "cint": { const x = bankers(n(0)); if (x < -32768 || x > 32767) throw new VbaError("Überlauf", 6); return x; }
    case "clng": case "clnglng": case "clngptr": { const x = bankers(n(0)); if (k === "clng" && (x < -2147483648 || x > 2147483647)) throw new VbaError("Überlauf", 6); return x; }
    case "cbyte": { const x = bankers(n(0)); if (x < 0 || x > 255) throw new VbaError("Überlauf", 6); return x; }
    case "cdbl": case "csng": case "cdec": return n(0);
    case "ccur": return Math.round(n(0) * 10000) / 10000;
    case "cvar": return a[0];
    case "cbool": return toBool(a[0]);
    case "cdate": case "datevalue": return k === "datevalue" ? new Date(new Date(toDate(a[0]).getTime()).setHours(0, 0, 0, 0)) : toDate(a[0]);
    case "timevalue": { const d = toDate(a[0]); return new Date(1899, 11, 30, d.getHours(), d.getMinutes(), d.getSeconds()); }
    case "cverr": return n(0);
    case "int": return Math.floor(n(0));
    case "fix": return Math.trunc(n(0));
    case "abs": return Math.abs(n(0));
    case "sgn": return Math.sign(n(0));
    case "sqr": { if (n(0) < 0) throw new VbaError("Ungültiger Prozeduraufruf oder ungültiges Argument", 5); return Math.sqrt(n(0)); }
    case "exp": return Math.exp(n(0));
    case "log": { if (n(0) <= 0) throw new VbaError("Ungültiger Prozeduraufruf oder ungültiges Argument", 5); return Math.log(n(0)); }
    case "sin": return Math.sin(n(0));
    case "cos": return Math.cos(n(0));
    case "tan": return Math.tan(n(0));
    case "atn": return Math.atan(n(0));
    case "round": { const d = has(1) ? n(1) : 0; const f = 10 ** d; return bankers(Number((n(0) * f).toPrecision(15))) / f; }
    case "isnumeric": { if (typeof a[0] === "number") return true; if (typeof a[0] === "boolean") return true; if (typeof a[0] !== "string") return false; const t = a[0].trim(); return t !== "" && !Number.isNaN(Number(t.replace(",", "."))); }
    case "isempty": return a[0] === undefined;
    case "isnull": return a[0] === null;
    case "isarray": return isArr(a[0]);
    case "isobject": return isObj(a[0]) || a[0] === null;
    case "ismissing": return a[0] === undefined;
    case "iserror": return false;
    case "isdate": { if (a[0] instanceof Date) return true; try { toDate(a[0]); return typeof a[0] === "string"; } catch { return false; } }
    case "ubound": case "lbound": {
      let arr = a[0];
      if (!isArr(arr)) throw new VbaError("Typen unverträglich", 13);
      const dim = has(1) ? n(1) : 1;
      for (let i = 1; i < dim; i++) { const next: VbaValue = arr.items[0]; if (!isArr(next)) throw new VbaError("Index außerhalb des gültigen Bereichs", 9); arr = next; }
      if (k === "ubound") return arr.lb + arr.items.length - 1;
      return arr.lb;
    }
    case "array": return { kind: "array", lb: base, items: [...a] };
    case "split": {
      const src = s(0);
      if (src === "") return { kind: "array", lb: 0, items: [] } as VbaArray;
      const d = has(1) ? s(1) : " ";
      let parts = d === "" ? [src] : src.split(d);
      if (has(2) && n(2) > 0 && parts.length > n(2)) parts = [...parts.slice(0, n(2) - 1), parts.slice(n(2) - 1).join(d)];
      return { kind: "array", lb: 0, items: parts };
    }
    case "join": { const arr = a[0]; return isArr(arr) ? arr.items.map(toStr).join(has(1) ? s(1) : " ") : ""; }
    case "filter": { const arr = a[0]; if (!isArr(arr)) throw new VbaError("Typen unverträglich", 13); const inc = has(2) ? toBool(a[2]) : true, t = cmpText(3); const m = t ? s(1).toLowerCase() : s(1); return { kind: "array", lb: 0, items: arr.items.filter((x) => (t ? toStr(x).toLowerCase() : toStr(x)).includes(m) === inc) }; }
    case "now": return new Date();
    case "date": return new Date(new Date().setHours(0, 0, 0, 0));
    case "time": { const d = new Date(); return new Date(1899, 11, 30, d.getHours(), d.getMinutes(), d.getSeconds()); }
    case "timer": { const d = new Date(); return d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds() + d.getMilliseconds() / 1000; }
    case "dateserial": return new Date(n(0) < 100 ? 1900 + n(0) + (n(0) < 30 ? 100 : 0) : n(0), n(1) - 1, n(2));
    case "timeserial": return new Date(1899, 11, 30, n(0), n(1), n(2));
    case "dateadd": return dateAdd(s(0), n(1), toDate(a[2]));
    case "datediff": return dateDiff(s(0), toDate(a[1]), toDate(a[2]));
    case "datepart": { const d = toDate(a[1]); switch (s(0).toLowerCase()) { case "yyyy": return d.getFullYear(); case "q": return Math.floor(d.getMonth() / 3) + 1; case "m": return d.getMonth() + 1; case "y": return days(d) - days(new Date(d.getFullYear(), 0, 1)) + 1; case "d": return d.getDate(); case "w": return d.getDay() + 1; case "h": return d.getHours(); case "n": return d.getMinutes(); case "s": return d.getSeconds(); } throw new VbaError("Ungültiger Prozeduraufruf oder ungültiges Argument", 5); }
    case "year": return toDate(a[0]).getFullYear();
    case "month": return toDate(a[0]).getMonth() + 1;
    case "day": return toDate(a[0]).getDate();
    case "hour": return toDate(a[0]).getHours();
    case "minute": return toDate(a[0]).getMinutes();
    case "second": return toDate(a[0]).getSeconds();
    case "weekday": { const w = toDate(a[0]).getDay() + 1; const first = has(1) && n(1) > 0 ? n(1) : 1; return ((w - first + 7) % 7) + 1; }
    case "monthname": { const m = MONTHS[n(0) - 1]; if (!m) throw new VbaError("Ungültiger Prozeduraufruf oder ungültiges Argument", 5); return has(1) && toBool(a[1]) ? m.slice(0, 3) : m; }
    case "weekdayname": { const first = has(2) && n(2) > 0 ? n(2) : 2; const w = WEEKDAYS[(n(0) - 1 + first - 1) % 7]; if (!w) throw new VbaError("Ungültiger Prozeduraufruf oder ungültiges Argument", 5); return has(1) && toBool(a[1]) ? w.slice(0, 2) : w; }
    case "format": return format(a[0], has(1) ? s(1) : "");
    case "formatnumber": { const d = has(1) ? n(1) : 2; return n(0).toLocaleString("de-DE", { minimumFractionDigits: d, maximumFractionDigits: d }); }
    case "formatpercent": { const d = has(1) ? n(1) : 2; return (n(0) * 100).toLocaleString("de-DE", { minimumFractionDigits: d, maximumFractionDigits: d }) + "%"; }
    case "formatcurrency": { const d = has(1) ? n(1) : 2; return n(0).toLocaleString("de-DE", { style: "currency", currency: "EUR", minimumFractionDigits: d, maximumFractionDigits: d }); }
    case "formatdatetime": { const f = has(1) ? n(1) : 0; return format(toDate(a[0]), ["general date", "long date", "short date", "long time", "short time"][f] ?? "general date"); }
    case "rgb": return (n(0) & 255) | ((n(1) & 255) << 8) | ((n(2) & 255) << 16);
    case "qbcolor": return [0, 0x800000, 0x8000, 0x808000, 0x80, 0x800080, 0x8080, 0xc0c0c0, 0x808080, 0xff0000, 0xff00, 0xffff00, 0xff, 0xff00ff, 0xffff, 0xffffff][n(0)] ?? 0;
    case "iif": return toBool(a[0]) ? a[1] : a[2];
    case "choose": { const i = Math.trunc(n(0)); return i >= 1 && i < a.length ? a[i] : null; }
    case "switch": { for (let i = 0; i + 1 < a.length; i += 2) if (toBool(a[i])) return a[i + 1]; return null; }
    case "nz": return a[0] === null || a[0] === undefined ? (has(1) ? a[1] : 0) : a[0];
    case "error": case "error$": return ERROR_TEXT[n(0)] ?? "Anwendungs- oder objektdefinierter Fehler";
    case "doevents": return 0;
    case "beep": return undefined;
    case "randomize": randomize(has(0) ? n(0) : timerNow()); return undefined;
    case "rnd": return rnd(has(0) ? n(0) : undefined);
    case "cvdate": return toDate(a[0]);
    case "serialdate": return serialToDate(n(0));
  }
  return NOT_BUILTIN;
}
