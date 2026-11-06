// Text functions, second round: BAHTTEXT, DM, MATRIXZUTEXT, WERTZUTEXT, TEXTTEILEN, the regular
// expression functions (REGEXTESTEN, REGEXEXTRAHIEREN, REGEXERSETZEN) and URLCODIEREN.

import { def, deref } from "./eval";
import { defS, n, roundTo, toArr } from "./fn-util";
import { groupDE } from "./fn-text";
import { ERROR_TEXT_DE, type CellError } from "../model";
import { Arr, E, isErr, numberText, toBool, toStr, type Any, type V } from "./values";
import type { Ctx } from "./eval";

// ----- BAHTTEXT -------------------------------------------------------------------------------

const TH_DIGITS = ["ศูนย์", "หนึ่ง", "สอง", "สาม", "สี่", "ห้า", "หก", "เจ็ด", "แปด", "เก้า"];
const TH_PLACES = ["", "สิบ", "ร้อย", "พัน", "หมื่น", "แสน"];

/** Thai words for a whole number below one million (empty for 0). */
function thaiGroup(x: number): string {
  let out = "";
  const digits = String(x).split("").map(Number);
  const len = digits.length;
  digits.forEach((d, i) => {
    const place = len - 1 - i;
    if (d === 0) return;
    if (place === 1) out += (d === 1 ? "" : d === 2 ? "ยี่" : TH_DIGITS[d]) + TH_PLACES[1];
    else if (place === 0 && d === 1 && x > 1) out += "เอ็ด";
    else out += TH_DIGITS[d] + TH_PLACES[place];
  });
  return out;
}

function thai(x: number): string {
  if (x >= 1e6) {
    const hi = Math.floor(x / 1e6), lo = x - hi * 1e6;
    return thai(hi) + "ล้าน" + thaiGroup(lo);
  }
  return thaiGroup(x);
}

defS("BAHTTEXT", 1, 1, ([a]) => {
  const x = n(a);
  if (isErr(x)) return x;
  const cents = Math.round(Math.abs(x) * 100);
  const baht = Math.floor(cents / 100), satang = cents % 100;
  let out = x < 0 && cents > 0 ? "ลบ" : "";
  if (baht === 0 && satang === 0) return "ศูนย์บาทถ้วน";
  if (baht > 0) out += thai(baht) + "บาท";
  out += satang ? thaiGroup(satang) + "สตางค์" : "ถ้วน";
  return out;
});

// ----- DM -------------------------------------------------------------------------------------

defS("DOLLAR", 1, 2, ([a, d0]) => {
  const x = n(a);
  if (isErr(x)) return x;
  let d = 2;
  if (d0 !== undefined && d0 !== null) {
    const v = n(d0);
    if (isErr(v)) return v;
    d = Math.trunc(v);
  }
  if (d > 127) return E.value;
  const r = roundTo(Math.abs(x), d, "half");
  return (x < 0 ? "-" : "") + groupDE(r, Math.max(0, d), true) + " €";
});

// ----- MATRIXZUTEXT, WERTZUTEXT ---------------------------------------------------------------

function valueText(v: V, strict: boolean): string {
  if (v === null) return "";
  if (typeof v === "number") return numberText(v);
  if (typeof v === "boolean") return v ? "WAHR" : "FALSCH";
  if (isErr(v)) return ERROR_TEXT_DE[v.code];
  return strict ? "\"" + v.replace(/"/g, "\"\"") + "\"" : v;
}

function formatArg(a: Any | undefined, ctx: Ctx): boolean | CellError {
  if (a === undefined || a === null) return false;
  const v = deref(a, ctx);
  const x = n(v instanceof Arr ? v.d[0] : v);
  if (isErr(x)) return x;
  if (x !== 0 && x !== 1) return E.value;
  return x === 1;
}

def("ARRAYTOTEXT", 1, 2, ([a0, f0], ctx) => {
  const strict = formatArg(f0, ctx);
  if (isErr(strict)) return strict;
  const a = isErr(a0) ? new Arr(1, 1, [a0]) : toArr(a0, ctx);
  if (isErr(a)) return a;
  if (!strict) return a.d.map((v) => valueText(v, false)).join("; ");
  const rows: string[] = [];
  for (let r = 0; r < a.rows; r++) {
    const cells: string[] = [];
    for (let c = 0; c < a.cols; c++) cells.push(valueText(a.get(r, c), true));
    rows.push(cells.join("."));
  }
  return "{" + rows.join(";") + "}";
});

def("VALUETOTEXT", 1, 2, ([a0, f0], ctx) => {
  const strict = formatArg(f0, ctx);
  if (isErr(strict)) return strict;
  const v = deref(a0, ctx);
  if (v instanceof Arr) return new Arr(v.rows, v.cols, v.d.map((x) => valueText(x, strict)));
  return valueText(v, strict);
});

// ----- TEXTTEILEN -----------------------------------------------------------------------------

/** The delimiters of an argument (several from an array); null when it is missing. */
function delims(a: Any | undefined, ctx: Ctx): string[] | null | CellError {
  if (a === undefined || a === null) return null;
  const v = deref(a, ctx);
  const out: string[] = [];
  for (const x of v instanceof Arr ? v.d : [v]) {
    const t = toStr(x);
    if (isErr(t)) return t;
    if (t === "") return E.value;
    out.push(t);
  }
  return out;
}

function splitBy(text: string, ds: string[], ci: boolean): string[] {
  const hay = ci ? text.toLowerCase() : text;
  const needles = ds.map((d) => (ci ? d.toLowerCase() : d));
  const out: string[] = [];
  let start = 0, i = 0;
  while (i <= text.length) {
    let hit = "";
    for (const d of needles) if (d.length > hit.length && hay.startsWith(d, i)) hit = d;
    if (hit) {
      out.push(text.slice(start, i));
      i += hit.length;
      start = i;
    } else i++;
  }
  out.push(text.slice(start));
  return out;
}

def("TEXTSPLIT", 2, 6, ([t0, c0, r0, i0, m0, p0], ctx) => {
  const tv = deref(t0, ctx);
  const text = toStr(tv instanceof Arr ? tv.d[0] : tv);
  if (isErr(text)) return text;
  const cd = delims(c0, ctx), rd = delims(r0, ctx);
  if (isErr(cd)) return cd;
  if (isErr(rd)) return rd;
  if (!cd && !rd) return E.value;
  const flag = (a: Any | undefined): boolean | CellError => {
    if (a === undefined || a === null) return false;
    const v = deref(a, ctx);
    return toBool(v instanceof Arr ? v.d[0] : v);
  };
  const ignore = flag(i0), ci = flag(m0);
  if (isErr(ignore)) return ignore;
  if (isErr(ci)) return ci;
  if (text === "") return E.value;
  const pad: V = p0 === undefined ? E.na : (deref(p0, ctx) as V);
  let rows = rd ? splitBy(text, rd, ci) : [text];
  if (ignore) rows = rows.filter((r) => r !== "");
  const cells = rows.map((r) => {
    let c = cd ? splitBy(r, cd, ci) : [r];
    if (ignore) c = c.filter((x) => x !== "");
    return c;
  }).filter((c) => c.length);
  if (!cells.length) return E.calc;
  const cols = Math.max(...cells.map((c) => c.length));
  return Arr.of(cells.map((c) => [...c, ...new Array(cols - c.length).fill(pad)]));
});

// ----- Regular expressions --------------------------------------------------------------------

function regex(p: V, caseArg: V | undefined, global: boolean): RegExp | CellError {
  const pat = toStr(p);
  if (isErr(pat)) return pat;
  let ci = false;
  if (caseArg !== undefined && caseArg !== null) {
    const c = n(caseArg);
    if (isErr(c)) return c;
    if (c !== 0 && c !== 1) return E.value;
    ci = c === 1;
  }
  try {
    return new RegExp(pat, (global ? "g" : "") + (ci ? "i" : ""));
  } catch {
    return E.value;
  }
}

defS("REGEXTEST", 2, 3, ([t0, p0, c0]) => {
  const t = toStr(t0);
  if (isErr(t)) return t;
  const re = regex(p0, c0, false);
  return isErr(re) ? re : re.test(t);
});

defS("REGEXEXTRACT", 2, 4, ([t0, p0, m0, c0]) => {
  const t = toStr(t0);
  if (isErr(t)) return t;
  let mode = 0;
  if (m0 !== undefined && m0 !== null) {
    const m = n(m0);
    if (isErr(m)) return m;
    mode = Math.trunc(m);
    if (mode < 0 || mode > 2) return E.value;
  }
  const re = regex(p0, c0, mode === 1);
  if (isErr(re)) return re;
  if (mode === 1) {
    const all = [...t.matchAll(re)].map((m) => m[0]);
    return all.length ? new Arr(1, all.length, all) : E.na;
  }
  const m = re.exec(t);
  if (!m) return E.na;
  if (mode === 0) return m[0];
  const groups = m.length > 1 ? m.slice(1).map((g) => g ?? "") : [m[0]];
  return new Arr(1, groups.length, groups);
});

defS("REGEXREPLACE", 3, 5, ([t0, p0, r0, o0, c0]) => {
  const t = toStr(t0), rep = toStr(r0);
  if (isErr(t)) return t;
  if (isErr(rep)) return rep;
  let occ = 0;
  if (o0 !== undefined && o0 !== null) {
    const o = n(o0);
    if (isErr(o)) return o;
    occ = Math.trunc(o);
  }
  const re = regex(p0, c0, true);
  if (isErr(re)) return re;
  if (occ === 0) return t.replace(re, rep);
  const ms = [...t.matchAll(re)];
  const m = occ > 0 ? ms[occ - 1] : ms[ms.length + occ];
  if (!m || m.index === undefined) return t;
  const one = new RegExp(re.source, re.flags.replace("g", ""));
  return t.slice(0, m.index) + m[0].replace(one, rep) + t.slice(m.index + m[0].length);
});

// ----- URLCODIEREN ----------------------------------------------------------------------------

defS("ENCODEURL", 1, 1, ([a]) => {
  const t = toStr(a);
  if (isErr(t)) return t;
  return encodeURIComponent(t).replace(/[!'()*~]/g, (c) => "%" + c.charCodeAt(0).toString(16).toUpperCase());
});
