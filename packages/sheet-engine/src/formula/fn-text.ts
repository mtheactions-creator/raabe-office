// Text functions.

import { def, deref } from "./eval";
import { defS, n, roundTo, s, values } from "./fn-util";
import { Arr, E, isErr, textToNumber, toStr, type V } from "./values";
import type { CellError } from "../model";
import { fileFormat, formatValue } from "../numfmt";

const MAX_TEXT = 32767;

function text1(name: string, f: (t: string) => V): void {
  defS(name, 1, 1, ([a]) => {
    const t = s(a);
    return isErr(t) ? t : f(t);
  });
}

/** An optional whole-number argument (truncated), `d` when omitted. */
function int(v: V | undefined, d: number): number | CellError {
  if (v === undefined) return d;
  const x = n(v);
  return isErr(x) ? x : Math.trunc(x);
}

defS("LEFT", 1, 2, ([a, k0]) => {
  const t = s(a);
  if (isErr(t)) return t;
  const k = int(k0, 1);
  if (isErr(k)) return k;
  return k < 0 ? E.value : t.slice(0, k);
});
defS("RIGHT", 1, 2, ([a, k0]) => {
  const t = s(a);
  if (isErr(t)) return t;
  const k = int(k0, 1);
  if (isErr(k)) return k;
  return k < 0 ? E.value : k === 0 ? "" : t.slice(-k);
});
defS("MID", 3, 3, ([a, st0, k0]) => {
  const t = s(a);
  if (isErr(t)) return t;
  const st = int(st0, 1), k = int(k0, 0);
  if (isErr(st)) return st;
  if (isErr(k)) return k;
  if (st < 1 || k < 0) return E.value;
  return t.slice(st - 1, st - 1 + k);
});
text1("LEN", (t) => t.length);
// Excel keeps ß in GROSS ("straße" → "STRAßE").
text1("UPPER", (t) => Array.from(t, (ch) => (ch === "ß" ? ch : ch.toUpperCase())).join(""));
text1("LOWER", (t) => t.toLowerCase());
text1("PROPER", (t) => {
  let out = "";
  let prevLetter = false;
  for (const ch of t) {
    const letter = /\p{L}/u.test(ch);
    out += letter ? (prevLetter ? ch.toLowerCase() : ch.toUpperCase()) : ch;
    prevLetter = letter;
  }
  return out;
});
text1("TRIM", (t) => t.replace(/ +/g, " ").replace(/^ | $/g, ""));
text1("CLEAN", (t) => t.replace(/[\u0000-\u001f]/g, ""));

/** Windows-1252, which ZEICHEN and CODE use on German Windows. */
const CP1252: Record<number, number> = {
  0x80: 0x20ac, 0x82: 0x201a, 0x83: 0x0192, 0x84: 0x201e, 0x85: 0x2026, 0x86: 0x2020, 0x87: 0x2021, 0x88: 0x02c6, 0x89: 0x2030, 0x8a: 0x0160,
  0x8b: 0x2039, 0x8c: 0x0152, 0x8e: 0x017d, 0x91: 0x2018, 0x92: 0x2019, 0x93: 0x201c, 0x94: 0x201d, 0x95: 0x2022, 0x96: 0x2013, 0x97: 0x2014,
  0x98: 0x02dc, 0x99: 0x2122, 0x9a: 0x0161, 0x9b: 0x203a, 0x9c: 0x0153, 0x9e: 0x017e, 0x9f: 0x0178,
};
const CP1252_BACK = new Map(Object.entries(CP1252).map(([k, v]) => [v, Number(k)]));

defS("CHAR", 1, 1, ([a]) => {
  const x = n(a);
  if (isErr(x)) return x;
  const c = Math.trunc(x);
  if (c < 1 || c > 255) return E.value;
  return String.fromCharCode(CP1252[c] ?? c);
});
text1("CODE", (t) => {
  if (!t) return E.value;
  const c = t.charCodeAt(0);
  if (c < 256) return c;
  return CP1252_BACK.get(c) ?? 63;
});
defS("UNICHAR", 1, 1, ([a]) => {
  const x = n(a);
  if (isErr(x)) return x;
  const c = Math.trunc(x);
  if (c < 1 || c > 0x10ffff) return E.value;
  return String.fromCodePoint(c);
});
text1("UNICODE", (t) => (t ? t.codePointAt(0)! : E.value));

defS("FIND", 2, 3, ([w0, t0, st0]) => {
  const w = s(w0), t = s(t0);
  if (isErr(w)) return w;
  if (isErr(t)) return t;
  const st = int(st0, 1);
  if (isErr(st)) return st;
  if (st < 1 || st > t.length + 1) return E.value;
  const i = t.indexOf(w, st - 1);
  return i < 0 ? E.value : i + 1;
});

/** A pattern with Excel's wildcards (*, ?, ~) as a regular expression source. */
export function wildcard(p: string, tildeTilde = true): string {
  let out = "";
  for (let i = 0; i < p.length; i++) {
    const c = p[i];
    // "~~" stands for "~" in lookups and SUCHEN, but stays "~~" in criteria (ZÄHLENWENN).
    if (c === "~" && i + 1 < p.length && (tildeTilde ? "*?~" : "*?").includes(p[i + 1])) {
      out += "\\" + p[++i];
    } else if (c === "*") out += "[\\s\\S]*";
    else if (c === "?") out += "[\\s\\S]";
    else out += c.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }
  return out;
}

defS("SEARCH", 2, 3, ([w0, t0, st0]) => {
  const w = s(w0), t = s(t0);
  if (isErr(w)) return w;
  if (isErr(t)) return t;
  const st = int(st0, 1);
  if (isErr(st)) return st;
  if (st < 1 || st > t.length + 1) return E.value;
  if (w === "") return st;
  const re = new RegExp(wildcard(w), "i");
  const m = re.exec(t.slice(st - 1));
  return m ? m.index + st : E.value;
});
defS("SUBSTITUTE", 3, 4, ([t0, o0, r0, k0]) => {
  const t = s(t0), o = s(o0), r = s(r0);
  if (isErr(t)) return t;
  if (isErr(o)) return o;
  if (isErr(r)) return r;
  if (o === "") return t;
  if (k0 === undefined) return t.split(o).join(r);
  const k = int(k0, 1);
  if (isErr(k)) return k;
  if (k < 1) return E.value;
  let i = -1;
  for (let c = 0; c < k; c++) {
    i = t.indexOf(o, i + 1);
    if (i < 0) return t;
  }
  return t.slice(0, i) + r + t.slice(i + o.length);
});
defS("REPLACE", 4, 4, ([t0, st0, k0, r0]) => {
  const t = s(t0), r = s(r0);
  if (isErr(t)) return t;
  if (isErr(r)) return r;
  const st = int(st0, 1), k = int(k0, 0);
  if (isErr(st)) return st;
  if (isErr(k)) return k;
  if (st < 1 || k < 0) return E.value;
  return t.slice(0, st - 1) + r + t.slice(st - 1 + k);
});
defS("REPT", 2, 2, ([t0, k0]) => {
  const t = s(t0);
  if (isErr(t)) return t;
  const k = int(k0, 0);
  if (isErr(k)) return k;
  if (k < 0 || t.length * k > MAX_TEXT) return E.value;
  return t.repeat(k);
});
defS("EXACT", 2, 2, ([a, b0]) => {
  const x = s(a), y = s(b0);
  if (isErr(x)) return x;
  if (isErr(y)) return y;
  return x === y;
});

// VERKETTEN works element by element on ranges (it spills); TEXTKETTE joins them.
defS("CONCATENATE", 1, 255, (vals) => {
  let out = "";
  for (const v of vals) {
    const t = s(v);
    if (isErr(t)) return t;
    out += t;
  }
  return out.length > MAX_TEXT ? E.value : out;
});
def("CONCAT", 1, 253, (args, ctx) => {
  let out = "";
  for (const a of args) {
    for (const v of values(a, ctx)) {
      const t = toStr(v);
      if (isErr(t)) return t;
      out += t;
    }
  }
  return out.length > MAX_TEXT ? E.value : out;
});
def("TEXTJOIN", 3, 252, ([d0, ig0, ...rest], ctx) => {
  const delims = values(d0, ctx).map((v) => toStr(v));
  for (const d of delims) if (isErr(d)) return d;
  const ig = deref(ig0, ctx);
  if (ig instanceof Arr) return E.value;
  if (isErr(ig)) return ig;
  const ignore = ig === null ? false : typeof ig === "string" ? E.value : Boolean(ig);
  if (isErr(ignore)) return ignore;
  const parts: string[] = [];
  for (const a of rest) {
    for (const v of values(a, ctx)) {
      const t = toStr(v);
      if (isErr(t)) return t;
      if (ignore && t === "") continue;
      parts.push(t);
    }
  }
  let out = "";
  parts.forEach((p, i) => {
    if (i > 0) out += delims.length ? (delims[(i - 1) % delims.length] as string) : "";
    out += p;
  });
  return out.length > MAX_TEXT ? E.value : out;
});

defS("VALUE", 1, 1, ([a]) => {
  if (typeof a === "number") return a;
  if (a === null) return 0;
  if (typeof a === "boolean") return E.value;
  if (isErr(a)) return a;
  const v = textToNumber(a);
  return v === null ? E.value : v;
});
defS("NUMBERVALUE", 1, 3, ([a, d0, g0]) => {
  const t = s(a);
  if (isErr(t)) return t;
  const d = d0 === undefined ? "," : s(d0), g = g0 === undefined ? "." : s(g0);
  if (isErr(d)) return d;
  if (isErr(g)) return g;
  let x = t.replace(/\s+/g, "");
  if (x === "") return 0;
  let pct = 0;
  while (x.endsWith("%")) { pct++; x = x.slice(0, -1); }
  const dc = d[0] ?? ",", gc = g[0] ?? ".";
  const di = x.indexOf(dc);
  let ip = di >= 0 ? x.slice(0, di) : x;
  const fp = di >= 0 ? x.slice(di + 1) : "";
  if (fp.includes(gc) || fp.includes(dc)) return E.value;
  ip = ip.split(gc).join("");
  const v = Number(ip + "." + fp);
  if (!/^[+-]?\d*\.?\d*(e[+-]?\d+)?$/i.test(ip + "." + fp) || Number.isNaN(v)) return E.value;
  return v / 100 ** pct;
});

/** German number text with grouping: 1234567,891 → "1.234.567,891". */
export function groupDE(x: number, decimals: number, group: boolean): string {
  const neg = x < 0;
  const fixed = Math.abs(x).toFixed(Math.max(0, decimals));
  let [ip, fp] = fixed.split(".");
  if (group) ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
  return (neg && Number(fixed) !== 0 ? "-" : "") + ip + (fp ? "," + fp : "");
}

defS("FIXED", 1, 3, ([a, d0, g0]) => {
  const x = n(a);
  if (isErr(x)) return x;
  const d = d0 === undefined || d0 === null ? 2 : int(d0, 2);
  if (isErr(d)) return d;
  const noGroup = g0 === undefined || g0 === null ? false : Boolean(n(g0));
  const r = roundTo(x, Math.min(d, 127), "half");
  return groupDE(r, Math.min(Math.max(d, 0), 127), !noGroup);
});

// TEXT formats with the cell formatter (numfmt.ts); its code is typed in German ("TT.MM.JJJJ").
defS("TEXT", 2, 2, ([a, f0]) => {
  if (isErr(a)) return a;
  const code = s(f0);
  if (isErr(code)) return code;
  let v: number | string | boolean = a === null ? 0 : a;
  if (typeof v === "string") {
    const num = textToNumber(v);
    if (num !== null) v = num;
  }
  if (code === "") return "";
  try {
    return formatValue(v, fileFormat(code)).text;
  } catch {
    return E.value;
  }
});

defS("TEXTBEFORE", 2, 6, ([t0, d0, k0, m0, e0, nf]) => textAround(t0, d0, k0, m0, e0, nf, true));
defS("TEXTAFTER", 2, 6, ([t0, d0, k0, m0, e0, nf]) => textAround(t0, d0, k0, m0, e0, nf, false));

function textAround(t0: V, d0: V, k0: V | undefined, m0: V | undefined, e0: V | undefined, nf: V | undefined, before: boolean): V {
  const t = s(t0), d = s(d0);
  if (isErr(t)) return t;
  if (isErr(d)) return d;
  const k = k0 === undefined || k0 === null ? 1 : int(k0, 1);
  if (isErr(k)) return k;
  if (k === 0 || Math.abs(k) > t.length + 1) return E.value;
  const ci = m0 !== undefined && m0 !== null && Boolean(n(m0));
  const matchEnd = e0 !== undefined && e0 !== null && Boolean(n(e0));
  const hay = ci ? t.toLowerCase() : t, needle = ci ? d.toLowerCase() : d;
  const pos: number[] = [];
  if (needle === "") {
    pos.push(k > 0 ? 0 : t.length);
  } else {
    for (let i = hay.indexOf(needle); i >= 0; i = hay.indexOf(needle, i + needle.length)) pos.push(i);
  }
  if (matchEnd) pos.push(t.length);
  let p: number | undefined;
  if (needle === "") p = pos[0];
  else p = k > 0 ? pos[k - 1] : pos[pos.length + k];
  if (p === undefined) return nf !== undefined ? nf : E.na;
  const len = p === t.length && matchEnd ? 0 : needle.length;
  return before ? t.slice(0, p) : t.slice(p + len);
}

