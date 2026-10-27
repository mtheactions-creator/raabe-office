// Date and time functions on Excel's 1900 serials (dates.ts).

import { def } from "./eval";
import { defS, n, values } from "./fn-util";
import { MAX_SERIAL, daysInMonth, hms, serial, weekday0, ymd } from "./dates";
import { E, dateValue, isErr, textToNumber, timeValue, type Any, type V } from "./values";
import type { CellError } from "../model";
import type { Ctx } from "./eval";

/** A date argument: numbers, or text read as a date (#WERT! otherwise); #ZAHL! when out of range. */
function dateArg(v: V | undefined): number | CellError {
  if (v === undefined || v === null) return 0;
  if (typeof v === "string") {
    const x = textToNumber(v);
    return x === null ? E.value : x;
  }
  const x = n(v);
  if (isErr(x)) return x;
  if (x < 0 || x >= MAX_SERIAL + 1) return E.num;
  return x;
}

function datePart(name: string, f: (d: { y: number; m: number; d: number }, x: number) => number): void {
  defS(name, 1, 1, ([a]) => {
    const x = dateArg(a);
    if (isErr(x)) return x;
    if (x < 0 || x >= MAX_SERIAL + 1) return E.num;
    return f(ymd(x), x);
  });
}
datePart("YEAR", (d) => d.y);
datePart("MONTH", (d) => d.m);
datePart("DAY", (d) => d.d);

function timePart(name: string, f: (t: { h: number; mi: number; s: number }) => number): void {
  defS(name, 1, 1, ([a]) => {
    const x = dateArg(a);
    if (isErr(x)) return x;
    return f(hms(x));
  });
}
timePart("HOUR", (t) => t.h);
timePart("MINUTE", (t) => t.mi);
timePart("SECOND", (t) => t.s);

defS("DATE", 3, 3, ([y0, m0, d0]) => {
  let y = n(y0);
  const m = n(m0), d = n(d0);
  if (isErr(y)) return y;
  if (isErr(m)) return m;
  if (isErr(d)) return d;
  // The month is carried into the year first; when the year given was below 1900, 1900 is
  // added after that (=DATUM(-1;13;1) is 1.1.1900, =DATUM(1899;13;1) is 1.1.3800,
  // =DATUM(0;0;31) and =DATUM(1900;-1;1) are #ZAHL!; recorded).
  const ym = Math.trunc(y) * 12 + Math.trunc(m) - 1;
  let yy = Math.floor(ym / 12);
  const mm = ym - yy * 12 + 1;
  if (yy < 0) return E.num;
  if (y < 1900) yy += 1900;
  if (yy < 1900 || yy >= 10000) return E.num;
  const r = serial(yy, mm, Math.trunc(d));
  if (r < 0 || r > MAX_SERIAL) return E.num;
  return r;
});
defS("TIME", 3, 3, ([h0, m0, s0]) => {
  const h = n(h0), m = n(m0), s = n(s0);
  for (const x of [h, m, s]) if (isErr(x)) return x;
  const H = Math.trunc(h as number), M = Math.trunc(m as number), S = Math.trunc(s as number);
  if (H > 32767 || M > 32767 || S > 32767) return E.num;
  const secs = H * 3600 + M * 60 + S;
  if (secs < 0) return E.num;
  return (secs % 86400) / 86400;
});

defS("WEEKDAY", 1, 2, ([a, t0]) => {
  const x = dateArg(a);
  if (isErr(x)) return x;
  const t = t0 === undefined || t0 === null ? 1 : n(t0);
  if (isErr(t)) return t;
  const w = weekday0(x); // 0 = Sunday
  switch (Math.trunc(t)) {
    case 1: case 17: return w + 1;
    case 2: case 11: return ((w + 6) % 7) + 1;
    case 3: return (w + 6) % 7;
    case 12: return ((w + 5) % 7) + 1;
    case 13: return ((w + 4) % 7) + 1;
    case 14: return ((w + 3) % 7) + 1;
    case 15: return ((w + 2) % 7) + 1;
    case 16: return ((w + 1) % 7) + 1;
    default: return E.num;
  }
});

/** First weekday of a week for KALENDERWOCHE's types (0 = Sunday). */
const WEEK_START: Record<number, number> = { 1: 0, 2: 1, 11: 1, 12: 2, 13: 3, 14: 4, 15: 5, 16: 6, 17: 0 };

function isoWeek(x: number): number {
  const { y } = ymd(x);
  // The Thursday of this week decides the year.
  const wd = (weekday0(x) + 6) % 7; // 0 = Monday
  const thu = Math.floor(x) - wd + 3;
  const ty = ymd(thu).y;
  const jan1 = serial(ty, 1, 1);
  void y;
  return Math.floor((thu - jan1) / 7) + 1;
}

defS("WEEKNUM", 1, 2, ([a, t0]) => {
  const x = dateArg(a);
  if (isErr(x)) return x;
  const t = t0 === undefined || t0 === null ? 1 : n(t0);
  if (isErr(t)) return t;
  const tt = Math.trunc(t);
  if (tt === 21) return isoWeek(x);
  const start = WEEK_START[tt];
  if (start === undefined) return E.num;
  const jan1 = serial(ymd(x).y, 1, 1);
  const off = (weekday0(jan1) - start + 7) % 7;
  return Math.floor((Math.floor(x) - jan1 + off) / 7) + 1;
});
defS("ISOWEEKNUM", 1, 1, ([a]) => {
  const x = dateArg(a);
  if (isErr(x)) return x;
  return isoWeek(x);
});

/** Days of a month in the real calendar (1900 is no leap year here). */
function realDays(y: number, m: number): number {
  return y === 1900 && m === 2 ? 28 : daysInMonth(y, m);
}

function addMonths(x: number, k: number): { y: number; m: number; d: number } {
  const { y, m, d } = ymd(x);
  const t = y * 12 + (m - 1) + k;
  const yy = Math.floor(t / 12), mm = t - yy * 12 + 1;
  // EDATUM and MONATSENDE know no 29 February 1900 (=EDATUM(60;0) is 59; recorded).
  return { y: yy, m: mm, d: Math.min(d, realDays(yy, mm)) };
}

defS("EDATE", 2, 2, ([a, k0]) => {
  const x = dateArg(a);
  if (isErr(x)) return x;
  const k = n(k0);
  if (isErr(k)) return k;
  const r = addMonths(x, Math.trunc(k));
  if (r.y < 1900 || r.y > 9999) return E.num;
  return serial(r.y, r.m, r.d);
});
defS("EOMONTH", 2, 2, ([a, k0]) => {
  const x = dateArg(a);
  if (isErr(x)) return x;
  const k = n(k0);
  if (isErr(k)) return k;
  const r = addMonths(x, Math.trunc(k));
  if (r.y < 1900 || r.y > 9999) return E.num;
  // MONATSENDE ends February 1900 on the 28th (recorded), unlike the serials.
  return serial(r.y, r.m, realDays(r.y, r.m));
});

defS("DATEDIF", 3, 3, ([a, b0, u0]) => {
  const s0 = dateArg(a), e0 = dateArg(b0);
  if (isErr(s0)) return s0;
  if (isErr(e0)) return e0;
  if (typeof u0 !== "string") return isErr(u0) ? u0 : E.num;
  const s = Math.floor(s0), e = Math.floor(e0);
  if (s > e) return E.num;
  const A = ymd(s), B = ymd(e);
  const months = (B.y - A.y) * 12 + (B.m - A.m) - (B.d < A.d ? 1 : 0);
  switch (u0.toUpperCase()) {
    case "D": return e - s;
    case "M": return months;
    case "Y": return Math.floor(months / 12);
    case "YM": return months % 12;
    case "MD": {
      if (B.d >= A.d) return B.d - A.d;
      // Days from the same day in the month before the end date (Excel's own arithmetic).
      const pm = B.m === 1 ? 12 : B.m - 1, py = B.m === 1 ? B.y - 1 : B.y;
      return daysInMonth(py, pm) - A.d + B.d;
    }
    case "YD": {
      // Excel moves the end date into the start date's year (or the next one).
      const y = B.m > A.m || (B.m === A.m && B.d >= A.d) ? A.y : A.y + 1;
      return serial(y, B.m, B.d) - s;
    }
    default: return E.num;
  }
});

defS("DAYS", 2, 2, ([a, b0]) => {
  const e = dateArg(a), s = dateArg(b0);
  if (isErr(e)) return e;
  if (isErr(s)) return s;
  return Math.floor(e) - Math.floor(s);
});
defS("DAYS360", 2, 3, ([a, b0, m0]) => {
  const s = dateArg(a), e = dateArg(b0);
  if (isErr(s)) return s;
  if (isErr(e)) return e;
  const eu = m0 !== undefined && m0 !== null && Boolean(n(m0));
  const A = ymd(s), B = ymd(e);
  let d1 = A.d, d2 = B.d;
  if (eu) {
    if (d1 === 31) d1 = 30;
    if (d2 === 31) d2 = 30;
  } else {
    // Here February 1900 has its 29 days (TAGE360(59;159) is 99), unlike in BRTEILJAHRE.
    const lastFeb = (d: { y: number; m: number; d: number }) => d.m === 2 && d.d === daysInMonth(d.y, 2);
    if (d1 === 31 || lastFeb(A)) d1 = 30;
    if (d2 === 31 && d1 >= 30) d2 = 30;
  }
  return (B.y - A.y) * 360 + (B.m - A.m) * 30 + (d2 - d1);
});

/** The holidays of an argument as a set of day serials. */
function holidays(a: Any | undefined, ctx: Ctx): Set<number> | CellError {
  const out = new Set<number>();
  if (a === undefined || a === null) return out;
  for (const v of values(a, ctx)) {
    if (v === null) continue;
    const x = dateArg(v);
    if (isErr(x)) return x;
    out.add(Math.floor(x));
  }
  return out;
}

/** Weekend days (0 = Sunday) of NETTOARBEITSTAGE.INTL's Wochenende argument. */
function weekend(v: V | undefined): boolean[] | CellError {
  const we = [true, false, false, false, false, false, true];
  if (v === undefined || v === null) return we;
  if (typeof v === "string") {
    if (!/^[01]{7}$/.test(v) || v === "1111111") return E.value;
    // Monday first in the string.
    return [v[6] === "1", v[0] === "1", v[1] === "1", v[2] === "1", v[3] === "1", v[4] === "1", v[5] === "1"];
  }
  const k = n(v);
  if (isErr(k)) return k;
  const code = Math.trunc(k);
  const out = [false, false, false, false, false, false, false];
  if (code >= 1 && code <= 7) {
    // 1: Sat+Sun, 2: Sun+Mon, … 7: Fri+Sat.
    out[(code + 5) % 7] = true;
    out[(code + 6) % 7] = true;
    return out;
  }
  if (code >= 11 && code <= 17) {
    // 11: Sunday only … 17: Saturday only.
    out[code - 11] = true;
    return out;
  }
  return E.num;
}

function networkdays(sd: number, ed: number, we: boolean[], hol: Set<number>): number {
  const sign = ed >= sd ? 1 : -1;
  const lo = Math.min(sd, ed), hi = Math.max(sd, ed);
  let c = 0;
  for (let d = lo; d <= hi; d++) if (!we[weekday0(d)] && !hol.has(d)) c++;
  return sign * c;
}

function workday(sd: number, k: number, we: boolean[], hol: Set<number>): number | CellError {
  if (we.every((x) => x)) return E.value;
  let d = sd;
  const step = k >= 0 ? 1 : -1;
  let left = Math.abs(k);
  while (left > 0) {
    d += step;
    if (d < 0 || d > MAX_SERIAL) return E.num;
    if (!we[weekday0(d)] && !hol.has(d)) left--;
  }
  return d;
}

def("NETWORKDAYS", 2, 3, ([a, b0, h], ctx) => {
  const s = dateArg(firstVal(a, ctx)), e = dateArg(firstVal(b0, ctx));
  if (isErr(s)) return s;
  if (isErr(e)) return e;
  const hol = holidays(h, ctx);
  if (isErr(hol)) return hol;
  return networkdays(Math.floor(s), Math.floor(e), weekend(undefined) as boolean[], hol);
});
def("NETWORKDAYS.INTL", 2, 4, ([a, b0, w, h], ctx) => {
  const s = dateArg(firstVal(a, ctx)), e = dateArg(firstVal(b0, ctx));
  if (isErr(s)) return s;
  if (isErr(e)) return e;
  const we = weekend(w === undefined ? undefined : firstVal(w, ctx));
  if (isErr(we)) return we;
  const hol = holidays(h, ctx);
  if (isErr(hol)) return hol;
  return networkdays(Math.floor(s), Math.floor(e), we, hol);
});
def("WORKDAY", 2, 3, ([a, k0, h], ctx) => {
  const s = dateArg(firstVal(a, ctx));
  if (isErr(s)) return s;
  const k = n(firstVal(k0, ctx));
  if (isErr(k)) return k;
  const hol = holidays(h, ctx);
  if (isErr(hol)) return hol;
  return workday(Math.floor(s), Math.trunc(k), weekend(undefined) as boolean[], hol);
});
def("WORKDAY.INTL", 2, 4, ([a, k0, w, h], ctx) => {
  const s = dateArg(firstVal(a, ctx));
  if (isErr(s)) return s;
  const k = n(firstVal(k0, ctx));
  if (isErr(k)) return k;
  const we = weekend(w === undefined ? undefined : firstVal(w, ctx));
  if (isErr(we)) return we;
  const hol = holidays(h, ctx);
  if (isErr(hol)) return hol;
  return workday(Math.floor(s), Math.trunc(k), we, hol);
});

function firstVal(a: Any, ctx: Ctx): V {
  const v = values(a, ctx);
  return v.length === 1 ? v[0] : v.length ? E.value : null;
}

defS("DATEVALUE", 1, 1, ([a]) => {
  if (typeof a !== "string") return isErr(a) ? a : E.value;
  const t = a.trim();
  let d = dateValue(t);
  if (d === null) {
    const sp = t.lastIndexOf(" ");
    if (sp > 0 && timeValue(t.slice(sp + 1)) !== null) d = dateValue(t.slice(0, sp));
  }
  return d === null ? E.value : d;
});
defS("TIMEVALUE", 1, 1, ([a]) => {
  if (typeof a !== "string") return isErr(a) ? a : E.value;
  const x = textToNumber(a);
  if (x === null) return E.value;
  return x - Math.floor(x);
});

def("NOW", 0, 0, (_a, ctx) => ctx.host.now(), { volatile: true });
def("TODAY", 0, 0, (_a, ctx) => Math.floor(ctx.host.now()), { volatile: true });

defS("YEARFRAC", 2, 3, ([a, b0, c0]) => {
  let s = dateArg(a), e = dateArg(b0);
  if (isErr(s)) return s;
  if (isErr(e)) return e;
  const basis = c0 === undefined || c0 === null ? 0 : n(c0);
  if (isErr(basis)) return basis;
  s = Math.floor(s);
  e = Math.floor(e);
  if (s > e) [s, e] = [e, s];
  const A = ymd(s), B = ymd(e);
  switch (Math.trunc(basis)) {
    case 0: {
      let d1 = A.d, d2 = B.d;
      if (d1 === 31 && d2 === 31) { d1 = 30; d2 = 30; }
      else if (d1 === 31) d1 = 30;
      else if (d1 === 30 && d2 === 31) d2 = 30;
      const lastFeb = (d: { y: number; m: number; d: number }) => d.m === 2 && d.d === realDays(d.y, 2);
      if (lastFeb(A) && lastFeb(B)) { d1 = 30; d2 = 30; }
      else if (lastFeb(A)) d1 = 30;
      return ((B.y - A.y) * 360 + (B.m - A.m) * 30 + (d2 - d1)) / 360;
    }
    case 1: {
      const leap = (y: number) => (y % 4 === 0 && (y % 100 !== 0 || y % 400 === 0));
      if (A.y === B.y || (B.y === A.y + 1 && (A.m > B.m || (A.m === B.m && A.d >= B.d)))) {
        let den = 365;
        if (A.y === B.y) den = leap(A.y) ? 366 : 365;
        else {
          const feb29 = (y: number) => leap(y) && serial(y, 2, 29);
          const f1 = feb29(A.y), f2 = feb29(B.y);
          if ((f1 && s <= f1) || (f2 && e >= f2)) den = 366;
        }
        return (e - s) / den;
      }
      let days = 0;
      for (let y = A.y; y <= B.y; y++) days += leap(y) ? 366 : 365;
      return (e - s) / (days / (B.y - A.y + 1));
    }
    case 2: return (e - s) / 360;
    case 3: return (e - s) / 365;
    case 4: {
      const d1 = Math.min(A.d, 30), d2 = Math.min(B.d, 30);
      return ((B.y - A.y) * 360 + (B.m - A.m) * 30 + (d2 - d1)) / 360;
    }
    default: return E.num;
  }
});
