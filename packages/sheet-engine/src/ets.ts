// Exponential triple smoothing for PROGNOSE.ETS and its companions (FORECAST.ETS, .CONFINT, .SEASONALITY,
// .STAT; D-092): the AAA model (additive error, trend and season) Excel names, on a timeline with a
// constant step (or calendar months), missing points filled by interpolation or zeros, duplicate time
// stamps aggregated.
//
// Excel's own parameter search is not documented; this one was built against Excel's results
// (probe_excel_ets.py, test/fixtures/excel-ets.json): the season and the start of level and trend come from
// a decomposition of the whole series (centred moving average, then a straight line through the
// seasonally adjusted values), alpha, beta and gamma are chosen to minimise the squared one-step errors on a
// grid refined down to 0.001, the season length is the strongest peak of the autocorrelation of the
// detrended series (when it reaches 0.4), the confidence interval is z * RMSE * sqrt(1 + sum c_j^2) with
// c_j = alpha (1 + j beta) + gamma [j a multiple of m]. Measured: forecasts within 2.2 % of Excel's, the
// season length as Excel's in all recorded series; Excel's alpha, beta and gamma themselves are not matched.

export type EtsAggregation = 1 | 2 | 3 | 4 | 5 | 6 | 7;

export interface EtsSeries {
  /** Values on the regular grid (index 0 = first time point). */
  y: number[];
  /** First time point and step (days or plain numbers); `months`: the grid steps by calendar months. */
  t0: number;
  step: number;
  months?: boolean;
}

export interface EtsModel {
  m: number;
  alpha: number;
  beta: number;
  gamma: number;
  level: number;
  trend: number;
  season: number[];
  /** One-step errors over the series. */
  errors: number[];
  n: number;
}

const MAX_SEASON = 8760;

/** The date serial number of a calendar month index (months since 1900-01, day of month kept). */
function excelDate(y: number, m: number, d: number): number {
  const ms = Date.UTC(y, m, d) - Date.UTC(1899, 11, 30);
  return Math.round(ms / 86400000);
}
function ymd(serial: number): [number, number, number] {
  const dt = new Date(Date.UTC(1899, 11, 30) + Math.round(serial) * 86400000);
  return [dt.getUTCFullYear(), dt.getUTCMonth(), dt.getUTCDate()];
}
const monthIndex = (serial: number) => { const [y, m] = ymd(serial); return y * 12 + m; };

/** The date `k` months after `serial` (day of month kept, as EDATUM). */
export function etsAddMonths(serial: number, k: number): number {
  const [y, m, d] = ymd(serial);
  const total = y * 12 + m + k;
  const ny = Math.floor(total / 12), nm = total - ny * 12;
  const last = new Date(Date.UTC(ny, nm + 1, 0)).getUTCDate();
  return excelDate(ny, nm, Math.min(d, last));
}

function aggregate(xs: number[], how: EtsAggregation): number {
  const s = [...xs].sort((a, b) => a - b);
  switch (how) {
    case 2: case 3: return xs.length;
    case 4: return s[s.length - 1];
    case 5: { const h = s.length >> 1; return s.length % 2 ? s[h] : (s[h - 1] + s[h]) / 2; }
    case 6: return s[0];
    case 7: return xs.reduce((a, b) => a + b, 0);
    default: return xs.reduce((a, b) => a + b, 0) / xs.length;
  }
}

/** The series on a regular grid; null when the timeline has no constant step (Excel: #ZAHL!). */
export function regularSeries(timeline: number[], values: number[], fillZero = false, how: EtsAggregation = 1): EtsSeries | null {
  const groups = new Map<number, number[]>();
  timeline.forEach((t, i) => { const g = groups.get(t); if (g) g.push(values[i]); else groups.set(t, [values[i]]); });
  const ts = [...groups.keys()].sort((a, b) => a - b);
  if (ts.length < 2) return null;
  const vs = ts.map((t) => aggregate(groups.get(t)!, how));
  const diffs = ts.slice(1).map((t, i) => t - ts[i]);
  // Calendar months: every step 28 to 31 days and the same day of the month (or the month's end).
  const months = diffs.every((d) => d >= 28 && d <= 31 * 3) && ts.every((t) => Number.isInteger(t)) && diffs.some((d) => d !== diffs[0])
    && ts.every((t) => { const [, , d] = ymd(t); const [, , d0] = ymd(ts[0]); return Math.abs(d - d0) <= 3; });
  let idx: number[];
  let step: number;
  if (months) {
    idx = ts.map((t) => monthIndex(t) - monthIndex(ts[0]));
    step = 1;
  } else {
    step = Math.min(...diffs);
    if (!(step > 0)) return null;
    idx = ts.map((t) => (t - ts[0]) / step);
    if (idx.some((k) => Math.abs(k - Math.round(k)) > 1e-6)) return null;
    idx = idx.map(Math.round);
  }
  const n = idx[idx.length - 1] + 1;
  // Excel refuses timelines with too many holes (more than 30 % of the points).
  if (n - ts.length > Math.max(1, Math.floor(0.3 * n))) return null;
  const y = new Array<number>(n).fill(NaN);
  idx.forEach((k, i) => { y[k] = vs[i]; });
  for (let k = 0; k < n; k++) {
    if (!Number.isNaN(y[k])) continue;
    if (fillZero) { y[k] = 0; continue; }
    let a = k - 1, b = k + 1;
    while (Number.isNaN(y[b])) b++;
    y[k] = y[a] + ((y[b] - y[a]) * (k - a)) / (b - a);
    a = k;
  }
  return { y, t0: ts[0], step, months };
}

/** The position of a target date on the series' grid (fractional between points). */
export function gridPosition(s: EtsSeries, target: number): number {
  if (!s.months) return (target - s.t0) / s.step;
  const [y, m, d] = ymd(target);
  const [y0, m0, d0] = ymd(s.t0);
  const whole = y * 12 + m - (y0 * 12 + m0);
  const base = etsAddMonths(s.t0, whole);
  const next = etsAddMonths(s.t0, whole + 1);
  void d; void d0;
  return whole + (target - base) / Math.max(1, next - base);
}

function linreg(ys: number[]): [number, number] {
  const n = ys.length;
  let sx = 0, sy = 0, sxx = 0, sxy = 0;
  ys.forEach((v, i) => { sx += i; sy += v; sxx += i * i; sxy += i * v; });
  const d = n * sxx - sx * sx;
  const b = d ? (n * sxy - sx * sy) / d : 0;
  return [(sy - b * sx) / n, b];
}

/** Season length found in the data: the strongest autocorrelation peak of the detrended series (0: none). */
export function detectSeason(y: number[]): number {
  const n = y.length;
  if (n < 4) return 0;
  const [a, b] = linreg(y);
  const r = y.map((v, i) => v - (a + b * i));
  const varr = r.reduce((s, v) => s + v * v, 0);
  if (varr < 1e-12 * Math.max(1, y.reduce((s, v) => s + v * v, 0))) return 0;
  const acf = (k: number) => { let s = 0; for (let i = 0; i + k < n; i++) s += r[i] * r[i + k]; return s / varr; };
  let best = 0, bv = 0.4;
  let prev = acf(1), cur = acf(2);
  for (let k = 2; k <= Math.floor(n / 2); k++) {
    const next = k + 1 < n ? acf(k + 1) : -Infinity;
    if (cur > prev && cur >= next && cur > bv) { best = k; bv = cur; }
    prev = cur; cur = next;
  }
  return best;
}

/** Start states: season from a centred moving average over the whole series, level and trend from the adjusted values. */
function startStates(y: number[], m: number): { level: number; trend: number; season: number[] } {
  if (m <= 1) {
    const [a, b] = linreg(y);
    return { level: a - b, trend: b, season: [0] };
  }
  const n = y.length;
  const half = Math.floor(m / 2);
  const det: (number | null)[] = new Array(n).fill(null);
  for (let i = half; i < n - half; i++) {
    let s = 0;
    if (m % 2 === 0) { s = 0.5 * y[i - half] + 0.5 * y[i + half]; for (let k = i - half + 1; k < i + half; k++) s += y[k]; }
    else for (let k = i - half; k <= i + half; k++) s += y[k];
    det[i] = y[i] - s / m;
  }
  let season = Array.from({ length: m }, (_, j) => {
    const xs: number[] = [];
    for (let i = j; i < n; i += m) if (det[i] !== null) xs.push(det[i]!);
    return xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : 0;
  });
  const mean = season.reduce((a, b) => a + b, 0) / m;
  season = season.map((v) => v - mean);
  const [a, b] = linreg(y.map((v, i) => v - season[i % m]));
  return { level: a - b, trend: b, season };
}

/** Runs the smoothing (Holt-Winters form) from the start states; returns the end states and one-step errors. */
function smooth(y: number[], m: number, alpha: number, beta: number, gamma: number, st: { level: number; trend: number; season: number[] }) {
  let l = st.level, t = st.trend;
  const s = [...st.season];
  const errors: number[] = [];
  for (let i = 0; i < y.length; i++) {
    const si = m > 1 ? s[i % m] : 0;
    errors.push(y[i] - (l + t + si));
    const ln = alpha * (y[i] - si) + (1 - alpha) * (l + t);
    t = beta * (ln - l) + (1 - beta) * t;
    if (m > 1) s[i % m] = gamma * (y[i] - ln) + (1 - gamma) * si;
    l = ln;
  }
  return { level: l, trend: t, season: s, errors };
}

/** Fits the model: `season` 1 = find it, 0 = none, else the length. */
export function fitEts(y: number[], season: number): EtsModel {
  let m = season === 1 ? detectSeason(y) : season;
  if (m < 2 || m > MAX_SEASON || m > y.length) m = 0;
  const st = startStates(y, m);
  const sse = (p: number[]) => smooth(y, m, p[0], p[1], p[2], st).errors.reduce((a, e) => a + e * e, 0);
  const grid = [0.001, 0.126, 0.251, 0.376, 0.501, 0.626, 0.751, 0.876];
  let best: { v: number; p: number[] } = { v: Infinity, p: [0.001, 0.001, 0.001] };
  for (const a of grid) for (const b of grid) for (const g of m > 1 ? grid : [0]) {
    const v = sse([a, b, g]);
    if (v < best.v - 1e-12) best = { v, p: [a, b, g] };
  }
  for (let step = 0.0625; step >= 0.0009; step /= 2) {
    let improved = true;
    while (improved) {
      improved = false;
      for (let k = 0; k < (m > 1 ? 3 : 2); k++) for (const d of [step, -step]) {
        const q = [...best.p];
        q[k] = Math.min(0.999, Math.max(0.001, Math.round((q[k] + d) * 1e6) / 1e6));
        const v = sse(q);
        if (v < best.v - 1e-12) { best = { v, p: q }; improved = true; }
      }
    }
  }
  const [alpha, beta, gamma] = best.p;
  const r = smooth(y, m, alpha, beta, m > 1 ? gamma : 0, st);
  return { m, alpha, beta, gamma: m > 1 ? gamma : 0, level: r.level, trend: r.trend, season: r.season, errors: r.errors, n: y.length };
}

/** The forecast `h` steps after the last point (fractional h: the season between the two steps' seasons). */
export function etsForecast(md: EtsModel, h: number): number {
  const at = (k: number) => (md.m > 1 ? md.season[(((md.n + k - 1) % md.m) + md.m) % md.m] : 0);
  const k = Math.floor(h + 1e-9);
  const s = Math.abs(h - k) < 1e-9 ? at(k) : at(k) + (at(k + 1) - at(k)) * (h - k);
  return md.level + h * md.trend + s;
}

export function etsRmse(md: EtsModel): number {
  return Math.sqrt(md.errors.reduce((a, e) => a + e * e, 0) / md.errors.length);
}

/** Half the width of the confidence interval `h` steps ahead (z from the normal distribution). */
export function etsConfint(md: EtsModel, h: number, z: number): number {
  let v = 1;
  for (let j = 1; j < Math.ceil(h - 1e-9); j++) {
    const c = md.alpha * (1 + j * md.beta) + (md.m > 1 && j % md.m === 0 ? md.gamma : 0);
    v += c * c;
  }
  return z * etsRmse(md) * Math.sqrt(v);
}

/** PROGNOSE.ETS.STAT: 1 alpha, 2 beta, 3 gamma, 4 MASE, 5 SMAPE, 6 MAE, 7 RMSE, 8 step size. */
export function etsStat(md: EtsModel, s: EtsSeries, y: number[], kind: number): number {
  const e = md.errors;
  const n = e.length;
  switch (kind) {
    case 1: return md.alpha;
    case 2: return md.beta;
    case 3: return md.m > 1 ? md.gamma : 2.220446049250313e-16;
    case 4: {
      let naive = 0;
      for (let i = 1; i < y.length; i++) naive += Math.abs(y[i] - y[i - 1]);
      naive /= Math.max(1, y.length - 1);
      return naive ? e.reduce((a, x) => a + Math.abs(x), 0) / n / naive : 0;
    }
    case 5: return e.reduce((a, x, i) => { const f = y[i] - x; const d = Math.abs(y[i]) + Math.abs(f); return a + (d ? (2 * Math.abs(x)) / d : 0); }, 0) / n / 2;
    case 6: return e.reduce((a, x) => a + Math.abs(x), 0) / n;
    case 7: return etsRmse(md);
    default: return s.months ? 1 : s.step;
  }
}
