// PROGNOSE.ETS, PROGNOSE.ETS.KONFINT, PROGNOSE.ETS.SAISONALITÄT, PROGNOSE.ETS.STAT (FORECAST.ETS …, D-092):
// argument rules as recorded in Excel (probe_excel_ets.py): Saisonalität 0 (none), 1 (find it) or up to
// 8760, Datenvervollständigung 0 (zeros) or 1 (interpolation), Aggregation 1 to 7, Konfidenzniveau between 0
// and 1, a target before the timeline's start and a timeline without a constant step give #ZAHL!, ranges of
// different sizes #NV. A target inside the timeline gets the forecast formula from the end states (h <= 0, D-124), one
// between two steps has the trend's share of the step and the season between the two steps' (recorded: half a step after
// the end). The model itself is ets.ts.

import { def, deref, type Ctx } from "./eval";
import { values } from "./fn-util";
import { normInv } from "./fn-stats";
import { Arr, E, isErr, num, toNum, type Any, type V } from "./values";
import type { CellError } from "../model";
import { etsConfint, etsForecast, etsStat, fitEts, gridPosition, regularSeries, type EtsAggregation, type EtsModel, type EtsSeries } from "../ets";

interface Fitted { s: EtsSeries; md: EtsModel }

const cache = new Map<string, Fitted>();

/** The series and its model for the arguments (kept while the same numbers come again). */
function fitted(vals: Any, tl: Any, ctx: Ctx, season: number, fill: number, agg: number): Fitted | CellError {
  const v = values(vals, ctx), t = values(tl, ctx);
  if (v.length !== t.length) return E.na;
  const ts: number[] = [], ys: number[] = [];
  for (let i = 0; i < v.length; i++) {
    const a = t[i], b = v[i];
    if (isErr(a)) return a;
    if (isErr(b)) return b;
    if (typeof a !== "number") return E.value;
    if (typeof b !== "number") continue;
    ts.push(a); ys.push(b);
  }
  if (ys.length < 2) return E.num;
  const key = JSON.stringify([ts, ys, season, fill, agg]);
  const hit = cache.get(key);
  if (hit) return hit;
  const s = regularSeries(ts, ys, fill === 0, agg as EtsAggregation);
  if (!s) return E.num;
  const md = fitEts(s.y, season);
  if (cache.size > 64) cache.clear();
  const f = { s, md };
  cache.set(key, f);
  return f;
}

/** An optional whole-number argument with its allowed range (undefined/empty: the default). */
function opt(v: V | undefined, d: number, ok: (x: number) => boolean, trunc = true): number | CellError {
  if (v === undefined || v === null) return d;
  const x = toNum(v);
  if (isErr(x)) return x;
  const y = trunc ? Math.trunc(x) : x;
  return ok(y) ? y : E.num;
}

/** Applies `f` to each value of the target argument (a range gives an array of results). */
function eachTarget(a: Any, ctx: Ctx, f: (x: number) => number | CellError): V | Arr {
  const t = deref(a, ctx);
  const one = (v: V): V => {
    if (isErr(v)) return v;
    const x = toNum(v);
    if (isErr(x)) return x;
    return f(x);
  };
  if (t instanceof Arr) return new Arr(t.rows, t.cols, t.d.map(one));
  return one(t);
}

function scalarArg(a: Any | undefined, ctx: Ctx): V | undefined {
  if (a === undefined) return undefined;
  const v = deref(a, ctx);
  return v instanceof Arr ? v.d[0] : v;
}

const seasonOk = (x: number) => x >= 0 && x <= 8760;
const fillOk = (x: number) => x === 0 || x === 1;
const aggOk = (x: number) => x >= 1 && x <= 7;

/** The value at a target: the series inside the timeline, the forecast after it. */
function valueAt(f: Fitted, target: number): number | CellError {
  const pos = gridPosition(f.s, target);
  const last = f.s.y.length - 1;
  if (pos < -1e-9) return E.num;
  // Inside the timeline too, Excel gives the forecast formula from the end states: level + h trend + season with
  // h <= 0 (recorded on every point of 99 fits, D-124).
  return etsForecast(f.md, pos - last);
}

def("FORECAST.ETS", 3, 6, (args, ctx) => {
  const season = opt(scalarArg(args[3], ctx), 1, seasonOk), fill = opt(scalarArg(args[4], ctx), 1, fillOk), agg = opt(scalarArg(args[5], ctx), 1, aggOk);
  for (const x of [season, fill, agg]) if (isErr(x)) return x;
  const f = fitted(args[1], args[2], ctx, season as number, fill as number, agg as number);
  if (isErr(f)) return f;
  return eachTarget(args[0], ctx, (x) => { const r = valueAt(f, x); return isErr(r) ? r : num(r); });
});

def("FORECAST.ETS.CONFINT", 3, 7, (args, ctx) => {
  const level = opt(scalarArg(args[3], ctx), 0.95, (x) => x > 0 && x < 1, false);
  const season = opt(scalarArg(args[4], ctx), 1, seasonOk), fill = opt(scalarArg(args[5], ctx), 1, fillOk), agg = opt(scalarArg(args[6], ctx), 1, aggOk);
  for (const x of [level, season, fill, agg]) if (isErr(x)) return x;
  const f = fitted(args[1], args[2], ctx, season as number, fill as number, agg as number);
  if (isErr(f)) return f;
  const z = normInv(1 - (1 - (level as number)) / 2);
  return eachTarget(args[0], ctx, (x) => {
    const pos = gridPosition(f.s, x);
    if (pos < -1e-9) return E.num;
    const h = Math.max(1, pos - (f.s.y.length - 1));
    return num(etsConfint(f.md, h, z));
  });
});

def("FORECAST.ETS.SEASONALITY", 2, 4, (args, ctx) => {
  const fill = opt(scalarArg(args[2], ctx), 1, fillOk), agg = opt(scalarArg(args[3], ctx), 1, aggOk);
  for (const x of [fill, agg]) if (isErr(x)) return x;
  const f = fitted(args[0], args[1], ctx, 1, fill as number, agg as number);
  if (isErr(f)) return f;
  return f.md.m;
});

def("FORECAST.ETS.STAT", 3, 6, (args, ctx) => {
  const kind = opt(scalarArg(args[2], ctx), 1, (x) => x >= 1 && x <= 8);
  const season = opt(scalarArg(args[3], ctx), 1, seasonOk), fill = opt(scalarArg(args[4], ctx), 1, fillOk), agg = opt(scalarArg(args[5], ctx), 1, aggOk);
  for (const x of [kind, season, fill, agg]) if (isErr(x)) return x;
  const f = fitted(args[0], args[1], ctx, season as number, fill as number, agg as number);
  if (isErr(f)) return f;
  return num(etsStat(f.md, f.s, f.s.y, kind as number));
});
