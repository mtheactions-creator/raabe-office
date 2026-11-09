"""A broad measurement set for Excel's exponential smoothing (D-124): many series (trend, seasons of several
lengths, noise, random walks, level shifts, short series, missing points with both kinds of Datenvervollständigung,
duplicate time stamps with every Aggregation, calendar months) and, for each, everything Excel lets us read of its
model: PROGNOSE.ETS.STAT 1-8, the season found, the one-step values inside the timeline (target = each time
point), forecasts and confidence intervals 1-24 steps ahead, and the same without a season (Saisonalität 0).

Our own invisible Excel, COM only, English formula names through Formula2. Written to
packages/sheet-engine/test/fixtures/excel-ets-broad.json (ets-broad.test.ts, fit_ets.py).

    .venv/Scripts/python.exe probe_excel_ets_broad.py
"""
from __future__ import annotations

import json
import math
from pathlib import Path

from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
OUT = HERE.parents[1] / "packages" / "sheet-engine" / "test" / "fixtures" / "excel-ets-broad.json"
ERRORS = {-2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?", -2146826288: "#NULL!",
          -2146826252: "#NUM!", -2146826265: "#REF!", -2146826273: "#VALUE!"}


def lcg(seed: int):
    s = seed
    while True:
        s = (s * 1103515245 + 12345) % 2 ** 31
        yield s / 2 ** 31


def gauss(r):
    # Box-Muller from the LCG (deterministic).
    u1 = max(1e-9, next(r))
    u2 = next(r)
    return math.sqrt(-2 * math.log(u1)) * math.cos(2 * math.pi * u2)


def make() -> dict:
    out: dict = {}
    seed = 100
    # Seasonal with trend and noise: m, n, trend, amplitude, noise.
    for m, n, tr, amp, noise in [(4, 16, 1.0, 8, 1), (4, 24, 0.5, 5, 2), (4, 40, -0.3, 10, 3), (6, 30, 2, 12, 2),
                                 (7, 35, 0.2, 6, 1), (7, 56, 1, 20, 5), (12, 36, 2, 15, 3), (12, 48, 0.5, 30, 6),
                                 (12, 60, -1, 10, 2), (5, 25, 3, 4, 0.5), (3, 18, 1, 5, 1), (2, 20, 0.5, 3, 0.5),
                                 (12, 30, 1, 20, 2), (24, 72, 0.1, 10, 1)]:
        seed += 1
        r = lcg(seed)
        shape = [math.sin(2 * math.pi * k / m) + 0.4 * math.cos(4 * math.pi * k / m) for k in range(m)]
        v = [round(100 + tr * i + amp * shape[i % m] + noise * gauss(r), 3) for i in range(n)]
        out[f"seas_m{m}_n{n}_{seed}"] = {"t": list(range(1, n + 1)), "v": v}
    # Exact seasons (no noise): the start states show themselves exactly.
    for m, n, tr in [(3, 12, 1), (4, 12, 2), (4, 20, 0), (5, 15, -1), (6, 24, 0.5), (12, 36, 1)]:
        shape = [((k * 7) % m) - (m - 1) / 2 for k in range(m)]
        v = [round(50 + tr * i + 3 * shape[i % m], 6) for i in range(n)]
        out[f"exact_m{m}_n{n}"] = {"t": list(range(1, n + 1)), "v": v}
    # Trend with noise, no season.
    for n, tr, noise in [(10, 1, 0.5), (20, 2, 3), (30, -1, 1), (50, 0.3, 2), (15, 5, 10), (8, 1, 0.2)]:
        seed += 1
        r = lcg(seed)
        out[f"trend_n{n}_{seed}"] = {"t": list(range(1, n + 1)), "v": [round(20 + tr * i + noise * gauss(r), 3) for i in range(n)]}
    # Random walks and level shifts.
    for n in (20, 40, 60):
        seed += 1
        r = lcg(seed)
        x, v = 50.0, []
        for _ in range(n):
            x += 2 * gauss(r)
            v.append(round(x, 3))
        out[f"walk_n{n}_{seed}"] = {"t": list(range(1, n + 1)), "v": v}
    for n in (20, 36):
        seed += 1
        r = lcg(seed)
        out[f"shift_n{n}_{seed}"] = {"t": list(range(1, n + 1)), "v": [round((10 if i < n // 2 else 25) + gauss(r), 3) for i in range(n)]}
    # Noise only, quadratic growth, positive multiplicative season.
    for n in (12, 25, 45):
        seed += 1
        r = lcg(seed)
        out[f"noise_n{n}_{seed}"] = {"t": list(range(1, n + 1)), "v": [round(30 + 5 * gauss(r), 3) for _ in range(n)]}
    out["quad_n20"] = {"t": list(range(1, 21)), "v": [round(5 + 0.3 * i * i, 3) for i in range(20)]}
    seed += 1
    r = lcg(seed)
    out["mult_m12_n48"] = {"t": list(range(1, 49)), "v": [round((100 + 3 * i) * (1 + 0.3 * math.sin(2 * math.pi * i / 12)) + 2 * gauss(r), 3) for i in range(48)]}
    # Short series.
    for n in (4, 5, 6, 7):
        seed += 1
        r = lcg(seed)
        out[f"short_n{n}_{seed}"] = {"t": list(range(1, n + 1)), "v": [round(10 + i + gauss(r), 3) for i in range(n)]}
    # Steps other than 1, weekly dates, calendar months (first and last day of the month).
    seed += 1
    r = lcg(seed)
    out["step5_m4"] = {"t": [5 * i for i in range(1, 25)], "v": [round(40 + 0.5 * i + [6, -2, -7, 3][i % 4] + gauss(r), 3) for i in range(24)]}
    import datetime as dt

    def serial(d: dt.date) -> int:
        return (d - dt.date(1899, 12, 30)).days

    months = [serial(dt.date(2021 + (k // 12), 1 + k % 12, 1)) for k in range(36)]
    seed += 1
    r = lcg(seed)
    out["months_first"] = {"t": months, "v": [round(200 + 2 * k + 25 * math.sin(2 * math.pi * k / 12) + 4 * gauss(r), 3) for k in range(36)]}
    ends = []
    for k in range(30):
        y, mo = 2022 + ((k + 1) // 12), 1 + (k + 1) % 12
        ends.append(serial(dt.date(y, mo, 1) - dt.timedelta(days=1)))
    seed += 1
    r = lcg(seed)
    out["months_last"] = {"t": ends, "v": [round(80 + k + 10 * math.cos(2 * math.pi * k / 6) + 2 * gauss(r), 3) for k in range(30)]}
    # Missing points (interpolated or zero) and duplicates (every aggregation).
    seed += 1
    r = lcg(seed)
    base = [round(60 + 0.8 * i + [5, 1, -4, -2][i % 4] + gauss(r), 3) for i in range(32)]
    keep = [i for i in range(32) if i not in (5, 6, 17, 25)]
    out["gaps_m4"] = {"t": [i + 1 for i in keep], "v": [base[i] for i in keep]}
    seed += 1
    r = lcg(seed)
    t, v = [], []
    for i in range(30):
        reps = 1 + (i % 5 == 2) + (i % 7 == 3)
        for k in range(reps):
            t.append(i + 1)
            v.append(round(15 + 0.5 * i + 4 * math.sin(2 * math.pi * i / 6) + gauss(r) + k, 3))
    out["dups_m6"] = {"t": t, "v": v}
    # Unsorted timeline.
    seed += 1
    r = lcg(seed)
    order = [7, 2, 9, 0, 4, 1, 11, 3, 8, 5, 10, 6, 15, 12, 14, 13, 19, 16, 18, 17]
    vals = [round(30 + i + [3, -3][i % 2] + 0.5 * gauss(r), 3) for i in range(20)]
    out["unsorted"] = {"t": [o + 1 for o in order], "v": [vals[o] for o in order]}
    return out


def col(n: int) -> str:
    s = ""
    while n:
        n, k = divmod(n - 1, 26)
        s = chr(65 + k) + s
    return s


def main() -> None:
    data = make()
    x = ExcelTarget(visible=False)
    out: dict = {"series": data, "cases": {}}
    try:
        ws = x.sheet
        x.app.Calculation = -4135  # manual while filling; restored below (an application option)
        for name, s in data.items():
            n = len(s["t"])
            ws.Cells.ClearContents()
            ws.Range(f"A1:A{n}").Value2 = tuple((t,) for t in s["t"])
            ws.Range(f"B1:B{n}").Value2 = tuple((v,) for v in s["v"])
            vals, tl = f"$B$1:$B${n}", f"$A$1:$A${n}"
            ts = sorted(set(s["t"]))
            step = ts[1] - ts[0] if not name.startswith("months") else None
            # Variants: (label, season argument, fill, aggregation).
            variants = [("auto", 1, 1, 1), ("none", 0, 1, 1)]
            if name.startswith("gaps"):
                variants.append(("zero", 1, 0, 1))
            if name.startswith("dups"):
                variants += [(f"agg{a}", 1, 1, a) for a in range(2, 8)]
            rec: dict = {}
            formulas: list[tuple[str, str]] = []
            last = ts[-1]

            def at(h: int):
                if step is not None:
                    return last + h * step
                return f"EDATE({last},{h})"

            for label, se, fi, ag in variants:
                tail = f",{se},{fi},{ag}"
                formulas.append((f"{label}|season", f"=FORECAST.ETS.SEASONALITY({vals},{tl},{fi},{ag})"))
                for k in range(1, 9):
                    formulas.append((f"{label}|stat{k}", f"=FORECAST.ETS.STAT({vals},{tl},{k}{tail})"))
                for i, t in enumerate(ts):
                    formulas.append((f"{label}|fit{i}", f"=FORECAST.ETS({t},{vals},{tl}{tail})"))
                for h in range(1, 25):
                    formulas.append((f"{label}|ets{h}", f"=FORECAST.ETS({at(h)},{vals},{tl}{tail})"))
                    formulas.append((f"{label}|conf{h}", f"=FORECAST.ETS.CONFINT({at(h)},{vals},{tl},0.95{tail})"))
            for k, (_, f) in enumerate(formulas):
                c = ws.Cells(1 + k % 1000, 4 + k // 1000)
                c.Formula2 = f
            x.app.Calculate()
            for k, (key, _) in enumerate(formulas):
                v = ws.Cells(1 + k % 1000, 4 + k // 1000).Value2
                if isinstance(v, int) and v in ERRORS:
                    v = ERRORS[v]
                label, what = key.split("|")
                rec.setdefault(label, {})[what] = v
            out["cases"][name] = {}
            for label, r in rec.items():
                nfit = len(ts)
                out["cases"][name][label] = {
                    "season": r["season"],
                    "stat": [r[f"stat{k}"] for k in range(1, 9)],
                    "fit": [r[f"fit{i}"] for i in range(nfit)],
                    "ets": [r[f"ets{h}"] for h in range(1, 25)],
                    "conf": [r[f"conf{h}"] for h in range(1, 25)],
                }
            a = out["cases"][name]["auto"]
            print(name, a["season"], [round(v, 4) if isinstance(v, float) else v for v in a["stat"]], flush=True)
    finally:
        try:
            x.app.Calculation = -4105
        except Exception:  # noqa: BLE001
            pass
        x.close()
    OUT.write_text(json.dumps(out, ensure_ascii=False, indent=0), encoding="utf-8", newline="\n")
    print("written", OUT)


if __name__ == "__main__":
    main()
