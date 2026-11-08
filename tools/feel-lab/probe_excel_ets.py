"""Records Excel's exponential smoothing functions (D-092): PROGNOSE.ETS, PROGNOSE.ETS.KONFINT,
PROGNOSE.ETS.SAISONALITÄT and PROGNOSE.ETS.STAT for fixed series (seasonal with trend, short, quarterly,
with a gap, with duplicate time stamps, dates, a flat one, bad arguments). Our own invisible Excel, COM
only. Written to packages/sheet-engine/test/fixtures/excel-ets.json (used by ets.test.ts).

    .venv/Scripts/python.exe probe_excel_ets.py
"""
from __future__ import annotations

import json
import math
from pathlib import Path

from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
OUT = HERE.parents[1] / "packages" / "sheet-engine" / "test" / "fixtures" / "excel-ets.json"
ERRORS = {-2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?", -2146826288: "#NULL!",
          -2146826252: "#NUM!", -2146826265: "#REF!", -2146826273: "#VALUE!"}


def lcg(seed: int):
    s = seed
    while True:
        s = (s * 1103515245 + 12345) % 2 ** 31
        yield s / 2 ** 31


def series() -> dict:
    out = {}
    r = lcg(7)
    out["seasonal12"] = {"t": list(range(1, 37)), "v": [round(100 + 2 * i + 15 * math.sin(2 * math.pi * i / 12) + 6 * (next(r) - 0.5), 2) for i in range(36)]}
    r = lcg(11)
    out["quarterly"] = {"t": list(range(1, 21)), "v": [round(50 + 1.5 * i + [8, -3, -9, 4][i % 4] + 3 * (next(r) - 0.5), 2) for i in range(20)]}
    out["short"] = {"t": list(range(1, 9)), "v": [3, 5, 4, 6, 8, 7, 9, 10]}
    r = lcg(3)
    out["noisy"] = {"t": list(range(1, 31)), "v": [round(20 + 10 * (next(r) - 0.5), 2) for _ in range(30)]}
    out["linear"] = {"t": list(range(1, 13)), "v": [10 + 3 * i for i in range(12)]}
    # Monthly dates (first of month) as serial numbers: 2024-01-01 is 45292.
    months = [45292, 45323, 45352, 45383, 45413, 45444, 45474, 45505, 45536, 45566, 45597, 45627]
    months += [m + 366 for m in months] + [m + 731 for m in months]
    r = lcg(5)
    out["dates"] = {"t": [45292 + 7 * i for i in range(30)], "v": [round(200 + 3 * i + 20 * math.cos(2 * math.pi * i / 4) + 8 * (next(r) - 0.5), 1) for i in range(30)]}
    gap = {"t": list(range(1, 25)), "v": [round(30 + i + [5, 0, -5][i % 3], 1) for i in range(24)]}
    del gap["t"][10], gap["v"][10]
    out["gap"] = gap
    dup = {"t": [1, 2, 3, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14], "v": [5, 7, 6, 8, 9, 8, 10, 12, 11, 13, 14, 13, 15, 17, 16]}
    out["dup"] = dup
    return out


def main() -> None:
    data = series()
    x = ExcelTarget(visible=False)
    out: dict = {"series": data, "cases": []}
    try:
        ws = x.sheet

        def put(name: str, s: dict) -> tuple[str, str]:
            n = len(s["t"])
            ws.Range("A1:B200").ClearContents()
            for i in range(n):
                ws.Cells(i + 1, 1).Value2 = s["t"][i]
                ws.Cells(i + 1, 2).Value2 = s["v"][i]
            return f"$B$1:$B${n}", f"$A$1:$A${n}"

        def ev(formula: str):
            c = ws.Range("D1")
            try:
                c.Formula2Local = formula
            except Exception:  # noqa: BLE001
                return {"rejected": True}
            v = c.Value2
            if isinstance(v, int) and v in ERRORS:
                return {"error": ERRORS[v]}
            return v

        for name, s in data.items():
            vals, tl = put(name, s)
            last = s["t"][-1]
            step = s["t"][-1] - s["t"][-2] if name != "dup" else 1
            rec = {"series": name, "seasonality": ev(f"=PROGNOSE.ETS.SAISONALITÄT({vals};{tl})"),
                   "stat": [ev(f"=PROGNOSE.ETS.STAT({vals};{tl};{k})") for k in range(1, 9)],
                   "ets": {}, "confint": {}, "variants": {}}
            for h in (1, 2, 3, 6, 12):
                t = last + h * step
                rec["ets"][str(h)] = ev(f"=PROGNOSE.ETS({t};{vals};{tl})")
                rec["confint"][str(h)] = ev(f"=PROGNOSE.ETS.KONFINT({t};{vals};{tl})")
            t1 = last + step
            for label, f in {
                "seas0": f"=PROGNOSE.ETS({t1};{vals};{tl};0)",
                "seas4": f"=PROGNOSE.ETS({t1};{vals};{tl};4)",
                "seas0_stat": f"=PROGNOSE.ETS.STAT({vals};{tl};1;0)",
                "seas0_stat2": f"=PROGNOSE.ETS.STAT({vals};{tl};2;0)",
                "seas0_rmse": f"=PROGNOSE.ETS.STAT({vals};{tl};7;0)",
                "seas4_stat": f"=PROGNOSE.ETS.STAT({vals};{tl};1;4)",
                "conf90": f"=PROGNOSE.ETS.KONFINT({t1};{vals};{tl};0,9)",
                "conf90_h3": f"=PROGNOSE.ETS.KONFINT({last + 3 * step};{vals};{tl};0,9)",
                "seas0_conf": f"=PROGNOSE.ETS.KONFINT({t1};{vals};{tl};0,95;0)",
                "half": f"=PROGNOSE.ETS({str(last + step / 2).replace('.', ',')};{vals};{tl})",
                "inside": f"=PROGNOSE.ETS({s['t'][len(s['t']) // 2]};{vals};{tl})",
                "zero_fill": f"=PROGNOSE.ETS({t1};{vals};{tl};1;0)",
                "agg_sum": f"=PROGNOSE.ETS({t1};{vals};{tl};1;1;7)",
                "agg_max": f"=PROGNOSE.ETS({t1};{vals};{tl};1;1;4)",
                "seasonality_zero": f"=PROGNOSE.ETS.SAISONALITÄT({vals};{tl};0)",
            }.items():
                rec["variants"][label] = ev(f)
            out["cases"].append(rec)
            print(name, rec["seasonality"], rec["stat"], rec["ets"]["1"], rec["confint"]["1"], flush=True)
        # Bad arguments.
        vals, tl = put("short", data["short"])
        out["errors"] = {f: ev(f) for f in [
            f"=PROGNOSE.ETS(20;{vals};{tl};9000)", f"=PROGNOSE.ETS(20;{vals};{tl};-1)", f"=PROGNOSE.ETS(20;{vals};$A$1:$A$7)",
            f"=PROGNOSE.ETS(20;$B$1:$B$2;$A$1:$A$2)", f"=PROGNOSE.ETS.KONFINT(20;{vals};{tl};1)", f"=PROGNOSE.ETS.STAT({vals};{tl};9)",
            f"=PROGNOSE.ETS(\"x\";{vals};{tl})", f"=PROGNOSE.ETS(20;{vals};{tl};1;2)", f"=PROGNOSE.ETS(20;{vals};{tl};1;1;8)",
            f"=PROGNOSE.ETS(0;{vals};{tl})"]}
        print(out["errors"])
        # English names as the file stores them.
        ws.Range("D1").Formula2Local = f"=PROGNOSE.ETS(20;{vals};{tl})"
        ws.Range("D2").Formula2Local = f"=PROGNOSE.ETS.KONFINT(20;{vals};{tl})"
        ws.Range("D3").Formula2Local = f"=PROGNOSE.ETS.SAISONALITÄT({vals};{tl})"
        ws.Range("D4").Formula2Local = f"=PROGNOSE.ETS.STAT({vals};{tl};1)"
        out["english"] = [ws.Range(a).Formula2 for a in ("D1", "D2", "D3", "D4")]
    finally:
        x.close()
    OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print("written", OUT)


if __name__ == "__main__":
    main()
