"""Start time and memory of RaabeCalc against Excel (FEEL.md F5, F7): from launching the program
until a workbook is ready to type into, five times each, and the working set of the whole process
tree with an empty workbook open. Excel starts as a process of its own (DispatchEx, visible).
Result: results/<date>-calc-startup.json.
"""
from __future__ import annotations

import datetime
import json
import statistics
import subprocess
import time
from pathlib import Path

from feellab.calc_catalogue import CalcSide


def working_set(pid: int) -> int:
    """Bytes of the process and its children (WebView2 runs in child processes)."""
    ps = (f"$p = Get-CimInstance Win32_Process; $ids = @({pid}); $more = $true; while ($more) {{ $more = $false; "
          f"foreach ($x in $p) {{ if ($ids -contains $x.ParentProcessId -and -not ($ids -contains $x.ProcessId)) {{ $ids += $x.ProcessId; $more = $true }} }} }}; "
          f"($p | Where-Object {{ $ids -contains $_.ProcessId }} | Measure-Object WorkingSetSize -Sum).Sum")
    out = subprocess.run(["powershell", "-NoProfile", "-Command", ps], capture_output=True, text=True).stdout.strip()
    return int(out or 0)


def main() -> None:
    res: dict = {"raabecalc": [], "excel": []}
    mem: dict = {}
    for i in range(5):
        t = time.perf_counter()
        c = CalcSide(port=9390 + i)
        res["raabecalc"].append(round((time.perf_counter() - t) * 1000))
        if i == 4:
            time.sleep(2)
            mem["raabecalc"] = working_set(c.proc.pid)
        c.close()
        time.sleep(1)
    from feellab.excel import ExcelTarget
    for i in range(5):
        t = time.perf_counter()
        x = ExcelTarget(visible=True, place=False)
        _ = x.sheet.Range("A1").Value2
        res["excel"].append(round((time.perf_counter() - t) * 1000))
        if i == 4:
            time.sleep(2)
            mem["excel"] = working_set(x.pid)
        x.close()
        time.sleep(2)
    out = {"startMs": {k: {"median": statistics.median(v), "all": v} for k, v in res.items()}, "workingSetMB": {k: round(v / 1e6) for k, v in mem.items()}}
    print(json.dumps(out, indent=1))
    Path(__file__).parent.joinpath("results", f"{datetime.date.today().isoformat()}-calc-startup.json").write_text(json.dumps(out, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
