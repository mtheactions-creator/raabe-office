"""Keystroke → presented frame inside RaabeWrite, without real input: keystrokes go in through
the DevTools protocol and the browser's Event Timing API reports, for each keydown, how long
it took until the next frame was presented (8 ms resolution). Runs off screen.

    .venv/Scripts/python.exe -m feellab.presentation [count] [fixture] [--trace out.json]
"""

from __future__ import annotations

import json
import statistics
import sys
import time

from . import targets

OBSERVE = """
window.__evt = [];
window.__loaf = [];
new PerformanceObserver((l) => { for (const e of l.getEntries()) if (e.name === 'keydown')
  __evt.push({start: e.startTime, procStart: e.processingStart, procEnd: e.processingEnd, duration: e.duration}); })
  .observe({type: 'event', durationThreshold: 16, buffered: false});
new PerformanceObserver((l) => { for (const e of l.getEntries()) __loaf.push({start: e.startTime, duration: e.duration,
  blocking: e.blockingDuration, render: e.renderStart, style: e.styleAndLayoutStart,
  scripts: e.scripts.map((s) => ({d: s.duration, f: s.sourceFunctionName, inv: s.invoker}))}); })
  .observe({type: 'long-animation-frame', buffered: false});
true
"""


def run(count: int = 200, fixture: str = "long", trace: str | None = None, background: bool = True, gap: float = 0.08) -> dict:
    t = targets.raabe_write(fixture, port=9333, background=background)
    try:
        text = t.text()
        at = text.find("Lichtenergie", len(text) // 4)
        t.set_selection(at, at)
        t.js(f"__probe.reveal({at})")
        t.js("__probe.focus()")
        time.sleep(1.0)
        t.js(OBSERVE)
        if trace:
            t.cdp("Tracing.start", {"traceConfig": {"includedCategories": [
                "devtools.timeline", "disabled-by-default-devtools.timeline", "disabled-by-default-devtools.timeline.frame",
                "benchmark", "latencyInfo", "viz", "gpu", "cc", "input", "blink"]}, "transferMode": "ReturnAsStream"})
        sample = "Das ist ein Test mit etwas mehr Text zum Messen. "
        for i in range(count):
            t.send_text(sample[i % len(sample)])
            time.sleep(gap)
        time.sleep(0.5)
        evt = t.js("__evt")
        loaf = t.js("__loaf")
        if trace:
            _save_trace(t, trace)
    finally:
        t.close()
    # Entries exist only above 16 ms; everything else was faster.
    durations = sorted(e["duration"] for e in evt)
    buckets = {f">{lim}ms": sum(d > lim for d in durations) for lim in (16, 40, 72, 104)}
    return {"keystrokes": count, "slow_entries": len(evt), "buckets": buckets,
            "slow_durations": durations[-20:], "median_slow": statistics.median(durations) if durations else 0,
            "processing_ms": sorted(round(e["procEnd"] - e["procStart"], 1) for e in evt)[-10:],
            "input_delay_ms": sorted(round(e["procStart"] - e["start"], 1) for e in evt)[-10:],
            "long_frames": len(loaf), "loaf": loaf[:15]}


def _save_trace(t: targets.WebViewTarget, path: str) -> None:
    import base64
    t._id += 1
    t.ws.send(json.dumps({"id": t._id, "method": "Tracing.end"}))
    handle = None
    while handle is None:
        msg = json.loads(t.ws.recv())
        if msg.get("method") == "Tracing.tracingComplete":
            handle = msg["params"]["stream"]
    chunks = []
    while True:
        r = t.cdp("IO.read", {"handle": handle})
        data = r.get("data", "")
        chunks.append(base64.b64decode(data).decode("utf-8") if r.get("base64Encoded") else data)
        if r.get("eof"):
            break
    t.cdp("IO.close", {"handle": handle})
    with open(path, "w", encoding="utf-8") as f:
        f.write("".join(chunks))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    trace = sys.argv[sys.argv.index("--trace") + 1] if "--trace" in sys.argv else None
    if trace in args:
        args.remove(trace)
    res = run(int(args[0]) if args else 200, args[1] if len(args) > 1 else "long", trace, background="--visible" not in sys.argv)
    print(json.dumps({k: v for k, v in res.items() if k != "loaf"}, indent=1))
    for f in res["loaf"]:
        print(round(f["duration"]), "ms frame, blocking", round(f["blocking"]), [(round(s["d"]), s["f"], s["inv"]) for s in f["scripts"]][:4])
