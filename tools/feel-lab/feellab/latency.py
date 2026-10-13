"""F1 keystroke-to-screen latency, measured from outside the application.

For every sample the screen region around the caret is captured until it is stable, a
real key event is injected with SendInput, and the region is polled until it differs
from every state seen before the key press (so a blinking caret is never mistaken for
the new character). Takes over keyboard and screen."""

from __future__ import annotations

import random
import statistics
import time

from . import winapi
from .targets import Target, WebViewTarget, WordTarget

SAMPLES = 60
VK_N = 0x4E


def _stats(values: list[float]) -> dict:
    v = sorted(values)
    if not v:
        return {}
    return {
        "n": len(v),
        "median_ms": round(statistics.median(v), 1),
        "mean_ms": round(statistics.fmean(v), 1),
        "p95_ms": round(v[int(0.95 * (len(v) - 1))], 1),
        "min_ms": round(v[0], 1),
        "max_ms": round(v[-1], 1),
    }


def measure(target: Target, caret_offset: int, samples: int = SAMPLES) -> dict:
    target.set_selection(caret_offset, caret_offset)
    target.ensure_visible(caret_offset)
    target.activate()
    time.sleep(0.4)
    x, y, h = target.point_at(caret_offset)
    grabber = winapi.Grabber(int(x) - 4, int(y) - 4, 760, int(h) + 8)
    results: list[float] = []
    timeouts = 0
    try:
        for _ in range(samples):
            time.sleep(0.12 + random.random() * 0.034)  # de-correlate from the display refresh
            seen = {grabber.signature() for _ in range(4)}
            t0 = time.perf_counter_ns()
            winapi.key_down(VK_N)
            detected = None
            while time.perf_counter_ns() - t0 < 500_000_000:
                if grabber.signature() not in seen:
                    detected = time.perf_counter_ns()
                    break
            winapi.key_up(VK_N)
            if detected is None:
                timeouts += 1
            else:
                results.append((detected - t0) / 1e6)
    finally:
        grabber.close()
    return {**_stats(results), "timeouts": timeouts, "samples_ms": [round(r, 2) for r in results]}


def _wrapped_line_start(t: Target, after: int) -> int:
    text = t.text()
    return next(s for s in t.lines() if s > after and text[s - 1] == " ")


def run() -> dict:
    out: dict = {}
    scenarios = {
        # Empty paragraph at the end of a one-page document.
        "short_end": ("short", lambda t: len(t.text())),
        # Start of a wrapped line inside a paragraph on page ~12 of the 50-page document:
        # every key reflows the paragraph, and the typed text stays on this line.
        "long_middle": ("long", lambda t: _wrapped_line_start(t, len(t.text()) // 4)),
    }
    makers = {
        "word": lambda fx: WordTarget(fx),
        "proto-a": lambda fx: WebViewTarget("a", fx, port=9340),
        "proto-b": lambda fx: WebViewTarget("b", fx, port=9341),
    }
    for scenario, (fx, where) in scenarios.items():
        out[scenario] = {}
        for name, make in makers.items():
            t = make(fx)
            try:
                offset = where(t)
                if scenario == "short_end":
                    t.set_selection(offset, offset)
                    t.activate()
                    winapi.press("enter")
                    offset += 1
                    time.sleep(0.3)
                res = measure(t, offset)
                if isinstance(t, WebViewTarget):
                    inside = t.js("__probe.latency.take()")
                    res["in_app_event_to_frame"] = _stats(inside)
                out[scenario][name] = res
            finally:
                t.close()
    return out
