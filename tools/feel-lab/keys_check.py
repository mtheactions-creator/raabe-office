"""Real-keystroke checks of paragraph shortcuts whose keyboard behaviour COM cannot show
(pressing Strg+E twice etc.), in Word and in RaabeWrite. Takes over keyboard and screen.

    .venv/Scripts/python.exe keys_check.py
"""

from __future__ import annotations

import sys
import time

from feellab import targets, winapi

ALIGN = {0: "left", 1: "center", 2: "right", 3: "both"}
CASES = [
    ("Strg+E once", ["ctrl+e"]),
    ("Strg+R once", ["ctrl+r"]),
    ("Strg+B once", ["ctrl+b"]),
    ("Strg+L twice", ["ctrl+l", "ctrl+l"]),
    ("Strg+E twice", ["ctrl+e", "ctrl+e"]),
    ("Strg+R twice", ["ctrl+r", "ctrl+r"]),
    ("Strg+B twice", ["ctrl+b", "ctrl+b"]),
    ("Strg+E then Strg+L", ["ctrl+e", "ctrl+l"]),
    ("Strg+L on a left paragraph", ["ctrl+l"]),
]


def word_state(w: targets.WordTarget) -> str:
    return ALIGN.get(w.app.Selection.ParagraphFormat.Alignment, "?")


def ours_state(t: targets.WebViewTarget) -> str:
    return t.js(f"__probe.formatAt({t.selection()[1]}).align")


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    word = targets.WordTarget("short")
    ours = targets.raabe_write("short", port=9333)
    ok = 0
    try:
        for name, keys in CASES:
            results = []
            for t, state, is_word in ((word, word_state, True), (ours, ours_state, False)):
                t.load("short")
                at = t.find("Algen")
                t.set_selection(at, at)
                t.activate()
                time.sleep(0.3)
                for k in keys:
                    winapi.press(k)
                    time.sleep(0.25)
                results.append(state(t))
            same = results[0] == results[1]
            ok += same
            print(("PASS " if same else "FAIL ") + f"{name}: Word {results[0]}, RaabeWrite {results[1]}", flush=True)
    finally:
        ours.close()
        word.close()
    print(f"\n{ok}/{len(CASES)} identical")


if __name__ == "__main__":
    main()
