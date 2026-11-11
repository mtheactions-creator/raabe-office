"""Helper for Step 11 probes: lists the ribbon controls of our own Office instance whose names match the given words
(type, name, rect, children), so menu-closing probes can find the dropdown arrows and galleries. Starts its own Excel or
Word (DispatchEx), shows the window (no input) and quits it again.

    .venv/Scripts/python.exe probe_office_ribbon_names.py excel Füllfarbe Zellenformatvorlagen Fett
"""
from __future__ import annotations

import sys
import time

from feellab import uia, winapi


def dump(root, words: list[str]) -> None:
    stack = [(root, 0)]
    while stack:
        e, d = stack.pop()
        try:
            dc = uia.describe(e)
            kids = uia.children(e) if d < 14 else []
        except Exception:  # noqa: BLE001
            continue
        if any(w.lower() in dc["name"].lower() for w in words):
            print(d, dc["type"], repr(dc["name"]), dc["rect"], dc["class"])
            for c in kids:
                cc = uia.describe(c)
                print("    ->", cc["type"], repr(cc["name"]), cc["rect"])
        stack.extend((c, d + 1) for c in reversed(kids))


def main() -> None:
    app, words = sys.argv[1], sys.argv[2:]
    if app == "excel":
        from feellab.excel import ExcelTarget
        x = ExcelTarget(visible=True)
        try:
            time.sleep(1.5)
            dump(uia.element_for(x.hwnd), words)
        finally:
            x.close()
    else:
        from feellab.targets import WordTarget
        w = WordTarget()
        try:
            time.sleep(1.5)
            dump(uia.element_for(w.hwnd), words)
        finally:
            w.close()


if __name__ == "__main__":
    main()
