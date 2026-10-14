"""Probe: opens a ribbon combo box or menu button in Word with a real click and captures the
list below it. Own Word instance. Takes over the mouse briefly.

    .venv/Scripts/python.exe probe_word_combo.py Überprüfen "Anzeige zur Überarbeitung" [out.png]
"""
import sys
import time
from pathlib import Path

from feellab import targets, uia, winapi

tab, name = sys.argv[1], sys.argv[2]
out = Path(sys.argv[3] if len(sys.argv) > 3 else Path(__file__).parent / "captures" / "word-popup.png")
w = targets.WordTarget("short")
try:
    winapi.place(w.hwnd, 0, 0, 3000, 1500)
    winapi.activate(w.hwnd)
    root = uia.element_for(w.hwnd)
    uia.select(uia.find(root, lambda d: d["type"] == "TabItem" and d["name"] == tab))
    time.sleep(0.8)
    el = uia.find(root, lambda d: d["type"] in ("ComboBox", "MenuItem", "SplitButton", "Button") and d["name"] == name, max_depth=16)
    x, y, wd, ht = uia.describe(el)["rect"]
    winapi.click(x + wd - 8, y + ht / 2)
    time.sleep(1.0)
    winapi.save_region_png(str(out), max(0, x - 300), y, 1100, 900)
    winapi.press("esc")
    print("saved", out)
finally:
    w.close()
