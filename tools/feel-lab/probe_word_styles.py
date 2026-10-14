"""Probe: Word's Formatvorlagen pane (Start > Formatvorlagen launcher): an entry's menu (right
click), the dialogs Formatvorlage ändern, Neue Formatvorlage von Formatierung erstellen and
Formatvorlagenbereichsoptionen. Real clicks in our own visible Word instance; records UI
Automation names and screenshots. Takes over mouse and screen.
Screenshots: captures/word-dialogs/ (local); names: results/<date>-word-dialogs.json.

    .venv/Scripts/python.exe probe_word_styles.py
"""
from __future__ import annotations

import json
import sys
import time

from feellab import uia, winapi
from probe_word_dialogs import CAP, OUT, Word, new_windows, snapshot


def shot(name: str, hwnd: int, out: dict, el=None) -> None:
    l, t, r, b = winapi.window_rect(hwnd)
    winapi.save_region_png(str(CAP / f"{name}.png"), l, t, r - l, b - t)
    out[name] = {"window": winapi.window_class(hwnd), "title": winapi.window_title(hwnd), "size": [r - l, b - t],
                 "items": snapshot(el if el is not None else uia.element_for(hwnd), (l, t))}
    print(name, winapi.window_class(hwnd), len(out[name]["items"]), flush=True)


def opened(before: set[int]) -> int | None:
    wins = new_windows(before)
    return next((h for h in wins if winapi.window_class(h) in ("bosa_sdm_msword", "NUIDialog", "#32770")), None) \
        or next((h for h in wins if winapi.window_class(h) == "Net UI Tool Window"), None)


def click_el(el) -> None:
    r = uia.describe(el)["rect"]
    winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    out: dict = json.loads(OUT.read_text(encoding="utf-8")) if OUT.exists() else {}
    w = Word()
    try:
        w.tab("Start")
        launcher = w.group_launcher("Formatvorlagen")
        if launcher is None:
            print("no launcher")
            return
        click_el(launcher)
        time.sleep(1.5)
        item = uia.find(w.root, lambda d: d["name"] == "Überschrift 1" and d["type"] in ("ListItem", "Button", "MenuItem") and d["rect"][0] > 1000, max_depth=16)
        if item is None:
            print("no entry")
            return
        r = uia.describe(item)["rect"]
        if "buttons" in sys.argv[1:]:
            mod = menu = None
        # The entry's menu (right click opens the same menu as its arrow).
        if "buttons" not in sys.argv[1:]:
          before = set(winapi.find_windows(lambda h: True))
          winapi.right_click(r[0] + r[2] / 2, r[1] + r[3] / 2)
          time.sleep(1.2)
          menu = opened(before)
          if menu:
            shot("styles pane entry menu", menu, out)
        # Ändern... from that menu.
        mod = None if "buttons" in sys.argv[1:] else uia.find(w.root, lambda d: d["name"] in ("Ändern...", "Ändern…") and d["type"] in ("MenuItem", "Button"), max_depth=20)
        if mod is None and menu:
            mod = uia.find(uia.element_for(menu), lambda d: d["name"] in ("Ändern...", "Ändern…"), max_depth=12)
        if mod is not None:
            before = set(winapi.find_windows(lambda h: True))
            click_el(mod)
            time.sleep(1.8)
            dlg = opened(before)
            if dlg:
                shot("modify style dialog", dlg, out)
                # Its Format menu.
                fmt = uia.find(uia.element_for(dlg), lambda d: d["name"].startswith("Format") and d["type"] in ("Button", "MenuItem", "SplitButton"), max_depth=12)
                if fmt is not None:
                    before2 = set(winapi.find_windows(lambda h: True))
                    click_el(fmt)
                    time.sleep(1.2)
                    m2 = opened(before2)
                    if m2:
                        shot("modify style format menu", m2, out)
                    winapi.press("esc")
                    time.sleep(0.5)
                winapi.press("esc")
                time.sleep(0.8)
        elif "buttons" not in sys.argv[1:]:
            winapi.press("esc")
            time.sleep(0.5)
        # The pane's buttons: Neue Formatvorlage, Optionen...
        for name, label in [("new style dialog", "Neue Formatvorlage"), ("styles pane options", "Optionen...")]:
            btn = uia.find(w.root, lambda d: d["name"] == label and d["type"] == "Button", max_depth=16)
            if btn is None:
                print("no button", label, flush=True)
                continue
            before = set(winapi.find_windows(lambda h: True))
            click_el(btn)
            time.sleep(1.8)
            dlg = opened(before)
            if dlg:
                l, t, rr, bb = winapi.window_rect(dlg)
                winapi.save_region_png(str(CAP / f"{name}.png"), l, t, rr - l, bb - t)
                print(name, flush=True)
                # Abbrechen (Esc may not reach a Word dialog that has no focus).
                cancel = uia.find(uia.element_for(dlg), lambda d: d["type"] == "Button" and d["name"] == "Abbrechen", max_depth=10)
                if cancel is not None:
                    click_el(cancel)
                else:
                    winapi.press("esc")
                time.sleep(1.0)
    finally:
        OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        w.close()


if __name__ == "__main__":
    main()
