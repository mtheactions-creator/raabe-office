"""Probe: Word's dialogs Neues Aufzählungszeichen definieren and Neues Zahlenformat definieren
(Step 4c F). In our own visible Word instance the arrow of Start > Aufzählungszeichen and
Nummerierung is clicked, then the dialog's menu entry; the dialog's controls (with their values)
and the entries of its drop-down lists are recorded, then it is closed with Esc. Takes over
mouse and screen. Result: results/<date>-word-list-define.json.

    .venv/Scripts/python.exe probe_word_list_define.py
"""
from __future__ import annotations

import datetime
import json
import sys
import time
from pathlib import Path

from feellab import uia, winapi
from probe_word_dialogs import Word, new_windows

HERE = Path(__file__).parent
OUT = HERE / "results" / f"{datetime.date.today()}-word-list-define.json"


def controls(dialog: int) -> list[dict]:
    l, t, _, _ = winapi.window_rect(dialog)
    el = uia.element_for(dialog)
    f = el.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
    out = []
    for i in range(f.Length):
        e = f.GetElement(i)
        d = uia.describe(e)
        if not d["name"] and d["type"] not in ("Edit", "ComboBox"):
            continue
        item = {"type": d["type"], "name": d["name"], "enabled": d["enabled"], "rect": [d["rect"][0] - l, d["rect"][1] - t, d["rect"][2], d["rect"][3]]}
        try:
            if d["type"] in ("Edit", "ComboBox"):
                item["value"] = e.GetCurrentPattern(uia.U.UIA_ValuePatternId).QueryInterface(uia.U.IUIAutomationValuePattern).CurrentValue
        except Exception:  # noqa: BLE001
            pass
        if d["type"] == "ComboBox":
            # The entries: expand the list, read its items, collapse it again.
            try:
                # Word's dialog lists open with a click on their arrow, in a popup of their own.
                x, y, cw, ch = d["rect"]
                before = set(winapi.find_windows(lambda h: True))
                winapi.click(x + cw - 12, y + ch / 2)
                time.sleep(0.8)
                names: list[str] = []
                for src in [dialog, *[h for h in winapi.find_windows(lambda h: True) if h not in before]]:
                    try:
                        g = uia.element_for(src).FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
                        names += [uia.describe(g.GetElement(k))["name"] for k in range(g.Length) if uia.describe(g.GetElement(k))["type"] == "ListItem"]
                    except Exception:  # noqa: BLE001
                        continue
                item["options"] = list(dict.fromkeys(names))
                winapi.press("esc")
                time.sleep(0.4)
            except Exception as ex:  # noqa: BLE001
                item["options_error"] = str(ex)[:80]
        out.append(item)
    return out


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    out: dict = {}
    w = Word()
    try:
        w.d.Content.Text = "Erster Punkt\rZweiter Punkt"
        w.d.Range(0, 0).Select()
        w.tab("Start")
        for control, entry in (("Aufzählungszeichen", "Neues Aufzählungszeichen definieren..."), ("Nummerierung", "Neues Zahlenformat definieren...")):
            el = w.control(control)
            if el is None:
                out[entry] = {"error": "control not found"}
                continue
            x, y, cw, ch = uia.describe(el)["rect"]
            before = set(winapi.find_windows(lambda h: True))
            winapi.click(x + cw - 6, y + ch / 2)
            time.sleep(1.2)
            hit = uia.find(w.root, lambda d: d["name"] == entry and d["type"] in ("MenuItem", "Button"), max_depth=22)
            if hit is None:
                for h in new_windows(before):
                    hit = uia.find(uia.element_for(h), lambda d: d["name"] == entry, max_depth=14) or hit
            if hit is None:
                out[entry] = {"error": "entry not found"}
                winapi.press("esc")
                continue
            r = uia.describe(hit)["rect"]
            before = set(winapi.find_windows(lambda h: True))
            winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
            time.sleep(1.5)
            dialog = next((h for h in new_windows(before) if winapi.window_class(h) in ("bosa_sdm_msword", "NUIDialog", "#32770")), None)
            if dialog is None:
                out[entry] = {"error": "no dialog"}
                continue
            l, t, rr, b = winapi.window_rect(dialog)
            out[entry] = {"title": winapi.window_title(dialog), "size": [rr - l, b - t], "items": controls(dialog)}
            print(entry, json.dumps(out[entry], ensure_ascii=False)[:600], flush=True)
            winapi.press("esc")
            time.sleep(0.8)
    finally:
        OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        w.close()


if __name__ == "__main__":
    main()
