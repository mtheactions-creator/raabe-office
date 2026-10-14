"""Probe: Word's "Quelle erstellen" dialog (Referenzen > Zitat einfügen > Neue Quelle
hinzufügen...): the field labels for every Quellentyp, with and without "Alle
Literaturverzeichnisfelder anzeigen". Own visible Word instance; takes over mouse and
keyboard. Adds to results/<date>-word-dialogs.json.
"""
from __future__ import annotations

import json
import sys
import time

from feellab import uia, winapi
from probe_word_dialogs import CAP, OUT, Word, new_windows, snapshot


def labels(d) -> list[str]:
    f = d.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
    out = []
    for i in range(f.Length):
        x = uia.describe(f.GetElement(i))
        if x["type"] == "Text" and x["name"] and x["rect"][2] > 0:
            out.append(x["name"])
    return out


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    out = json.loads(OUT.read_text(encoding="utf-8")) if OUT.exists() else {}
    w = Word()
    res: dict = {}
    try:
        w.tab("Referenzen")
        el = w.control("Zitat einfügen")
        x, y, cw, ch = uia.describe(el)["rect"]
        winapi.click(x + cw / 2, y + ch / 2)
        time.sleep(1.0)
        item = uia.find(w.root, lambda d: d["name"] == "Neue Quelle hinzufügen..." and d["type"] in ("MenuItem", "Button"), max_depth=20)
        r = uia.describe(item)["rect"]
        before = set(winapi.find_windows(lambda h: True))
        winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
        time.sleep(1.5)
        dlg = next(h for h in new_windows(before) if winapi.window_class(h) in ("bosa_sdm_msword", "NUIDialog", "#32770"))
        d = uia.element_for(dlg)
        l, t, rr, bb = winapi.window_rect(dlg)
        winapi.save_region_png(str(CAP / "source dialog.png"), l, t, rr - l, bb - t)
        res["first"] = snapshot(d, (l, t))
        combo = uia.find(d, lambda q: q["type"] == "ComboBox" and q["name"].startswith("Quellentyp"), max_depth=8)
        types: list[str] = []
        if combo is not None:
            cr = uia.describe(combo)["rect"]
            for k in range(20):
                winapi.click(cr[0] + 40, cr[1] + cr[3] / 2)
                time.sleep(0.4)
                winapi.press("esc")
                time.sleep(0.2)
                name = uia.describe(combo)
                f = combo.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
                value = ""
                try:
                    pat = combo.GetCurrentPattern(uia.U.UIA_ValuePatternId).QueryInterface(uia.U.IUIAutomationValuePattern)
                    value = pat.CurrentValue
                except Exception:  # noqa: BLE001
                    value = name["name"]
                if value in types:
                    break
                types.append(value)
                res[f"type {value}"] = labels(d)
                print(value, res[f"type {value}"][:30], flush=True)
                if k == 0:
                    winapi.save_region_png(str(CAP / f"source dialog {k}.png"), l, t, rr - l, bb - t)
                # The next type: focus the combo box and press Down.
                winapi.click(cr[0] + 40, cr[1] + cr[3] / 2)
                time.sleep(0.4)
                winapi.press("down")
                time.sleep(0.2)
                winapi.press("enter")
                time.sleep(0.6)
        res["types"] = types
        # All fields of the first type.
        allf = uia.find(d, lambda q: q["type"] == "CheckBox" and q["name"].startswith("Alle Literaturverzeichnisfelder"), max_depth=8)
        if allf is not None:
            ar = uia.describe(allf)["rect"]
            winapi.click(ar[0] + 8, ar[1] + ar[3] / 2)
            time.sleep(0.8)
            l, t, rr, bb = winapi.window_rect(dlg)
            winapi.save_region_png(str(CAP / "source dialog all.png"), l, t, rr - l, bb - t)
            res["all fields"] = labels(d)
        winapi.press("esc")
        time.sleep(0.6)
    finally:
        out["source dialog"] = res
        OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        w.close()


if __name__ == "__main__":
    main()
