"""Probe: Word's Symbol dialog in detail: the Sonderzeichen tab (characters, names, shortcuts),
the subsets of "(normaler Text)" and the "von:" choices. Own visible Word instance; takes over
mouse and keyboard. Adds to results/<date>-word-dialogs.json; screenshots in
captures/word-dialogs/ (local).
"""
from __future__ import annotations

import json
import sys
import time

from feellab import uia, winapi
from probe_word_dialogs import CAP, OUT, Word, new_windows, snapshot


def items_of(el) -> list[dict]:
    f = el.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
    return [uia.describe(f.GetElement(i)) for i in range(f.Length)]


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    out = json.loads(OUT.read_text(encoding="utf-8")) if OUT.exists() else {}
    w = Word()
    try:
        w.tab("Einfügen")
        el = w.control("Symbol")
        x, y, cw, ch = uia.describe(el)["rect"]
        winapi.click(x + cw / 2, y + ch / 2)
        time.sleep(1.2)
        more = uia.find(w.root, lambda d: d["name"] == "Weitere Symbole..." and d["type"] == "MenuItem", max_depth=20)
        r = uia.describe(more)["rect"]
        before = set(winapi.find_windows(lambda h: True))
        winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
        time.sleep(1.5)
        dlg = next(h for h in new_windows(before) if winapi.window_class(h) == "bosa_sdm_msword")
        d = uia.element_for(dlg)
        l, t, rr, bb = winapi.window_rect(dlg)
        # The font "(normaler Text)": typed into the font box.
        font = uia.find(d, lambda x: x["type"] == "Edit" and x["name"] == "Schriftart:", max_depth=8)
        fr = uia.describe(font)["rect"]
        winapi.click(fr[0] + fr[2] / 2, fr[1] + fr[3] / 2)
        winapi.press("ctrl+a")
        winapi.type_text("(normaler Text)")
        winapi.press("enter")
        time.sleep(1.0)
        winapi.save_region_png(str(CAP / "symbol dialog normal.png"), l, t, rr - l, bb - t)
        out["symbol dialog normal"] = snapshot(d, (l, t))
        # The subset list: its drop button, then the list's entries.
        sub = uia.find(d, lambda x: x["type"] == "ComboBox" and x["name"].startswith("Subset"), max_depth=8)
        subsets: list[str] = []
        if sub is not None:
            sr = uia.describe(sub)["rect"]
            before2 = set(winapi.find_windows(lambda h: True))
            winapi.click(sr[0] + sr[2] - 10, sr[1] + sr[3] / 2)
            time.sleep(1.0)
            for h in new_windows(before2):
                subsets += [x["name"] for x in items_of(uia.element_for(h)) if x["type"] == "ListItem"]
            if not subsets:
                subsets = [x["name"] for x in items_of(sub) if x["type"] == "ListItem"]
            winapi.press("esc")
            time.sleep(0.4)
        out["symbol subsets"] = subsets
        print("subsets", len(subsets), subsets[:8], flush=True)
        # von: choices.
        von = uia.find(d, lambda x: x["type"] == "ComboBox" and x["name"] == "von:", max_depth=8)
        if von is not None:
            vr = uia.describe(von)["rect"]
            before3 = set(winapi.find_windows(lambda h: True))
            winapi.click(vr[0] + vr[2] - 10, vr[1] + vr[3] / 2)
            time.sleep(0.8)
            vals = []
            for h in new_windows(before3):
                vals += [x["name"] for x in items_of(uia.element_for(h)) if x["type"] == "ListItem"]
            if not vals:
                vals = [x["name"] for x in items_of(von) if x["type"] == "ListItem"]
            out["symbol von"] = vals
            print("von", vals, flush=True)
            winapi.press("esc")
            time.sleep(0.4)
        # The Sonderzeichen tab.
        tab = uia.find(d, lambda x: x["type"] == "TabItem" and x["name"] == "Sonderzeichen", max_depth=8)
        tr = uia.describe(tab)["rect"]
        winapi.click(tr[0] + tr[2] / 2, tr[1] + tr[3] / 2)
        time.sleep(1.0)
        winapi.save_region_png(str(CAP / "symbol dialog special.png"), l, t, rr - l, bb - t)
        out["symbol special"] = snapshot(d, (l, t))
        print("special", len(out["symbol special"]), flush=True)
        winapi.press("esc")
        time.sleep(0.6)
    finally:
        OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        w.close()


if __name__ == "__main__":
    main()
