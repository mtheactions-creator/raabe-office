"""Probe: Word's docked task panes and the menus still missing for Step 4b E (Überarbeitungsbereich,
Formatvorlagen, Zwischenablage, Auswahl, Zeilennummern, Groß-/Kleinschreibung, Deckblatt …).
Opens each with a real click in our own visible Word instance (a document with tracked changes
and a comment), records what UI Automation sees and a screenshot, then closes it. Takes over
mouse and screen. The Office user name is replaced by "Autor" in the results (public repo).
Screenshots: captures/word-dialogs/ (local); names: results/<date>-word-dialogs.json.

    .venv/Scripts/python.exe probe_word_panes.py [name filter …]
"""
from __future__ import annotations

import json
import sys
import time

from feellab import uia, winapi
from probe_word_dialogs import CAP, OUT, Word, new_windows, snapshot

# (name, tab, control, how, pane name or ""): how is "click", "arrow" (right part of a split
# button), "launcher" (group's dialog launcher) or "menu:<item>" (after opening the control).
ITEMS = [
    ("reviewing pane", "Überprüfen", "Überarbeitungsbereich", "click", "Überarbeitungen"),
    ("reviewing pane menu", "Überprüfen", "Überarbeitungsbereich", "arrow", ""),
    ("reviewing pane horizontal", "Überprüfen", "Überarbeitungsbereich", "arrow:Überarbeitungsbereich horizontal...", "Überarbeitungen"),
    ("styles pane", "Start", "Formatvorlagen", "launcher", "Formatvorlagen"),
    ("clipboard pane", "Start", "Zwischenablage", "launcher", "Zwischenablage"),
    ("selection pane", "Start", "Bearbeiten", "click:Markieren>Auswahlbereich...", "Auswahl"),
    ("line numbers menu", "Layout", "Zeilennummern", "click", ""),
    ("case menu", "Start", "Groß-/Kleinschreibung", "click", ""),
    ("line spacing menu", "Start", "Zeilen- und Absatzabstand", "click", ""),
    ("cover page gallery", "Einfügen", "Deckblatt", "click", ""),
]


def scrub(x, user: str, initials: str):
    if isinstance(x, str):
        return x.replace(user, "Autor").replace(f"[{initials}", "[AU") if user else x
    if isinstance(x, list):
        return [scrub(v, user, initials) for v in x]
    if isinstance(x, dict):
        return {k: scrub(v, user, initials) for k, v in x.items()}
    return x


def pane_named(w: Word, name: str):
    return uia.find(w.root, lambda d: d["name"] == name and d["type"] in ("Pane", "Window", "Group", "Custom", "ToolBar") and d["rect"][2] > 60, max_depth=10)


def close_pane(w: Word, pane) -> None:
    btn = uia.find(pane, lambda d: d["type"] == "Button" and d["name"] in ("Schließen", "Aufgabenbereich schließen", "Close"), max_depth=8)
    if btn is not None:
        r = uia.describe(btn)["rect"]
        winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
        time.sleep(0.8)


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    only = sys.argv[1:]
    out: dict = json.loads(OUT.read_text(encoding="utf-8")) if OUT.exists() else {}
    w = Word()
    d = w.d
    user, initials = w.app.UserName, w.app.UserInitials
    # Tracked changes of each kind and a comment for the Überarbeitungsbereich.
    d.TrackRevisions = True
    p2 = d.Paragraphs(2).Range
    d.Range(p2.Start + 8, p2.Start + 8).InsertAfter("neuer ")
    d.Paragraphs(3).Range.Words(1).Delete()
    d.Paragraphs(2).Range.Words(4).Bold = True
    d.TrackRevisions = False
    d.Comments.Add(d.Paragraphs(2).Range.Words(7), "Ein Kommentar")
    # Two shapes for the Auswahlbereich.
    d.Shapes.AddShape(1, 300, 200, 120, 60, d.Paragraphs(2).Range)
    d.Shapes.AddTextbox(1, 300, 300, 120, 60, d.Paragraphs(2).Range)
    d.Range(0, 0).Select()
    time.sleep(0.5)
    try:
        for name, tab, control, how, pane in ITEMS:
            if only and not any(o in name for o in only):
                continue
            w.tab(tab)
            el = w.group_launcher(control) if how == "launcher" else w.control(control)
            if el is None:
                print("not found", name, flush=True)
                out[name] = {"error": "not found"}
                continue
            x, y, cw, ch = uia.describe(el)["rect"]
            kind = uia.describe(el)["type"]
            before = set(winapi.find_windows(lambda h: True))
            if how.startswith("arrow"):
                winapi.click(x + cw - 6, y + ch / 2)
            else:
                winapi.click(x + cw / 2, y + ch - 8 if kind == "SplitButton" and ch > 60 else y + ch / 2)
            time.sleep(1.5)
            missing = False
            for item in (how.split(":", 1)[1].split(">") if ":" in how else []):
                hit = uia.find(w.root, lambda dd: dd["name"] == item and dd["type"] in ("MenuItem", "Button"), max_depth=20)
                if hit is None:
                    print("no menu item", item, flush=True)
                    winapi.press("esc")
                    missing = True
                    break
                r = uia.describe(hit)["rect"]
                winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
                time.sleep(1.5)
            if missing:
                continue
            if pane:
                el2 = pane_named(w, pane)
                l, t, r, b = winapi.window_rect(w.hwnd)
                winapi.save_region_png(str(CAP / f"{name}.png"), l, t, r - l, b - t)
                if el2 is None:
                    print("no pane", name, flush=True)
                    out[name] = {"error": "no pane"}
                    continue
                pr = uia.describe(el2)["rect"]
                winapi.save_region_png(str(CAP / f"{name} only.png"), pr[0], pr[1], pr[2], pr[3])
                out[name] = scrub({"rect": [pr[0] - l, pr[1] - t, pr[2], pr[3]], "items": snapshot(el2, (pr[0], pr[1]))}, user, initials)
                print(name, len(out[name]["items"]), flush=True)
                close_pane(w, el2)
                continue
            wins = new_windows(before)
            target = next((h for h in wins if winapi.window_class(h) in ("bosa_sdm_msword", "NUIDialog", "#32770")), None) \
                or next((h for h in wins if winapi.window_class(h) == "Net UI Tool Window"), None)
            if target is None:
                print("nothing opened", name, [winapi.window_class(h) for h in wins], flush=True)
                out[name] = {"error": "nothing opened"}
                winapi.press("esc")
                continue
            l, t, r, b = winapi.window_rect(target)
            winapi.save_region_png(str(CAP / f"{name}.png"), l, t, r - l, b - t)
            items = snapshot(el, (l, t)) or snapshot(uia.element_for(target), (l, t))
            out[name] = scrub({"window": winapi.window_class(target), "size": [r - l, b - t], "items": items}, user, initials)
            print(name, len(items), flush=True)
            winapi.press("esc")
            time.sleep(0.6)
    finally:
        OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        w.close()


if __name__ == "__main__":
    main()
