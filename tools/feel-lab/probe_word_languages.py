"""Probe: Word's proofing languages for Step 4b F. Records through COM every language Word
knows (local name, id) and which have spelling, grammar, hyphenation and thesaurus tools on
this PC, and whether Word detects languages by itself; then, in our own visible Word
instance with real clicks and keys, the Sprache menu and dialog, the Thesaurus pane
(Umschalt+F7), the context menu's Synonyme, Manuelle Silbentrennung and what F7 opens.
Takes over mouse and screen. The Office user name is replaced by "Autor" (public repo).
Screenshots: captures/word-dialogs/ (local); names: results/<date>-word-languages.json.

    .venv/Scripts/python.exe probe_word_languages.py [part …]   (parts: com ui detect options)
"""
from __future__ import annotations

import datetime
import json
import sys
import time
from pathlib import Path

from feellab import uia, winapi
from probe_word_context import capture, items_of, popups
from probe_word_dialogs import CAP, Word, new_windows, snapshot
from probe_word_panes import close_pane, pane_named, scrub

HERE = Path(__file__).parent
OUT = HERE / "results" / f"{datetime.date.today()}-word-languages.json"
DIALOG_CLASSES = ("bosa_sdm_msword", "NUIDialog", "#32770")


def com_part(app) -> dict:
    """Every language with its proofing tools; Word's automatic language detection."""
    langs = []
    for L in app.Languages:
        e = {"id": L.ID, "name": L.NameLocal, "en": L.Name}
        for key, attr in (("spelling", "ActiveSpellingDictionary"), ("grammar", "ActiveGrammarDictionary"),
                          ("hyphenation", "ActiveHyphenationDictionary"), ("thesaurus", "ActiveThesaurusDictionary")):
            try:
                dic = getattr(L, attr)
                e[key] = dic.Name if dic is not None else None
            except Exception:  # noqa: BLE001
                e[key] = None
        langs.append(e)
    return {"check_language": bool(app.CheckLanguage), "languages": langs}


def dialog_after(before: set[int], timeout: float = 4.0) -> int | None:
    end = time.time() + timeout
    while time.time() < end:
        hit = next((h for h in new_windows(before) if winapi.window_class(h) in DIALOG_CLASSES), None)
        if hit:
            time.sleep(0.6)
            return hit
        time.sleep(0.2)
    return None


def record_dialog(hwnd: int, name: str) -> dict:
    l, t, r, b = winapi.window_rect(hwnd)
    winapi.save_region_png(str(CAP / f"{name}.png"), l, t, r - l, b - t)
    return {"title": winapi.window_title(hwnd), "size": [r - l, b - t], "items": snapshot(uia.element_for(hwnd), (l, t))}


def region(w: Word, title_el) -> dict:
    """A docked pane: everything below its title bar within its width (UIA gives the title
    bar and the content as separate elements)."""
    tx, ty, tw, th = uia.describe(title_el)["rect"]
    l, t, r, b = winapi.window_rect(w.hwnd)
    f = w.root.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
    items = []
    for i in range(f.Length):
        dd = uia.describe(f.GetElement(i))
        x, y, ww, hh = dd["rect"]
        if ww <= 0 or x < tx - 12 or x + ww > tx + tw + 30 or y < ty or (not dd["name"] and dd["type"] not in ("Edit", "ComboBox", "List", "Tree")):
            continue
        items.append({"type": dd["type"], "name": dd["name"], "key": dd["key"], "enabled": dd["enabled"], "rect": [x - tx, y - ty, ww, hh]})
    bottom = max((it["rect"][1] + it["rect"][3] for it in items), default=th)
    winapi.save_region_png(str(CAP / f"{uia.describe(title_el)['name'] or 'pane'} region.png"), tx - 12, ty, tw + 40, min(bottom, b - ty))
    return {"rect": [tx - l, ty - t, tw, th], "items": items}


def ribbon_control(w: Word, name: str):
    return uia.find(w.root, lambda x: x["name"] == name and x["type"] in ("Button", "SplitButton", "MenuItem") and x["rect"][2] > 0, max_depth=18)


def menu_of_control(w: Word, control: str) -> tuple[list, dict]:
    ctl = ribbon_control(w, control)
    if ctl is None:
        return [], {"error": "not found"}
    before = set(popups())
    r = uia.describe(ctl)["rect"]
    winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
    hs = []
    end = time.time() + 3.0
    while time.time() < end and not hs:
        time.sleep(0.3)
        hs = [h for h in popups() if h not in before]
    time.sleep(0.5)
    hs = [h for h in popups() if h not in before]
    items = []
    # UIA hangs the open menu under the ribbon control (as in probe_word_dialogs).
    for root in [ctl] + [uia.element_for(h) for h in hs]:
        if items:
            break
        f = root.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
        for i in range(f.Length):
            el = f.GetElement(i)
            dd = uia.describe(el)
            if dd["name"] and dd["type"] in ("MenuItem", "Button", "CheckBox", "ListItem"):
                items.append((el, dd))
    if hs:
        l, t, rr, b = winapi.window_rect(hs[0])
        winapi.save_region_png(str(CAP / f"{control} menu.png"), l, t, rr - l, b - t)
    return items, {"items": [{"type": d["type"], "name": d["name"], "key": d["key"], "enabled": d["enabled"]} for _, d in items]}


def under(el) -> list:
    """Named menu entries below an element (an open submenu hangs below its item)."""
    out = []
    f = el.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
    for i in range(f.Length):
        e = f.GetElement(i)
        dd = uia.describe(e)
        if dd["name"] and dd["type"] in ("MenuItem", "Button", "CheckBox", "ListItem", "SplitButton"):
            out.append((e, dd))
    return out


def click_el(el) -> None:
    r = uia.describe(el)["rect"]
    winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)


def ui_part(w: Word, out: dict) -> None:
    app, d = w.app, w.d
    win = d.Windows(1)
    d.Content.Text = "\r".join([
        "Das Haus ist schön und groß.",
        "Die Fotosynthese ist ein Vorgang, bei dem Pflanzen mithilfe von Lichtenergie aus Kohlenstoffdioxid und Wasser energiereiche Kohlenhydrate herstellen.",
        "Das ist ein Satz mit einem Fehlr und das sind falsch Satz.",
    ])
    d.ShowSpellingErrors = True
    d.ShowGrammaticalErrors = True

    def point(rg) -> tuple[float, float]:
        left, top, width, height = win.GetPoint(0, 0, 0, 0, rg)
        return float(left) + 4, float(top) + float(height) / 2

    def find(text: str):
        rg = d.Content.Duplicate
        rg.Find.Execute(FindText=text)
        return rg

    # Überprüfen > Sprache: the menu and the dialog.
    w.tab("Überprüfen")
    items, rec = menu_of_control(w, "Sprache")
    out["language menu"] = rec
    print("language menu", [x["name"] for x in rec.get("items", [])], flush=True)
    # In a narrow window Word folds Übersetzen and Sprache into one menu: Sprache is a submenu.
    target = next((el for el, dd in items if dd["name"].startswith("Sprache für die Korrekturhilfen")), None)
    sub_el = next((el for el, dd in items if dd["name"] == "Sprache"), None)
    if target is None and sub_el is not None:
        click_el(sub_el)
        time.sleep(1.3)
        sub = under(sub_el)
        rec["submenu"] = [{"type": dd["type"], "name": dd["name"], "key": dd["key"], "enabled": dd["enabled"]} for _, dd in sub]
        print("language submenu", [x["name"] for x in rec["submenu"]], flush=True)
        target = next((el for el, dd in sub if dd["name"].startswith("Sprache für die Korrekturhilfen")), None)
    if target is not None:
        before = set(winapi.find_windows(lambda h: True))
        click_el(target)
        dlg = dialog_after(before)
        if dlg:
            out["language dialog"] = record_dialog(dlg, "language dialog")
            print("language dialog", out["language dialog"]["title"], len(out["language dialog"]["items"]), flush=True)
            winapi.press("esc")
            time.sleep(0.8)
    else:
        winapi.press("esc")

    # Thesaurus (Umschalt+F7) on "schön".
    rg = find("schön")
    rg.Select()
    time.sleep(0.4)
    winapi.press("shift+f7")
    time.sleep(3.0)
    pane = pane_named(w, "Thesaurus")
    l, t, r, b = winapi.window_rect(w.hwnd)
    winapi.save_region_png(str(CAP / "thesaurus pane.png"), l, t, r - l, b - t)
    if pane is not None:
        out["thesaurus pane"] = region(w, pane)
        print("thesaurus pane", len(out["thesaurus pane"]["items"]), flush=True)
        # A synonym's menu (hover shows its arrow: Einfügen, Kopieren).
        entry = next((it for it in out["thesaurus pane"]["items"] if it["name"] == "hübsch"), None)
        if entry is not None:
            tx, ty = uia.describe(pane)["rect"][:2]
            x, y, ww, hh = entry["rect"]
            winapi.move(tx + x + ww / 2, ty + y + hh / 2)
            time.sleep(0.8)
            before = set(popups())
            winapi.click(tx + x + ww - 8, ty + y + hh / 2)
            time.sleep(1.2)
            hs = [h for h in popups() if h not in before]
            if hs:
                out["thesaurus entry menu"] = [{"type": dd["type"], "name": dd["name"]} for h in hs for _, dd in items_of(h)]
                winapi.press("esc")
                time.sleep(0.4)
        close_pane(w, pane)
    else:
        out["thesaurus pane"] = {"error": "no pane"}
        print("no thesaurus pane", flush=True)

    # The context menu on "schön": Synonyme opens a submenu.
    x, y = point(d.Range(find("schön").Start + 2, find("schön").Start + 2))
    winapi.click(x, y)
    time.sleep(0.4)
    before = set(popups())
    winapi.right_click(x, y)
    time.sleep(1.4)
    menu = [h for h in popups() if h not in before]
    syn = None
    for h in menu:
        for el, dd in items_of(h):
            if dd["name"] == "Synonyme":
                syn = (el, dd)
    if syn is not None:
        click_el(syn[0])
        time.sleep(1.4)
        subs = [h for h in popups() if h not in before and h not in menu]
        sub = [{"type": dd["type"], "name": dd["name"], "enabled": dd["enabled"]} for _, dd in under(syn[0])]
        if not sub:
            for h in subs:
                sub += [{"type": dd["type"], "name": dd["name"], "enabled": dd["enabled"]} for _, dd in under(uia.element_for(h))]
        out["context synonyms"] = sub
        if subs:
            l2, t2, r2, b2 = winapi.window_rect(subs[0])
            winapi.save_region_png(str(CAP / "synonyms submenu.png"), l2 - 10, t2 - 10, r2 - l2 + 20, b2 - t2 + 20)
        print("synonyms", [x["name"] for x in sub][:30], flush=True)
    winapi.press("esc")
    time.sleep(0.3)
    winapi.press("esc")
    time.sleep(0.4)

    # Layout > Silbentrennung > Manuell: the dialog for the first word Word proposes.
    w.tab("Layout")
    items, rec = menu_of_control(w, "Silbentrennung")
    manual = next((el for el, dd in items if dd["name"] == "Manuell"), None)
    if manual is not None:
        d.Range(0, 0).Select()
        before = set(winapi.find_windows(lambda h: True))
        click_el(manual)
        dlg = dialog_after(before, 6.0)
        if dlg:
            out["manual hyphenation"] = record_dialog(dlg, "manual hyphenation")
            print("manual hyphenation", out["manual hyphenation"]["title"], flush=True)
            winapi.press("esc")
            time.sleep(0.8)
            # Word may ask whether to go on or report that it is done.
            dlg2 = dialog_after(before | {dlg}, 2.0)
            if dlg2:
                out["manual hyphenation end"] = record_dialog(dlg2, "manual hyphenation end")
                winapi.press("esc")
                time.sleep(0.6)
        else:
            out["manual hyphenation"] = {"error": "no dialog"}
    else:
        winapi.press("esc")

    # F7: what Word opens for the first error.
    d.Range(0, 0).Select()
    before = set(winapi.find_windows(lambda h: True))
    winapi.press("f7")
    time.sleep(3.5)
    dlg = dialog_after(before, 1.0)
    if dlg:
        out["f7"] = record_dialog(dlg, "f7 dialog")
        winapi.press("esc")
    else:
        l, t, r, b = winapi.window_rect(w.hwnd)
        winapi.save_region_png(str(CAP / "f7 pane.png"), l, t, r - l, b - t)
        for name in ("Editor", "Rechtschreibung", "Grammatik", "Rechtschreibung und Grammatik"):
            pane = pane_named(w, name)
            if pane is not None:
                out["f7"] = {"pane": name, **region(w, pane)}
                close_pane(w, pane)
                break
        else:
            out["f7"] = {"error": "nothing found"}
    print("f7", out["f7"].get("title") or out["f7"].get("pane") or out["f7"], flush=True)
    time.sleep(0.8)
    # Word's grammar and spelling errors in the sample (offline checker).
    out["errors"] = {
        "spelling": [e.Text for e in d.SpellingErrors],
        "grammar": [e.Text for e in d.GrammaticalErrors],
    }


def detect_part(w: Word, out: dict) -> None:
    """Typed French, Spanish and English sentences in a German document: what language Word gives them."""
    d = w.d
    d.Content.Text = "Ein deutscher Satz zu Beginn des Dokuments."
    d.Content.LanguageID = 1031
    end = d.Content.End - 1
    d.Range(end, end).Select()
    winapi.activate(w.hwnd)
    time.sleep(0.5)
    sentences = {
        "fr": "Nous sommes allés au marché avec nos amis pour acheter des légumes frais.",
        "es": "Mañana vamos a la playa con nuestros amigos porque hace mucho calor.",
        "en": "We went to the market with our friends to buy some fresh vegetables.",
    }
    res = {}
    for key, text in sentences.items():
        winapi.press("enter")
        winapi.type_text(text, gap=0.02)
        time.sleep(4.0)
        p = d.Paragraphs(d.Paragraphs.Count).Range
        res[key] = {"language": p.LanguageID, "detected": bool(p.LanguageDetected)}
        print("detect", key, res[key], flush=True)
    out["detection"] = res


def options_part(w: Word, out: dict) -> None:
    """Datei > Optionen: the pages Dokumentprüfung and Sprache."""
    # Alt, D opens Datei; O is Optionen.
    winapi.activate(w.hwnd)
    winapi.press("alt+d")
    time.sleep(2.0)
    before = set(winapi.find_windows(lambda h: True))
    winapi.press("o")
    dlg = dialog_after(before, 8.0)
    if not dlg:
        out["options"] = {"error": "no dialog"}
        winapi.press("esc")
        return
    root = uia.element_for(dlg)
    res = {}
    for page in ("Dokumentprüfung", "Sprache"):
        item = uia.find(root, lambda x: x["name"] == page and x["type"] in ("ListItem", "TabItem", "Button"), max_depth=14)
        if item is None:
            res[page] = {"error": "no page"}
            continue
        click_el(item)
        time.sleep(1.5)
        res[page] = record_dialog(dlg, f"options {page}")
        print("options", page, len(res[page]["items"]), flush=True)
    out["options"] = res
    winapi.press("esc")
    time.sleep(1.0)


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    parts = sys.argv[1:] or ["com", "ui", "detect"]
    out: dict = json.loads(OUT.read_text(encoding="utf-8")) if OUT.exists() else {}
    w = Word()
    user, initials = w.app.UserName, w.app.UserInitials
    try:
        if "com" in parts:
            out["com"] = com_part(w.app)
            print("languages", len(out["com"]["languages"]), "check_language", out["com"]["check_language"], flush=True)
        if "ui" in parts:
            ui_part(w, out)
        if "detect" in parts:
            detect_part(w, out)
        if "options" in parts:
            options_part(w, out)
    finally:
        OUT.write_text(json.dumps(scrub(out, user, initials), ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        w.close()


if __name__ == "__main__":
    main()
