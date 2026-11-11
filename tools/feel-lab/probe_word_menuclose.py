"""Probe (Step 11, part 1): how Word's menus close. With the first word selected, opens a popup (context menu on the
selection, the Schriftfarbe palette, the Aufzählungszeichen gallery), then does one thing (left click into the text,
right click there, a ribbon tab, Fett, Esc, a click into another program's window) and records: is a popup still open,
where the selection is, which ribbon tab is active, did Fett act. Own visible Word, real input under the lock.

    .venv/Scripts/python.exe probe_word_menuclose.py
"""
from __future__ import annotations

import datetime
import os
import json
import time
from pathlib import Path

from feellab import uia, winapi
from feellab.inputlock import real_input
from feellab.popups import ForeignWindow, find_named, own_popups, pick, rect_center, selected_ribbon_tab, split_arrow, wait_popup
from feellab.targets import WordTarget

HERE = Path(__file__).parent

OPENERS = ["ctx_text", "dd_fontcolor", "gallery_bullets"]
ACTIONS = ["click_text", "rclick_text", "click_ribbontab", "click_bold", "esc", "click_other_app"]


class Probe:
    def __init__(self):
        self.w = WordTarget()
        self.hwnd = int(self.w.app.ActiveWindow.Hwnd)
        top = winapi.user32.GetAncestor(self.hwnd, 2)  # GA_ROOT
        self.top = top or self.hwnd
        self.pid = winapi.window_pid(self.top)
        self.root = uia.element_for(self.top)

    def ours(self) -> bool:
        return winapi.window_pid(winapi.foreground()) == self.pid

    def reset(self) -> None:
        for _ in range(3):
            if own_popups(self.pid):
                winapi.activate(self.top)
                winapi.press("esc")
                time.sleep(0.3)
        # Enough text below and right of the popups: a click meant for "outside" must never land on a menu.
        para = "Die Schule liegt am Rand der Stadt und hat einen großen Hof mit alten Bäumen und vielen Bänken."
        self.w.doc.Content.Text = chr(13).join([para] * 24)
        self.w.doc.Content.Font.Reset()
        self.w.doc.UndoClear()
        self.text = self.w.text()
        self.w.set_selection(0, 5)
        uia.select(find_named(self.root, "TabItem", "Start"))
        winapi.activate(self.top)
        time.sleep(0.4)
        # After a click into another program Word needs one real click before its ribbon opens a dropdown.
        winapi.click(*self.w.char_center(len(para) * 3 + 10))
        time.sleep(0.3)
        self.w.set_selection(0, 5)
        n = len(self.text)
        self.pts = {
            "sel": self.w.char_center(2),
            "fontcolor": rect_center(split_arrow(self.root, "Schriftfarbe")),
            "bullets": rect_center(split_arrow(self.root, "Aufzählungszeichen")),
        }
        offs = [o for o in (n // 2, n // 3, (2 * n) // 3, n - 10) if 0 < o < n and self.text[o] not in "\r "]
        self.targets = {
            "click_text": {f"@{o}": self.w.char_center(o) for o in offs},
            "rclick_text": {f"@{o}": self.w.char_center(o) for o in offs},
            "click_ribbontab": {n: rect_center(find_named(self.root, "TabItem", n)) for n in ("Einfügen", "Layout", "Ansicht")},
            "click_bold": {"Fett": rect_center(find_named(self.root, "Button", "Fett"))},
        }

    def open(self, opener: str) -> bool:
        if not self.ours():
            winapi.activate(self.top)
        if opener == "ctx_text":
            winapi.right_click(*self.pts["sel"])
        elif opener == "dd_fontcolor":
            winapi.click(*self.pts["fontcolor"])
        elif opener == "gallery_bullets":
            winapi.click(*self.pts["bullets"])
        if wait_popup(self.pid):
            return True
        # The first click after another window had the focus sometimes only activates Word: once more.
        winapi.activate(self.top)
        time.sleep(0.3)
        return self.open_again(opener)

    def open_again(self, opener: str) -> bool:
        if opener == "ctx_text":
            winapi.right_click(*self.pts["sel"])
        elif opener == "dd_fontcolor":
            winapi.click(*self.pts["fontcolor"])
        else:
            winapi.click(*self.pts["bullets"])
        return wait_popup(self.pid)

    def act(self, action: str, other: ForeignWindow) -> str | None:
        if not self.ours():
            raise RuntimeError("our Word is not in front")
        if action == "esc":
            winapi.press("esc")
            return "esc"
        if action == "click_other_app":
            winapi.click(*other.point)
            return "other"
        name, pt = pick(self.targets[action], own_popups(self.pid))
        if name is None:
            return None
        (winapi.right_click if action == "rclick_text" else winapi.click)(*pt)
        return name

    def state(self) -> dict:
        time.sleep(0.7)
        pops = own_popups(self.pid)
        st = {"popup_open": bool(pops), "popups": pops, "foreground_ours": self.ours(),
              "ribbon_tab": selected_ribbon_tab(self.root)}
        for _ in range(3):
            if own_popups(self.pid):
                winapi.activate(self.top)
                winapi.press("esc")
                time.sleep(0.3)
        st["selection"] = list(self.w.selection())
        st["first_word_bold"] = bool(self.w.doc.Range(0, 5).Font.Bold == -1 or self.w.doc.Range(0, 5).Font.Bold is True)
        return st


def main() -> None:
    out: dict = {"date": datetime.date.today().isoformat(), "app": "Word", "cases": {}}
    with real_input("probe_word_menuclose"):
        pr = Probe()
        try:
            other = ForeignWindow()
        except Exception:
            pr.w.close()
            raise
        try:
            import os
            only = set(filter(None, os.environ.get("ONLY", "").split(",")))
            for opener in OPENERS:
                for action in ACTIONS:
                    if only and f"{opener}+{action}" not in only:
                        continue
                    pr.reset()
                    key = f"{opener}+{action}"
                    if not pr.open(opener):
                        out["cases"][key] = {"error": "popup did not open"}
                        print(key, "NOT OPENED")
                        continue
                    target = pr.act(action, other)
                    if target is None:
                        out["cases"][key] = {"error": "every target under the popup"}
                        print(key, "NO TARGET")
                        continue
                    st = pr.state()
                    st["target"] = target
                    out["cases"][key] = st
                    print(key, {k: st[k] for k in ("target", "popup_open", "selection", "ribbon_tab", "first_word_bold")})
                    winapi.move(1900, 1950)
        finally:
            other.close()
            pr.w.close()
    path = HERE / "results" / f"{out['date']}-word-menuclose.json"
    if os.environ.get("ONLY") and path.exists():
        old = json.loads(path.read_text(encoding="utf-8"))
        old["cases"].update(out["cases"])
        out = old
    path.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print("wrote", path.name)


if __name__ == "__main__":
    main()
