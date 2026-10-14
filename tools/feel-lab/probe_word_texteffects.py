"""Probe: Word's Texteffekte und Typografie (Start). Every entry of the preset gallery and of
the Schatten, Spiegelung and Leuchten submenus is applied to a word in our own visible Word
instance (real clicks; nested menus are read with ElementFromPoint), the document is saved
and the run's w14 XML recorded. Also records the menu's other submenus (names only).
Takes over mouse and screen. Result: results/<date>-word-texteffects.json.

    .venv/Scripts/python.exe probe_word_texteffects.py <scratch folder> [part …]
"""
from __future__ import annotations

import datetime
import json
import re
import sys
import time
import zipfile
from pathlib import Path

from feellab import uia, winapi
from probe_word_dialogs import Word, new_windows

HERE = Path(__file__).parent
OUT = HERE / "results" / f"{datetime.date.today()}-word-texteffects.json"


def at_point(x: int, y: int) -> dict | None:
    try:
        import ctypes.wintypes as W
        el = uia._uia.ElementFromPoint(W.POINT(int(x), int(y)))
    except Exception:  # noqa: BLE001
        return None
    return uia.describe(el) if el else None


def scan(rect: tuple[int, int, int, int], step: int = 14) -> list[dict]:
    """The items of a popup window, found by asking for the element under a grid of points."""
    l, t, r, b = rect
    seen: dict[str, dict] = {}
    for y in range(t + 4, b - 2, step):
        for x in range(l + 4, r - 2, step):
            d = at_point(x, y)
            if d and d["name"] and d["type"] in ("ListItem", "MenuItem", "Button") and d["name"] not in seen:
                seen[d["name"]] = d
    return list(seen.values())


class Probe:
    def __init__(self, scratch: Path):
        self.w = Word()
        self.d = self.w.d
        self.scratch = scratch
        self.d.Content.Text = "Beispiel Effekt Text"
        self.n = 0

    def select_word(self) -> None:
        self.d.Range(9, 15).Select()  # "Effekt"
        time.sleep(0.2)

    def reset(self) -> None:
        self.d.Content.Font.Reset()
        time.sleep(0.2)

    def open_menu(self) -> tuple[int, int, int, int] | None:
        self.w.tab("Start")
        ctl = self.w.control("Texteffekte und Typografie")
        if ctl is None:
            return None
        r = uia.describe(ctl)["rect"]
        before = set(winapi.find_windows(lambda h: True))
        winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
        time.sleep(1.2)
        win = next((h for h in new_windows(before) if winapi.window_class(h) == "Net UI Tool Window"), None)
        return winapi.window_rect(win) if win else None

    def open_sub(self, name: str) -> tuple[int, int, int, int] | None:
        rect = self.open_menu()
        if not rect:
            return None
        item = uia.find(self.w.root, lambda d: d["name"] == name and d["type"] == "MenuItem", max_depth=22)
        if item is None:
            winapi.press("esc")
            return None
        r = uia.describe(item)["rect"]
        before = set(winapi.find_windows(lambda h: True))
        winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
        time.sleep(1.2)
        wins = [h for h in new_windows(before) if winapi.window_class(h) == "Net UI Tool Window"]
        return winapi.window_rect(wins[-1]) if wins else None

    def xml(self) -> str:
        self.n += 1
        p = self.scratch / f"fx{self.n}.docx"
        self.d.SaveAs2(str(p), 16)
        doc = zipfile.ZipFile(p).read("word/document.xml").decode("utf-8")
        runs = re.findall(r"<w:r\b[^>]*>(?:(?!</w:r>).)*?Effekt(?:(?!</w:r>).)*?</w:r>", doc, re.S)
        rpr = re.search(r"<w:rPr>(.*?)</w:rPr>", runs[0], re.S) if runs else None
        return rpr.group(1) if rpr else ""

    def gallery(self, sub: str | None) -> list[dict]:
        """Applies each entry of the presets (sub None) or of a submenu; returns names and XML."""
        rect = self.open_menu() if sub is None else self.open_sub(sub)
        if not rect:
            print("cannot open", sub)
            return []
        items = [d for d in scan(rect) if d["type"] == "ListItem"]
        names = [d["name"] for d in items]
        print(sub or "presets", len(names), names[:4], flush=True)
        winapi.press("esc")
        time.sleep(0.3)
        winapi.press("esc")
        time.sleep(0.4)
        out = []
        for d in items:
            self.reset()
            self.select_word()
            rect2 = self.open_menu() if sub is None else self.open_sub(sub)
            if not rect2:
                continue
            x, y, w_, h_ = d["rect"]
            winapi.click(x + w_ / 2, y + h_ / 2)
            time.sleep(0.9)
            out.append({"name": d["name"], "xml": self.xml()})
            print("  ", d["name"], len(out[-1]["xml"]), flush=True)
        return out

    def names(self, sub: str) -> list[str]:
        rect = self.open_sub(sub)
        if not rect:
            return []
        n = [d["name"] for d in scan(rect, 10)]
        winapi.press("esc")
        time.sleep(0.3)
        winapi.press("esc")
        time.sleep(0.3)
        return n


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    scratch = Path(sys.argv[1])
    parts = sys.argv[2:] or ["presets", "Schatten", "Spiegelung", "Leuchten", "names"]
    out: dict = json.loads(OUT.read_text(encoding="utf-8")) if OUT.exists() else {}
    p = Probe(scratch)
    try:
        if "presets" in parts:
            out["presets"] = p.gallery(None)
        for sub in ("Schatten", "Spiegelung", "Leuchten"):
            if sub in parts:
                out[sub] = p.gallery(sub)
                OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))
        if "names" in parts:
            out["names"] = {sub: p.names(sub) for sub in ("Kontur", "Schatten", "Spiegelung", "Leuchten", "Zahlenformatvorlagen", "Ligaturen", "Stil-Sets")}
    finally:
        OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        p.w.close()


if __name__ == "__main__":
    main()
