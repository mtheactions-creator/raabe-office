"""Probe: Word's Bildformat galleries. Word's gallery popups expose no items to UI
Automation, so this opens each gallery with a click, takes a screenshot, and applies chosen
entries by clicking at their position (given as fractions of the popup) to record the XML
Word writes for them (blipFill, spPr, effect extent), undoing after each. Own Word instance;
takes over mouse and screen. Screenshots go to captures/word-picture/ (local only), the XML
to results/<date>-word-picture.json.

    .venv/Scripts/python.exe probe_word_picture.py shots            # screenshots of the popups
    .venv/Scripts/python.exe probe_word_picture.py apply <gallery>    # entries listed in CLICKS
"""
from __future__ import annotations

import datetime
import json
import re
import struct
import sys
import tempfile
import time
import zlib
from pathlib import Path

from feellab import uia, winapi
from feellab.targets import WINDOW_RECT

HERE = Path(__file__).parent
CAP = HERE / "captures" / "word-picture"
CAP.mkdir(parents=True, exist_ok=True)
PNG = Path(tempfile.gettempdir()) / "raabe-picture-probe.png"

# Ribbon buttons of the Bildformat tab (UIA names) whose popups are recorded.
BUTTONS = {
    "corrections": "Korrekturen", "color": "Farbe", "transparency": "Transparenz", "styles": "Schnellformatvorlagen",
    "border": "Bildrahmen", "effects": "Bildeffekte", "position": "Position", "wrap": "Textumbruch", "rotate": "Drehen",
    "align": "Ausrichten", "crop": "Zuschneiden",
}

# Entries to apply: gallery → list of (label, x fraction, y fraction) inside the popup.
CLICKS: dict[str, list[tuple[str, float, float]]] = {}


def make_png(path: Path) -> None:
    """A colourful test picture (gradients), so colour effects change it visibly."""
    w, h = 160, 120
    rows = []
    for y in range(h):
        row = bytearray([0])
        for x in range(w):
            row += bytes((int(255 * x / w), int(255 * y / h), 128))
        rows.append(bytes(row))
    def chunk(t, d):
        return struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
    path.write_bytes(b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)) +
                     chunk(b"IDAT", zlib.compress(b"".join(rows))) + chunk(b"IEND", b""))


def popups() -> set[int]:
    return set(winapi.find_windows(lambda h: winapi.window_class(h) == "Net UI Tool Window"))


def picture_xml(d) -> dict:
    x = d.Content.WordOpenXML
    def grab(tag):
        m = re.search(rf"<{tag}[ >].*?</{tag}>", x, re.S)
        return m.group(0) if m else ""
    docpr = re.search(r"<wp:docPr [^>]*/?>", x)
    extent = re.search(r"<wp:effectExtent [^>]*/>", x)
    anchor = re.search(r"<wp:anchor [^>]*>.*?<wp:extent", x, re.S)
    return {"blipFill": re.sub(r'r:embed="[^"]*"', 'r:embed="rId"', grab("pic:blipFill")), "spPr": grab("pic:spPr"),
            "effectExtent": extent.group(0) if extent else "", "docPr": docpr.group(0) if docpr else "",
            "anchor": anchor.group(0) if anchor else ""}


class Word:
    def __init__(self):
        import win32com.client
        make_png(PNG)
        self.app = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx("Word.Application"))
        self.app.Visible = True
        d = self.app.Documents.Add()
        d.Content.Text = "Bild:"
        d.Paragraphs(1).Range.InsertParagraphAfter()
        d.Paragraphs(2).Range.InlineShapes.AddPicture(str(PNG), False, True)
        self.d = d
        win = d.Windows(1)
        win.View.Type = 3
        winapi.place(win.Hwnd, *WINDOW_RECT)
        winapi.activate(win.Hwnd)
        time.sleep(0.8)
        self.root = uia.element_for(win.Hwnd)
        self.base = picture_xml(d)

    def select(self) -> None:
        if self.d.InlineShapes.Count:
            self.d.InlineShapes(1).Select()
        elif self.d.Shapes.Count:
            self.d.Shapes(1).Select()
        time.sleep(0.3)
        if not hasattr(self, "tab"):
            tab = uia.find(self.root, lambda x: x["type"] == "TabItem" and x["name"] == "Bildformat", max_depth=14)
            self.tab = uia.describe(tab)["rect"] if tab is not None else None
            if tab is not None:
                uia.select(tab)
                time.sleep(0.6)
            # The tab's buttons, once (searching Word's whole tree is slow).
            self.buttons = {}
            names = set(BUTTONS.values())
            stack = [(self.root, 0)]
            while stack:
                e, depth = stack.pop()
                dd = uia.describe(e)
                if dd["name"] in names and dd["type"] in ("MenuItem", "SplitButton", "Button") and dd["name"] not in self.buttons:
                    self.buttons[dd["name"]] = (dd["type"], dd["rect"])
                if depth < 18:
                    stack.extend((c, depth + 1) for c in uia.children(e))
            print("buttons", {k: v[1] for k, v in self.buttons.items()})
        elif self.tab:
            x, y, w, h = self.tab
            winapi.click(x + w / 2, y + h / 2)
            time.sleep(0.5)

    def open(self, name: str):
        self.select()
        if name not in self.buttons:
            return None
        kind, (bx, by, bw, bh) = self.buttons[name]
        before = popups()
        if kind == "SplitButton":
            winapi.click(bx + bw / 2, by + bh - 8)
        else:
            winapi.click(bx + bw / 2, by + bh / 2)
        time.sleep(1.3)
        new = [h for h in popups() if h not in before]
        if not new:
            return None
        return max(new, key=lambda h: (winapi.window_rect(h)[2] - winapi.window_rect(h)[0]) * (winapi.window_rect(h)[3] - winapi.window_rect(h)[1]))

    def close(self) -> None:
        self.d.Close(0)
        self.app.Quit(0)


def shots(w: Word, keys: list[str]) -> None:
    for key, name in BUTTONS.items():
        if keys and key not in keys:
            continue
        pop = w.open(name)
        if pop is None:
            print("no popup for", name)
            winapi.press("esc")
            continue
        l, t, r, b = winapi.window_rect(pop)
        winapi.save_region_png(str(CAP / f"{key}.png"), l, t, r - l, b - t)
        print(key, "popup", (l, t, r - l, b - t))
        winapi.press("esc")
        time.sleep(0.4)


def subshots(w: Word, key: str, rows: list[tuple[str, float, float]]) -> None:
    """Submenus of a menu (Bildeffekte > Schatten …): hover each row, screenshot the new popup."""
    for label, fx, fy in rows:
        pop = w.open(BUTTONS[key])
        if pop is None:
            continue
        l, t, r, b = winapi.window_rect(pop)
        known = popups()
        winapi.click(l + fx * (r - l), t + fy * (b - t))
        time.sleep(1.3)
        new = [h for h in popups() if h not in known and (winapi.window_rect(h)[3] - winapi.window_rect(h)[1]) > 150]
        if new:
            h = max(new, key=lambda h: (winapi.window_rect(h)[2] - winapi.window_rect(h)[0]) * (winapi.window_rect(h)[3] - winapi.window_rect(h)[1]))
            l2, t2, r2, b2 = winapi.window_rect(h)
            winapi.save_region_png(str(CAP / f"{key}-{label}.png"), l2, t2, r2 - l2, b2 - t2)
            print(key, label, "submenu", (l2, t2, r2 - l2, b2 - t2))
        winapi.press("esc")
        time.sleep(0.2)
        winapi.press("esc")
        time.sleep(0.3)


def apply_sub(w: Word, key: str, row: tuple[float, float], clicks: list[tuple[str, float, float]]) -> list[dict]:
    """Entries of a submenu: open the menu, hover the row, click the entry in the submenu."""
    out = []
    for label, fx, fy in clicks:
        pop = w.open(BUTTONS[key])
        if pop is None:
            out.append({"label": label, "error": "no popup"})
            continue
        l, t, r, b = winapi.window_rect(pop)
        known = popups()
        winapi.click(l + row[0] * (r - l), t + row[1] * (b - t))
        time.sleep(1.2)
        new = [h for h in popups() if h not in known and (winapi.window_rect(h)[3] - winapi.window_rect(h)[1]) > 150]
        if not new:
            out.append({"label": label, "error": "no submenu"})
            winapi.press("esc")
            continue
        h = max(new, key=lambda h: (winapi.window_rect(h)[2] - winapi.window_rect(h)[0]) * (winapi.window_rect(h)[3] - winapi.window_rect(h)[1]))
        l2, t2, r2, b2 = winapi.window_rect(h)
        # Into the submenu horizontally first, so it stays open.
        winapi.move(l2 + 10, t + row[1] * (b - t))
        time.sleep(0.3)
        winapi.click(l2 + fx * (r2 - l2), t2 + fy * (b2 - t2))
        time.sleep(1.0)
        xml = picture_xml(w.d)
        out.append({"label": label, "xml": xml, "changed": xml != w.base})
        print(key, label, "changed" if xml != w.base else "unchanged")
        winapi.press("esc")
        time.sleep(0.2)
        for _ in range(3):
            if picture_xml(w.d) == w.base:
                break
            w.d.Undo(1)
            time.sleep(0.3)
    return out


def tooltip_text(known: set[int]) -> str:
    """The text of a tooltip that appeared (a small new popup), through UI Automation."""
    for h in popups():
        if h in known:
            continue
        el = uia.element_for(h)
        names = [uia.describe(el)["name"]]
        try:
            allc = el.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
            names += [uia.describe(allc.GetElement(i))["name"] for i in range(allc.Length)]
        except Exception:  # noqa: BLE001
            pass
        text = " | ".join(n for n in names if n)
        if text:
            return text
    return ""


def tips(w: Word, key: str, clicks: list[tuple[str, float, float]], row: tuple[float, float] | None = None) -> dict:
    """Tooltips (Word's names) of a gallery's entries: hover each and read the tooltip."""
    out = {}
    pop = w.open(BUTTONS[key])
    if pop is None:
        return out
    l, t, r, b = winapi.window_rect(pop)
    if row is not None:
        known = popups()
        winapi.click(l + row[0] * (r - l), t + row[1] * (b - t))
        time.sleep(1.2)
        new = [h for h in popups() if h not in known and (winapi.window_rect(h)[3] - winapi.window_rect(h)[1]) > 150]
        if not new:
            return out
        winapi.move(winapi.window_rect(new[0])[0] + 10, t + row[1] * (b - t))
        l, t, r, b = winapi.window_rect(new[0])
    for label, fx, fy in clicks:
        known = popups()
        winapi.move(l + fx * (r - l), t + fy * (b - t) + 2)
        time.sleep(0.2)
        winapi.move(l + fx * (r - l), t + fy * (b - t))
        time.sleep(1.4)
        out[label] = tooltip_text(known)
        print(key, label, out[label])
    winapi.press("esc")
    time.sleep(0.2)
    winapi.press("esc")
    return out


def apply(w: Word, key: str, clicks: list[tuple[str, float, float]]) -> list[dict]:
    out = []
    for label, fx, fy in clicks:
        pop = w.open(BUTTONS[key])
        if pop is None:
            out.append({"label": label, "error": "no popup"})
            continue
        l, t, r, b = winapi.window_rect(pop)
        winapi.click(l + fx * (r - l), t + fy * (b - t))
        time.sleep(1.0)
        xml = picture_xml(w.d)
        out.append({"label": label, "xml": xml, "changed": xml != w.base})
        print(key, label, "changed" if xml != w.base else "unchanged")
        winapi.press("esc")
        time.sleep(0.2)
        for _ in range(3):
            if picture_xml(w.d) == w.base:
                break
            w.d.Undo(1)
            time.sleep(0.3)
    return out


def main() -> None:
    mode = sys.argv[1] if len(sys.argv) > 1 else "shots"
    w = Word()
    results = {}
    try:
        results["base"] = w.base
        if mode == "shots":
            shots(w, sys.argv[2:])
        elif mode == "subshots":
            spec = json.loads((HERE / "probe_word_picture.json").read_text(encoding="utf-8"))
            for key in sys.argv[2:]:
                subshots(w, key, [tuple(x) for x in spec[key + "_rows"]])
        elif mode == "tips":
            spec = json.loads((HERE / "probe_word_picture.json").read_text(encoding="utf-8"))
            for name in sys.argv[2:]:
                sub = spec[name]
                if isinstance(sub, dict):
                    results[f"tips_{name}"] = tips(w, sub["menu"], [tuple(x) for x in sub["clicks"]], tuple(sub["row"]))
                else:
                    results[f"tips_{name}"] = tips(w, name, [tuple(x) for x in sub])
        elif mode == "applysub":
            spec = json.loads((HERE / "probe_word_picture.json").read_text(encoding="utf-8"))
            for name in sys.argv[2:]:
                sub = spec[name]
                results[name] = apply_sub(w, sub["menu"], tuple(sub["row"]), [tuple(x) for x in sub["clicks"]])
        else:
            for key in sys.argv[2:]:
                spec = json.loads((HERE / "probe_word_picture.json").read_text(encoding="utf-8"))[key]
                results[key] = apply(w, key, [tuple(x) for x in spec])
    finally:
        w.close()
    if mode not in ("shots", "subshots"):
        out = HERE / "results" / f"{datetime.date.today().isoformat()}-word-picture.json"
        old = json.loads(out.read_text(encoding="utf-8")) if out.exists() else {}
        old.update(results)
        out.write_text(json.dumps(old, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
