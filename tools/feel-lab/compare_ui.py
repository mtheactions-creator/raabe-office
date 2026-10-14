"""Screen comparison (D-038, package A): Word and RaabeWrite on the screen at the same size,
Word above, RaabeWrite below. For every ribbon tab it captures both ribbons, matches the
controls by name (Word through UI Automation, RaabeWrite through the DOM) and lists size
and position differences; with --hover it also captures every button with a menu in the
normal and hover states, on its main part and on its arrow. Takes over mouse and screen.

    .venv/Scripts/python.exe compare_ui.py [--hover] [Start Einfügen …]

Output in captures/compare-ui/ (not committed: Microsoft screenshots stay local); the
numbers also go to results/<date>-compare-ui.json.
"""

from __future__ import annotations

import datetime as dt
import json
import sys
import time
from pathlib import Path

from PIL import Image, ImageDraw

from feellab import targets, uia, winapi

OUT = Path(__file__).parent / "captures" / "compare-ui"
RESULTS = Path(__file__).parent / "results"
TABS = ["Start", "Einfügen", "Entwurf", "Layout", "Referenzen", "Überprüfen", "Ansicht"]
WORD_RECT = (0, 0, 3000, 1000)
RAABE_RECT = (0, 1060, 3000, 1000)
CONTROL_TYPES = {"Button", "SplitButton", "MenuItem", "ComboBox", "CheckBox", "Edit", "ListItem"}


def grab(path: Path, rect: tuple[int, int, int, int]) -> Image.Image:
    winapi.save_region_png(str(path), *rect)
    return Image.open(path)


def norm(name: str) -> str:
    return " ".join(name.replace("…", "").replace("...", "").split()).lower()


class WordSide:
    def __init__(self) -> None:
        self.t = targets.WordTarget("short")
        winapi.place(self.t.hwnd, *WORD_RECT)
        winapi.activate(self.t.hwnd)
        time.sleep(1.0)
        self.root = uia.element_for(self.t.hwnd)

    def origin(self) -> tuple[int, int]:
        l, t, _r, _b = winapi.window_rect(self.t.hwnd)
        return l, t

    def select_tab(self, name: str) -> None:
        el = uia.find(self.root, lambda d: d["type"] == "TabItem" and d["name"] == name)
        if el is None:
            raise KeyError(name)
        uia.select(el)
        time.sleep(0.9)

    def ribbon_bottom(self) -> int:
        el = uia.find(self.root, lambda d: d["name"] == "Unteres Menüband")
        r = uia.describe(el)["rect"]
        return r[1] + r[3]

    def controls(self) -> list[dict]:
        el = uia.find(self.root, lambda d: d["name"] == "Unteres Menüband")
        out: list[dict] = []

        def walk(node: dict, group: str) -> None:
            if node["type"] == "Group" and node["name"]:
                group = node["name"]
            if node["type"] in CONTROL_TYPES and node["name"] and node["rect"][2] > 0:
                kids = [c for c in node.get("children", []) if c["type"] in ("Button", "MenuItem") and c["rect"][2] > 0]
                out.append({"name": node["name"], "type": node["type"], "group": group, "rect": node["rect"],
                            "parts": [c["rect"] for c in kids], "enabled": node["enabled"]})
                return
            for c in node.get("children", []):
                walk(c, group)
        walk(uia.tree(el, max_depth=14), "")
        return out

    def close(self) -> None:
        self.t.close()


class RaabeSide:
    def __init__(self) -> None:
        self.t = targets.raabe_write("short", port=9371)
        winapi.place(self.t.hwnd, *RAABE_RECT)
        winapi.activate(self.t.hwnd)
        time.sleep(1.0)

    def origin(self) -> tuple[int, int]:
        l, t, _r, _b = winapi.window_rect(self.t.hwnd)
        return l, t

    def select_tab(self, name: str) -> None:
        ok = self.t.js("(() => { const b = [...document.querySelectorAll('.tabrow .tab')].find(x => x.textContent.trim() === %s); "
                       "if (!b) return false; b.click(); return true; })()" % json.dumps(name))
        if not ok:
            raise KeyError(name)
        time.sleep(0.6)

    def _to_screen(self, r: list[float]) -> list[int]:
        dpr = self.t.js("devicePixelRatio")
        ox, oy = self.t._viewport_origin()
        return [round(ox + r[0] * dpr), round(oy + r[1] * dpr), round(r[2] * dpr), round(r[3] * dpr)]

    def ribbon_bottom(self) -> int:
        r = self.t.js("(() => { const b = document.querySelector('.ribbon-card').getBoundingClientRect(); return [b.x, b.y, b.width, b.height]; })()")
        s = self._to_screen(r)
        return s[1] + s[3]

    def controls(self) -> list[dict]:
        raw = self.t.js("""(() => {
          const out = [];
          const rect = (e) => { const b = e.getBoundingClientRect(); return [b.x, b.y, b.width, b.height]; };
          const name = (e) => (e.getAttribute('aria-label') || (e.dataset.tip || '').split('\\n')[0] || e.textContent || '').trim();
          for (const e of document.querySelectorAll('.ribbon-card .rsplit, .ribbon-card .rcombo, .ribbon-card button.rb, .ribbon-card .rgallery, .ribbon-card .rspinner')) {
            if (e.closest('.rsplit') && !e.classList.contains('rsplit')) continue;
            const g = e.closest('.rgroup');
            const parts = e.classList.contains('rsplit') ? [...e.children].map(rect) : [];
            const main = e.classList.contains('rsplit') ? e.querySelector('button') : e;
            const input = e.querySelector('input');
            out.push({ name: e.classList.contains('rcombo') ? (input && input.getAttribute('aria-label')) || name(e) : name(main), type: e.classList.contains('rsplit') ? 'SplitButton' : e.classList.contains('rcombo') ? 'ComboBox' : 'Button',
                       group: g ? g.getAttribute('aria-label') || '' : '', rect: rect(e), parts, enabled: !e.classList.contains('disabled') && !main.disabled });
          }
          return out;
        })()""")
        for c in raw:
            c["rect"] = self._to_screen(c["rect"])
            c["parts"] = [self._to_screen(p) for p in c["parts"]]
        return raw

    def close(self) -> None:
        self.t.close()


def relative(rect: list[int], origin: tuple[int, int]) -> list[int]:
    return [rect[0] - origin[0], rect[1] - origin[1], rect[2], rect[3]]


def match(word: list[dict], ours: list[dict], wo, oo) -> dict:
    by_name: dict[str, list[dict]] = {}
    for c in ours:
        by_name.setdefault(norm(c["name"]), []).append(c)
    rows, only_word = [], []
    used: set[int] = set()
    for w in word:
        cands = [c for c in by_name.get(norm(w["name"]), []) if id(c) not in used]
        if not cands:
            only_word.append({"name": w["name"], "group": w["group"], "type": w["type"]})
            continue
        # The nearest candidate (names repeat, e.g. "Weitere Optionen").
        wr = relative(w["rect"], wo)
        c = min(cands, key=lambda c: abs(relative(c["rect"], oo)[0] - wr[0]))
        used.add(id(c))
        cr = relative(c["rect"], oo)
        rows.append({"name": w["name"], "group": w["group"], "word": wr, "ours": cr,
                     "dx": cr[0] - wr[0], "dy": cr[1] - wr[1], "dw": cr[2] - wr[2], "dh": cr[3] - wr[3],
                     "word_type": w["type"], "our_type": c["type"]})
    only_ours = [{"name": c["name"], "group": c["group"]} for c in ours if id(c) not in used]
    return {"matched": rows, "only_word": only_word, "only_ours": only_ours}


def hover_shots(side, controls: list[dict], tag: str, folder: Path) -> list[dict]:
    """For each control with an arrow: crops normal, hovering the main part, hovering the arrow."""
    shots = []
    for i, c in enumerate(controls):
        if c["type"] not in ("SplitButton",) and not c["parts"]:
            continue
        x, y, w, h = c["rect"]
        box = (x - 6, y - 6, w + 12, h + 12)
        winapi.move(x + w + 400, y - 120)  # away, so nothing is hovered
        time.sleep(0.35)
        base = folder / f"{tag}-{i:03d}-normal.png"
        grab(base, box)
        parts = c["parts"] or [[x, y, w // 2, h], [x + w // 2, y, w - w // 2, h]]
        files = [str(base)]
        for k, p in enumerate(parts[:2]):
            winapi.move(p[0] + p[2] / 2, p[1] + p[3] / 2)
            time.sleep(0.3)
            f = folder / f"{tag}-{i:03d}-hover{k}.png"
            grab(f, box)
            files.append(str(f))
        shots.append({"name": c["name"], "files": files})
    winapi.move(10, 1040)
    return shots


def side_by_side(word_img: Image.Image, our_img: Image.Image, path: Path) -> None:
    w = max(word_img.width, our_img.width)
    out = Image.new("RGB", (w, word_img.height + our_img.height + 8), "white")
    out.paste(word_img, (0, 0))
    ImageDraw.Draw(out).rectangle([0, word_img.height, w, word_img.height + 7], fill="#ff00ff")
    out.paste(our_img, (0, word_img.height + 8))
    out.save(path)


def hover_sheet(word_shots: list[dict], our_shots: list[dict], path: Path) -> None:
    """One row per control: Word normal, hover main, hover arrow | ours the same."""
    ours = {norm(s["name"]): s for s in our_shots}
    rows = []
    for ws in word_shots:
        os_ = ours.get(norm(ws["name"]))
        if not os_:
            continue
        rows.append([Image.open(f) for f in ws["files"]] + [Image.open(f) for f in os_["files"]])
    if not rows:
        return
    cell_w = max(im.width for r in rows for im in r)
    cell_h = max(im.height for r in rows for im in r)
    sheet = Image.new("RGB", ((cell_w + 6) * 6 + 20, (cell_h + 6) * len(rows)), "white")
    for y, r in enumerate(rows):
        for x, im in enumerate(r[:6]):
            sheet.paste(im, (x * (cell_w + 6) + (20 if x >= 3 else 0), y * (cell_h + 6)))
    sheet.save(path)


def main() -> None:
    hover = "--hover" in sys.argv
    tabs = [a for a in sys.argv[1:] if not a.startswith("--")] or TABS
    OUT.mkdir(parents=True, exist_ok=True)
    word, ours = WordSide(), RaabeSide()
    report: dict = {}
    try:
        for tab in tabs:
            word.select_tab(tab)
            ours.select_tab(tab)
            winapi.move(10, 1040)
            time.sleep(0.4)
            wo, oo = word.origin(), ours.origin()
            wimg = grab(OUT / f"word-{tab}.png", (wo[0], wo[1], WORD_RECT[2], word.ribbon_bottom() - wo[1] + 8))
            oimg = grab(OUT / f"ours-{tab}.png", (oo[0], oo[1], RAABE_RECT[2], ours.ribbon_bottom() - oo[1] + 8))
            side_by_side(wimg, oimg, OUT / f"compare-{tab}.png")
            wc, oc = word.controls(), ours.controls()
            report[tab] = match(wc, oc, wo, oo)
            if hover:
                folder = OUT / f"hover-{tab}"
                folder.mkdir(exist_ok=True)
                winapi.activate(word.t.hwnd)
                ws = hover_shots(word, wc, "word", folder)
                winapi.activate(ours.t.hwnd)
                os_ = hover_shots(ours, oc, "ours", folder)
                hover_sheet(ws, os_, OUT / f"hover-{tab}.png")
            m = report[tab]
            big = [r for r in m["matched"] if max(abs(r["dw"]), abs(r["dh"])) > 4]
            print(f"{tab}: {len(m['matched'])} matched, {len(big)} differ in size by > 4 px, "
                  f"{len(m['only_word'])} only in Word, {len(m['only_ours'])} only in RaabeWrite", flush=True)
    finally:
        word.close()
        ours.close()
    RESULTS.mkdir(exist_ok=True)
    path = RESULTS / f"{dt.date.today().isoformat()}-compare-ui.json"
    path.write_text(json.dumps({"meta": {"date": dt.datetime.now().isoformat(timespec="seconds")}, "results": report},
                               ensure_ascii=False, indent=1), encoding="utf-8")
    print("saved to", path)


if __name__ == "__main__":
    main()
