"""Real-mouse checks of RaabeWrite's formatting controls (palettes, lists, menus, gallery,
dialogs, format painter). Takes over keyboard, mouse and screen.

    .venv/Scripts/python.exe ui_check.py
"""

from __future__ import annotations

import json
import sys
import time

from feellab import targets, winapi

RESULTS: list[tuple[str, bool, str]] = []


def check(name: str, ok: bool, detail: str = "") -> None:
    RESULTS.append((name, bool(ok), detail))
    print(("PASS " if ok else "FAIL ") + name + (f"  [{detail}]" if detail and not ok else ""), flush=True)


class UI:
    def __init__(self, t: targets.WebViewTarget):
        self.t = t

    def center(self, expr: str) -> tuple[float, float] | None:
        """Screen point in the middle of the element `expr` evaluates to."""
        r = self.t.js(f"(() => {{ const e = {expr}; if (!e) return null; e.scrollIntoView({{block: 'nearest'}}); "
                      "const b = e.getBoundingClientRect(); return [b.x + b.width / 2, b.y + b.height / 2]; })()")
        if not r:
            return None
        dpr = self.t.js("devicePixelRatio")
        ox, oy = self.t._viewport_origin()
        return ox + r[0] * dpr, oy + r[1] * dpr

    def click(self, expr: str, wait: float = 0.35, count: int = 1) -> bool:
        p = self.center(expr)
        if not p:
            return False
        winapi.click(*p, count=count)
        time.sleep(wait)
        return True

    def fmt(self, offset: int) -> dict:
        return self.t.js(f"__probe.formatAt({offset})")


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    t = targets.raabe_write("short", port=9333)
    ui = UI(t)
    try:
        t.activate()
        time.sleep(0.5)
        algen = t.find("Algen")

        def select_algen():
            t.set_selection(algen, algen + 5)
            t.js("__probe.focus()")

        select_algen()
        ui.click("document.querySelector('[aria-label=\"Schriftfarbe\"]').parentElement.querySelector('.arrow')")
        ok = ui.click("document.querySelector('.palette [aria-label=\"Rot\"]')")
        check("font colour palette: click Rot", ok and ui.fmt(algen + 1)["color"] == "FF0000", json.dumps(ui.fmt(algen + 1)))

        select_algen()
        ui.click("document.querySelector('[aria-label=\"Texthervorhebungsfarbe\"]').parentElement.querySelector('.arrow')")
        ok = ui.click("document.querySelector('.palette [aria-label=\"Türkis\"]')")
        check("highlight palette: click Türkis", ok and ui.fmt(algen + 1)["highlight"] == "cyan", json.dumps(ui.fmt(algen + 1)))

        select_algen()
        ui.click("document.querySelectorAll('.rcombo')[0].querySelector('button')")
        ok = ui.click("[...document.querySelectorAll('.combo-list button')].find(b => b.textContent === 'Cambria')")
        check("font list: click Cambria", ok and ui.fmt(algen + 1)["font"] == "Cambria", json.dumps(ui.fmt(algen + 1)))
        shown = t.js("document.querySelectorAll('.rcombo input')[0].value")
        check("font box shows the new font", shown == "Cambria", str(shown))

        select_algen()
        ui.click("document.querySelectorAll('.rcombo')[1].querySelector('button')")
        ok = ui.click("[...document.querySelectorAll('.combo-list button')].find(b => b.textContent === '14')")
        check("size list: click 14", ok and ui.fmt(algen + 1)["size"] == 14, json.dumps(ui.fmt(algen + 1)))

        select_algen()
        ui.click("document.querySelector('[aria-label=\"Unterstreichen\"]').parentElement.querySelector('.arrow')")
        ok = ui.click("[...document.querySelectorAll('.menu button, .menu .item')].find(b => b.textContent.includes('Doppelt unterstrichen'))")
        check("underline menu: Doppelt unterstrichen", ok and ui.fmt(algen + 1)["underline"] == "double", json.dumps(ui.fmt(algen + 1)))

        select_algen()
        ok = ui.click("[...document.querySelectorAll('.rgallery .item')].find(b => b.textContent === 'Überschrift 1')")
        check("style gallery: Überschrift 1", ok and ui.fmt(algen)["style"] == "Überschrift 1", json.dumps(ui.fmt(algen)))
        checked = t.js("document.querySelector('.rgallery .item.checked')?.textContent")
        check("gallery marks the current style", checked == "Überschrift 1", str(checked))

        # Absatz dialog: left indent 1 cm, OK.
        t.load("short")
        algen = t.find("Algen")
        t.set_selection(algen, algen)
        t.js("__probe.focus()")
        ui.click("[...document.querySelectorAll('.launcher')].find(b => (b.dataset.tip || '').startsWith('Absatz'))", wait=0.5)
        ok = ui.click("[...document.querySelectorAll('.dialog input')][0]")
        winapi.press("ctrl+a")
        winapi.type_text("1 cm")
        ui.click("[...document.querySelectorAll('.dialog button')].find(b => b.textContent === 'OK')")
        ind = ui.fmt(algen)["indLeft"]
        check("Absatz dialog: Einzug links 1 cm", ok and abs(ind - 28.35) < 0.1, str(ind))
        check("dialog closes", t.js("document.querySelectorAll('.dialog').length") == 0)

        # Schriftart dialog: Fett via Schriftschnitt, OK.
        t.set_selection(algen, algen + 5)
        t.js("__probe.focus()")
        ui.click("[...document.querySelectorAll('.launcher')].find(b => (b.dataset.tip || '').startsWith('Schriftart'))", wait=0.5)
        t.js("(() => { const s = document.querySelector('.dialog select'); s.selectedIndex = 2; s.dispatchEvent(new Event('input')); })()")
        ui.click("[...document.querySelectorAll('.dialog button')].find(b => b.textContent === 'OK')")
        check("Schriftart dialog: Fett", ui.fmt(algen + 1)["bold"], json.dumps(ui.fmt(algen + 1)))

        # Format painter: bold "Algen" → drag over "Bakterien".
        t.set_selection(algen + 2, algen + 2)
        t.js("__probe.focus()")
        ui.click("document.querySelector('[aria-label=\"Format übertragen\"]')")
        bak = t.find("Bakterien")
        x1, y1 = t.caret_point(bak)
        x2, y2 = t.caret_point(bak + 9)
        winapi.drag(x1 + 1, y1, x2 - 1, y2)
        time.sleep(0.4)
        check("format painter copies bold onto dragged text", ui.fmt(bak + 3)["bold"] and not ui.fmt(bak - 2)["bold"], json.dumps(ui.fmt(bak + 3)))
        check("painter switches off after one use", not t.js("document.querySelector('.workspace, .editor-workspace')?.style.cursor?.includes('url')"))
    finally:
        t.close()
    passed = sum(ok for _, ok, _ in RESULTS)
    print(f"\n{passed}/{len(RESULTS)} checks passed")


if __name__ == "__main__":
    main()
