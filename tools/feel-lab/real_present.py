"""The real-input part of the by-hand audit (D-137): the controls where the OS mouse and keyboard matter (a palette
opened and a swatch picked, a combo box's list, typing into the font box, a spinner's arrow, the Format pane's slider
and its native list box, Tab through a dialog's fields, a shape drawn from the gallery by dragging, a transition from
the ribbon gallery), done with SendInput in the release exe, our own instance. Takes mouse and keyboard: runs only
while holding the real-input lock (feellab/inputlock.py) and sends input only while our window is in front.

    .venv/Scripts/python.exe real_present.py

Result: results/<date>-present-real-input.json.
"""
from __future__ import annotations

import datetime
import json
import time
from pathlib import Path

from feellab import winapi
from feellab.inputlock import real_input, touch
from sweep_present import start

HERE = Path(__file__).parent


class Real:
    def __init__(self):
        self.page = start(False, "", 9488)
        self.hwnd = self.page.hwnd
        winapi.place(self.hwnd, 0, 0, 3000, 1900)
        time.sleep(0.5)

    def js(self, e):
        return self.page.js(e)

    def ours(self) -> bool:
        fg = winapi.foreground()
        return winapi.window_pid(fg) == self.page.proc.pid

    def front(self) -> None:
        winapi.activate(self.hwnd)
        time.sleep(0.2)
        if not self.ours():
            raise RuntimeError("our window is not in front: no input sent")

    def origin(self):
        for child in winapi.child_windows(self.hwnd):
            if winapi.window_class(child) == "Chrome_RenderWidgetHostHWND":
                ox, oy = winapi.client_origin(child)
                return ox, oy, self.js("devicePixelRatio")
        ox, oy = winapi.client_origin(self.hwnd)
        return ox, oy, self.js("devicePixelRatio")

    def screen(self, r: dict) -> tuple[float, float]:
        ox, oy, dpr = self.origin()
        return ox + r["x"] * dpr, oy + r["y"] * dpr

    def click(self, r: dict, count: int = 1) -> None:
        if not self.ours():
            raise RuntimeError("our window is not in front: no input sent")
        x, y = self.screen(r)
        winapi.click(x, y, count)
        time.sleep(0.35)

    def control(self, name: str, kind: str | None = None) -> dict:
        c = next(x for x in self.js("window.__sw.controls()") if x["name"] == name and (kind is None or x["kind"] == kind))
        return self.js(f"window.__sw.findControl({json.dumps(c['key'])})")

    def item(self, name: str) -> dict:
        it = next(x for x in reversed(self.js("window.__sw.items()")) if x["name"] == name)
        return self.js(f"window.__sw.itemRect({json.dumps(it['path'])})")

    def setup(self, ctx: str, tab: str | None) -> None:
        for _ in range(2):
            winapi.press("esc")
            time.sleep(0.1)
        self.js(f"window.__present.sweepContext({json.dumps(ctx)})")
        time.sleep(0.2)
        if tab:
            self.click(self.js(f"window.__sw.tabRect({json.dumps(tab)})"))

    def slide_xml(self) -> str:
        return self.js("(() => { const a = window.__present.app; return new XMLSerializer().serializeToString(a.deck.pptx.doc(a.slidePath())); })()")


def cases(r: Real) -> list[dict]:
    out = []

    def case(name, fn, check):
        try:
            touch()
            fn()
            ok, got = check()
        except Exception as e:  # noqa: BLE001
            ok, got = False, f"error: {str(e)[:300]}"
        errs = r.js("window.__sw.take()")
        out.append({"name": name, "ok": bool(ok) and not errs, "got": got, "errs": errs})
        print("ok  " if ok and not errs else "FAIL", name, got, errs, flush=True)

    # 1. Schriftfarbe: the palette opened with the mouse, Rot picked.
    def font_color():
        r.setup("shape-text", "Start")
        r.click(r.control("Schriftfarbe (Pfeil)", "color-arrow"))
        r.click(r.item("Rot"))
    case("Schriftfarbe-Palette > Rot", font_color, lambda: ('val="FF0000"' in r.slide_xml(), "FF0000 in the run" if 'val="FF0000"' in r.slide_xml() else "no red"))

    # 2. Schriftgrad: the combo box's list, 24 picked.
    def size_list():
        r.setup("shape-text", "Start")
        r.click(r.control("Schriftgrad (Öffnen)", "combo-arrow"))
        r.click(r.item("24"))
    case("Schriftgrad-Liste > 24", size_list, lambda: ('sz="2400"' in r.slide_xml(), "sz=2400" if 'sz="2400"' in r.slide_xml() else "no sz=2400"))

    # 3. Schriftart: a click into the box, typed with the keyboard, Enter.
    def font_typed():
        r.setup("shape-text", "Start")
        r.click(r.control("Schriftart", "combo-input"))
        winapi.press("ctrl+a")
        winapi.type_text("Arial")
        winapi.press("enter")
        time.sleep(0.4)
    case("Schriftart getippt: Arial", font_typed, lambda: ('typeface="Arial"' in r.slide_xml(), "Arial" if 'typeface="Arial"' in r.slide_xml() else "no Arial"))

    # 4. Formenhöhe: the spinner's up arrow, three clicks.
    h0 = {}

    def spinner():
        r.setup("shape", "Formformat")
        h0["h"] = r.js("window.__present.state().shapes.filter((s) => !s.ph)[0].h")
        for _ in range(3):
            r.click(r.control("Formenhöhe: (Mehr)", "spinner-up"))
    case("Formenhöhe: Pfeil nach oben x3", spinner, lambda: ((lambda h: (h > h0["h"], f"{h0['h']} -> {h} EMU"))(r.js("window.__present.state().shapes.filter((s) => !s.ph)[0].h"))))

    # 5. Format pane: Einfarbige Füllung, then the Transparenz slider dragged; Strichtyp from the native list.
    def pane_slider():
        r.setup("shape", "Formformat")
        r.click(r.control("Objekt formatieren...", "launcher"))
        radio = r.js("(() => { const l = Array.from(document.querySelectorAll('aside.taskpane label')).find((x) => x.textContent.trim() === 'Einfarbige Füllung'); return l ? window.__sw.rect(l.querySelector('input')) : null; })()")
        r.click(radio)
        s = r.js("(() => { const e = document.querySelector('aside.taskpane .tp-slider'); e.scrollIntoView({ block: 'center' }); const b = e.getBoundingClientRect(); return { x: b.left + 4, y: b.top + b.height / 2, w: b.width }; })()")
        x1, y1 = r.screen(s)
        x2, _ = r.screen({"x": s["x"] + s["w"] * 0.5, "y": s["y"]})
        if not r.ours():
            raise RuntimeError("not in front")
        winapi.drag(x1, y1, x2, y1)
        time.sleep(0.5)
    case("Formatbereich: Transparenz-Schieberegler gezogen", pane_slider, lambda: ("alpha" in r.slide_xml(), "alpha written" if "alpha" in r.slide_xml() else "no alpha"))

    def pane_select():
        r.setup("shape", "Formformat")
        r.click(r.control("Objekt formatieren...", "launcher"))
        # Linie opened, Einfarbige Linie, then Strichtyp's native list: a click, two steps down with the arrow keys, Enter.
        head = r.js("(() => { const b = Array.from(document.querySelectorAll('aside.taskpane .tp-sec-head')).find((x) => x.textContent.includes('Linie')); return b && b.getAttribute('aria-expanded') === 'false' ? window.__sw.rect(b) : null; })()")
        if head:
            r.click(head)
        radio = r.js("(() => { const l = Array.from(document.querySelectorAll('aside.taskpane label')).find((x) => x.textContent.trim() === 'Einfarbige Linie'); return l ? window.__sw.rect(l.querySelector('input')) : null; })()")
        r.click(radio)
        sel = r.js("(() => { const s = Array.from(document.querySelectorAll('aside.taskpane select')).find((x) => Array.from(x.options).some((o) => o.text === 'Strich')); s.scrollIntoView({ block: 'center' }); return window.__sw.rect(s); })()")
        r.click(sel)
        time.sleep(0.3)
        winapi.press("down")
        winapi.press("down")
        winapi.press("enter")
        time.sleep(0.4)
    case("Formatbereich: Strichtyp aus der Liste", pane_select, lambda: ("prstDash" in r.slide_xml(), "prstDash written" if "prstDash" in r.slide_xml() else "no prstDash"))

    # 6. Absatz dialog: Tab moves through the fields and stays inside; a value typed; Enter.
    tabs: dict = {}

    def dialog_tab():
        r.setup("shape-text", "Start")
        launcher = r.js("(() => { const g = Array.from(document.querySelectorAll('.ribbon-card .rgroup')).find((x) => x.getAttribute('aria-label') === 'Absatz'); return window.__sw.rect(g.querySelector('.launcher')); })()")
        r.click(launcher)
        seen = []
        for _ in range(14):
            seen.append(r.js("(() => { const e = document.activeElement; return (e.closest('[role=dialog]') ? 'in:' : 'out:') + e.tagName + ':' + (e.getAttribute('aria-label') || e.textContent || '').trim().slice(0, 20); })()"))
            winapi.press("tab")
            time.sleep(0.08)
        tabs["seen"] = seen
        winapi.press("esc")
    case("Absatz-Dialog: Tab bleibt im Dialog", dialog_tab, lambda: (all(s.startswith("in:") for s in tabs["seen"]) and len(set(tabs["seen"])) > 3, tabs["seen"]))

    # 7. Einfügen > Formen > Ellipse, then drawn by dragging on the slide.
    def draw_shape():
        r.setup("nothing", "Einfügen")
        r.click(r.control("Formen", "dropdown"))
        r.click(r.item("Ellipse"))
        a = r.js("window.__present.slidePointToClient(300, 150)")
        b = r.js("window.__present.slidePointToClient(500, 300)")
        x1, y1 = r.screen(a)
        x2, y2 = r.screen(b)
        if not r.ours():
            raise RuntimeError("not in front")
        winapi.drag(x1, y1, x2, y2)
        time.sleep(0.4)
    case("Formen > Ellipse gezogen", draw_shape, lambda: ('prst="ellipse"' in r.slide_xml(), [(s["kind"], s["w"], s["h"]) for s in r.js("window.__present.state().shapes")]))

    # 8. Übergänge: Verblassen from the ribbon's gallery.
    def transition():
        r.setup("slides", "Übergänge")
        r.click(r.control("Verblassen", "gallery-item"))
    case("Übergänge-Katalog > Verblassen", transition, lambda: ("<p:fade" in r.slide_xml() or ":fade" in r.slide_xml(), "fade" if "fade" in r.slide_xml() else "no fade"))

    # 9. Bildschirmpräsentation: F5, a click to go on, Esc ends it.
    st: dict = {}

    def show():
        r.setup("slides", None)
        r.front()
        winapi.press("f5")
        time.sleep(1.2)
        st["in"] = r.js("window.__present.inShow()")
        st["s1"] = (r.js("window.__present.showState()") or {}).get("slide")
        x, y = r.screen({"x": 600, "y": 500})
        winapi.click(x, y)
        time.sleep(0.8)
        st["s2"] = (r.js("window.__present.showState()") or {}).get("slide")
        winapi.press("esc")
        time.sleep(0.6)
        st["out"] = not r.js("window.__present.inShow()")
    case("F5, Klick weiter, Esc", show, lambda: (st.get("in") and st.get("out"), st))
    return out


def main() -> None:
    r = Real()
    try:
        with real_input("audit-present real_present.py", wait=1800):
            r.front()
            results = cases(r)
    finally:
        r.page.close()
    out = {"date": datetime.date.today().isoformat(), "passed": sum(1 for x in results if x["ok"]), "total": len(results), "cases": results}
    p = HERE / "results" / f"{out['date']}-present-real-input.json"
    p.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))
    print(out["passed"], "/", out["total"], "->", p.name)


if __name__ == "__main__":
    main()
