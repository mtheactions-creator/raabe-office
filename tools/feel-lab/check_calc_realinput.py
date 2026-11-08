"""Real-input pass of the RaabeCalc audit (D-144): the controls where the operating system's mouse and keyboard
matter (drop-down lists, galleries, colour pickers, submenus, a dialog's focus moving with Tab, a context menu, a
spinner, the Sortierwarnung) are used with real clicks and keys on our own release exe, placed on screen. Every
click goes only to a window checked to be our exe in front. Takes the real-input lock for the run (about two
minutes); setup and checks go through DevTools.

    .venv/Scripts/python.exe check_calc_realinput.py

Result: results/2026-11-09-calc-realinput.json.
"""
from __future__ import annotations

import json
import time
from pathlib import Path

from feellab import winapi
from feellab.cdp_app import AppExe
from feellab.inputlock import real_input
import sweep_calc as SW

HERE = Path(__file__).parent


class Real:
    def __init__(self):
        self.exe = AppExe("calc", port=9446, background=False)
        self.p = self.exe.page
        while not self.p.js("typeof window.__calc === 'object'"):
            time.sleep(0.1)
        self.p.js(SW.JS)
        winapi.activate(self.exe.hwnd)
        time.sleep(0.4)

    def origin(self) -> tuple[float, float, float]:
        for child in winapi.child_windows(self.exe.hwnd):
            if winapi.window_class(child) == "Chrome_RenderWidgetHostHWND":
                ox, oy = winapi.client_origin(child)
                return ox, oy, self.p.js("devicePixelRatio")
        ox, oy = winapi.client_origin(self.exe.hwnd)
        return ox, oy, self.p.js("devicePixelRatio")

    def ours(self) -> bool:
        return winapi.window_pid(winapi.foreground()) == self.exe.proc.pid

    def at(self, js_rect: str) -> tuple[float, float]:
        r = self.p.js(f"(() => {{ const e = {js_rect}; if (!e) return null; e.scrollIntoView?.({{block: 'nearest'}}); const r = e.getBoundingClientRect(); return {{x: r.left + r.width / 2, y: r.top + r.height / 2}}; }})()")
        if not r:
            raise LookupError(js_rect)
        ox, oy, dpr = self.origin()
        return ox + r["x"] * dpr, oy + r["y"] * dpr

    def click(self, js_el: str, right: bool = False) -> None:
        x, y = self.at(js_el)
        if not self.ours():
            winapi.activate(self.exe.hwnd)
        if not self.ours():
            raise RuntimeError("our RaabeCalc is not in front")
        if right:
            winapi.right_click(x, y)
        else:
            winapi.click(x, y)
        time.sleep(0.35)

    def keys(self, combo: str) -> None:
        if not self.ours():
            raise RuntimeError("our RaabeCalc is not in front")
        winapi.press(combo)
        time.sleep(0.25)

    def type(self, text: str) -> None:
        if not self.ours():
            raise RuntimeError("our RaabeCalc is not in front")
        winapi.type_text(text)
        time.sleep(0.2)

    def reset(self, sel: str = "A2:B4") -> None:
        sw = SW.Sweep.__new__(SW.Sweep)
        sw.p, sw.exe = self.p, self.exe
        SW.Sweep.reset(sw, sel)

    def close(self) -> None:
        self.exe.close()


CTRL = "__sweep.controls().find(c => c.id === {id!r})"


def ctrl(i: str) -> str:
    return f"(() => {{ const c = __sweep.controls().find(c => c.id === {json.dumps(i)}); return c ? document.elementFromPoint(c.x, c.y) : null; }})()"


def tab(name: str) -> str:
    return f"[...document.querySelectorAll('.tabrow .tab')].find(t => t.textContent === {json.dumps(name)})"


def item(label: str) -> str:
    return f"[...document.querySelectorAll('.menu button, .menu [role=menuitem], .combo-list button, .tile-panel button')].filter(b => b.offsetParent).find(b => (b.getAttribute('aria-label') || b.textContent).trim() === {json.dumps(label)})"


def swatch(label: str) -> str:
    return f"[...document.querySelectorAll('.menu.palette [aria-label], .menu.palette [data-tip]')].find(b => (b.getAttribute('aria-label') || b.dataset.tip) === {json.dumps(label)})"


STYLE = "(a => { const r = __calc.parse(a); const s = __calc.app.wb.styles.get(__calc.app.sheet.get(r.r0, r.c0)?.s); return s; })"


def cases(r: Real) -> list[tuple[str, callable, str]]:
    """(name, actions, check expression that is true when the option worked)."""
    return [
        ("Schriftart: Liste öffnen, Arial wählen", lambda: (r.click(tab("Start")), r.click(ctrl("Schriftart|Schriftart")), r.click(item("Arial"))),
         f"{STYLE}('A2').font.name === 'Arial' && {STYLE}('B4').font.name === 'Arial'"),
        ("Schriftgrad: tippen und Enter", lambda: (r.click(tab("Start")), r.click("document.querySelector('.rcombo input[aria-label=Schriftgrad]')"), r.keys("ctrl+a"), r.type("14"), r.keys("enter")),
         f"{STYLE}('A2').font.size === 14"),
        ("Füllfarbe: Palette, Gelb", lambda: (r.click(tab("Start")), r.click(ctrl("Schriftart|Füllfarbe ▾")), r.click(swatch("Gelb"))),
         f"/^#?FFFF00$/i.test({STYLE}('A2').fill || '')"),
        ("Füllfarbe: Keine Füllung", lambda: (r.click(tab("Start")), r.click(ctrl("Schriftart|Füllfarbe ▾")), r.click(swatch("Gelb")),
                                              r.click(ctrl("Schriftart|Füllfarbe ▾")), r.click(item("Keine Füllung"))),
         f"!{STYLE}('A2').fill"),
        ("Schriftfarbe: Palette, Rot", lambda: (r.click(tab("Start")), r.click(ctrl("Schriftart|Schriftfarbe ▾")), r.click(swatch("Rot"))),
         f"/^#?FF0000$/i.test({STYLE}('A2').font.color || '')"),
        ("Zellenformatvorlagen: Gut", lambda: (r.click(tab("Start")), r.click(ctrl("Formatvorlagen|Zellenformatvorlagen")), r.click(item("Gut"))),
         f"{STYLE}('A2').fill === '#C6EFCE' && {STYLE}('A2').font.color === '#006100'"),
        ("Rahmen: Alle Rahmenlinien", lambda: (r.click(tab("Start")), r.click(ctrl("Schriftart|Rahmen ▾")), r.click(item("Alle Rahmenlinien"))),
         f"!!{STYLE}('A3').borders.top && !!{STYLE}('B3').borders.left"),
        ("Bedingte Formatierung: Untermenü, Größer als, Wert, Enter", lambda: (r.reset("B2:B5"), r.click(tab("Start")), r.click(ctrl("Formatvorlagen|Bedingte Formatierung")),
                                                                         r.click(item("Regeln zum Hervorheben von Zellen")), r.click(item("Größer als...")),
                                                                         r.keys("ctrl+a"), r.type("10"), r.keys("enter")),
         "(__calc.app.sheet.conditional || []).some(f => f.rules.some(x => x.operator === 'greaterThan' && x.formulas[0] === '10'))"),
        ("Zellen formatieren: Strg+1, Tab bleibt im Dialog, Esc", lambda: (r.keys("ctrl+1"), *[r.keys("tab") for _ in range(25)]),
         "(() => { const d = document.querySelector('.dialog-layer .dialog'); const ok = !!d && d.contains(document.activeElement); return ok; })()"),
        ("Zellen formatieren: Esc schließt", lambda: (r.keys("ctrl+1"), r.keys("esc")), "!document.querySelector('.dialog-layer')"),
        ("Blattregister: Kontextmenü, Registerfarbe", lambda: (r.click("document.querySelector('.sheettab')", right=True), r.click(item("Registerfarbe")),
                                                              r.click(swatch("Rot"))),
         "!!__calc.app.wb.sheets[0].tabColor"),
        ("Skalierung: Drehfeld Mehr", lambda: (r.click(tab("Seitenlayout")), r.click(ctrl("An Format anpassen|Skalierung: Mehr"))),
         "(__calc.app.sheet.page?.scale ?? 100) > 100"),
        ("Sortieren: Spalte markiert, Sortierwarnung, Enter erweitert", lambda: (r.reset("B2:B5"), r.click(tab("Start")), r.click(ctrl("Bearbeiten|Sortieren und Filtern")),
                                                                         r.click(item("Nach Größe sortieren (aufsteigend)")), r.keys("enter")),
         "__calc.app.sheet.get(1, 0)?.v === 'Ben' && __calc.app.sheet.get(1, 1)?.v === 7.5"),
        ("Rückgängig: Strg+Z nach dem Sortieren", lambda: r.keys("ctrl+z"), "__calc.app.sheet.get(1, 0)?.v === 'Anna'"),
    ]


def main() -> None:
    results = []
    with real_input("audit-calc check_calc_realinput", wait=3600):
        r = Real()
        try:
            for name, act, check in cases(r):
                if not name.startswith("Rückgängig"):
                    r.reset()
                    r.p.errors.clear()
                    winapi.activate(r.exe.hwnd)
                res: dict = {"name": name}
                try:
                    act()
                    time.sleep(0.3)
                    res["pass"] = bool(r.p.js(check))
                except Exception as e:  # noqa: BLE001
                    res["pass"] = False
                    res["error"] = str(e)[:200]
                res["errors"] = list(r.p.errors)
                if res["errors"]:
                    res["pass"] = False
                results.append(res)
                print(name, "ok" if res["pass"] else "FAILED", res.get("error", ""), res["errors"][:2], flush=True)
                if not name.startswith("Zellen formatieren: Strg+1"):
                    continue
                r.keys("esc")
        finally:
            r.close()
    out = HERE / "results" / "2026-11-09-calc-realinput.json"
    summary = {"cases": len(results), "passed": sum(1 for x in results if x["pass"])}
    out.write_text(json.dumps({"summary": summary, "results": results}, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(json.dumps(summary))


if __name__ == "__main__":
    main()
