"""The formula editor in RaabeWrite with real input (D-047): a new equation (Alt+Umschalt+0),
typing with build-up, a click into the numerator, a gallery entry wrapping the selection, a
symbol from the ribbon, the context menu, Konvertieren, Enter, a click back into the equation,
the Formeloptionen arrow, saving — and Word reading the result. Screenshots of the Formel tab,
its galleries and the equation box go to captures/math-ui/ (local only), for comparison with
Word's (captures/word-equation/). Takes over mouse and keyboard.

    .venv/Scripts/python.exe check_math_ui.py
"""
from __future__ import annotations

import json
import tempfile
import time
from pathlib import Path

from feellab import winapi
from feellab.targets import WINDOW_RECT, raabe_write

HERE = Path(__file__).parent
CAP = HERE / "captures" / "math-ui"
CAP.mkdir(parents=True, exist_ok=True)
OUT = Path(tempfile.gettempdir()) / "raabe-math-ui.docx"

STATE = """(() => { const e = __app.editor(); const paras = e.doc.main;
  const eqs = paras.flatMap((p, i) => (p.objects ?? []).filter((o) => o.kind === 'math').map((o) => ({ para: i, latex: o.latex, inline: !!o.mathInline })));
  const caret = [...document.querySelectorAll('.page > div')].find((d) => d.style.zIndex === '5' && d.style.background === 'rgb(0, 0, 0)');
  const cr = caret && caret.style.display !== 'none' ? caret.getBoundingClientRect() : null;
  const m = document.querySelector('.page math'); const mr = m ? m.getBoundingClientRect() : null;
  return JSON.stringify({ eqs, texts: paras.map((p) => p.text.replace(/\\ufffc/g, '[F]')), editing: __app.enabled('math.normalText'),
    caret: cr && [Math.round(cr.left), Math.round(cr.top), Math.round(cr.height)], math: mr && [Math.round(mr.left), Math.round(mr.top), Math.round(mr.width), Math.round(mr.height)],
    box: !!document.querySelector('.math-session .math-box'), sel: !!document.querySelector('.math-session .math-sel'), tab: __app.activeRibbonTab() }); })()"""

RECT = """(sel) => { const b = document.querySelector(sel); if (!b) return null; const r = b.getBoundingClientRect(); return [r.left + r.width / 2, r.top + r.height / 2, r.left, r.top, r.width, r.height]; }"""


def main() -> None:
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    out: dict = {}
    t = raabe_write("short", port=9378)
    try:
        winapi.place(t.hwnd, *WINDOW_RECT)
        winapi.activate(t.hwnd)
        t.js("__app.newDoc()")
        time.sleep(0.6)
        t.js("__app.closeOthers && __app.closeOthers()")
        t.js("__probe.focus()")
        scale = winapi.dpi_scale(t.hwnd)
        ox, oy = winapi.client_origin(t.hwnd)
        to_screen = lambda x, y: (ox + x * scale, oy + y * scale)  # noqa: E731
        state = lambda: json.loads(t.js(STATE))  # noqa: E731

        def shot(name: str, x: float, y: float, w: float, h: float) -> None:
            sx, sy = to_screen(x, y)
            winapi.save_region_png(str(CAP / f"{name}.png"), int(sx), int(sy), int(w * scale), int(h * scale))

        def box_shot(name: str) -> None:
            m = t.js(f"({RECT})('.page math')")
            if m:
                shot(name, m[2] - 50, m[3] - 30, m[4] + 120, m[5] + 60)

        def click_css(sel: str, right: bool = False, at: str = "center") -> bool:
            p = t.js(f"({RECT})({json.dumps(sel)})")
            if not p:
                print("not found", sel, flush=True)
                return False
            x = p[2] + p[4] - 1 if at == "right" else p[2] + 1 if at == "left" else p[0]
            sx, sy = to_screen(x, p[1])
            if right:
                winapi.right_click(sx, sy)
            else:
                winapi.click(sx, sy)
            time.sleep(0.7)
            return True

        def ribbon(label: str) -> bool:
            return click_css(f".ribbon-card button[aria-label='{label}']")

        # 1. A new equation, typed with build-up.
        winapi.press("alt+shift+0")
        time.sleep(0.6)
        out["new"] = state()
        box_shot("1-new")
        winapi.type_text("a/b ", gap=0.06)
        time.sleep(0.5)
        out["typed"] = state()
        box_shot("2-typed")
        print("typed", out["typed"], flush=True)
        # 2. A click at the right of the numerator, then typing.
        click_css(".page math [data-a^='0.num|']", at="right")
        winapi.type_text("x", gap=0.06)
        time.sleep(0.4)
        out["numerator"] = state()
        print("numerator", out["numerator"]["eqs"], flush=True)
        # 3. Umschalt+Links selects the x; Formel > Wurzel > Quadratwurzel puts it under a root.
        winapi.press("shift+left")
        time.sleep(0.3)
        out["selected"] = state()
        box_shot("3-selected")
        t.js("[...document.querySelectorAll('[role=tab]')].find((x) => x.textContent.trim() === 'Formel')?.click()")
        time.sleep(0.5)
        shot("tab", 0, 0, 2000 / scale * 0.98, 190)
        if ribbon("Wurzel"):
            shot("gallery-Wurzel", 900, 150, 520, 600)
            click_css(".rpanel .mtile[aria-label='Quadratwurzel']")
        out["root"] = state()
        print("root", out["root"]["eqs"], flush=True)
        # 4. A symbol from the ribbon's gallery (Unendlich), at the end.
        winapi.press("end")
        click_css(".msymtiles .item[data-tip='Unendlich']")
        out["symbol"] = state()
        print("symbol", out["symbol"]["eqs"], flush=True)
        # 5. Each gallery as a screenshot (the panel under its button).
        for label in ("Bruch", "Hoch/Tief", "Integral", "Großer Operator", "Eckige Klammer", "Funktion", "Akzent", "Grenzwert und Logarithmus", "Operator", "Matrix", "Formel"):
            if ribbon(label):
                p = t.js("(() => { const r = document.querySelector('.rpanel')?.getBoundingClientRect(); return r && [r.left, r.top, r.width, r.height]; })()")
                if p:
                    shot(f"gallery-{label.replace('/', '_')}", p[0], p[1], p[2], min(p[3], 1100))
                winapi.press("esc")
                time.sleep(0.3)
        click_css(".msymtiles .scroll button[aria-label='Weitere']")
        p = t.js("(() => { const r = document.querySelector('.rpanel')?.getBoundingClientRect(); return r && [r.left, r.top, r.width, r.height]; })()")
        if p:
            shot("gallery-Formelsymbole", p[0], p[1], p[2], min(p[3], 900))
        winapi.press("esc")
        time.sleep(0.3)
        t.js("__probe.focus()")
        # 6. The context menu on the fraction's numerator.
        click_css(".page math [data-a^='0.num|']", right=True)
        time.sleep(0.6)
        out["context"] = t.js("[...document.querySelectorAll('.menu button')].map((b) => b.textContent.trim()).filter(Boolean)")
        p = t.js("(() => { const r = document.querySelector('.menu')?.getBoundingClientRect(); return r && [r.left, r.top, r.width, r.height]; })()")
        if p:
            shot("context", p[0], p[1], p[2], p[3])
        print("context", out["context"], flush=True)
        winapi.press("esc")
        time.sleep(0.3)
        # 7. Konvertieren: linear, then professional again.
        before = state()["eqs"]
        t.js("__app.run('math.toLinear')")
        time.sleep(0.5)
        out["linear"] = state()
        box_shot("4-linear")
        t.js("__app.run('math.toProfessional')")
        time.sleep(0.5)
        out["professional"] = state()
        out["round trip"] = out["professional"]["eqs"] == before
        print("linear", out["linear"]["eqs"], "back", out["round trip"], flush=True)
        # 8. Enter leaves the equation: a new paragraph; then a click back into the equation.
        winapi.press("end")
        winapi.press("enter")
        winapi.type_text("Text", gap=0.06)
        time.sleep(0.4)
        out["enter"] = state()
        click_css(".page math [data-a^='0.den|']")
        out["click back"] = state()
        box_shot("5-click-back")
        # 9. The Formeloptionen arrow and its menu; In 'Inline' ändern.
        if click_css(".math-session .math-options"):
            out["box menu"] = t.js("[...document.querySelectorAll('.menu button')].map((b) => b.textContent.trim()).filter(Boolean)")
            p = t.js("(() => { const r = document.querySelector('.menu')?.getBoundingClientRect(); return r && [r.left, r.top, r.width, r.height]; })()")
            if p:
                shot("box-menu", p[0], p[1], p[2], p[3])
            click_css(".menu button[role=menuitem]:nth-child(6), .menu button:nth-of-type(6)")
            out["inline"] = state()
            print("inline", out["inline"]["eqs"], flush=True)
        # 10. Saved and read by Word.
        if OUT.exists():
            OUT.unlink()
        t.js(f"__app.saveTo({str(OUT)!r})")
        time.sleep(1.5)
        import win32com.client
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        try:
            d = word.Documents.Open(str(OUT), ReadOnly=True)
            out["word"] = {"omaths": d.OMaths.Count, "text": d.Content.Text[:80]}
            d.Close(0)
        finally:
            word.Quit(0)
        print("word", out["word"], flush=True)
    finally:
        (HERE / "results" / "math-ui.json").write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        t.close()


if __name__ == "__main__":
    main()
