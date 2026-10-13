"""Screenshots of RaabeWrite's Step 4 features (comments, tracked changes, equations, table
of contents, the Formel tab, the Neu gallery), taken through the DevTools protocol from the
off-screen app window. Writes PNGs to tools/feel-lab/captures/ (not in git).

    .venv/Scripts/python.exe capture_step4.py
"""

from __future__ import annotations

import base64
import time

from feellab import targets, winapi

OUT = targets.ROOT / "tools/feel-lab/captures"
API = "window.__app"
ED = "window.__app.editor()"


def shot(t, name: str) -> None:
    time.sleep(0.4)
    data = t.cdp("Page.captureScreenshot", {"format": "png"})["data"]
    (OUT / f"step4-{name}.png").write_bytes(base64.b64decode(data))
    print("captured", name)


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    t = targets.raabe_write("short", background=True)
    try:
        winapi.place_quietly(t.hwnd, -4000, 0, 1600, 1000)
        t.js(f"{API}.newDoc()")
        t.js("__probe.focus()")
        t.send_text("Photosynthese")
        t.js(f"{API}.run('style.h1')")
        t.send_key("End")
        t.send_key("Enter")
        t.send_text("Pflanzen wandeln Lichtenergie in chemische Energie um. Dabei entsteht Sauerstoff.")
        t.send_key("f", ctrl=True, alt=True)
        t.send_text("Vgl. Biologiebuch, S. 42.")
        t.js(f"(() => {{ const e = {ED}; e.selectRange(1, 0, 8); }})()")
        t.js(f"{API}.run('review.newComment')")
        time.sleep(0.5)
        t.js("(() => { const a = document.querySelector('.ccard textarea'); a.value = 'Welche Pflanzen genau?'; "
             "a.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true, bubbles: true })); })()")
        t.js(f"(() => {{ const e = {ED}; const o = e.doc.main[1].text.length; e.selectRange(1, o, o); }})()")
        t.js("__probe.focus()")
        t.send_key("Enter")
        t.send_text("Für die Energie eines Photons gilt ")
        t.js(f"{API}.run('insert.equation')")
        time.sleep(1.0)
        t.js("(() => { const m = document.querySelector('math-field'); m.insert('E=h\\\\cdot f'); m.dispatchEvent(new InputEvent('input', { bubbles: true })); })()")
        shot(t, "equation-editing")
        t.js("document.querySelector('math-field').dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))")
        t.js("__probe.focus()")
        t.send_text(".")
        t.send_key("Enter")
        t.js(f"{API}.run('insert.equation')")
        time.sleep(0.8)
        t.js("(() => { const m = document.querySelector('math-field'); m.insert('\\\\sum_{k=1}^{n}k=\\\\frac{n(n+1)}{2}'); m.dispatchEvent(new InputEvent('input', { bubbles: true })); })()")
        t.js("document.querySelector('math-field').dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))")
        t.js("__probe.focus()")
        t.send_key("Enter")
        t.send_key("e", ctrl=True, shift=True)
        t.send_text("Diese Zeile wurde später ergänzt.")
        t.js(f"(() => {{ const e = {ED}; e.selectRange(1, 55, 81); }})()")
        t.js("__probe.focus()")
        t.send_key("Delete")
        t.js(f"{ED}.selectRange(0, 0, 0)")
        t.js(f"{API}.run('toc.auto1')")
        t.js("[...document.querySelectorAll('.tabrow .tab')].find(b => b.textContent === 'Überprüfen').click()")
        shot(t, "review")
        t.js(f"{API}.newFromTemplate('facharbeit')")
        shot(t, "facharbeit")
        t.js("[...document.querySelectorAll('.tabrow button, .tabrow [role=tab]')].find(b => b.textContent.trim() === 'Datei').click()")
        shot(t, "new-gallery")
    finally:
        t.close()


if __name__ == "__main__":
    main()
