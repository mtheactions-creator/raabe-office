"""The macro editor and UserForms in use (D-106), in a headless Edge on the dev server (no real input: CDP
events): a UserForm is drawn in the designer (a CommandButton and a TextBox from the Werkzeugsammlung), its
Click procedure written, a module with a breakpoint run (F5), stepped (F8) with the Lokal-Fenster read, the form
shown and its button clicked; then the project is saved and read back. Screenshots in captures/vbaide.

    .venv/Scripts/python.exe check_present_vbaide.py      (npm run dev:present must be running)

Result: results/<date>-present-vbaide.json.
"""
from __future__ import annotations

import datetime
import json
import time
from pathlib import Path

from feellab.present_page import PresentPage

HERE = Path(__file__).parent
CAP = HERE / "captures" / "vbaide"


def main() -> None:
    CAP.mkdir(parents=True, exist_ok=True)
    out: dict = {"date": datetime.date.today().isoformat(), "checks": {}}
    page = PresentPage(1400, 900, 1.0)

    def key(k: str, code: str, vk: int, mods: int = 0) -> None:
        for t in ("rawKeyDown", "keyUp"):
            page.cdp("Input.dispatchKeyEvent", {"type": t, "key": k, "code": code, "windowsVirtualKeyCode": vk, "modifiers": mods})
        time.sleep(0.4)

    def click(sel: str, dbl: bool = False, dx: float = 0.5, dy: float = 0.5) -> None:
        r = page.js(f"(() => {{ const e = {sel}; const r = e.getBoundingClientRect(); return [r.left + r.width * {dx}, r.top + r.height * {dy}]; }})()")
        for n in ((1, 2) if dbl else (1,)):
            for t in ("mousePressed", "mouseReleased"):
                page.cdp("Input.dispatchMouseEvent", {"type": t, "x": r[0], "y": r[1], "button": "left", "clickCount": n})
        time.sleep(0.4)

    def shot(name: str) -> None:
        (CAP / f"{name}.png").write_bytes(page.shot(0, 0, 1400, 900))

    try:
        page.js("(async () => { const P = window.__present; P.newDoc(); const m = await import('/src/macros.ts'); m.enableMacros(P.app); void m.editor(P.app); await new Promise((r) => setTimeout(r, 500)); })()")
        out["checks"]["editor_open"] = page.js("!!document.querySelector('.vba-ide')")
        # Einfügen > UserForm, then a CommandButton and a TextBox from the Werkzeugsammlung.
        page.js("(() => { const s = document.querySelector('.vba-ide select[aria-label=Einfügen]'); s.value = 'form'; s.dispatchEvent(new Event('change')); })()")
        time.sleep(0.3)
        for tool, (fx, fy) in (("Befehlsschaltfläche", (0.15, 0.2)), ("Textfeld", (0.15, 0.6))):
            click(f"[...document.querySelectorAll('.vi-tools button')].find((b) => b.textContent === '{tool}')")
            click("document.querySelector('.vi-canvas .vf-body')", dx=fx, dy=fy)
        names = page.js("[...document.querySelectorAll('.vi-canvas .vf-c')].map((e) => e.dataset.name)")
        out["designer_controls"] = names
        out["checks"]["designer_adds_controls"] = names == ["CommandButton1", "TextBox1"]
        shot("1-designer")
        # A double-click on the button opens its Click procedure; the code is typed into it.
        click("document.querySelector('.vi-canvas .vf-c[data-name=CommandButton1]')", dbl=True)
        code = page.js("document.querySelector('.vi-text').value")
        out["checks"]["dblclick_opens_click_proc"] = "Private Sub CommandButton1_Click()" in code
        page.js("""(() => { const a = document.querySelector('.vi-text');
            a.value = 'Private Sub CommandButton1_Click()\\r\\n    TextBox1.Text = \"geklickt\"\\r\\n    Me.Caption = \"Fertig\"\\r\\nEnd Sub\\r\\n'; a.dispatchEvent(new Event('input')); })()""")
        # Einfügen > Modul with a procedure; a breakpoint on its second statement; F5.
        page.js("(() => { const s = document.querySelector('.vba-ide select[aria-label=Einfügen]'); s.value = 'standard'; s.dispatchEvent(new Event('change')); })()")
        page.js("""(() => { const a = document.querySelector('.vi-text');
            a.value = 'Option Explicit\\r\\nSub Zeige()\\r\\n    Dim n As Integer\\r\\n    n = 41\\r\\n    n = n + 1\\r\\n    Debug.Print \"n ist\"; n\\r\\n    UserForm1.Show vbModeless\\r\\nEnd Sub\\r\\n';
            a.dispatchEvent(new Event('input')); a.focus(); a.selectionStart = a.selectionEnd = 30; })()""")
        click("document.querySelectorAll('.vi-gutter div')[4]")  # line 5: n = n + 1
        out["checks"]["breakpoint_set"] = page.js("document.querySelectorAll('.vi-gutter div')[4].classList.contains('bp')")
        key("F5", "F5", 116)
        time.sleep(0.6)
        out["status_at_break"] = page.js("document.querySelector('.vi-status').textContent")
        out["locals_at_break"] = page.js("[...document.querySelectorAll('.vi-locals tr')].slice(1).map((r) => [...r.children].map((c) => c.textContent))")
        out["checks"]["stops_at_breakpoint"] = "Zeile 5" in (out["status_at_break"] or "")
        out["checks"]["locals_show_n"] = ["n", "41", "Integer"] in out["locals_at_break"]
        shot("2-breakpoint")
        key("F8", "F8", 119)
        time.sleep(0.5)
        out["status_after_step"] = page.js("document.querySelector('.vi-status').textContent")
        out["checks"]["single_step"] = "Zeile 6" in (out["status_after_step"] or "")
        key("F5", "F5", 116)
        time.sleep(0.8)
        out["immediate"] = page.js("document.querySelector('.vi-log').textContent")
        out["checks"]["debug_print"] = "n ist 42" in out["immediate"]
        out["checks"]["form_shown"] = page.js("!!document.querySelector('.vba-form:not(.vi-canvas .vba-form)')")
        shot("3-form")
        click("document.querySelector('body > .vba-form .vf-c[data-name=CommandButton1]')")
        time.sleep(0.5)
        out["form_after_click"] = page.js("[document.querySelector('body > .vba-form .vf-c[data-name=TextBox1]')?.value, document.querySelector('body > .vba-form .vf-title span')?.textContent]")
        out["checks"]["click_runs_code"] = out["form_after_click"] == ["geklickt", "Fertig"]
        shot("4-clicked")
        # Speichern: the project with its form goes into the presentation; read back from the saved bytes.
        key("s", "KeyS", 83, 2)
        out["read_back"] = page.js("""(async () => { const P = window.__present; const m = await import('/src/macros.ts');
            return m.modulesOf(P.app).map((x) => ({ name: x.name, kind: x.kind, form: x.form ? x.form.controls.map((c) => c.type + ':' + c.name) : null })); })()""")
        out["checks"]["saved_with_form"] = any(x["form"] == ["CommandButton:CommandButton1", "TextBox:TextBox1"] for x in out["read_back"])
    finally:
        page.close()
    out["summary"] = f"{sum(1 for v in out['checks'].values() if v)}/{len(out['checks'])}"
    p = HERE / "results" / f"{out['date']}-present-vbaide.json"
    p.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))
    print(json.dumps(out["checks"]), out["summary"])


if __name__ == "__main__":
    main()
