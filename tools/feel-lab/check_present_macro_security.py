"""Macros' safety in use (D-079, D-100, D-106), in a headless Edge on the dev server (no real input: clicks through
the page): a .pptm opens with its macros off and the message bar (Inhalt aktivieren); a macro run before
asks first; after Inhalt aktivieren it runs. A macro that writes a file asks every time (Makro: Dateizugriff):
Nicht zulassen ends it with Laufzeitfehler '70' and writes nothing. A runtime error shows Office's box
(Laufzeitfehler '13':, Beenden, Debuggen). Uses captures/macro-suite/suite.pptm (check_present_macro_suite.py).

    .venv/Scripts/python.exe check_present_macro_security.py      (npm run dev:present must be running)

Result: results/<date>-present-macro-security.json.
"""
from __future__ import annotations

import datetime
import json
import time
from pathlib import Path

from feellab.present_page import PresentPage

HERE = Path(__file__).parent
PPTM = HERE / "captures" / "macro-suite" / "suite.pptm"


def main() -> None:
    out: dict = {"date": datetime.date.today().isoformat(), "checks": {}}
    page = PresentPage()

    def dialog() -> dict | None:
        return page.js("""(() => { const d = [...document.querySelectorAll('.dialog-layer .dialog')].pop();
            return d ? { text: d.innerText.replace(/\\s+/g, ' ').slice(0, 300), buttons: [...d.querySelectorAll('button')].map((b) => b.textContent) } : null; })()""")

    def press(label: str) -> None:
        page.js(f"""(() => {{ const d = [...document.querySelectorAll('.dialog-layer .dialog')].pop();
            const b = d && [...d.querySelectorAll('button')].find((x) => x.textContent === {json.dumps(label)}); if (b) b.click(); return !!b; }})()""")
        time.sleep(0.4)

    try:
        page.open_file(PPTM)
        time.sleep(0.5)
        out["bar"] = page.js("document.querySelector('.macro-bar')?.innerText.replace(/\\s+/g, ' ') ?? null")
        out["checks"]["opens_with_macros_off"] = bool(out["bar"]) and "Inhalt aktivieren" in out["bar"]
        # A macro started before Inhalt aktivieren: the security notice comes first.
        page.js("(async () => { const m = await import('/src/macros.ts'); window.__run = m.runMacro(window.__present.app, 'Quiz'); })()")
        time.sleep(0.6)
        d = dialog()
        out["notice"] = d
        out["checks"]["asks_before_running"] = bool(d) and "Inhalt aktivieren" in d["buttons"]
        if d:
            press("Abbrechen")
        out["checks"]["cancel_runs_nothing"] = page.js("window.__present.app.deck.view(window.__present.app.deck.slides[4].path).shapes.find((s) => s.name === 'Punkte').text.paras[0].runs[0].text") == "0"
        page.js("document.querySelector('.macro-bar button.btn').click()")
        out["checks"]["bar_enables"] = page.js("(async () => (await import('/src/macros.ts')).macrosEnabled(window.__present.app))()")
        # A file: the question names the file; Nicht zulassen gives Laufzeitfehler '70'.
        page.js("(async () => { const m = await import('/src/macros.ts'); window.__run = m.runMacro(window.__present.app, 'NotizenExportieren'); })()")
        time.sleep(0.6)
        d = dialog()
        out["file_question"] = d
        out["checks"]["asks_for_file"] = bool(d) and "Makro: Dateizugriff" in d["text"] and "notizen.txt" in d["text"] and [b for b in d["buttons"] if b] == ["Zulassen", "Nicht zulassen"]
        if d:
            press("Nicht zulassen")
        time.sleep(0.4)
        d = dialog()
        out["refused"] = d
        out["checks"]["refusal_is_error_70"] = bool(d) and "Laufzeitfehler '70'" in d["text"] and [b for b in d["buttons"] if b] == ["Beenden", "Debuggen"]
        if d:
            press("Beenden")
        # A runtime error of the macro's own: Office's box.
        page.js("""(async () => { const P = window.__present; const m = await import('/src/macros.ts');
            const mods = m.modulesOf(P.app); mods.push({ name: 'Kaputt', kind: 'standard', code: 'Sub Kaputt()\\r\\n    Dim x As Integer\\r\\n    x = \"abc\"\\r\\nEnd Sub' }); m.setModules(P.app, mods);
            window.__run = m.runMacro(P.app, 'Kaputt'); })()""")
        time.sleep(0.6)
        d = dialog()
        out["runtime_error"] = d
        out["checks"]["runtime_error_box"] = bool(d) and "Laufzeitfehler '13'" in d["text"] and "Zeile 3" in d["text"]
        if d:
            press("Beenden")
    finally:
        page.close()
    out["summary"] = f"{sum(1 for v in out['checks'].values() if v)}/{len(out['checks'])}"
    p = HERE / "results" / f"{out['date']}-present-macro-security.json"
    import tempfile as _t
    text = json.dumps(out, ensure_ascii=False, indent=1)
    for tmp in {_t.gettempdir(), str(Path(_t.gettempdir()).resolve())}:
        text = text.replace(json.dumps(tmp)[1:-1], "%TEMP%").replace(tmp, "%TEMP%")
    p.write_text(text, encoding="utf-8", newline=chr(10))
    print(json.dumps(out["checks"]), out["summary"])


if __name__ == "__main__":
    main()
