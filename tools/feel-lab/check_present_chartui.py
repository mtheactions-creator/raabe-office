"""RaabePresent's selected chart with the shared chart UI (D-132), tried by hand: the release exe off screen, driven
with trusted DevTools input (no real mouse or keyboard, no lock). A column chart is inserted through Einfügen >
Diagramm, clicked, and then: Format > Auswahl formatieren and the context menu's Diagrammbereich formatieren... must
open the chart's Format pane (before: the shape pane, whose controls did nothing to the chart); Aktuelle Auswahl >
Diagrammtitel opens the title's pane; a fill chosen in the pane changes the chart; WordArt-Formate > Textfüllung
colours the chart's text; the three buttons stand beside the chart and Diagrammelemente switches the legend off;
Strg+Z takes it back; the written file keeps the changes (the chart part read again).

    .venv/Scripts/python.exe check_present_chartui.py        (RAABE_PRESENT_EXE points at another build)

Result: results/2026-11-09-present-chartui.json.
"""
from __future__ import annotations

import json
import sys
import time
import traceback
from pathlib import Path

from feellab.cdp_app import AppExe

HERE = Path(__file__).parent
RESULT = HERE / "results" / "2026-11-09-present-chartui.json"


class Run:
    def __init__(self) -> None:
        self.cases: list[dict] = []

    def check(self, name: str, ok: bool, detail="") -> None:
        self.cases.append({"case": name, "passed": bool(ok), "detail": "" if ok else str(detail)[:500]})
        print(("PASS " if ok else "FAIL ") + name + ("" if ok else f"  {str(detail)[:300]}"), flush=True)


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    run = Run()
    a = AppExe("present", port=9483, size=(3840, 2000))
    # The page loads once more after the window shows: connect to the page that is there then.
    import urllib.request
    from feellab.cdp_app import CdpPage
    deadline = time.time() + 40
    while True:
        time.sleep(1.0)
        try:
            with urllib.request.urlopen("http://127.0.0.1:9483/json/list", timeout=2) as r:
                pages = [x for x in json.load(r) if x.get("type") == "page"]
            a.page = CdpPage(pages[0]["webSocketDebuggerUrl"])
            if a.page.js("typeof window.__present === 'object' && !!window.__present.app"):
                break
        except Exception:  # noqa: BLE001
            pass
        if time.time() > deadline:
            raise TimeoutError("RaabePresent's page did not come up")
    p = a.page

    def rect(sel: str, text: str | None = None):
        r = p.js(f"""(() => {{ const all = [...document.querySelectorAll({json.dumps(sel)})].filter((x) => x.getClientRects().length);
          const e = {json.dumps(text)} === null ? all[0] : all.find((x) => (x.getAttribute('aria-label') ?? '') === {json.dumps(text)} || x.textContent.trim() === {json.dumps(text)});
          if (!e) return null; const r = e.getBoundingClientRect(); return [r.left + r.width / 2, r.top + r.height / 2]; }})()""")
        if not r:
            raise RuntimeError(f"not found: {sel} {text!r}")
        return r

    def click(sel: str, text: str | None = None, count: int = 1) -> None:
        x, y = rect(sel, text)
        p.click(x, y, count=count)
        time.sleep(0.4)

    model = lambda: p.js("(() => { const m = __present.app.chartUi.model(); if (!m) return null; const { raw, ...r } = m; return JSON.parse(JSON.stringify(r)); })()")  # noqa: E731
    title = lambda: p.js("__present.app.chartUi.paneTitle()")  # noqa: E731
    try:
        time.sleep(1.5)
        # Einfügen > Diagramm > Gruppierte Säulen, by hand.
        click(".tab", "Einfügen")
        click(".ribbon-card [aria-label]", "Diagramm")
        click(".menu .menu-item, .menu button", "Gruppierte Säulen")
        time.sleep(1.5)
        cid = p.js("(() => { const a = __present.app; return a.deck.view(a.slidePath()).shapes.find((s) => s.frame === 'chart')?.id ?? null; })()")
        run.check("Einfügen > Diagramm > Gruppierte Säulen inserts a chart", cid is not None, cid)
        p.js("__present.app.editor.select([])")
        time.sleep(0.3)
        x, y = rect(f'.stage [data-id="{cid}"]')
        p.click(x, y)
        time.sleep(0.6)
        run.check("a click selects the chart (Diagrammentwurf and Format show)", p.js("__present.app.chartUi.selected()"), "")
        btns = p.js("[...document.querySelectorAll('.pchart-overlay:not([hidden]) .chart-btn')].map((b) => b.getAttribute('aria-label'))")
        run.check("the three chart buttons beside the chart", btns == ["Diagrammelemente", "Diagrammformatvorlagen", "Diagrammfilter"], btns)
        click(".tab", "Format")
        click(".ribbon-card [data-group=current] [data-keytip=U]")  # Auswahl formatieren
        run.check("Format > Auswahl formatieren opens the chart's pane (Diagrammbereich formatieren)", title() == "Diagrammbereich formatieren", title())
        # A fill in the pane: Einfarbige Füllung changes the chart area.
        click(".taskpane:not([hidden]) label", "Einfarbige Füllung")
        m = model()
        fill = ((m or {}).get("format") or {}).get("chart", {}).get("fill")
        run.check("the pane changes the chart: Diagrammbereich > Einfarbige Füllung", bool(fill) and fill.get("type") == "solid", fill)
        # Aktuelle Auswahl > Diagrammtitel.
        click(".ribbon-card [aria-label]", "Diagrammelemente")
        click(".menu button", "Diagrammtitel")
        run.check("Aktuelle Auswahl > Diagrammtitel opens the title's pane", title() == "Diagrammtitel formatieren", title())
        # WordArt-Formate > Textfüllung: a colour from the palette (Rot of the standard colours).
        before = json.dumps(model(), sort_keys=True)
        x, y = rect(".ribbon-card [data-group=wordart] [aria-label='Weitere Optionen']")
        p.click(x, y)
        time.sleep(0.5)
        click("button.swatch", "Rot")
        m = model()
        col = (((m or {}).get("format") or {}).get("title") or {}).get("text", {}).get("color")
        run.check("WordArt-Formate > Textfüllung > Rot colours the chart's text", col is not None and col.upper().endswith("FF0000") and json.dumps(m, sort_keys=True) != before, col)
        # Diagrammelemente > Legende off, Strg+Z back.
        p.key("Escape")
        if not p.js("__present.app.chartUi.selected()"):
            x, y = rect(f'.stage [data-id="{cid}"]')
            p.click(x, y)
            time.sleep(0.5)
        click(".pchart-overlay .chart-btn", "Diagrammelemente")
        click(".cb-fly .cb-row", "Legende")
        off = (model() or {}).get("legend")
        p.key("Escape")
        p.key("z", "ctrl")
        time.sleep(0.6)
        back = (model() or {}).get("legend")
        run.check("Diagrammelemente > Legende off, Strg+Z brings it back", off == "none" and back not in (None, "none"), [off, back])
        # The file keeps the changes: the written deck's chart part read again.
        kept = p.js("""(() => { const a = __present.app; const path = a.deck.view(a.slidePath()).shapes.find((s) => s.frame === 'chart').chart; const x = a.deck.pptx.text(path);
          return [x.includes('FF0000'), /<c:spPr><a:solidFill>/.test(x.slice(x.lastIndexOf('</c:chart>')))]; })()""")
        run.check("the chart part keeps the text colour and the chart area's fill", kept == [True, True], kept)
        # The context menu's Diagrammbereich formatieren...
        p.key("Escape")
        x, y = rect(f'.stage [data-id="{cid}"]')
        p.click(x, y)
        p.click(x, y, button="right")
        time.sleep(0.4)
        click(".menu button", "Diagrammbereich formatieren...")
        run.check("context menu > Diagrammbereich formatieren... opens the chart's pane", title() == "Diagrammbereich formatieren", title())
        errs = [e for e in a.page.errors if "chart" in e.lower()]
        run.check("no chart errors in the console", not errs, errs[:3])
    except Exception:  # noqa: BLE001
        tb = traceback.format_exc()
        print(tb)
        run.check("run completed", False, tb.splitlines()[-1])
    finally:
        a.close()
    passed = sum(1 for c in run.cases if c["passed"])
    RESULT.write_text(json.dumps({"passed": passed, "total": len(run.cases), "cases": run.cases}, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(f"{passed}/{len(run.cases)} passed")


if __name__ == "__main__":
    main()
