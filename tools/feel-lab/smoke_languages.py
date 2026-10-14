"""Smoke test of Step 4b F's languages in the real RaabeWrite app: French and Spanish
spelling, Sprache automatisch erkennen, the Thesaurus pane and Synonyme menu, grammar marks, the Editor pane
(F7), Sprache / Rechtschreibung nicht prüfen written as Word reads them (checked through COM
in a separate, invisible Word instance). Drives the app through the DevTools protocol; the
app runs off screen, so this runs in the background.

    .venv/Scripts/python.exe smoke_languages.py
"""
from __future__ import annotations

import json
import shutil
import tempfile
import time
import traceback
from pathlib import Path

from feellab import targets, winapi

RESULTS: list[tuple[str, bool, str]] = []
TMP = Path(tempfile.mkdtemp(prefix="raabe-smoke-lang-"))
API = "window.__app"
ED = "window.__app.editor()"


def check(name: str, ok: bool, detail: str = "") -> None:
    RESULTS.append((name, bool(ok), detail))
    print(("PASS " if ok else "FAIL ") + name + (f"  [{detail}]" if detail and not ok else ""), flush=True)


def wait(t, expr: str, timeout: float = 6.0):
    end = time.perf_counter() + timeout
    while time.perf_counter() < end:
        v = t.js(expr)
        if v:
            return v
        time.sleep(0.15)
    return t.js(expr)


def marked(t, p: int) -> list[str]:
    """The words of paragraph p with a red wavy line (once the dictionary has answered)."""
    expr = (f"(() => {{ const e = {ED}; const para = e.doc.main[{p}]; e.sel.head; const m = e.extraMarks('main', para) || []; "
            f"return m.filter(x => x.kind === 'spell').map(x => para.text.slice(x.start, x.end)); }})()")
    last = None
    for _ in range(30):
        v = t.js(expr)
        if v == last and v is not None:
            return v
        last = v
        time.sleep(0.2)
    return last or []


def word_langs(path: Path) -> dict:
    import win32com.client
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    try:
        d = app.Documents.Open(str(path), ReadOnly=True, AddToRecentFiles=False)
        out = {"paras": [(p.Range.LanguageID, bool(p.Range.NoProofing)) for p in d.Paragraphs]}
        d.Close(0)
        return out
    finally:
        app.Quit(0)


def main() -> None:
    t = targets.raabe_write("short", background=True)
    try:
        winapi.place_quietly(t.hwnd, -4000, 0, 2000, 2050)
        t.js(f"{API}.newDoc()")
        t.js("__probe.focus()")
        t.send_text("Das Haus ist schön und hat einen Feher im Dach.")
        t.send_key("Enter")
        t.send_text("Nous sommes allés au marchè avec nos amis pour acheter des légumes.")
        t.send_key("Enter")
        t.send_text("Mañana vamos a la plaia con nuestros amigos porque hace calor.")
        t.send_key("Enter")

        # ----- spelling in three languages, the language found by itself -----------------------
        check("German: the misspelled word is marked", marked(t, 0) == ["Feher"], json.dumps(marked(t, 0)))
        check("French text is checked in French", marked(t, 1) == ["marchè"], json.dumps(marked(t, 1)))
        check("Spanish text is checked in Spanish", marked(t, 2) == ["plaia"], json.dumps(marked(t, 2)))
        t.js(f"{ED}.selectRange(1, 10, 10)")
        t.js(f"{API}.refresh?.()")
        lang = wait(t, "document.querySelector('.statusbar').textContent.includes('Französisch (Frankreich)')", 4)
        check("the status bar shows the French sentence's language", lang, t.js("document.querySelector('.statusbar').textContent"))
        fr = t.js("window.__TAURI_INTERNALS__.invoke('spell_suggest', { lang: 'fr-FR', word: 'marchè' })") or []
        check("French suggestions", "marché" in fr, json.dumps(fr, ensure_ascii=False))

        # ----- Thesaurus (Umschalt+F7) and Synonyme ---------------------------------------------
        t.js(f"{ED}.selectRange(0, 14, 14)")
        t.js(f"{API}.run('review.thesaurus')")
        heads = wait(t, "[...document.querySelectorAll('.thespane .thes-heading')].map(x => x.textContent)", 6) or []
        items = t.js("[...document.querySelectorAll('.thespane .thes-item')].map(x => x.textContent)") or []
        check("the Thesaurus pane shows meanings of 'schön'", len(heads) >= 3 and any("hübsch" in x for x in items), json.dumps(heads[:5], ensure_ascii=False))
        check("the pane looks up the word at the caret", t.js("document.querySelector('.thespane .thes-input').value") == "schön")
        t.js("[...document.querySelectorAll('.thespane .thes-item')].find(x => x.textContent.startsWith('hübsch'))?.click()")
        after = wait(t, "document.querySelector('.thespane .thes-input').value === 'hübsch'", 4)
        check("a click on a synonym looks it up (Zurück enabled)", after and t.js("!document.querySelector('.thespane .thes-back').disabled"))
        t.js(f"{API}.run('review.thesaurus')")
        menu = t.js(f"""(async () => {{
            const e = {ED}; const r = e.view.caretRect ? null : null;
            e.selectRange(0, 15, 15);
            e.onContextMenu(400, 300, {{ p: 0, o: 15 }});
            await new Promise(r => setTimeout(r, 800));
            const syn = [...document.querySelectorAll('.menu [role=menuitem], .menu button')].find(b => b.textContent.trim().startsWith('Synonyme'));
            if (!syn) return null;
            syn.dispatchEvent(new MouseEvent('mouseenter', {{ bubbles: true }}));
            syn.click();
            await new Promise(r => setTimeout(r, 600));
            const menus = [...document.querySelectorAll('.menu')];
            return menus[menus.length - 1].innerText.split('\\n').map(s => s.trim()).filter(Boolean);
        }})()""")
        check("the context menu's Synonyme lists synonyms and Thesaurus...", bool(menu) and "Thesaurus..." in menu and len(menu) >= 4, json.dumps(menu, ensure_ascii=False))
        t.send_key("Escape")
        t.send_key("Escape")

        # ----- Editor pane (F7) --------------------------------------------------------------------
        t.js(f"{ED}.selectRange(0, 0, 0)")
        t.js(f"{API}.run('review.spelling')")
        title = wait(t, "document.querySelector('.edpane .ed-title')?.textContent", 10)
        check("F7 opens the Editor pane on the first error", title == "Rechtschreibung", str(title))
        left = t.js("document.querySelector('.edpane .ed-bar > span')?.textContent")
        check("the pane counts the errors", left == "3 verbleibend", str(left))
        word = t.js("document.querySelector('.edpane .ed-word')?.textContent")
        sugg = t.js("[...document.querySelectorAll('.edpane .ed-s')].map(x => x.textContent)") or []
        check("the error and Word-like suggestions", word == "Feher" and "Fehler" in sugg, json.dumps([word, sugg], ensure_ascii=False))
        sim = t.js("document.querySelector('.edpane .ed-sim')?.textContent") or ""
        check("suggestions come with synonyms (ähnlich wie)", len(sim) > 2, sim)
        t.js("[...document.querySelectorAll('.edpane .ed-card')].find(x => x.querySelector('.ed-s').textContent === 'Fehler')?.click()")
        nxt = wait(t, "document.querySelector('.edpane .ed-word')?.textContent === 'marchè'", 6)
        check("Ändern replaces and goes to the next error", nxt and "Fehler im Dach" in t.js(f"{ED}.doc.main[0].text"), t.js(f"{ED}.doc.main[0].text"))
        t.js("[...document.querySelectorAll('.edpane .ed-action')].find(x => x.textContent === 'Einmal ignorieren')?.click()")
        nxt = wait(t, "document.querySelector('.edpane .ed-word')?.textContent === 'plaia'", 6)
        check("Einmal ignorieren goes on to the next error", nxt, str(t.js("document.querySelector('.edpane .ed-word')?.textContent")))
        t.js("document.querySelector('.edpane .tp-head button[aria-label=\"Bereich schließen\"]')?.click()")

        # ----- Grammatik (D-051): blue double underline, suggestions, Editor pane -----------------------
        t.js(f"(() => {{ const e = {ED}; const n = e.doc.main.length - 1; e.selectRange(n, e.doc.main[n].text.length, e.doc.main[n].text.length); }})()")
        t.js("__probe.focus()")
        t.send_key("Enter")
        t.send_text("Ich glaube, das er morgen kommt. Die Kinder spielt im Garten.")
        g = t.js(f"{ED}.doc.main.length - 1")
        expr = (f"(() => {{ const e = {ED}; const para = e.doc.main[{g}]; e.selectRange({g}, 0, 0); const m = e.extraMarks('main', para) || []; "
                f"return m.filter(x => x.kind === 'grammar').map(x => para.text.slice(x.start, x.end)); }})()")
        found = wait(t, f"{expr}.length >= 2 ? {expr} : null", 10)
        check("grammar errors are marked", bool(found) and "das" in found and "spielt" in found,
              json.dumps([t.js(expr), t.js(f"{ED}.doc.main.map(p => p.text)"), t.js(f"{ED}.doc.main[{g}].runs.map(r => r.props.lang)")], ensure_ascii=False))
        menu = t.js(f"""(async () => {{
            const e = {ED}; e.selectRange({g}, 13, 13);
            e.onContextMenu(400, 300, {{ p: {g}, o: 13 }});
            await new Promise(r => setTimeout(r, 800));
            const menus = [...document.querySelectorAll('.menu')];
            return menus.length ? menus[menus.length - 1].innerText.split('\\n').map(s => s.trim()).filter(Boolean).slice(0, 4) : null;
        }})()""")
        check("the context menu suggests 'dass'", bool(menu) and menu[0] == "dass", json.dumps(menu, ensure_ascii=False))
        t.send_key("Escape")
        t.js(f"{ED}.selectRange({g}, 0, 0)")
        t.js(f"{API}.run('review.spelling')")
        title = wait(t, "document.querySelector('.edpane .ed-title')?.textContent", 10)
        check("the Editor pane shows the grammar finding", title == "Grammatik", str(title))
        t.js("document.querySelector('.edpane .tp-head button[aria-label=\"Bereich schließen\"]')?.click()")
        t.js(f"{ED}.replaceRange({g}, 0, {ED}.doc.main[{g}].text.length, '')")

        # ----- Sprache: marked text in a language, nicht prüfen; Word reads both ------------------
        t.js(f"(() => {{ const e = {ED}; e.selectRange(2, 0, e.doc.main[2].text.length); e.setProofing('es-ES', false); e.selectRange(3, 0, 0); }})()")
        t.send_text("Ein Abschnitt ohne Prüfung: Feher.")
        t.js(f"(() => {{ const e = {ED}; e.selectRange(3, 0, e.doc.main[3].text.length); e.setProofing('de-DE', true); }})()")
        check("text marked 'nicht prüfen' is not marked", marked(t, 3) == [], json.dumps(marked(t, 3)))
        out = TMP / "languages.docx"
        check("saves the document", t.js(f"{API}.saveTo({json.dumps(str(out))})") and out.exists())
        w = word_langs(out)
        check("Word: the Spanish paragraph is Spanish", w["paras"][2][0] == 3082, json.dumps(w))
        check("Word: the last paragraph is not proofed", w["paras"][3][1] is True, json.dumps(w))
    except Exception:
        check("no exception", False, traceback.format_exc()[-800:])
    finally:
        t.close()
        shutil.rmtree(TMP, ignore_errors=True)
    passed = sum(1 for _, ok, _ in RESULTS if ok)
    print(f"\n{passed}/{len(RESULTS)} checks passed")


if __name__ == "__main__":
    main()
