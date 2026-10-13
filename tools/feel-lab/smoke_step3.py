"""Functional smoke test of RaabeWrite's Step 3 features in the real app: lists, tables,
page setup, sections, headers and footers with page numbers, pictures, find and replace,
spell checking, formatted copy and paste, print preview. Files are checked with Microsoft
Word (separate, invisible instance). The app runs off screen without the focus (D-028).

    .venv/Scripts/python.exe smoke_step3.py
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
TMP = Path(tempfile.mkdtemp(prefix="raabe-smoke3-"))
OFF = -4000
API = "window.__app"
ED = "window.__app.editor()"
# A 4 x 3 px PNG (red), small enough to write inline.
PNG = "iVBORw0KGgoAAAANSUhEUgAAAAQAAAADCAIAAAA7ljmRAAAAEklEQVR4nGP4z8AARAwM+FgAuv8L9e6ZtbAAAAAASUVORK5CYII="


def check(name: str, ok: bool, detail: str = "") -> None:
    RESULTS.append((name, bool(ok), detail))
    print(("PASS " if ok else "FAIL ") + name + (f"  [{detail}]" if detail and not ok else ""), flush=True)


def word_info(path: Path) -> dict:
    import win32com.client
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    try:
        d = app.Documents.Open(str(path), ReadOnly=True, AddToRecentFiles=False)
        sec = d.Sections(1)
        lists = []
        for p in d.Paragraphs:
            lf = p.Range.ListFormat
            if lf.ListType:
                lists.append((p.Range.Text.rstrip("\r\x07"), lf.ListString, lf.ListLevelNumber))
        info = {
            "paras": [p.Range.Text.rstrip("\r\x07") for p in d.Paragraphs],
            "tables": d.Tables.Count,
            "table": [[d.Tables(1).Cell(r, c).Range.Text[:-2] for c in range(1, d.Tables(1).Columns.Count + 1)] for r in range(1, d.Tables(1).Rows.Count + 1)] if d.Tables.Count else [],
            "lists": lists,
            "sections": d.Sections.Count,
            "orient": [d.Sections(k).PageSetup.Orientation for k in range(1, d.Sections.Count + 1)],
            "columns": [d.Sections(k).PageSetup.TextColumns.Count for k in range(1, d.Sections.Count + 1)],
            "footer": sec.Footers(1).Range.Text.strip(),
            "header": sec.Headers(1).Range.Text.strip(),
            "pictures": d.InlineShapes.Count,
            "pages": d.ComputeStatistics(2),
            "bold": [bool(d.Range(i, i + 1).Font.Bold) for i in range(min(60, d.Content.End - 1))],
        }
        d.Close(0)
        return info
    finally:
        app.Quit(0)


def wait(t, expr: str, timeout: float = 5.0):
    end = time.perf_counter() + timeout
    while time.perf_counter() < end:
        v = t.js(expr)
        if v:
            return v
        time.sleep(0.1)
    return t.js(expr)


def main() -> None:
    t = targets.raabe_write("short", background=True)
    try:
        winapi.place_quietly(t.hwnd, OFF, 0, 2000, 2050)
        t.js(f"{API}.newDoc()")
        t.js("__probe.focus()")

        # ----- lists: AutoFormat, Enter continues, Tab demotes, Enter on empty item ends --------
        t.send_text("Aufgaben")
        t.send_key("Enter")
        t.send_text("1. Lies den Text.")
        t.send_key("Enter")
        t.send_text("Beantworte die Fragen.")
        t.send_key("Enter")
        t.send_key("Tab")
        t.send_text("Frage a")
        t.send_key("Enter")
        t.send_key("Enter")
        t.send_key("Tab", shift=True)
        t.send_key("Enter")
        info = t.js(f"{API}.docInfo()")
        items = [p for p in info["paras"] if (p.get("numId") or 0) > 0]
        check("typing \"1. \" starts a numbered list", len(items) >= 3 and items[0]["text"] == "Lies den Text.", json.dumps(info["paras"][:5]))
        check("Tab makes a list item one level deeper", any(p.get("ilvl") == 1 and p["text"] == "Frage a" for p in items), json.dumps(items))
        check("list items get the Listenabsatz style", all(p["style"] == "Listenabsatz" for p in items), json.dumps([p["style"] for p in items]))

        # ----- table: insert, type, Tab to the next cell, new row after the last cell ------------
        t.send_text("Material:")
        t.send_key("Enter")
        t.js(f"{ED}.insertTable(2, 3)")
        for word in ["Stoff", "Menge", "Hinweis", "Wasser", "100 ml", "kalt"]:
            t.send_text(word)
            t.send_key("Tab")
        t.send_text("Salz")
        info = t.js(f"{API}.docInfo()")
        cells = [p for p in info["paras"] if p["cell"]]
        check("Tab in the last cell adds a row", len(cells) == 9 and cells[6]["text"] == "Salz", json.dumps([c["text"] for c in cells]))
        tabs = wait(t, "[...document.querySelectorAll('.tabrow .tab.contextual')].map(b => b.textContent).join(',')")
        check("Tabellenentwurf and Layout appear in a table", tabs == "Tabellenentwurf,Layout", str(tabs))
        t.js(f"{API}.run('table.insertRight')")
        cells = [p for p in t.js(f"{API}.docInfo()")["paras"] if p["cell"]]
        check("Rechts einfügen adds a column", len(cells) == 12, str(len(cells)))
        t.js(f"{API}.run('edit.undo')")

        # ----- headers, footers, page numbers, page setup, sections ---------------------------
        t.js(f"{API}.run('pageNumber.ofPages')")
        check("Seite X von Y opens the footer", wait(t, f"{API}.activeRibbonTab() === 'headerFooter'"), str(t.js(f"{API}.activeRibbonTab()")))
        t.js(f"{API}.run('hf.close')")
        check("closing the footer returns to the text", t.js(f"{API}.docInfo().story") == "main")
        n = len(t.text())
        t.set_selection(n, n)
        t.js(f"{API}.run('section.nextPage')")
        t.js("__probe.focus()")
        t.send_text("Ein Abschnitt im Querformat.")
        t.js(f"{API}.run('orientation.landscape')")
        info = t.js(f"{API}.docInfo()")
        check("a section break starts a new page", info["pages"] == 2 and sum(p["sect"] for p in info["paras"]) == 1, json.dumps({"pages": info["pages"]}))

        # ----- pictures --------------------------------------------------------------------------
        t.set_selection(0, 0)
        t.js(f"(() => {{ const b = Uint8Array.from(atob('{PNG}'), c => c.charCodeAt(0)); {ED}.insertImage({{ bytes: b, type: 'image/png' }}, 60, 45); }})()")
        check("an inserted picture is selected", t.js(f"!!{ED}.selectedImage()"))
        check("Bildformat appears for a picture", wait(t, "[...document.querySelectorAll('.tabrow .tab.contextual')].some(b => b.textContent === 'Bildformat')"))
        t.js(f"{ED}.setImageSize(120, 90)")
        check("a picture can be resized", abs((t.js(f"{ED}.selectedImage().obj.width") or 0) - 120) < 0.1)

        # ----- find and replace ---------------------------------------------------------------
        n_found = t.js(f"{ED}.findAll('Text').length")
        check("Suchen finds the word", n_found == 1, str(n_found))
        replaced = t.js(f"{ED}.replaceAll('Fragen', 'Aufgaben')")
        check("Alle ersetzen replaces", replaced == 1 and "Beantworte die Aufgaben." in t.text(), str(replaced))

        # ----- spell checking --------------------------------------------------------------------
        n = len(t.text())
        t.set_selection(n, n)
        t.js("__probe.focus()")
        t.send_text(" Ein Feher im Satz. ")
        squiggles = wait(t, "[...document.querySelectorAll('.page div')].filter(d => (d.style.background || '').includes('svg')).length", 8)
        check("a misspelled word gets a red wavy line", (squiggles or 0) >= 1, str(squiggles))
        sugg = t.js("window.__TAURI_INTERNALS__.invoke('spell_suggest', { lang: 'de-DE', word: 'Feher' })") or []
        check("suggestions for a misspelled word", "Fehler" in sugg, json.dumps(sugg))
        ok_words = t.js("window.__TAURI_INTERNALS__.invoke('spell_check', { lang: 'de-DE', words: ['Fehler', 'Feher', 'Kohlenstoffdioxid'] })")
        check("the German dictionary knows words", ok_words == [True, False, True], json.dumps(ok_words))

        # ----- formatted copy and paste -----------------------------------------------------------
        html = t.js(f"(() => {{ const e = {ED}; e.selectRange(0, 0, 8); return e.selectionHtml(); }})()") or ""
        check("copying puts HTML with formatting on the clipboard", "<span style=" in html and "font-family" in html, html[:80])
        n = len(t.text())
        t.set_selection(n, n)
        t.js(f"{ED}.pasteHtml('<p>Aus <b>Word</b> kopiert</p>')")
        last = t.text().split("\n")[-1]
        check("HTML from Word is pasted with formatting", last.endswith("Aus Word kopiert"), last)

        # ----- save, Word reads the structure -----------------------------------------------------
        out = TMP / "step3.docx"
        check("saves the document", t.js(f"{API}.saveTo({json.dumps(str(out))})") and out.exists())
        w = word_info(out)
        check("Word: the table with its cells", w["tables"] == 1 and w["table"][0][:2] == ["Stoff", "Menge"] and w["table"][2][0] == "Salz", json.dumps(w["table"]))
        check("Word: the numbered list", len(w["lists"]) >= 2 and w["lists"][0][1] == "1." and any(l[2] == 2 for l in w["lists"]), json.dumps(w["lists"]))
        check("Word: two sections, the second in landscape", w["sections"] == 2 and w["orient"] == [0, 1], json.dumps(w["orient"]))
        check("Word: footer \"Seite 1 von 2\"", w["footer"].replace("\r", "") == "Seite 1 von 2", repr(w["footer"]))
        check("Word: the picture", w["pictures"] == 1, str(w["pictures"]))
        check("Word: two pages", w["pages"] == 2, str(w["pages"]))

        # ----- print preview ---------------------------------------------------------------------
        t.js(f"{API}.run('file.print')")
        check("Drucken shows the page preview", wait(t, "!!document.querySelector('.print-preview .print-page')"), "")
        check("the preview counts the pages", "von 2" in (t.js("document.querySelector('.print-nav')?.textContent") or ""))
        t.send_key("Escape")
    except Exception:
        check("no exception", False, traceback.format_exc()[-400:])
    finally:
        t.close()
        shutil.rmtree(TMP, ignore_errors=True)
    passed = sum(ok for _, ok, _ in RESULTS)
    print(f"\n{passed}/{len(RESULTS)} checks passed")


if __name__ == "__main__":
    main()
