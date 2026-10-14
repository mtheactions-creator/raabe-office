"""Functional smoke test of RaabeWrite's Step 4 features in the real app: foot- and endnotes,
table of contents, comments with replies, tracked changes, equations and the Raabeschule
templates. Saved files are checked with Microsoft Word (separate, invisible instance): Word
must see every element, and its own update of the table of contents must give the page
numbers RaabeWrite wrote. The app runs off screen without the focus (D-028).

    .venv/Scripts/python.exe smoke_step4.py
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
TMP = Path(tempfile.mkdtemp(prefix="raabe-smoke4-"))
OFF = -4000
API = "window.__app"
ED = "window.__app.editor()"


def check(name: str, ok: bool, detail: str = "") -> None:
    RESULTS.append((name, bool(ok), detail))
    print(("PASS " if ok else "FAIL ") + name + (f"  [{detail}]" if detail and not ok else ""), flush=True)


def toc_entries(d) -> list[str]:
    """Entry texts with page numbers (Word's update leaves an empty paragraph at the end)."""
    toc = d.TablesOfContents(1)
    return [x for x in (p.Range.Text.strip("\r\x07 ") for p in toc.Range.Paragraphs) if x]


def plain_math(s: str) -> str:
    """Word's equation text uses mathematical italic letters and a paragraph mark per box."""
    import unicodedata
    return unicodedata.normalize("NFKC", s).replace("\r", "").replace(" ", "")


def word_info(path: Path) -> dict:
    import win32com.client
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    try:
        d = app.Documents.Open(str(path), ReadOnly=True, AddToRecentFiles=False)
        info: dict = {
            "footnotes": d.Footnotes.Count,
            "footnote": d.Footnotes(1).Range.Text if d.Footnotes.Count else "",
            "endnotes": d.Endnotes.Count,
            "comments": d.Comments.Count,
            "comment": [d.Comments(i).Range.Text for i in range(1, d.Comments.Count + 1)],
            "scope": [d.Comments(i).Scope.Text for i in range(1, d.Comments.Count + 1)],
            "replies": d.Comments(1).Replies.Count if d.Comments.Count else 0,
            "revisions": [(d.Revisions(i).Type, d.Revisions(i).Range.Text) for i in range(1, d.Revisions.Count + 1)],
            "track": bool(d.TrackRevisions),
            "omaths": d.OMaths.Count,
            "omath": [(d.OMaths(i).Type, d.OMaths(i).Range.Text) for i in range(1, d.OMaths.Count + 1)],
            "tocs": d.TablesOfContents.Count,
            "pages": d.ComputeStatistics(2),
        }
        if d.TablesOfContents.Count:
            info["toc"] = toc_entries(d)
            d.TablesOfContents(1).Update()
            info["toc_word"] = toc_entries(d)
        d.Close(0)
        return info
    finally:
        app.Quit(0)


def word_template(path: Path) -> dict:
    import win32com.client
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    try:
        d = app.Documents.Open(str(path), ReadOnly=True, AddToRecentFiles=False)
        heads = []
        for p in d.Paragraphs:
            if p.OutlineLevel <= 3:
                heads.append((p.Range.ListFormat.ListString, p.Range.Text.strip("\r\x07 ")))
        info = {
            "pages": d.ComputeStatistics(2), "tocs": d.TablesOfContents.Count, "tables": d.Tables.Count, "omaths": d.OMaths.Count,
            "heads": heads, "footer": d.Sections(1).Footers(1).Range.Text.strip(), "first": bool(d.Sections(1).PageSetup.DifferentFirstPageHeaderFooter),
            "font": d.Styles(-1).Font.Name, "spacing": d.Styles(-1).ParagraphFormat.LineSpacing,
        }
        if d.TablesOfContents.Count:
            info["toc"] = toc_entries(d)
            d.TablesOfContents(1).Update()
            info["toc_word"] = toc_entries(d)
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


def end_of_text(t) -> None:
    """Caret at the end of the main text (also coming from a note)."""
    t.js(f"(() => {{ const e = {ED}; const n = e.doc.main.length - 1; const o = e.doc.main[n].text.length; e.selectRange(n, o, o); }})()")
    t.js("__probe.focus()")


def main() -> None:
    t = targets.raabe_write("short", background=True)
    try:
        winapi.place_quietly(t.hwnd, OFF, 0, 2000, 2050)
        t.js(f"{API}.newDoc()")
        t.js("__probe.focus()")

        # ----- headings and a footnote (Strg+Alt+F) ---------------------------------------------
        t.send_text("Einleitung")
        t.js(f"{API}.run('style.h1')")
        t.send_key("End")
        t.send_key("Enter")
        t.send_text("Fotosynthese wandelt Licht in chemische Energie um.")
        t.send_key("f", ctrl=True, alt=True)
        info = t.js(f"{API}.docInfo()")
        check("Strg+Alt+F inserts a footnote and opens it", info["story"] == "fn:1" and "fn:1" in info["stories"], json.dumps({k: info[k] for k in ("story", "stories")}))
        t.send_text("Vgl. Biologiebuch, S. 42.")
        note = t.js(f"{ED}.doc.story('fn:1').map(p => p.text).join('|')")
        check("typing goes into the footnote", note.endswith("Vgl. Biologiebuch, S. 42."), repr(note))
        lines = t.js(f"{ED}.layout.pages[0].frames.map(i => {ED}.layout.frames[i]).filter(f => f.kind === 'note').length")
        check("the footnote is laid out at the page bottom", (lines or 0) >= 1, str(lines))
        end_of_text(t)
        t.send_key("Enter")
        t.send_text("Hauptteil")
        t.js(f"{API}.run('style.h1')")
        t.send_key("End")
        t.send_key("Enter")
        t.send_text("Die Lichtreaktion findet in den Thylakoiden statt.")
        t.js(f"{API}.run('ref.endnote')")
        t.send_text("Quelle: Unterrichtsmaterial.")
        check("Endnote einfügen adds an endnote", "en:1" in t.js(f"{API}.docInfo()")["stories"])
        end_of_text(t)

        # ----- comments ----------------------------------------------------------------------
        t.js(f"{ED}.selectRange(1, 0, 12)")
        t.send_key("k", ctrl=True, alt=True)
        ok = wait(t, "!!document.querySelector('.ccard textarea')")
        check("Strg+Alt+K opens a new comment card", ok)
        t.js("(() => { const a = document.querySelector('.ccard textarea'); a.value = 'Begriff erklären!'; "
             "a.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true, bubbles: true })); })()")
        comments = t.js(f"{ED}.comments().map(c => c.text)")
        check("posting the comment stores its text", comments == ["Begriff erklären!"], json.dumps(comments))
        t.js("(() => { const a = document.querySelector('.ccard textarea'); a.value = 'Erledigt.'; "
             "a.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true, bubbles: true })); })()")
        replies = t.js(f"{ED}.comments().filter(c => c.parent).length")
        check("a reply joins the thread", replies == 1, str(replies))
        check("the card sits beside the page", wait(t, f"{ED}.markupLayer.querySelectorAll('.ccard').length === 1"))

        # ----- tracked changes (Strg+Umschalt+E) -----------------------------------------------
        t.js("__probe.focus()")
        t.send_key("e", ctrl=True, shift=True)
        check("Strg+Umschalt+E turns tracking on", t.js(f"{ED}.trackChanges") is True)
        end_of_text(t)
        t.send_text(" Neu eingefügt.")
        t.js(f"{ED}.selectRange(3, 4, 17)")
        t.js("__probe.focus()")
        t.send_key("Delete")
        p3 = t.js(f"{ED}.doc.main[3]")
        kinds = [r["props"].get("rev", {}).get("kind") for r in p3["runs"]]
        check("deleted text stays, marked as deleted", "Lichtreaktion" in p3["text"] and "del" in kinds, json.dumps(kinds))
        check("typed text is marked as inserted", "ins" in kinds, json.dumps(kinds))
        t.js(f"{ED}.setMarkup('original')")
        orig = t.js(f"{ED}.layout.paras[3].lines.map(l => l.width).reduce((a, b) => a + b, 0)")
        t.js(f"{ED}.setMarkup('all')")
        full = t.js(f"{ED}.layout.paras[3].lines.map(l => l.width).reduce((a, b) => a + b, 0)")
        check("Markup: Original hides the insertion", orig < full, f"{orig} < {full}")
        t.js(f"{API}.run('review.trackChanges')")

        # ----- equations ---------------------------------------------------------------------
        end_of_text(t)
        t.send_key("Enter")
        t.send_text("Für die Energie gilt ")
        t.send_key("0", alt=True, shift=True, vk=48)
        ok = wait(t, "!!document.querySelector('.math-session')", 8)
        check("Alt+Umschalt+0 opens the equation editor", ok)
        check("the Formel tab appears", wait(t, f"{API}.activeRibbonTab() === 'equation'"), str(t.js(f"{API}.activeRibbonTab()")))
        # Our own formula editor (D-039): typed in the line, built up on the space as in Word.
        t.send_text("E=mc^2 ")
        t.send_key("End")
        t.send_key("ArrowRight")
        wait(t, "!document.querySelector('.math-session')", 4)
        latex = t.js(f"{ED}.doc.main.flatMap(p => p.objects ?? []).filter(o => o.kind === 'math').map(o => o.latex)")
        check("the equation is stored with its superscript", [x.replace("{", "").replace("}", "") for x in latex or []] == ["E=mc^2"], json.dumps(latex))
        maths = t.js("document.querySelectorAll('.page math').length")
        check("the equation is drawn as MathML", (maths or 0) >= 1, str(maths))
        t.js("__probe.focus()")
        t.send_text(".")
        t.send_key("Enter")
        t.js(f"{API}.run('insert.equation')")
        wait(t, "!!document.querySelector('.math-session')", 8)
        t.send_text("x=(-b+\\sqrt (b^2-4ac))/2a ")
        t.send_key("End")
        t.send_key("ArrowRight")
        wait(t, "!document.querySelector('.math-session')", 4)
        check("a display equation in its own paragraph", t.js(f"{ED}.doc.main[{ED}.doc.main.length - 1].text") == "\ufffc",
              json.dumps(t.js(f"{ED}.doc.main.slice(-2).map(p => p.text)")))

        # ----- table of contents ----------------------------------------------------------------
        t.js(f"{ED}.selectRange(0, 0, 0)")
        t.js(f"{API}.run('toc.auto1')")
        info = t.js(f"{API}.docInfo()")
        styles = [p["style"] for p in info["paras"][:3]]
        check("Automatische Tabelle 1 inserts a table of contents", styles == ["Inhaltsverzeichnisberschrift", "Verzeichnis1", "Verzeichnis1"], json.dumps(styles))

        # ----- save, Word reads it ---------------------------------------------------------------
        out = TMP / "step4.docx"
        check("saves the document", t.js(f"{API}.saveTo({json.dumps(str(out))})") and out.exists())
        w = word_info(out)
        check("Word: one footnote with its text", w["footnotes"] == 1 and "Biologiebuch" in w["footnote"], json.dumps(w["footnote"]))
        check("Word: one endnote", w["endnotes"] == 1, str(w["endnotes"]))
        check("Word: the comment and its reply", w["comments"] == 2 and w["replies"] == 1 and w["comment"][0] == "Begriff erklären!", json.dumps(w["comment"]))
        check("Word: the comment covers the commented text", w["scope"][0] == "Fotosynthese", json.dumps(w["scope"]))
        kinds = sorted({k for k, _ in w["revisions"]})
        check("Word: tracked insertion and deletion", kinds == [1, 2], json.dumps(w["revisions"]))
        check("Word: two equations, one display", w["omaths"] == 2 and sorted(k for k, _ in w["omath"]) == [0, 1], json.dumps(w["omath"]))
        check("Word: the equation text", plain_math(w["omath"][0][1]) == "E=mc2", json.dumps(w["omath"][0][1]))
        check("Word: a table of contents", w["tocs"] == 1, str(w["tocs"]))
        check("Word: its own TOC update gives the same entries and page numbers", w.get("toc") == w.get("toc_word"),
              json.dumps({"ours": w.get("toc"), "word": w.get("toc_word")}, ensure_ascii=False))

        # ----- reopen: everything is editable -----------------------------------------------------
        t.js(f"{API}.openPath({json.dumps(str(out))})")
        wait(t, f"{API}.docs().some(d => d.path === {json.dumps(str(out))} && d.active)", 8)
        reopened = t.js(f"(() => {{ const e = {ED}; return {{ unsupported: {API}.docs().find(d => d.active).unsupported, comments: e.comments().length, "
                        f"math: e.doc.main.flatMap(p => p.objects ?? []).filter(o => o.kind === 'math').map(o => o.latex) }}; }})()")
        check("reopened: nothing left unsupported", reopened["unsupported"] == [], json.dumps(reopened["unsupported"]))
        check("reopened: comments and equations are back", reopened["comments"] == 2 and len(reopened["math"]) == 2, json.dumps(reopened))

        # ----- templates --------------------------------------------------------------------------
        for name in ["arbeitsblatt", "versuchsprotokoll", "handout", "facharbeit"]:
            t.js(f"{API}.newFromTemplate('{name}')")
            dirty = t.js(f"{API}.docs().find(d => d.active).dirty")
            path = TMP / f"{name}.docx"
            saved = t.js(f"{API}.saveTo({json.dumps(str(path))})")
            pages = t.js(f"{ED}.layout.pageCount")
            wt = word_template(path)
            check(f"template {name}: opens unchanged and saves", dirty is False and saved and path.exists(), f"dirty={dirty}")
            check(f"template {name}: Word counts the same pages", wt["pages"] == pages, f"word {wt['pages']} vs ours {pages}")
            if name == "facharbeit":
                check("Facharbeit: numbered headings", [h for h in wt["heads"] if h[0]][:3] == [("1", "Einleitung"), ("2", "Hauptteil"), ("2.1", "Unterkapitel")], json.dumps(wt["heads"], ensure_ascii=False))
                check("Facharbeit: Times New Roman, 1,5-zeilig", wt["font"] == "Times New Roman" and abs(wt["spacing"] - 18) < 0.1, f"{wt['font']} {wt['spacing']}")
                check("Facharbeit: no page number on the title page", wt["first"], "")
                check("Facharbeit: Word's TOC update gives the same page numbers", wt.get("toc") == wt.get("toc_word"),
                      json.dumps({"ours": wt.get("toc"), "word": wt.get("toc_word")}, ensure_ascii=False))
            if name == "versuchsprotokoll":
                check("Versuchsprotokoll: tables and the example equation", wt["tables"] == 2 and wt["omaths"] == 1, json.dumps(wt, ensure_ascii=False)[:300])
    except Exception:
        check("no exception", False, traceback.format_exc()[-600:])
    finally:
        t.close()
        shutil.rmtree(TMP, ignore_errors=True)
    passed = sum(ok for _, ok, _ in RESULTS)
    print(f"\n{passed}/{len(RESULTS)} checks passed")


if __name__ == "__main__":
    main()
