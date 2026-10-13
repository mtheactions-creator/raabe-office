"""Functional smoke test of the real RaabeWrite app (Step 1 shell, Step 2 formatting).

Drives the app through the DevTools protocol (input goes into the page only, the real
keyboard and mouse are never used) and checks files with Microsoft Word through COM in a
separate, invisible Word instance. The app runs off screen and never takes the focus, so
the test can run while someone works at the PC.

    .venv/Scripts/python.exe smoke_write.py
"""

from __future__ import annotations

import json
import os
import shutil
import tempfile
import time
import traceback
from pathlib import Path

from feellab import targets, winapi

RESULTS: list[tuple[str, bool, str]] = []
TMP = Path(tempfile.mkdtemp(prefix="raabe-smoke-"))
RECOVERY = Path(tempfile.gettempdir()) / "raabe-write-feel-recovery"  # test mode's own folder
OFF = -4000  # x of the off-screen window


def check(name: str, ok: bool, detail: str = "") -> None:
    RESULTS.append((name, bool(ok), detail))
    print(("PASS " if ok else "FAIL ") + name + (f"  [{detail}]" if detail and not ok else ""), flush=True)


def wait(t: targets.WebViewTarget, expr: str, timeout: float = 5.0):
    end = time.perf_counter() + timeout
    while time.perf_counter() < end:
        v = t.js(expr)
        if v:
            return v
        time.sleep(0.05)
    return t.js(expr)


def word_doc(path: Path, build) -> None:
    """Creates a .docx with Word itself (isolated instance)."""
    import win32com.client
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    try:
        d = app.Documents.Add()
        build(d)
        d.SaveAs2(str(path), 16)
        d.Close(0)
    finally:
        app.Quit(0)


def word_read(path: Path) -> dict:
    import win32com.client
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    try:
        d = app.Documents.Open(str(path), ReadOnly=True, AddToRecentFiles=False)
        paras = [(p.Range.Text.rstrip("\r"), p.Style.NameLocal) for p in d.Paragraphs]
        info = {"paras": paras, "compat": d.CompatibilityMode, "pages": d.ComputeStatistics(2),
                "tables": d.Tables.Count, "cell": d.Tables(1).Cell(1, 1).Range.Text[:-2] if d.Tables.Count else None,
                "bold": [bool(d.Range(i, i + 1).Font.Bold) for i in range(min(40, d.Content.End - 1))],
                "color": [d.Range(i, i + 1).Font.Color for i in range(min(40, d.Content.End - 1))],
                "size": [d.Range(i, i + 1).Font.Size for i in range(min(40, d.Content.End - 1))]}
        d.Close(0)
        return info
    finally:
        app.Quit(0)


def main() -> None:
    before_recovery = set(RECOVERY.glob("*")) if RECOVERY.exists() else set()
    t = targets.raabe_write("short", background=True)
    try:
        winapi.place_quietly(t.hwnd, OFF, 0, 2000, 2050)
        api = "window.__app"
        docs = t.js(f"{api}.docs()")
        check("starts with one document", len(docs) == 1 and docs[0]["name"] == "Dokument1", json.dumps(docs))
        check("Start tab is active", t.js(f"{api}.activeRibbonTab()") == "start")
        text0 = t.text()

        # Typing and undo.
        t.set_selection(len(text0), len(text0))
        t.js("__probe.focus()")
        t.send_text(" Test")
        check("typing inserts text", t.text() == text0 + " Test", t.text()[-20:])
        check("document becomes dirty", t.js(f"{api}.docs()[0].dirty"))
        t.send_key("z", ctrl=True)
        check("Strg+Z undoes the typing", t.text() == text0, t.text()[-20:])

        # Styles from the gallery.
        t.set_selection(40, 40)
        t.js(f"{api}.run('style.h1')")
        time.sleep(0.2)
        checked = t.js("document.querySelector('.rgallery .item.checked')?.textContent")
        check("style gallery applies Überschrift 1", checked == "Überschrift 1", str(checked))
        t.js(f"{api}.run('style.normal')")

        # Save and read back with Word.
        out = TMP / "smoke.docx"
        t.set_selection(0, 0)
        t.send_text("Neu: ")
        ok = t.js(f"{api}.saveTo({json.dumps(str(out))})")
        docs = t.js(f"{api}.docs()")
        check("saves as .docx", ok and out.exists() and not docs[0]["dirty"] and docs[0]["name"] == "smoke", json.dumps(docs))
        info = word_read(out)
        expected = t.text().split("\n")
        check("Word reads the saved text identically", [p[0] for p in info["paras"]] == expected,
              f"{[p[0][:20] for p in info['paras'][:3]]}")
        check("Word sees the heading style", info["paras"][0][1] == "Überschrift 1", info["paras"][0][1])
        check("Word opens it in current mode (not compatibility mode)", info["compat"] == 15, str(info["compat"]))

        # Formatting (Step 2): Strg+Umschalt+F in a word, colour and size; Word reads it back.
        t.js(f"{api}.newDoc()")
        t.js("__probe.focus()")
        sample = "Ein fettes Wort und roter Text."
        t.send_text(sample)
        t.set_selection(6, 6)  # inside "fettes"
        t.send_key("F", ctrl=True, shift=True)
        bold = [t.js(f"__probe.formatAt({i}).bold") for i in range(len(sample))]
        check("Strg+Umschalt+F formats the whole word at the caret", bold == [4 <= i < 10 for i in range(len(sample))], str(bold))
        t.set_selection(6, 6)
        t.send_key("C", ctrl=True, alt=True)  # Format kopieren
        w = sample.index("Wort")
        t.set_selection(w, w + 4)
        t.send_key("V", ctrl=True, alt=True)  # Format einfügen
        check("Strg+Alt+C / Strg+Alt+V copy formatting", t.js(f"__probe.formatAt({w + 1}).bold"))
        bold = [t.js(f"__probe.formatAt({i}).bold") for i in range(len(sample))]
        red = sample.index("roter Text")
        t.set_selection(red, len(sample))
        t.js(f"{api}.run('font.color')")
        t.send_key("<", ctrl=True, shift=True, vk=226)  # Strg+Umschalt+< (German keyboard)
        fmt = t.js(f"__probe.formatAt({red + 2})")
        check("Schriftfarbe applies red", fmt["color"] == "FF0000", json.dumps(fmt))
        check("Strg+Umschalt+< grows 11 pt to 12 pt", fmt["size"] == 12, json.dumps(fmt))
        t.set_selection(len(sample), len(sample))
        t.send_key("Enter")
        t.send_text("Weiter")
        # Word: the new paragraph takes the paragraph mark's formatting, and the selection did
        # not include the mark.
        check("Enter takes the formatting of the paragraph mark", t.js("__probe.typingFormat().color") == "auto")
        fmt_file = TMP / "formatierung.docx"
        t.js(f"{api}.saveTo({json.dumps(str(fmt_file))})")
        winfo = word_read(fmt_file)
        check("Word reads the bold word", winfo["bold"][:len(sample)] == bold, str(winfo["bold"][:len(sample)]))
        check("Word reads red 12 pt text", winfo["color"][red] == 0x0000FF and winfo["size"][red] == 12,
              f"{winfo['color'][red]} {winfo['size'][red]}")
        t.js(f"{api}.run('file.close')")
        time.sleep(0.3)

        # A document Word made with a table and bold text: the table's cells are edited like
        # any text, and saving goes to the original (D-024).
        rich = TMP / "mit-tabelle.docx"
        def build(d):
            d.Content.Text = "Arbeitsblatt"
            d.Paragraphs(1).Range.Font.Bold = True
            r = d.Content
            r.InsertParagraphAfter()
            d.Tables.Add(d.Paragraphs(2).Range, 2, 2)
            d.Tables(1).Cell(1, 1).Range.Text = "Zelle"
        word_doc(rich, build)
        t.js(f"{api}.openPath({json.dumps(str(rich))})")
        docs = wait(t, f"{api}.docs().some(d => d.name === 'mit-tabelle') && {api}.docs()")
        active = [d for d in docs if d["active"]][0] if docs else {}
        check("opens a Word document with a table in a new tab", active.get("name") == "mit-tabelle", json.dumps(docs))
        check("shows the table as cells (nothing left unsupported)", not active.get("unsupported"), json.dumps(active.get("unsupported")))
        text = t.js("__probe.text()") or ""
        check("the cell text is part of the document", "\nZelle\n" in text, repr(text))
        check("reads the bold heading", t.js("__probe.formatAt(0).bold"))
        mtime = rich.stat().st_mtime
        cell_end = text.index("Zelle") + len("Zelle")
        t.set_selection(cell_end, cell_end)
        t.js("__probe.focus()")
        t.send_text("n")
        t.set_selection(len("Arbeitsblatt"), len("Arbeitsblatt"))
        t.send_text(" 1")
        ok = t.js(f"{api}.save()")
        check("saves over the original", ok and rich.stat().st_mtime != mtime)
        rinfo = word_read(rich)
        check("Word: edited heading, still bold", rinfo["paras"][0][0] == "Arbeitsblatt 1" and all(rinfo["bold"][:14]),
              str(rinfo["paras"][:1]))
        check("Word: the table is still there, with the typed cell text", rinfo["tables"] == 1 and rinfo["cell"] == "Zellen", f"{rinfo['tables']} {rinfo['cell']!r}")

        # Keytips (Alt), then Ansicht (F) and 100 % (0).
        t.js(f"{api}.run('view.pageWidth')")
        t.js("__probe.focus()")
        t.send_key("Alt", alt=True)
        tips = wait(t, "(() => { const r = Array.from(document.querySelectorAll('.keytip')).map(k => k.textContent); return r.length ? r : null; })()")
        check("Alt shows keytips on the tabs", tips and {"D", "R", "I", "S", "F"} <= set(tips), str(tips))
        t.send_text("F")
        check("keytip F opens Ansicht", wait(t, f"{api}.activeRibbonTab() === 'view'"))
        t.send_text("0")
        time.sleep(0.2)
        check("keytip 0 sets 100 % zoom", (t.js("document.querySelector('.statusbar .pct')?.textContent") or "").startswith("100"))
        check("keytips are gone afterwards", t.js("document.querySelectorAll('.keytip').length") == 0)

        # Search box: "Wörter" finds "Wörter zählen" and runs it.
        t.js("document.querySelector('.search input').focus()")
        t.send_text("Wörter")
        results = wait(t, "(() => { const r = Array.from(document.querySelectorAll('.search-results button')).map(b => b.textContent); return r.length ? r : null; })()")
        check("search finds commands", results and any("Wörter zählen" in r for r in results), str(results))
        t.send_key("Enter")
        title = wait(t, "document.querySelector('.dialog .dlg-title .t')?.textContent")
        check("search runs the command (Wörter zählen dialog)", title == "Wörter zählen", str(title))
        t.send_key("Escape")

        # Backstage.
        t.js("document.querySelector('.tabrow .tab').click()")
        nav = wait(t, "!document.querySelector('.backstage').hidden && Array.from(document.querySelectorAll('.backstage nav .item')).map(i => i.textContent)")
        check("Datei opens the backstage", nav and "Speichern unter" in nav and "Drucken" in nav, str(nav))
        t.send_key("Escape")
        check("Escape closes the backstage", wait(t, "document.querySelector('.backstage').hidden"))

        # PDF export.
        t.js("document.querySelectorAll('.doctab')[0].click()")
        pdf = TMP / "smoke.pdf"
        t.js(f"{api}.exportPdfTo({json.dumps(str(pdf))})")
        end = time.perf_counter() + 15
        while time.perf_counter() < end and not (pdf.exists() and pdf.stat().st_size > 1000):
            time.sleep(0.2)
        data = pdf.read_bytes() if pdf.exists() else b""
        pages = data.count(b"/Type /Page") - data.count(b"/Type /Pages")
        check("exports a PDF", data.startswith(b"%PDF"), f"{len(data)} bytes")
        check("PDF has one page per document page", pages == info["pages"], f"pdf {pages}, word {info['pages']}")

        # Ribbon shrinking at a narrow window.
        t.js("document.querySelectorAll('.tabrow .tab')[1].click()")
        winapi.place_quietly(t.hwnd, OFF, 0, 1100, 1400)
        time.sleep(0.8)
        collapsed = t.js("Array.from(document.querySelectorAll('.ribbon-card .rgroup.collapsed')).map(g => g.dataset.group)")
        order = ["styles", "voice", "editing", "font", "paragraph", "clipboard"]
        check("narrow window collapses groups in Word's order",
              collapsed and set(collapsed) == set(order[:len(collapsed)]), str(collapsed))
        overflow = t.js("(() => { const c = document.querySelector('.ribbon-card'); return c.scrollWidth - c.clientWidth; })()")
        check("ribbon fits the narrow window", overflow <= 1, str(overflow))
        winapi.place_quietly(t.hwnd, OFF, 0, 2000, 2050)

        # AutoRecover: change, back up, crash, restart.
        t.send_text("Absturztest ")
        t.js(f"{api}.autoRecoverNow()")
        time.sleep(0.5)
        new_files = (set(RECOVERY.glob("*")) - before_recovery) if RECOVERY.exists() else set()
        check("AutoRecover writes a backup", any(f.suffix == ".docx" for f in new_files), str(new_files))
        t.proc.kill()
        t.proc.wait(5)
        t.ws.close()
        t = targets.raabe_write("short", port=9371, keep_recovery=True, background=True)
        body = wait(t, "document.querySelector('.dialog .dlg-body')?.textContent", 8)
        check("after a crash, recovery is offered", body and "nicht ordnungsgemäß beendet" in body, str(body))
        t.js("document.querySelector('.dialog .dlg-buttons .primary').click()")
        docs = wait(t, f"{api}.docs().some(d => d.name.includes('Wiederhergestellt')) && {api}.docs()")
        check("recovered document opens", bool(docs), json.dumps(docs))

        # Closing a dirty document asks first.
        t.js("document.querySelector('.doctab.active .x').click()")
        q = wait(t, "document.querySelector('.dialog .dlg-body')?.textContent")
        check("closing a changed document asks to save", q and q.startswith("Möchten Sie die Änderungen an"), str(q))
        t.js("Array.from(document.querySelectorAll('.dialog .dlg-buttons button')).find(b => b.textContent === 'Nicht speichern').click()")
        time.sleep(0.3)
        check("Nicht speichern closes the tab", not any("Wiederhergestellt" in d["name"] for d in t.js(f"{api}.docs()")))
    except Exception:
        traceback.print_exc()
        check("no exception", False)
    finally:
        try:
            t.close()
        except Exception:
            pass
        if RECOVERY.exists():
            for f in set(RECOVERY.glob("*")) - before_recovery:
                f.unlink(missing_ok=True)
        shutil.rmtree(TMP, ignore_errors=True)
    passed = sum(ok for _, ok, _ in RESULTS)
    print(f"\n{passed}/{len(RESULTS)} checks passed")


if __name__ == "__main__":
    main()
