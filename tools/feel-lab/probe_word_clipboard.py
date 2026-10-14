"""Probe: what Word puts on the clipboard (HTML Format) for lists and tables, and how Word
pastes HTML given to it. Own invisible Word instance; uses the system clipboard, no other
input. Writes the HTML to captures/word-clipboard/ (local only).

    .venv/Scripts/python.exe probe_word_clipboard.py copy          # Word's HTML for samples
    .venv/Scripts/python.exe probe_word_clipboard.py paste file.html  # paste HTML into Word, report
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

import win32clipboard

HERE = Path(__file__).parent
CAP = HERE / "captures" / "word-clipboard"
CAP.mkdir(parents=True, exist_ok=True)


def cf_html() -> int:
    return win32clipboard.RegisterClipboardFormat("HTML Format")


def read_html() -> str:
    for _ in range(20):
        try:
            win32clipboard.OpenClipboard()
            break
        except Exception:
            time.sleep(0.1)
    try:
        raw = win32clipboard.GetClipboardData(cf_html())
    finally:
        win32clipboard.CloseClipboard()
    return raw.decode("utf-8") if isinstance(raw, bytes) else raw


def write_html(fragment: str) -> None:
    """Puts an HTML fragment on the clipboard in the CF_HTML envelope, as Chromium does."""
    head = "Version:0.9\r\nStartHTML:{:010d}\r\nEndHTML:{:010d}\r\nStartFragment:{:010d}\r\nEndFragment:{:010d}\r\n"
    pre = "<html><body><!--StartFragment-->"
    post = "<!--EndFragment--></body></html>"
    if "<!--StartFragment-->" in fragment:
        pre, post = "", ""
    body = pre + fragment + post
    h0 = len(head.format(0, 0, 0, 0).encode("utf-8"))
    b = body.encode("utf-8")
    sf = h0 + body.index("<!--StartFragment-->") + len("<!--StartFragment-->") if "<!--StartFragment-->" in body else h0
    sf = h0 + len(body[:body.index("<!--StartFragment-->") + len("<!--StartFragment-->")].encode("utf-8"))
    ef = h0 + len(body[:body.index("<!--EndFragment-->")].encode("utf-8"))
    data = head.format(h0, h0 + len(b), sf, ef).encode("utf-8") + b
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(cf_html(), data)
        win32clipboard.SetClipboardText("x", win32clipboard.CF_UNICODETEXT)
    finally:
        win32clipboard.CloseClipboard()


def word():
    import win32com.client
    app = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx("Word.Application"))
    app.Visible = False
    return app


def describe(d) -> list[str]:
    out = []
    for i in range(1, d.Paragraphs.Count + 1):
        p = d.Paragraphs(i)
        lf = p.Range.ListFormat
        pf = p.Format
        cell = ""
        if p.Range.Information(31):  # end-of-row mark
            continue
        if p.Range.Information(12):  # wdWithInTable
            try:
                c = p.Range.Cells(1)
                cell = f" cell({c.RowIndex},{c.ColumnIndex}) w={c.Width:.1f}"
            except Exception:  # noqa: BLE001
                continue
        lvl = ""
        if lf.ListType:
            lt = lf.ListTemplate
            lv = lt.ListLevels(lf.ListLevelNumber) if lt else None
            lvl = f" list={lf.ListString!r} lvl={lf.ListLevelNumber}"
            if lv:
                lvl += f" fmt={lv.NumberFormat!r} font={lv.Font.Name} pos={lv.NumberPosition:.1f} text={lv.TextPosition:.1f} tab={lv.TabPosition:.1f} trail={lv.TrailingCharacter}"
        out.append(f"{i}: {p.Range.Text.rstrip(chr(13) + chr(7))!r} style={p.Style.NameLocal} left={pf.LeftIndent:.1f} first={pf.FirstLineIndent:.1f}{lvl}{cell}")
    return out


def copy_samples() -> None:
    app = word()
    try:
        d = app.Documents.Add()
        d.Content.Text = "\r".join(["Becherglas", "Wasser", "Leitungswasser", "Salz", "Wasser einfüllen", "Salz zugeben", "Ende"])
        gal = app.ListGalleries
        d.Range(d.Paragraphs(1).Range.Start, d.Paragraphs(4).Range.End).ListFormat.ApplyListTemplate(gal(1).ListTemplates(1), False, 0, 2)
        d.Paragraphs(3).Range.ListFormat.ListIndent()
        d.Range(d.Paragraphs(5).Range.Start, d.Paragraphs(6).Range.End).ListFormat.ApplyListTemplate(gal(2).ListTemplates(1), False, 0, 2)
        r = d.Range(d.Paragraphs(7).Range.Start, d.Paragraphs(7).Range.Start)
        t = d.Tables.Add(r, 2, 3, 1, 0)
        t.Columns(1).Width = 60
        t.Columns(2).Width = 200
        t.Columns(3).Width = 100
        for i in range(1, 3):
            for k in range(1, 4):
                t.Cell(i, k).Range.Text = f"Z{i}S{k}"
        print("\n".join(describe(d)))
        d.Content.Copy()
        html = read_html()
        (CAP / "word-list-table.html").write_text(html, encoding="utf-8", newline="\n")
        print("saved", CAP / "word-list-table.html", len(html))
        # Round trip: paste Word's own HTML into a new document.
        d2 = app.Documents.Add()
        write_html(html[html.index("<html"):])
        d2.Content.Paste()
        print("--- pasted back")
        print("\n".join(describe(d2)))
        d2.Close(0)
        d.Close(0)
    finally:
        app.Quit(0)


def paste_file(path: str) -> None:
    app = word()
    try:
        d = app.Documents.Add()
        write_html(Path(path).read_text(encoding="utf-8"))
        d.Content.Paste()
        print("\n".join(describe(d)))
        d.Close(0)
    finally:
        app.Quit(0)


if __name__ == "__main__":
    if sys.argv[1] == "copy":
        copy_samples()
    else:
        paste_file(sys.argv[2])
