"""Copy and paste between Word and RaabeWrite (D-038, package B): a Word-made document with
bulleted, numbered and roman lists on two levels, an indented paragraph and a table with
uneven column widths goes through the clipboard both ways:

  word-to-raabe  Word copies everything, RaabeWrite pastes into an empty document.
  raabe-to-word  RaabeWrite opens the document, copies everything, Word pastes it.

Each result is saved and read by Word (paragraph text, style, list label and level, indents,
table cell widths) and compared with the original. Uses the clipboard and an own (invisible)
Word instance; RaabeWrite runs in its own window.

    .venv/Scripts/python.exe clipboard_roundtrip.py
"""
from __future__ import annotations

import datetime
import json
import tempfile
import time
from pathlib import Path

from feellab.targets import raabe_write

HERE = Path(__file__).parent
TMP = Path(tempfile.gettempdir()) / "raabe-clipboard"


def build(app, path: Path) -> None:
    d = app.Documents.Add()
    d.Content.Text = "\r".join([
        "Material", "Becherglas", "Wasser", "Leitungswasser", "Salz", "Ablauf", "Wasser einfüllen", "Salz zugeben", "Umrühren",
        "Gliederung", "Einleitung", "Hauptteil", "Eingerückter Absatz mit Text.", "Tabelle", "Ende",
    ])
    gal = app.ListGalleries
    d.Range(d.Paragraphs(2).Range.Start, d.Paragraphs(5).Range.End).ListFormat.ApplyListTemplate(gal(1).ListTemplates(1), False, 0, 2)
    d.Paragraphs(4).Range.ListFormat.ListIndent()
    d.Range(d.Paragraphs(7).Range.Start, d.Paragraphs(9).Range.End).ListFormat.ApplyListTemplate(gal(2).ListTemplates(1), False, 0, 2)
    d.Paragraphs(8).Range.ListFormat.ListIndent()
    # Roman numerals (gallery entry "I. II. III.").
    rom = None
    for i in range(1, 8):
        t = gal(2).ListTemplates(i)
        if t.ListLevels(1).NumberStyle == 1:  # wdListNumberStyleUppercaseRoman
            rom = t
            break
    d.Range(d.Paragraphs(11).Range.Start, d.Paragraphs(12).Range.End).ListFormat.ApplyListTemplate(rom or gal(2).ListTemplates(1), False, 0, 2)
    d.Paragraphs(13).LeftIndent = 42.5
    d.Paragraphs(13).FirstLineIndent = 14.2
    r = d.Paragraphs(15).Range
    r.InsertParagraphBefore()
    r = d.Paragraphs(15).Range
    t = d.Tables.Add(r, 2, 3, 1, 0)
    for k, w in enumerate((60, 220, 120), start=1):
        t.Columns(k).Width = w
    for i in range(1, 3):
        for k in range(1, 4):
            t.Cell(i, k).Range.Text = f"Z{i}S{k}"
    d.SaveAs2(str(path), 16)
    d.Close(0)


def signature(app, path: Path) -> list[dict]:
    d = app.Documents.Open(str(path), ReadOnly=True, AddToRecentFiles=False)
    try:
        rows = []
        for p in d.Paragraphs:
            r = p.Range
            if r.Information(31):
                continue
            text = r.Text.rstrip("\r\x07")
            if not text.strip():
                continue
            lf = r.ListFormat
            row = {"text": text, "style": p.Style.NameLocal, "list": lf.ListString if lf.ListType else "",
                   "level": lf.ListLevelNumber if lf.ListType else 0,
                   "indent": round(p.Format.LeftIndent, 1), "first": round(p.Format.FirstLineIndent, 1)}
            if r.Information(12):
                row["width"] = round(r.Cells(1).Width, 1)
            if lf.ListType:
                lv = lf.ListTemplate.ListLevels(lf.ListLevelNumber)
                row["font"] = lv.Font.Name
            rows.append(row)
        return rows
    finally:
        d.Close(0)


def compare(a: list[dict], b: list[dict]) -> list[dict]:
    diffs = []
    by_text = {x["text"]: x for x in b}
    for x in a:
        y = by_text.get(x["text"])
        if y is None:
            diffs.append({"text": x["text"], "missing": True})
            continue
        d = {k: (x.get(k), y.get(k)) for k in x if k != "style" and x.get(k) != y.get(k)
             and not (isinstance(x.get(k), float) and isinstance(y.get(k), float) and abs(x[k] - y[k]) <= 0.6)}
        if d:
            diffs.append({"text": x["text"], **{k: {"word": v[0], "ours": v[1]} for k, v in d.items()}})
    return diffs


def main() -> None:
    import win32com.client
    TMP.mkdir(parents=True, exist_ok=True)
    app = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx("Word.Application"))
    app.Visible = False
    results = {}
    try:
        src = TMP / "source.docx"
        build(app, src)
        original = signature(app, src)

        # Word → RaabeWrite
        d = app.Documents.Open(str(src), ReadOnly=True, AddToRecentFiles=False)
        d.Content.Copy()
        d.Close(0)
        rw = raabe_write("short", port=9373)
        try:
            rw.js("__app.newDoc()")
            time.sleep(0.8)
            rw.js("__app.closeOthers && __app.closeOthers()")
            rw.js("__probe.focus()")
            rw.js("__app.run('edit.paste')")
            time.sleep(1.5)
            out = TMP / "word-to-raabe.docx"
            out.unlink(missing_ok=True)
            rw.js(f"__app.saveTo({str(out)!r})")
            deadline = time.time() + 10
            while time.time() < deadline and not out.exists():
                time.sleep(0.1)
            time.sleep(0.3)
            results["word-to-raabe"] = compare(original, signature(app, out))

            # RaabeWrite → Word
            rw.js(f"__app.openPath({str(src)!r})")
            time.sleep(2.0)
            rw.js("__app.closeOthers && __app.closeOthers()")
            rw.js("__probe.focus()")
            rw.js("__app.editor().selectAll()")
            rw.js("__app.run('edit.copy')")
            time.sleep(1.0)
        finally:
            rw.close()
        d2 = app.Documents.Add()
        d2.Content.Paste()
        out2 = TMP / "raabe-to-word.docx"
        d2.SaveAs2(str(out2), 16)
        d2.Close(0)
        results["raabe-to-word"] = compare(original, signature(app, out2))
        results["original"] = original
    finally:
        app.Quit(0)
    for k in ("word-to-raabe", "raabe-to-word"):
        print(f"== {k}: {len(results[k])} differences")
        for x in results[k]:
            print("  ", json.dumps(x, ensure_ascii=False))
    out = HERE / "results" / f"{datetime.date.today().isoformat()}-clipboard.json"
    out.write_text(json.dumps(results, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
