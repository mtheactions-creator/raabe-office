"""Pictures, shapes, text boxes and hyperlinks through the clipboard between Word and RaabeWrite
(Step 4c F). A Word-made document with an inline picture, a floating picture, a rectangle with
text, a text box and two hyperlinks goes through the clipboard both ways:

  word-to-raabe  Word copies everything, RaabeWrite pastes into an empty document.
  raabe-to-word  RaabeWrite opens the document, copies everything, Word pastes it.
  picture        Word copies only the floating picture (and back: RaabeWrite copies it).

Each result is saved and summarised (text, inline pictures, floating pictures, shapes, text
boxes, link targets) and compared with the original. Uses the clipboard and an own (invisible)
Word instance; RaabeWrite runs in its own window. Result: results/<date>-clipboard-objects.json.

    .venv/Scripts/python.exe clipboard_objects.py
"""
from __future__ import annotations

import datetime
import json
import re
import tempfile
import time
import zipfile
from pathlib import Path

from feellab.targets import raabe_write
from probe_word_context import make_png

HERE = Path(__file__).parent
TMP = Path(tempfile.gettempdir()) / "raabe-clipboard-objects"


def build(app, path: Path) -> None:
    png = TMP / "picture.png"
    make_png(png)
    d = app.Documents.Add()
    d.Content.Text = "\r".join([
        "Ein Absatz mit einem Bild  im Text.",
        "Hier steht ein Link zu einer Webseite und ein zweiter zur Schule.",
        "Ein Absatz mit einem schwebenden Bild, einer Form und einem Textfeld.",
        "Ende.",
    ])
    p1 = d.Paragraphs(1).Range
    at = p1.Start + p1.Text.index("Bild ") + 5
    d.InlineShapes.AddPicture(str(png), False, True, d.Range(at, at))
    # The later link first: Word's positions count a link's hidden field code, its text does not.
    p2 = d.Paragraphs(2).Range
    t = p2.Text
    a = p2.Start + t.index("Link")
    b = p2.Start + t.index("zweiter")
    d.Hyperlinks.Add(d.Range(b, b + 7), "https://www.example.com/schule")
    d.Hyperlinks.Add(d.Range(a, a + 4), "https://www.example.org/")
    anchor = d.Paragraphs(3).Range
    pic = d.Shapes.AddPicture(str(png), False, True, 300, 120, 90, 60, anchor)
    pic.WrapFormat.Type = 0
    shp = d.Shapes.AddShape(1, 80, 200, 120, 50, anchor)
    shp.TextFrame.TextRange.Text = "Form"
    box = d.Shapes.AddTextbox(1, 240, 220, 140, 40, anchor)
    box.TextFrame.TextRange.Text = "Textfeld"
    d.SaveAs2(str(path), 16)
    d.Close(0)


def summary(path: Path) -> dict:
    with zipfile.ZipFile(path) as z:
        doc = z.read("word/document.xml").decode("utf-8")
        rels = z.read("word/_rels/document.xml.rels").decode("utf-8") if "word/_rels/document.xml.rels" in z.namelist() else ""
    body = re.sub(r"<mc:Fallback>.*?</mc:Fallback>", "", doc, flags=re.S)
    targets = {}
    for rel in re.findall(r"<Relationship\b[^>]*>", rels):
        i, t = re.search(r'Id="([^"]+)"', rel), re.search(r'Target="([^"]+)"', rel)
        if i and t:
            targets[i.group(1)] = t.group(1)
    links = [targets.get(i, "?") for i in re.findall(r'<w:hyperlink[^>]*r:id="([^"]+)"', body)]
    links += re.findall(r'HYPERLINK\s+&quot;([^&]+)&quot;', body)
    inline = len(re.findall(r"<wp:inline\b", body))
    anchors = re.findall(r"<wp:anchor\b.*?</wp:anchor>", body, re.S)
    text = re.sub(r"<w:txbxContent>.*?</w:txbxContent>", "", body, flags=re.S)
    words = " ".join("".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", p)) for p in re.findall(r"<w:p\b.*?</w:p>", text, re.S)).split()
    return {
        "text": " ".join(words),
        "inline pictures": inline,
        "floating pictures": sum(1 for a in anchors if "<pic:pic" in a),
        "shapes": sum(1 for a in anchors if "<wps:wsp" in a and "Textfeld" not in a),
        "text boxes": sum(1 for a in anchors if "Textfeld" in a),
        "links": sorted(set(links)),
    }


def compare(a: dict, b: dict) -> list[str]:
    return [f"{k}: Word {a[k]!r}, ours {b[k]!r}" for k in a if a[k] != b[k]]


def main() -> None:
    import sys
    import win32com.client
    sys.stdout.reconfigure(encoding="utf-8")
    TMP.mkdir(parents=True, exist_ok=True)
    app = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx("Word.Application"))
    app.Visible = False
    results: dict = {}
    try:
        src = TMP / "source.docx"
        print("building", flush=True)
        build(app, src)
        print("built", flush=True)
        original = summary(src)
        results["original"] = original

        def save_from_raabe(rw, name: str) -> Path:
            out = TMP / f"{name}.docx"
            out.unlink(missing_ok=True)
            rw.js(f"__app.saveTo({str(out)!r})")
            deadline = time.time() + 10
            while time.time() < deadline and not out.exists():
                time.sleep(0.1)
            time.sleep(0.3)
            return out

        rw = raabe_write("short", port=9375)
        try:
            # Word → RaabeWrite: everything.
            d = app.Documents.Open(str(src), ReadOnly=True, AddToRecentFiles=False)
            d.Content.Copy()
            d.Close(0)
            rw.js("__app.newDoc()")
            time.sleep(0.8)
            rw.js("__app.closeOthers && __app.closeOthers()")
            rw.js("__probe.focus()")
            rw.js("__app.run('edit.paste')")
            time.sleep(2.5)
            results["word-to-raabe"] = summary(save_from_raabe(rw, "word-to-raabe"))
            print("word-to-raabe", results["word-to-raabe"], flush=True)

            # Word → RaabeWrite: the floating picture alone.
            d = app.Documents.Open(str(src), ReadOnly=True, AddToRecentFiles=False)
            pics = [s for s in d.Shapes if s.Type == 13]
            pics[0].Select()
            app.Selection.Copy()
            # Word renders some clipboard formats only when asked: its document stays open (as a user's would).
            rw.js("__app.newDoc()")
            time.sleep(0.8)
            rw.js("__app.closeOthers && __app.closeOthers()")
            rw.js("__probe.focus()")
            rw.js("__app.run('edit.paste')")
            time.sleep(2.0)
            d.Close(0)
            results["picture word-to-raabe"] = summary(save_from_raabe(rw, "picture-word-to-raabe"))

            # RaabeWrite → Word: everything, then the floating picture alone.
            rw.js(f"__app.openPath({str(src)!r})")
            time.sleep(2.0)
            rw.js("__app.closeOthers && __app.closeOthers()")
            rw.js("__probe.focus()")
            rw.js("__app.editor().selectAll()")
            rw.js("__app.run('edit.copy')")
            time.sleep(1.0)
            d2 = app.Documents.Add()
            d2.Content.Paste()
            out2 = TMP / "raabe-to-word.docx"
            d2.SaveAs2(str(out2), 16)
            d2.Close(0)
            results["raabe-to-word"] = summary(out2)
            rw.js("(() => { const e = __app.editor(); const ps = e.paragraphs(); for (let p = 0; p < ps.length; p++) { const fs = ps[p].floats ?? []; "
                  "for (let i = 0; i < fs.length; i++) if (!fs[i].shape) { e.selectFloat(p, i); return p; } } return -1; })()")
            rw.js("__app.run('edit.copy')")
            time.sleep(1.0)
            d3 = app.Documents.Add()
            d3.Content.Paste()
            out3 = TMP / "picture-raabe-to-word.docx"
            d3.SaveAs2(str(out3), 16)
            d3.Close(0)
            results["picture raabe-to-word"] = summary(out3)
        finally:
            rw.close()
    finally:
        app.Quit(0)
    print("original", json.dumps(original, ensure_ascii=False))
    for k in ("word-to-raabe", "raabe-to-word"):
        results[f"{k} differences"] = compare(original, results[k])
        print(f"== {k}: {len(results[f'{k} differences'])} differences")
        for x in results[f"{k} differences"]:
            print("   ", x)
    for k in ("picture word-to-raabe", "picture raabe-to-word"):
        print(f"== {k}:", json.dumps(results[k], ensure_ascii=False))
    out = HERE / "results" / f"{datetime.date.today().isoformat()}-clipboard-objects.json"
    out.write_text(json.dumps(results, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
