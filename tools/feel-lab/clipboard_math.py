"""Equations through the clipboard between Word and RaabeWrite (Step 4c C): a Word-made
document with inline and display equations (fraction and power, the quadratic formula, a
function with normal text, a sum) goes through the clipboard both ways:

  word-to-raabe  Word copies everything, RaabeWrite pastes into an empty document.
  raabe-to-word  RaabeWrite opens the document, copies everything, Word pastes it.

Each result is saved, its equations normalised like catalogue_math.py does and compared with
the original. Uses the clipboard and an own (invisible) Word instance; RaabeWrite runs in its
own window. Result: results/<date>-clipboard-math.json.

    .venv/Scripts/python.exe clipboard_math.py
"""
from __future__ import annotations

import datetime
import json
import tempfile
import time
from pathlib import Path

from catalogue_math import equations, merge_runs
from feellab.targets import raabe_write

HERE = Path(__file__).parent
TMP = Path(tempfile.gettempdir()) / "raabe-clipboard-math"


def build(app, path: Path) -> None:
    d = app.Documents.Add()
    d.Content.Text = "Vor der Formel  danach.\rx=(-b±√(b^2-4ac))/2a\r\rIm Text ∑_(i=1)^n i steht eine Summe.\rEnde"
    # Inline: in the text of the first paragraph.
    r = d.Range(15, 15)
    r.Text = "a/b+x^2"
    d.OMaths.Add(r).OMaths(1).BuildUp()
    # Display: the second and third paragraphs.
    p2 = d.Paragraphs(2).Range
    d.OMaths.Add(d.Range(p2.Start, p2.End - 1)).OMaths(1).BuildUp()
    p3 = d.Paragraphs(3).Range
    r3 = d.Range(p3.Start, p3.Start)
    r3.Text = 'sin x="wenn" a'
    d.OMaths.Add(r3).OMaths(1).BuildUp()
    # Inline sum in text.
    p4 = d.Paragraphs(4).Range
    t = p4.Text
    a = p4.Start + t.index("∑")
    b = p4.Start + t.index(" steht")
    d.OMaths.Add(d.Range(a, b)).OMaths(1).BuildUp()
    d.SaveAs2(str(path), 16)
    d.Close(0)


def compare(a: list[str], b: list[str]) -> list[dict]:
    a2 = [merge_runs(x) for x in a]
    b2 = [merge_runs(x) for x in b]
    out = []
    for i in range(max(len(a2), len(b2))):
        x = a2[i] if i < len(a2) else None
        y = b2[i] if i < len(b2) else None
        if x != y:
            out.append({"para": i + 1, "word": x, "ours": y})
    return out


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
        build(app, src)
        original = equations(src)
        results["original"] = original

        # Word → RaabeWrite
        d = app.Documents.Open(str(src), ReadOnly=True, AddToRecentFiles=False)
        d.Content.Copy()
        d.Close(0)
        rw = raabe_write("short", port=9374)
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
            results["word-to-raabe"] = compare(original, equations(out))
            results["word-to-raabe-result"] = equations(out)

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
        results["raabe-to-word"] = compare(original, equations(out2))
        results["raabe-to-word-result"] = equations(out2)
    finally:
        app.Quit(0)
    for k in ("word-to-raabe", "raabe-to-word"):
        print(f"== {k}: {len(results[k])} differences")
        for x in results[k]:
            print("  ", json.dumps(x, ensure_ascii=False))
    out = HERE / "results" / f"{datetime.date.today().isoformat()}-clipboard-math.json"
    out.write_text(json.dumps(results, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
