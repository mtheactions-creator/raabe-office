"""Font sizes inside equations, Word against RaabeWrite (Step 4d): Word builds one equation per
paragraph (exponents, exponents of exponents, indices, a sum with limits, a root with an index,
fractions inline and displayed) at 11 pt, saves it and exports a PDF; RaabeWrite (in the
background) opens the same file and exports its PDF. The glyph sizes of each line are read from
both PDFs (feellab/pdftext.py). Runs invisibly. Result: results/<date>-math-sizes.json.

    .venv/Scripts/python.exe probe_math_sizes.py
"""
from __future__ import annotations

import datetime
import json
import sys
import tempfile
import time
from pathlib import Path

from feellab import pdftext
from feellab.targets import raabe_write

HERE = Path(__file__).parent
OUT = HERE / "results" / f"{datetime.date.today()}-math-sizes.json"
TMP = Path(tempfile.gettempdir()) / "raabe-math-sizes"
CASES = [
    ("power", "x^2", True),
    ("power of power", "x^(2^3)", True),
    ("power of power of power", "x^(2^(3^4))", True),
    ("index", "x_i", True),
    ("index of index", "x_(i_j)", True),
    ("sum with limits", "∑_(i=1)^n i", True),
    ("root with index", "√(3&x)", True),
    ("fraction displayed", "a/b", True),
    ("fraction inline", "a/b", False),
    ("power inline", "x^2", False),
    ("power in fraction", "x^2/y^3", True),
]


def lines(pdf: Path) -> list[list[float]]:
    """Per text line (grouped by baseline), the sorted distinct glyph sizes."""
    out = []
    for page in pdftext.read(pdf):
        runs = sorted((r for r in page.runs if r.chars and r.size > 0), key=lambda r: r.y)
        groups: list[list] = []
        for r in runs:
            if groups and abs(groups[-1][-1].y - r.y) < 14:
                groups[-1].append(r)
            else:
                groups.append([r])
        for g in groups:
            out.append(sorted({round(r.size, 2) for r in g}, reverse=True))
    return out


def main() -> None:
    import win32com.client
    sys.stdout.reconfigure(encoding="utf-8")
    TMP.mkdir(parents=True, exist_ok=True)
    src, word_pdf, our_pdf = TMP / "sizes.docx", TMP / "sizes-word.pdf", TMP / "sizes-ours.pdf"
    app = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx("Word.Application"))
    app.Visible = False
    try:
        d = app.Documents.Add()
        d.Content.Font.Name = "Calibri"
        d.Content.Font.Size = 11
        # Each case on its own paragraph with blank paragraphs between (lines stay apart in the PDF).
        d.Content.Text = "\r\r".join(f"{name}:" for name, _l, _d in CASES)
        for k, (name, lin, display) in reversed(list(enumerate(CASES))):
            p = d.Paragraphs(2 * k + 1).Range
            at = p.End - 1
            r = d.Range(at, at)
            if display:
                r.InsertParagraphAfter()
                r = d.Range(at + 1, at + 1)
            r.Text = lin
            m = d.OMaths.Add(r)
            m.OMaths(1).BuildUp()
            if not display:
                m.OMaths(1).Type = 1  # inline
        d.SaveAs2(str(src), 16)
        d.ExportAsFixedFormat(str(word_pdf), 17)
        d.Close(0)
    finally:
        app.Quit(0)
    rw = raabe_write("short", port=9379, background=True)
    try:
        rw.js(f"__app.openPath({str(src)!r})")
        time.sleep(2.5)
        rw.js("__app.closeOthers && __app.closeOthers()")
        our_pdf.unlink(missing_ok=True)
        rw.js(f"__app.exportPdfTo({str(our_pdf)!r})")
        deadline = time.time() + 20
        while time.time() < deadline and not (our_pdf.exists() and our_pdf.stat().st_size > 1000):
            time.sleep(0.2)
        time.sleep(0.5)
    finally:
        rw.close()
    w, o = lines(word_pdf), lines(our_pdf)
    res = {"word": w, "ours": o}
    OUT.write_text(json.dumps(res, indent=1), encoding="utf-8", newline="\n")
    for i in range(max(len(w), len(o))):
        print(i, "word", w[i] if i < len(w) else None, "| ours", o[i] if i < len(o) else None)


if __name__ == "__main__":
    main()
