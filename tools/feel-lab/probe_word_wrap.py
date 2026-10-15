"""Where Word starts to put text beside a square-wrapped picture (Step 4c): paragraphs of short
words ("ab") with a picture at growing distances from the left margin; Word's PDF shows on which
lines text stands left of the picture. In Word's own invisible instance; nothing is kept.

    .venv/Scripts/python.exe probe_word_wrap.py
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

import win32com.client

sys.path.insert(0, str(Path(__file__).parent))
from feellab import pdftext  # noqa: E402
from make_fixtures import Builder, base_doc  # noqa: E402

OFFSETS = [14.5, 15.5, 16.5, 17.5]


def main() -> None:
    tmp = Path(tempfile.gettempdir())
    png = tmp / "raabe-fixture-wrap-opaque.png"
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    pdf = tmp / "raabe-probe-wrap.pdf"
    try:
        d = base_doc(app)
        b = Builder(d)
        for off in OFFSETS:
            p = b.para(" ".join(["ab"] * 160))
            il = d.InlineShapes.AddPicture(str(png), False, True, d.Range(p.Range.Start, p.Range.Start))
            il.Width, il.Height = 120, 50
            sh = il.ConvertToShape()
            sh.WrapFormat.Type = 0
            sh.RelativeHorizontalPosition = 0
            sh.Left = off + 9  # the picture's left edge; its text distance is 9 pt
            sh.RelativeVerticalPosition = 2
            sh.Top = 0
            d.Range(p.Range.End - 1, p.Range.End - 1).InsertParagraphAfter()
        d.SaveAs2(str(pdf), 17)
        d.Close(0)
    finally:
        app.Quit(0)
    pages = pdftext.read(pdf)
    for off in OFFSETS:
        pass
    rows: dict[float, list] = {}
    for pi, pg in enumerate(pages):
        for r in pg.runs:
            rows.setdefault((pi, round(r.y, 1)), []).append(r)
    for k in sorted(rows):
        xs = sorted(round(r.x, 1) for r in rows[k])
        print(k, xs[:3], "…" if len(xs) > 3 else "")


if __name__ == "__main__":
    main()
