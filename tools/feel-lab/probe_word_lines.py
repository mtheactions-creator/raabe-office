"""Measures how Word places lines: builds paragraphs with different fonts, sizes and line
spacing rules and reads each line's vertical position (points from the page top) through
COM. Used to derive the layout rules in packages/text-engine/src/layout.ts.

    .venv/Scripts/python.exe probe_word_lines.py
"""

from __future__ import annotations

import json

import win32com.client

LOREM = ("Die Lichtenergie wird in chemische Energie umgewandelt, und es entstehen Zucker und "
         "Sauerstoff aus Wasser und Kohlenstoffdioxid in den Blättern der Pflanzen. ") * 2

WD_LINE = {"single": 0, "1.5": 1, "double": 2, "atLeast": 3, "exactly": 4, "multiple": 5}


def line_tops(doc, para) -> list[float]:
    """Vertical position of every line start in a paragraph."""
    app = doc.Application
    sel = app.Selection
    sel.SetRange(para.Range.Start, para.Range.Start)
    tops = []
    end = para.Range.End
    while sel.Start < end - 1:
        sel.HomeKey(5)
        tops.append(round(sel.Information(6), 2))  # wdVerticalPositionRelativeToPage
        before = sel.Start
        if sel.MoveDown(5, 1) == 0 or sel.Start <= before or sel.Start >= end:
            break
    return tops


CASES = [
    # name, font, size, rule, value, extra
    ("calibri11 multiple1.08", "Calibri", 11, "multiple", 12.96, None),
    ("calibri11 single", "Calibri", 11, "single", 12, None),
    ("calibri11 1.5", "Calibri", 11, "1.5", 18, None),
    ("calibri11 double", "Calibri", 11, "double", 24, None),
    ("calibri11 multiple3", "Calibri", 11, "multiple", 36, None),
    ("calibri11 exactly20", "Calibri", 11, "exactly", 20, None),
    ("calibri11 exactly10", "Calibri", 11, "exactly", 10, None),
    ("calibri11 atLeast20", "Calibri", 11, "atLeast", 20, None),
    ("calibri11 atLeast5", "Calibri", 11, "atLeast", 5, None),
    ("calibri20 single", "Calibri", 20, "single", 12, None),
    ("times12 single", "Times New Roman", 12, "single", 12, None),
    ("times12 multiple1.08", "Times New Roman", 12, "multiple", 12.96, None),
    ("arial10 single", "Arial", 10, "single", 12, None),
    ("cambria11 single", "Cambria", 11, "single", 12, None),
    ("aptos11 single", "Aptos", 11, "single", 12, None),
    ("segoe10 single", "Segoe UI", 10, "single", 12, None),
    ("comic12 single", "Comic Sans MS", 12, "single", 12, None),
    ("calibri11 single + 20pt word on line 2", "Calibri", 11, "single", 12, "big"),
    ("calibri11 multiple1.08 + 20pt word on line 2", "Calibri", 11, "multiple", 12.96, "big"),
    ("calibri11 single + times20 word on line 2", "Calibri", 11, "single", 12, "bigtimes"),
    ("calibri11 single + superscript on line 2", "Calibri", 11, "single", 12, "super"),
]


def main() -> None:
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    out = {}
    try:
        doc = app.Documents.Add()
        doc.PageSetup.TopMargin = 72
        for name, font, size, rule, value, extra in CASES:
            doc.Content.Delete()
            doc.Content.Text = LOREM
            p = doc.Paragraphs(1)
            r = p.Range
            r.Font.Name, r.Font.Size = font, size
            r.Font.Kerning, r.Font.Ligatures = 0, 0
            pf = p.Format
            pf.SpaceBefore, pf.SpaceAfter = 0, 0
            pf.LineSpacingRule = WD_LINE[rule]
            if rule in ("multiple", "exactly", "atLeast"):
                pf.LineSpacing = value
            if extra:
                # Put a word on the second line in a different size, font or position.
                tops0 = line_tops(doc, p)
                sel = app.Selection
                sel.SetRange(p.Range.Start, p.Range.Start)
                sel.MoveDown(5, 1)
                sel.EndKey(5)
                sel.MoveLeft(2, 2)  # wdWord: a word near the end of line 2
                sel.MoveRight(2, 1, 1)
                if extra == "big":
                    sel.Font.Size = 20
                elif extra == "bigtimes":
                    sel.Font.Name, sel.Font.Size = "Times New Roman", 20
                elif extra == "super":
                    sel.Font.Superscript = True
            tops = line_tops(doc, p)
            gaps = [round(b - a, 2) for a, b in zip(tops, tops[1:])]
            out[name] = {"first_top": tops[0] if tops else None, "line_gaps": gaps}
            print(f"{name:45s} first {tops[0] if tops else None:7}  gaps {gaps}")

        # Paragraph spacing: two paragraphs, 6 pt after the first and 12 pt before the second.
        doc.Content.Delete()
        doc.Content.Text = "Erster Absatz mit einer Zeile.\rZweiter Absatz mit einer Zeile.\rDritter Absatz."
        for i in (1, 2, 3):
            p = doc.Paragraphs(i)
            p.Range.Font.Name, p.Range.Font.Size = "Calibri", 11
            p.Format.LineSpacingRule = 0
            p.Format.SpaceBefore, p.Format.SpaceAfter = 0, 0
        doc.Paragraphs(1).Format.SpaceAfter = 6
        doc.Paragraphs(2).Format.SpaceBefore = 12
        tops = [round(doc.Paragraphs(i).Range.Information(6), 2) for i in (1, 2, 3)]
        out["paragraph spacing 6 after + 12 before"] = {"tops": tops}
        print("paragraph spacing 6 after + 12 before          tops", tops)
        doc.Paragraphs(1).Format.SpaceAfter = 10
        doc.Paragraphs(2).Format.SpaceBefore = 0
        doc.Paragraphs(1).Format.Style = doc.Styles(-1)
        doc.Close(0)
    finally:
        app.Quit(0)
    with open("captures/word-lines.json", "w", encoding="utf-8") as f:
        json.dump(out, f, indent=1)


if __name__ == "__main__":
    main()
