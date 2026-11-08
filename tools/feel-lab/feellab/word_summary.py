"""The by-hand audit's view of a .docx through Word (D-133): the base document both sides start from, and a summary
of what an option can change (character and paragraph formatting of the target, lists, tables, page setup, objects,
notes, fields ...). Comparing the change a RaabeWrite option made with the change Word's same option made needs only
these summaries. Word runs in an invisible instance of our own (DispatchEx); nothing else is touched."""
from __future__ import annotations

import struct
import zlib
from pathlib import Path

# The base document: German school text, a list, a table and a picture. TARGET is what the text options act on.
PARAS = [
    ("h1", "Ein Bericht über Pflanzen"),
    ("p", "Der erste Absatz erklärt das Thema dieses Berichts in einigen ruhigen Sätzen. Pflanzen brauchen Licht, Wasser und Nährstoffe, damit sie wachsen können."),
    ("p", "Im zweiten Absatz stehen weitere Beobachtungen aus dem Schulgarten, die wir im Frühling notiert haben."),
    ("li", "Sonnenblumen"),
    ("li", "Tomaten"),
    ("li", "Kürbisse"),
    ("p", "Die folgende Tabelle zeigt die Messwerte."),
    ("table", [["Pflanze", "Höhe", "Woche"], ["Sonnenblume", "120", "8"], ["Tomate", "80", "6"]]),
    ("p", "Am Ende fassen wir die Ergebnisse kurz zusammen."),
    ("pic", ""),
    ("p", "Schluss."),
]
TARGET = "weitere Beobachtungen"
TARGET_PARA = 3  # 1-based paragraph of the target in Word (heading, first, second)


def png(path: Path, w: int = 160, h: int = 100) -> Path:
    """A small picture (a blue gradient with a red band), written without any library."""
    rows = []
    for y in range(h):
        row = bytearray([0])
        for x in range(w):
            if 40 <= y < 60:
                row += bytes([220, 40, 40])
            else:
                row += bytes([30, 90 + x // 2, 160 + y // 2])
        rows.append(bytes(row))
    raw = zlib.compress(b"".join(rows), 9)

    def chunk(t: bytes, d: bytes) -> bytes:
        return struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
    data = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)) + chunk(b"IDAT", raw) + chunk(b"IEND", b"")
    path.write_bytes(data)
    return path


def word_app():
    import win32com.client
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    return app


def build_base(app, path: Path, picture: Path) -> None:
    """Word writes the base document itself (no file of the owner's is used)."""
    d = app.Documents.Add()
    try:
        d.Content.Text = chr(13).join("" if k in ("table", "pic") else t for k, t in PARAS)
        d.Content.Style = -1
        for i, (kind, _) in enumerate(PARAS, 1):
            if kind == "h1":
                d.Paragraphs(i).Style = -2
        lis = [i for i, (k, _) in enumerate(PARAS, 1) if k == "li"]
        d.Range(d.Paragraphs(lis[0]).Range.Start, d.Paragraphs(lis[-1]).Range.End).ListFormat.ApplyBulletDefault()
        # Picture first (the table adds paragraphs, so indexes after it move).
        pi = next(i for i, (k, _) in enumerate(PARAS, 1) if k == "pic")
        r = d.Paragraphs(pi).Range
        r.Collapse(1)
        d.InlineShapes.AddPicture(str(picture), False, True, r)
        ti = next(i for i, (k, _) in enumerate(PARAS, 1) if k == "table")
        rows = next(t for k, t in PARAS if k == "table")
        r = d.Paragraphs(ti).Range
        r.Collapse(1)
        t = d.Tables.Add(r, len(rows), len(rows[0]))
        t.Style = "Tabellenraster"
        for i, row in enumerate(rows, 1):
            for j, cell in enumerate(row, 1):
                t.Cell(i, j).Range.Text = cell
        d.SaveAs2(str(path), 16)
    finally:
        d.Close(0)


def _font(f) -> dict:
    return {"bold": f.Bold, "italic": f.Italic, "underline": f.Underline, "strike": f.StrikeThrough,
            "dstrike": f.DoubleStrikeThrough, "sub": f.Subscript, "sup": f.Superscript, "size": f.Size, "name": f.Name,
            "color": f.TextColor.RGB if f.Color != 9999999 else f.Color, "caps": f.AllCaps, "smallcaps": f.SmallCaps,
            "hidden": f.Hidden, "spacing": round(f.Spacing, 1),
            "lig": f.Ligatures, "numForm": f.NumberForm, "numSpacing": f.NumberSpacing, "styleSet": f.StylisticSet}


def _para(p) -> dict:
    f = p.Format
    lf = p.Range.ListFormat
    out = {"style": p.Style.NameLocal if p.Style else None, "align": f.Alignment, "left": round(f.LeftIndent, 1),
           "right": round(f.RightIndent, 1), "first": round(f.FirstLineIndent, 1), "before": round(f.SpaceBefore, 1),
           "after": round(f.SpaceAfter, 1), "lineRule": f.LineSpacingRule, "line": round(f.LineSpacing, 1),
           "list": lf.ListString if lf.ListType else "", "listType": lf.ListType, "level": lf.ListLevelNumber if lf.ListType else 0,
           "shading": f.Shading.BackgroundPatternColor, "keepNext": f.KeepWithNext, "pageBreakBefore": f.PageBreakBefore,
           "borders": sum(1 for i in (-1, -2, -3, -4) if f.Borders(i).LineStyle != 0)}
    return out


def summarize(app, path: Path) -> dict:
    """What the audit compares, read by Word from the saved file."""
    d = app.Documents.Open(str(path), ReadOnly=True, AddToRecentFiles=False, Visible=False)
    try:
        s: dict = {}
        c = d.Content
        c.TextRetrievalMode.IncludeHiddenText = True
        text = c.Text
        s["text"] = text.replace("\r", "\n").replace("\x07", "|")[:4000]
        s["paras"] = d.Paragraphs.Count
        f = c.Duplicate.Find
        f.ClearFormatting()
        target = None
        rng = d.Content
        if rng.Find.Execute(FindText=TARGET, MatchCase=True):
            target = rng
        if target is not None:
            s["target"] = _font(target.Font)
            s["targetText"] = target.Text
            s["targetPara"] = _para(target.Paragraphs(1))
            s["targetHighlight"] = target.HighlightColorIndex
        n = d.Paragraphs.Count
        for i in range(1, min(n, 14) + 1):
            p = d.Paragraphs(i)
            if p.Range.Information(12):  # in a table: the table summary covers it
                continue
            s[f"p{i}"] = _para(p)
        ps = d.Sections(1).PageSetup
        s["page"] = {"orient": ps.Orientation, "w": round(ps.PageWidth, 1), "h": round(ps.PageHeight, 1),
                     "top": round(ps.TopMargin, 1), "bottom": round(ps.BottomMargin, 1), "left": round(ps.LeftMargin, 1),
                     "right": round(ps.RightMargin, 1), "cols": ps.TextColumns.Count, "lineNumbers": ps.LineNumbering.Active,
                     "valign": ps.VerticalAlignment}
        s["sections"] = d.Sections.Count
        s["tables"] = d.Tables.Count
        if d.Tables.Count:
            t = d.Tables(1)
            s["table1"] = {"rows": t.Rows.Count, "cols": t.Columns.Count, "style": t.Style.NameLocal if t.Style else None,
                           "cell11": t.Cell(1, 1).Range.Text[:-2], "headRow": t.Rows(1).HeadingFormat,
                           "shade11": t.Cell(1, 1).Shading.BackgroundPatternColor,
                           "hdr": t.ApplyStyleHeadingRows, "lastRow": t.ApplyStyleLastRow, "bandRows": t.ApplyStyleRowBands,
                           "firstCol": t.ApplyStyleFirstColumn, "lastCol": t.ApplyStyleLastColumn, "bandCols": t.ApplyStyleColumnBands,
                           "valign21": t.Cell(2, 1).VerticalAlignment, "align21": t.Cell(2, 1).Range.ParagraphFormat.Alignment,
                           "autofit": t.AllowAutoFit, "w11": round(t.Cell(1, 1).Width, 0)}
        s["inlineShapes"] = d.InlineShapes.Count
        s["shapes"] = d.Shapes.Count
        if d.InlineShapes.Count:
            ish = d.InlineShapes(1)
            s["pic1"] = {"w": round(ish.Width, 1), "h": round(ish.Height, 1)}
        s["footnotes"] = d.Footnotes.Count
        s["endnotes"] = d.Endnotes.Count
        s["comments"] = d.Comments.Count
        s["revisions"] = d.Revisions.Count
        s["track"] = d.TrackRevisions
        s["fields"] = sorted(fl.Type for fl in d.Fields)
        s["toc"] = d.TablesOfContents.Count
        s["hyperlinks"] = d.Hyperlinks.Count
        s["bookmarks"] = d.Bookmarks.Count
        s["oMaths"] = d.OMaths.Count
        hdr = d.Sections(1).Headers(1)
        ftr = d.Sections(1).Footers(1)
        s["header"] = hdr.Range.Text.strip() if hdr.Exists else ""
        s["footer"] = ftr.Range.Text.strip() if ftr.Exists else ""
        s["headerShapes"] = hdr.Shapes.Count if hdr.Exists else 0
        s["hyphenation"] = d.AutoHyphenation
        try:
            s["pageColor"] = d.Background.Fill.ForeColor.RGB if d.Background.Fill.Visible else None
        except Exception:  # noqa: BLE001
            s["pageColor"] = None
        s["pages"] = d.ComputeStatistics(2)
        s["words"] = d.ComputeStatistics(0)
        return s
    finally:
        d.Close(0)


def delta(base: dict, after: dict) -> dict:
    """Keys (flattened one level) whose value differs from the base."""
    out = {}
    for k in set(base) | set(after):
        a, b = base.get(k), after.get(k)
        if isinstance(a, dict) or isinstance(b, dict):
            a, b = a or {}, b or {}
            for kk in set(a) | set(b):
                if a.get(kk) != b.get(kk):
                    out[f"{k}.{kk}"] = b.get(kk)
        elif a != b:
            out[k] = b
    return out
