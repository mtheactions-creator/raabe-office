"""Builds test documents with Microsoft Word itself (isolated, invisible instance) into
packages/ooxml/test/fixtures/. The texts are written for this project.

    .venv/Scripts/python.exe make_fixtures.py
"""

from __future__ import annotations

from pathlib import Path

import win32com.client
from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "packages" / "ooxml" / "test" / "fixtures"
CAPTURES = Path(__file__).parent / "captures"

WD_STYLE = {"normal": -1, "h1": -2, "h2": -3, "title": -63, "quote": -181, "strong": -88, "list_bullet": -49, "list_number": -50}


def para(doc, text: str):
    r = doc.Content
    r.InsertParagraphAfter()
    p = doc.Paragraphs(doc.Paragraphs.Count)
    p.Range.Text = text
    return doc.Paragraphs(doc.Paragraphs.Count - 1) if doc.Paragraphs.Count > 1 else p


def formatting(app) -> None:
    d = app.Documents.Add()
    d.Content.Text = "\r".join([
        "Formatierungen im Überblick",                                     # 1 H1
        "Ein Satz mit fetten, kursiven und unterstrichenen Wörtern sowie einer roten und einer gelb hervorgehobenen Stelle.",  # 2
        "Wasser ist H2O, und die Fläche misst 25 m2. Doppelt unterstrichen und durchgestrichen.",  # 3
        "Dieser Absatz ist zentriert.",                                    # 4
        "Dieser Absatz ist rechtsbündig.",                                 # 5
        "Dieser Absatz steht im Blocksatz und ist lang genug, damit er über mehrere Zeilen läuft und man sieht, wie die Abstände zwischen den Wörtern verteilt werden, bis die Zeile voll ist.",  # 6
        "Ein Absatz mit Einzug links von zwei Zentimetern und einem hängenden Einzug, der ebenfalls über mehrere Zeilen laufen soll, damit man den Unterschied sieht.",  # 7
        "Name:\tKlasse:\tDatum:",                                          # 8 tabs
        "Zeilenabstand anderthalbfach: Dieser Text braucht ein wenig Länge, damit er umbricht und der Abstand sichtbar wird.",  # 9
        "Genau 20 Punkt Zeilenabstand, ebenfalls mit genug Text für einen Umbruch, damit es sich lohnt.",  # 10
        "Kleine Schrift (9 pt) in Times New Roman.",                       # 11
        "Große Schrift (18 pt) in Arial.",                                 # 12
        "Ein Zitat in der Formatvorlage Zitat.",                           # 13
        "Zeile eins\vZeile zwei nach manuellem Zeilenumbruch.",           # 14
        "Ende der ersten Seite.",                                          # 15 then page break
        "Zweite Seite nach einem Seitenumbruch.",                          # 16
        "Titel mit eigener Formatvorlage",                                 # 17 title
    ])
    P = d.Paragraphs
    P(1).Style = d.Styles(WD_STYLE["h1"])
    r = P(2).Range
    def word(rng, w):
        f = rng.Duplicate
        f.Find.Execute(FindText=w)
        return f
    word(r, "fetten").Font.Bold = True
    word(r, "kursiven").Font.Italic = True
    word(r, "unterstrichenen").Font.Underline = 1
    word(r, "roten").Font.Color = 0x0000FF
    word(r, "gelb").HighlightColorIndex = 7
    r3 = P(3).Range
    w = word(r3, "H2O"); c = w.Duplicate; c.SetRange(w.Start + 1, w.Start + 2); c.Font.Subscript = True
    w = word(r3, "m2"); c = w.Duplicate; c.SetRange(w.Start + 1, w.Start + 2); c.Font.Superscript = True
    word(r3, "Doppelt unterstrichen").Font.Underline = 3
    word(r3, "durchgestrichen").Font.StrikeThrough = True
    P(4).Alignment = 1
    P(5).Alignment = 2
    P(6).Alignment = 3
    P(7).LeftIndent = 56.7
    P(7).FirstLineIndent = -28.35
    P(8).TabStops.Add(Position=170, Alignment=0)
    P(8).TabStops.Add(Position=453, Alignment=2)
    P(9).LineSpacingRule = 1
    P(10).LineSpacingRule = 4
    P(10).LineSpacing = 20
    P(11).Range.Font.Name = "Times New Roman"; P(11).Range.Font.Size = 9
    P(12).Range.Font.Name = "Arial"; P(12).Range.Font.Size = 18
    P(13).Style = d.Styles(WD_STYLE["quote"])
    P(15).Range.InsertBreak(7) if False else None
    end15 = P(15).Range
    end15.SetRange(end15.End - 1, end15.End - 1)
    end15.InsertBreak(7)  # wdPageBreak
    last = d.Paragraphs(d.Paragraphs.Count)
    last.Style = d.Styles(WD_STYLE["title"])
    d.SaveAs2(str(OUT / "word-formatting.docx"), 16)
    d.Close(0)


def rich(app) -> None:
    CAPTURES.mkdir(exist_ok=True)
    png = CAPTURES / "fixture-image.png"
    im = Image.new("RGB", (240, 120), "#fde8e9")
    ImageDraw.Draw(im).ellipse((20, 10, 220, 110), outline="#ed1c24", width=6)
    im.save(png)
    d = app.Documents.Add()
    d.Content.Text = "\r".join([
        "Arbeitsblatt Photosynthese",          # 1
        "Lies den Text und beantworte die Fragen.",  # 2 footnote + comment + link
        "Material:",                            # 3 then table after
        "Erste Aufgabe",                        # 4 list
        "Zweite Aufgabe",                       # 5 list
        "Ein Punkt",                            # 6 bullets
        "Noch ein Punkt",                       # 7 bullets
        "Die Formel lautet: ",                  # 8 equation
        "Hier wurde etwas geändert.",           # 9 track changes
        "Ein Bild im Text: ",                   # 10 inline image
        "Ein Textfeld schwebt neben diesem Absatz.",  # 11 shape
        "Letzter Absatz mit Textmarke.",        # 12 bookmark
    ])
    P = d.Paragraphs
    P(1).Style = d.Styles(WD_STYLE["h1"])
    r2 = P(2).Range
    fn_at = r2.Duplicate; fn_at.SetRange(r2.End - 1, r2.End - 1)
    d.Footnotes.Add(fn_at, Text="Der Text steht im Biologiebuch auf Seite 42.")
    d.Comments.Add(P(2).Range.Words(1), "Bitte genau lesen.")
    link = P(2).Range.Duplicate
    link.Find.Execute(FindText="Fragen")
    d.Hyperlinks.Add(link, Address="https://www.raabeschule.de")
    after3 = P(3).Range
    after3.InsertParagraphAfter()
    t_range = P(4).Range
    tbl = d.Tables.Add(t_range, 3, 2)
    for i, (a, b) in enumerate([("Stoff", "Menge"), ("Wasser", "100 ml"), ("Licht", "viel")], start=1):
        tbl.Cell(i, 1).Range.Text = a
        tbl.Cell(i, 2).Range.Text = b
    P = d.Paragraphs
    def find_para(text):
        for k in range(1, P.Count + 1):
            if P(k).Range.Text.startswith(text):
                return P(k)
        raise KeyError(text)
    find_para("Erste Aufgabe").Range.ListFormat.ApplyNumberDefault()
    find_para("Zweite Aufgabe").Range.ListFormat.ApplyNumberDefault()
    p6 = find_para("Ein Punkt"); p7 = find_para("Noch ein Punkt")
    rng = d.Range(p6.Range.Start, p7.Range.End); rng.ListFormat.ApplyBulletDefault()
    p8 = find_para("Die Formel")
    eq = p8.Range.Duplicate; eq.SetRange(p8.Range.End - 1, p8.Range.End - 1)
    eq.Text = "a^2+b^2=c^2"
    d.OMaths.Add(eq).OMaths(1).BuildUp()
    p9 = find_para("Hier wurde")
    d.TrackRevisions = True
    w = p9.Range.Duplicate; w.Find.Execute(FindText="etwas"); w.Text = "einiges"
    d.TrackRevisions = False
    p10 = find_para("Ein Bild")
    at = p10.Range.Duplicate; at.SetRange(p10.Range.End - 1, p10.Range.End - 1)
    d.InlineShapes.AddPicture(str(png), False, True, at)
    p11 = find_para("Ein Textfeld")
    box = d.Shapes.AddTextbox(1, 300, 400, 140, 50, p11.Range)
    box.TextFrame.TextRange.Text = "Merke: Licht ist nötig."
    p12 = find_para("Letzter Absatz")
    d.Bookmarks.Add("Ende", p12.Range)
    sec = d.Sections(1)
    sec.Headers(1).Range.Text = "Raabeschule – Biologie"
    sec.Footers(1).PageNumbers.Add(1)
    # A table of contents at the top.
    top = d.Range(0, 0)
    top.InsertParagraphBefore()
    d.TablesOfContents.Add(d.Range(0, 0), UseHeadingStyles=True, UpperHeadingLevel=1, LowerHeadingLevel=3)
    d.SaveAs2(str(OUT / "word-rich.docx"), 16)
    d.Close(0)


LAYOUT_FILL = (
    "Im Wald leben Pflanzen, Tiere, Pilze und Bakterien eng zusammen. Jede Art hat ihren Platz, und "
    "viele sind aufeinander angewiesen: Bäume liefern Schatten und Laub, Pilze zersetzen totes Holz, "
    "und Regenwürmer lockern den Boden, damit Wasser und Luft zu den Wurzeln gelangen."
)


def layout(app) -> None:
    """A worksheet-like document for layout fidelity (line and page breaks, positions). All
    fonts are named explicitly so the comparison does not depend on theme fonts."""
    d = app.Documents.Add()
    ps = d.PageSetup
    ps.PageWidth, ps.PageHeight = 595.3, 841.9
    ps.TopMargin, ps.BottomMargin, ps.LeftMargin, ps.RightMargin = 70.85, 56.7, 70.85, 70.85
    for sid, font, size in ((-1, "Calibri", 11), (-2, "Calibri Light", 16), (-3, "Calibri Light", 13), (-63, "Calibri Light", 28)):
        s = d.Styles(sid)
        s.Font.Name, s.Font.Size = font, size
    n = d.Styles(-1).ParagraphFormat
    n.SpaceAfter, n.LineSpacingRule, n.LineSpacing = 8, 5, 12 * 1.08
    texts = [
        ("title", "Arbeitsblatt: Ökosystem Wald"),
        ("normal", "Ein Wald ist mehr als eine Ansammlung von Bäumen. Im " "Mischwald wachsen Buchen, Eichen und Fichten "
                   "nebeneinander, und jede Schicht – vom Boden bis zur Krone – bietet anderen Lebewesen Raum. Wer genau hinsieht, "
                   "entdeckt fette Käfer, kursive Spuren und unterstrichene Besonderheiten."),
        ("h1", "1. Die Stockwerke des Waldes"),
        ("justify", "Die Waldbodenlebewesengemeinschaft ist erstaunlich vielfältig. In einer Hand voll Erde leben mehr Lebe\u00adwesen, "
                    "als es Menschen auf der Erde gibt. Sie zer\u00adsetzen Laub und Holz und machen die Nährstoffe wieder für die "
                    "Pflanzen verfügbar."),
        ("list", "a)\tBodenschicht: Moose, Pilze und Flechten"),
        ("list", "b)\tKrautschicht: Farne, Gräser und Blumen bis etwa einen Meter Höhe, die im Frühjahr blühen, bevor die Bäume Blätter tragen"),
        ("list", "c)\tStrauchschicht: Holunder, Haselnuss und junge Bäume"),
        ("h2", "1.1 Die Baumschicht"),
        ("exact", "Genau 12 Punkt Zeilenabstand bei elf Punkt Schrift: Die Zeilen stehen enger als sonst, und der Text muss "
                  "trotzdem gut lesbar bleiben, auch wenn er über mehrere Zeilen läuft."),
        ("atleast", "Mindestens 18 Punkt Zeilenabstand: Hier ist mehr Luft zwischen den Zeilen, weil der Wert größer ist als die "
                    "natürliche Zeilenhöhe der Schrift."),
        ("onehalf", "Anderthalbfacher Zeilenabstand mit 12 Punkt Abstand davor und 6 Punkt danach. " + LAYOUT_FILL),
        ("firstline", "Ein Absatz mit Erstzeileneinzug von einem Zentimeter. " + LAYOUT_FILL),
        ("tabs", "Name:\tKlasse:\tDatum:"),
        ("decimal", "Fichte\t12,5\tMeter"),
        ("decimal", "Buche\t8,25\tMeter"),
        ("mixed", "Gemischte Schriften: Calibri, dann Times New Roman, dann Arial und zuletzt Cambria in einer Zeile, "
                  "die so lang ist, dass sie umbricht und die Zeilenhöhe zeigt."),
        ("chem", "Bei der Fotosynthese entsteht aus CO2 und H2O Traubenzucker; ein Baum braucht dafür 10 m2 Blattfläche."),
        ("empty", ""),
        ("bigmark", ""),
        ("h1", "2. Nahrungsbeziehungen"),
        ("normal", LAYOUT_FILL + " " + LAYOUT_FILL),
        ("nbsp", "Am 3.\u00a0Oktober zählten wir 25\u00a0Rehe und 4\u00a0Wild\u2011schweine; der Förster sagte: „Das ist ein gutes Jahr.“"),
        ("longword", "Lange Adresse: www.raabeschule.de/unterricht/biologie/oekosystem-wald/arbeitsblatt-stockwerke-und-nahrungsketten-klasse7"),
        ("normal", LAYOUT_FILL),
        ("normal", LAYOUT_FILL + " Am Ende wiederholt sich der Kreislauf."),
        ("normal", LAYOUT_FILL),
        ("normal", "Ein letzter Absatz vor der Überschrift, der die Seite fast füllt. " + LAYOUT_FILL),
        ("h1", "3. Der Wald im Jahreslauf"),
        ("keeplines", "Dieser Absatz wird zusammengehalten. " + LAYOUT_FILL + " " + LAYOUT_FILL),
        ("normal", LAYOUT_FILL + " " + LAYOUT_FILL + " " + LAYOUT_FILL),
        ("normal", LAYOUT_FILL),
        ("normal", LAYOUT_FILL + " " + LAYOUT_FILL),
        ("normal", LAYOUT_FILL),
        ("pagebefore", "Aufgaben (Seitenumbruch oberhalb)"),
        ("normal", "Aufgabe 1: Beschreibe die Stockwerke des Waldes mit eigenen Worten."),
        ("normal", "Aufgabe 2: Erkläre, warum Pilze für den Wald wichtig sind."),
    ]
    # Word's own optional and non-breaking hyphens (what Strg+- and Strg+Umschalt+- insert)
    # are chr(31) and chr(30) in Range.Text; U+00AD/U+2011 would stay literal characters.
    d.Content.Text = "\r".join(t.replace(chr(0xAD), chr(31)).replace(chr(0x2011), chr(30)) for _, t in texts)
    P = d.Paragraphs

    def find(rng, w):
        f = rng.Duplicate
        assert f.Find.Execute(FindText=w), w
        return f

    for i, (kind, _) in enumerate(texts, start=1):
        p = P(i)
        r = p.Range
        if kind == "title":
            p.Style = d.Styles(-63)
        elif kind == "h1":
            p.Style = d.Styles(-2)
        elif kind == "h2":
            p.Style = d.Styles(-3)
        elif kind == "normal" and "Mischwald" in r.Text:
            f = find(r, "Mischwald"); f.Font.Size = 16
            find(r, "fette").Font.Bold = True
            find(r, "kursive").Font.Italic = True
            find(r, "unterstrichene").Font.Underline = 1
        elif kind == "justify":
            p.Alignment = 3
        elif kind == "list":
            p.Style = d.Styles(-180)  # Listenabsatz (contextual spacing)
            p.LeftIndent, p.FirstLineIndent = 36, -18
        elif kind == "exact":
            p.LineSpacingRule, p.LineSpacing = 4, 12
        elif kind == "atleast":
            p.LineSpacingRule, p.LineSpacing = 3, 18
        elif kind == "onehalf":
            p.LineSpacingRule = 1
            p.SpaceBefore, p.SpaceAfter = 12, 6
        elif kind == "firstline":
            p.FirstLineIndent = 28.35
        elif kind == "tabs":
            p.TabStops.Add(Position=226.8, Alignment=1)  # centre, 8 cm
            p.TabStops.Add(Position=453.6, Alignment=2)  # right, 16 cm
        elif kind == "decimal":
            p.TabStops.Add(Position=170.1, Alignment=3)  # decimal, 6 cm
            p.TabStops.Add(Position=255.15, Alignment=0)
            p.SpaceAfter = 0
        elif kind == "mixed":
            find(r, "Times New Roman").Font.Name = "Times New Roman"
            f = find(r, "Arial"); f.Font.Name = "Arial"
            f = find(r, "Cambria"); f.Font.Name = "Cambria"
        elif kind == "chem":
            for w, k, sup in (("CO2", 2, False), ("H2O", 1, False), ("m2", 1, True)):
                f = find(r, w); c = f.Duplicate; c.SetRange(f.Start + k, f.Start + k + 1)
                if sup:
                    c.Font.Superscript = True
                else:
                    c.Font.Subscript = True
        elif kind == "bigmark":
            r.Font.Size = 20
        elif kind == "keeplines":
            p.KeepTogether = True
        elif kind == "pagebefore":
            p.Style = d.Styles(-3)
            p.PageBreakBefore = True
    d.SaveAs2(str(OUT / "word-layout.docx"), 16)
    d.Close(0)


FONT_SAMPLES = ["Calibri", "Times New Roman", "Arial", "Cambria", "Cambria Math", "Georgia", "Verdana", "Comic Sans MS",
                "Segoe UI", "Courier New", "Century Gothic", "Book Antiqua", "Trebuchet MS"]
FONT_TEXT = "Beispieltext mit Unterlängen und Oberlängen, gyp ÄÖÜ, lang genug für einen Umbruch in die zweite Zeile hinein."


def fonts(app) -> None:
    """Vertical metrics: one two-line paragraph per font and line-spacing rule, plus mixed lines.
    Spacing after is 0 so every distance is a line height."""
    d = app.Documents.Add()
    ps = d.PageSetup
    ps.PageWidth, ps.PageHeight = 595.3, 841.9
    ps.TopMargin, ps.BottomMargin, ps.LeftMargin, ps.RightMargin = 70.85, 56.7, 70.85, 70.85
    s = d.Styles(-1)
    s.Font.Name, s.Font.Size = "Calibri", 11
    s.ParagraphFormat.SpaceAfter, s.ParagraphFormat.SpaceBefore = 0, 0
    s.ParagraphFormat.LineSpacingRule = 0  # single
    rows: list[tuple[str, int, str, float, str]] = []  # font, size, rule, value, text
    for f in FONT_SAMPLES:
        rows.append((f, 12, "single", 0, f"{f} 12 einfach: {FONT_TEXT}"))
    for f in ("Calibri", "Times New Roman", "Cambria", "Comic Sans MS"):
        for rule, value in (("exact", 12), ("exact", 20), ("atleast", 22), ("multiple", 1.5)):
            rows.append((f, 12, rule, value, f"{f} 12 {rule} {value}: {FONT_TEXT}"))
    for f in ("Calibri", "Times New Roman"):
        for size in (8, 20, 36):
            rows.append((f, size, "single", 0, f"{f} {size}: Kurzer Text mit gyp, der umbricht, wenn er lang genug ist."))
    rows.append(("mixed1", 11, "single", 0, "Calibri elf mit einem großen Times-Wort in derselben Zeile und weiterem Text bis zum Umbruch der Zeile hier."))
    rows.append(("mixed2", 11, "single", 0, "Times elf mit einem großen Calibri-Wort in derselben Zeile und weiterem Text bis zum Umbruch der Zeile hier."))
    rows.append(("mixed3", 11, "single", 0, "Calibri elf mit Comic Sans in der Mitte und weiterem Text, damit die Zeile umbricht und man es sieht."))
    rows.append(("super", 11, "single", 0, "Hochgestellt: x2 und Tiefgestellt: H2O in einer Zeile, die so lang ist, dass sie in die nächste Zeile umbricht."))
    rows.append(("bigsuper", 11, "single", 0, "Großes Hochgestellt: 10 hoch 23 in einer Zeile mit genug weiterem Text für einen Umbruch der Zeile."))
    d.Content.Text = "\r".join(r[4] for r in rows)
    P = d.Paragraphs
    for i, (f, size, rule, value, _text) in enumerate(rows, start=1):
        p = P(i)
        r = p.Range
        if f in ("mixed1", "mixed2", "mixed3", "super", "bigsuper"):
            base = "Times New Roman" if f == "mixed2" else "Calibri"
            r.Font.Name, r.Font.Size = base, size
            def word(w):
                x = r.Duplicate
                assert x.Find.Execute(FindText=w), w
                return x
            if f == "mixed1":
                w = word("Times-Wort"); w.Font.Name, w.Font.Size = "Times New Roman", 20
            elif f == "mixed2":
                w = word("Calibri-Wort"); w.Font.Name, w.Font.Size = "Calibri", 20
            elif f == "mixed3":
                w = word("Comic Sans"); w.Font.Name = "Comic Sans MS"
            elif f == "super":
                for needle, k, sup in (("x2", 1, True), ("H2O", 1, False)):
                    w = word(needle); c = w.Duplicate; c.SetRange(w.Start + k, w.Start + k + 1)
                    if sup:
                        c.Font.Superscript = True
                    else:
                        c.Font.Subscript = True
            else:
                w = word("23"); w.Font.Superscript = True; w.Font.Size = 20
            continue
        r.Font.Name, r.Font.Size = f, size
        if rule == "exact":
            p.LineSpacingRule, p.LineSpacing = 4, value
        elif rule == "atleast":
            p.LineSpacingRule, p.LineSpacing = 3, value
        elif rule == "multiple":
            p.LineSpacingRule, p.LineSpacing = 5, 12 * value
    d.SaveAs2(str(OUT / "word-fonts.docx"), 16)
    d.Close(0)


def breaks(app) -> None:
    """Space before at the top of a page after each kind of break."""
    d = app.Documents.Add()
    ps = d.PageSetup
    ps.PageWidth, ps.PageHeight = 595.3, 841.9
    ps.TopMargin, ps.BottomMargin, ps.LeftMargin, ps.RightMargin = 70.85, 56.7, 70.85, 70.85
    n = d.Styles(-1)
    n.Font.Name, n.Font.Size = "Calibri", 11
    n.ParagraphFormat.SpaceAfter, n.ParagraphFormat.LineSpacingRule, n.ParagraphFormat.LineSpacing = 8, 5, 12 * 1.08
    for sid in (-2, -3):
        d.Styles(sid).Font.Name = "Calibri Light"
    texts = [
        ("normal", "Erste Seite, der Absatz endet mit einem manuellen Seitenumbruch."),
        ("h1", "Überschrift nach manuellem Umbruch"),
        ("before", "Normaler Absatz mit 18 Punkt Abstand davor, nach einem manuellen Umbruch."),
        ("h1", "Überschrift mit Seitenumbruch oberhalb"),
        ("normal", "Text nach der Überschrift."),
        ("mid", "Text vor dem Umbruch in der Mitte des Absatzes und Text danach auf der neuen Seite."),
        ("normal", "Ein Absatz nach dem geteilten Absatz."),
    ] + [("normal", LAYOUT_FILL + " " + LAYOUT_FILL)] * 9 + [
        ("h1", "Überschrift am natürlichen Seitenanfang"),
        ("normal", "Text danach."),
    ]
    d.Content.Text = "\r".join(t for _, t in texts)
    P = d.Paragraphs
    for i, (kind, _) in enumerate(texts, start=1):
        p = P(i)
        if kind == "h1":
            p.Style = d.Styles(-2)
        if kind == "before":
            p.SpaceBefore = 18
    # Manual page breaks: at the end of paragraph 1, after the paragraph with 18 pt before,
    # and in the middle of the "mid" paragraph; "page break before" on the second heading.
    # Found by text, since inserting a break can add paragraphs.
    for needle in (" und Text danach", "manuellen Seitenumbruch.", "nach einem manuellen Umbruch."):
        f = d.Content
        assert f.Find.Execute(FindText=needle), needle
        at = f.Start if needle.startswith(" und") else f.End
        d.Range(at, at).InsertBreak(7)
    for i in range(1, P.Count + 1):
        if P(i).Range.Text.startswith("Überschrift mit Seitenumbruch"):
            P(i).PageBreakBefore = True
    # The heading at the natural page top: pad with filler until it lands at the top of a page.
    target = next(i for i in range(1, P.Count + 1) if P(i).Range.Text.startswith("Überschrift am natürlichen"))
    for _ in range(40):
        r = P(target).Range
        if r.Information(3) != P(target - 1).Range.Information(3):
            break
        P(target - 1).Range.InsertAfter(" Noch mehr Text.")
    d.SaveAs2(str(OUT / "word-breaks.docx"), 16)
    d.Close(0)


def top(app) -> None:
    """Space before on the very first paragraph and directly after a page-break character."""
    d = app.Documents.Add()
    ps = d.PageSetup
    ps.PageWidth, ps.PageHeight = 595.3, 841.9
    ps.TopMargin, ps.BottomMargin, ps.LeftMargin, ps.RightMargin = 70.85, 56.7, 70.85, 70.85
    d.Styles(-1).Font.Name = "Calibri"
    d.Styles(-2).Font.Name = "Calibri Light"
    d.Content.Text = "\r".join(["Überschrift ganz oben", "Text.", "Vor dem Umbruch", "Absatz mit 24 Punkt davor", "Text."])
    P = d.Paragraphs
    P(1).Style = d.Styles(-2)
    P(1).SpaceBefore = 24
    P(4).SpaceBefore = 24
    # A page-break character at the end of paragraph 3 (inside it, no extra paragraph).
    r = P(3).Range
    d.Range(r.End - 1, r.End - 1).InsertAfter(chr(12))
    d.SaveAs2(str(OUT / "word-top.docx"), 16)
    d.Close(0)


TEST_AUTHOR = "Raabe Office Test"


def anonymize(path: Path) -> None:
    """The repository is public: replace the Office user's name, initials and account id that
    Word writes into documents (core properties, comments, revisions, people.xml)."""
    import re
    import zipfile
    with zipfile.ZipFile(path) as z:
        entries = [(i, z.read(i.filename)) for i in z.infolist()]
    core = next(d for i, d in entries if i.filename == "docProps/core.xml").decode("utf-8")
    m = re.search(r"<dc:creator>([^<]+)</dc:creator>", core)
    name = m.group(1) if m else None
    out = []
    for info, data in entries:
        if info.filename.endswith((".xml", ".rels")):
            text = data.decode("utf-8")
            if name and name != TEST_AUTHOR:
                text = text.replace(name, TEST_AUTHOR)
            text = re.sub(r'w:initials="[^"]*"', 'w:initials="RT"', text)
            text = re.sub(r'w15:userId="[^"]*"', 'w15:userId="0000000000000000"', text)
            data = text.encode("utf-8")
        out.append((info, data))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for info, data in out:
            z.writestr(info, data)


def main() -> None:
    import sys
    if "--anonymize" in sys.argv:
        for f in sorted(OUT.glob("*.docx")):
            anonymize(f)
            print("anonymized", f.name)
        return
    OUT.mkdir(parents=True, exist_ok=True)
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    try:
        only = [a[2:] for a in sys.argv[1:] if a.startswith("--")]
        if only:
            for name in only:
                {"layout": layout, "fonts": fonts, "breaks": breaks, "top": top, "formatting": formatting, "rich": rich}[name](app)
        else:
            formatting(app)
            rich(app)
            layout(app)
            fonts(app)
            breaks(app)
            top(app)
    finally:
        app.Quit(0)
    for f in sorted(OUT.glob("*.docx")):
        anonymize(f)
        print(f.name, f.stat().st_size)


if __name__ == "__main__":
    main()
