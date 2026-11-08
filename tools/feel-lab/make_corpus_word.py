"""The look corpus, Word part (D-080 D): school-like documents built by our own invisible Word in
TEMP/raabe-corpus/word (never committed, never the owner's files), compared with RaabeWrite by compare_corpus.py.

    .venv/Scripts/python.exe make_corpus_word.py [name ...]
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw

OUT = Path(tempfile.gettempdir()) / "raabe-corpus" / "word"
TEXT = ("Die Photosynthese ist ein Prozess, bei dem Pflanzen mit Hilfe von Lichtenergie aus Kohlenstoffdioxid und Wasser "
        "Traubenzucker und Sauerstoff herstellen. Sie findet in den Chloroplasten statt und ist die Grundlage fast allen Lebens auf der Erde.")
WD_STORY = 6


def doc(app):
    d = app.Documents.Add()
    d.Content.Text = ""
    return d


def p(d, text: str = "", style: int | str | None = None):
    r = d.Content
    r.InsertParagraphAfter()
    para = d.Paragraphs(d.Paragraphs.Count)
    para.Range.Text = text
    if style is not None:
        para.Style = style
    return d.Paragraphs(d.Paragraphs.Count)


def arbeitsblatt(app) -> None:
    d = doc(app)
    d.Paragraphs(1).Range.Text = "Arbeitsblatt Biologie – Photosynthese"
    d.Paragraphs(1).Style = -2  # Überschrift 1
    p(d, "Name: ____________________   Klasse: ______   Datum: __________")
    p(d, "Aufgabe 1", -3)
    p(d, TEXT)
    p(d, "Lies den Text und beantworte die Fragen:")
    for q in ["Wo findet die Photosynthese statt?", "Welche Stoffe werden benötigt?", "Welche Stoffe entstehen?"]:
        para = p(d, q)
        para.Range.ListFormat.ApplyNumberDefault()
    p(d, "Aufgabe 2", -3)
    p(d, "Ergänze die Wortgleichung:")
    p(d, "Kohlenstoffdioxid + Wasser → ______________ + ______________")
    d.SaveAs2(str(OUT / "arbeitsblatt.docx"), 16)
    d.Close(0)


def klausur(app) -> None:
    d = doc(app)
    d.Paragraphs(1).Range.Text = "Klassenarbeit Mathematik Nr. 2"
    d.Paragraphs(1).Style = -1  # Titel
    p(d, "Bearbeitungszeit: 45 Minuten. Hilfsmittel: Taschenrechner, Formelsammlung.")
    t = d.Tables.Add(p(d).Range, 2, 5)
    t.Borders.Enable = True
    for i, h in enumerate(["Aufgabe", "1", "2", "3", "Summe"]):
        t.Cell(1, i + 1).Range.Text = h
    t.Cell(2, 1).Range.Text = "Punkte"
    p(d, "1. Löse die Gleichung.", -3)
    para = p(d, "")
    r = para.Range
    r.Text = "x^2-5x+6=0"
    d.OMaths.Add(r)
    d.OMaths(d.OMaths.Count).BuildUp()
    p(d, "2. Berechne den Flächeninhalt eines Kreises mit dem Radius r = 4 cm.", None)
    para = p(d, "")
    r = para.Range
    r.Text = "A=πr^2"
    d.OMaths.Add(r)
    d.OMaths(d.OMaths.Count).BuildUp()
    p(d, "3. Zeichne den Graphen der Funktion f(x) = 2x − 1 in ein Koordinatensystem.")
    sec = d.Sections(1)
    sec.Footers(1).PageNumbers.Add(2)
    d.SaveAs2(str(OUT / "klausur.docx"), 16)
    d.Close(0)


def elternbrief(app) -> None:
    d = doc(app)
    d.Paragraphs(1).Range.Text = "Raabeschule · Musterstraße 1 · 12345 Musterstadt"
    d.Paragraphs(1).Range.Font.Size = 8
    p(d, "")
    p(d, "An die Eltern der Klasse 7a")
    p(d, "")
    para = p(d, "Musterstadt, 6. November 2026")
    para.Alignment = 2
    p(d, "Klassenfahrt nach Hamburg", -3)
    p(d, "Liebe Eltern,")
    for _ in range(3):
        p(d, TEXT + " " + TEXT)
    p(d, "Mit freundlichen Grüßen")
    p(d, "")
    p(d, "Klassenleitung")
    for para in list(d.Paragraphs)[5:]:
        para.Format.SpaceAfter = 6
        para.Alignment = 3  # Blocksatz
    d.SaveAs2(str(OUT / "elternbrief.docx"), 16)
    d.Close(0)


def tabelle(app) -> None:
    d = doc(app)
    d.Paragraphs(1).Range.Text = "Vokabeltest Englisch"
    d.Paragraphs(1).Style = -2
    rows = [("Englisch", "Deutsch", "Wortart"), ("the attempt", "der Versuch", "Nomen"), ("to achieve", "erreichen", "Verb"),
            ("reliable", "zuverlässig", "Adjektiv"), ("although", "obwohl", "Konjunktion"), ("the purpose", "der Zweck", "Nomen")]
    t = d.Tables.Add(p(d).Range, len(rows), 3)
    t.Style = "Gitternetztabelle 4 – Akzent 1"
    for i, row in enumerate(rows):
        for k, v in enumerate(row):
            t.Cell(i + 1, k + 1).Range.Text = v
    p(d, "")
    t2 = d.Tables.Add(p(d).Range, 3, 4)
    t2.Borders.Enable = True
    t2.Cell(1, 1).Merge(t2.Cell(1, 4))
    t2.Cell(1, 1).Range.Text = "Verbundene Überschrift"
    t2.Cell(1, 1).Shading.BackgroundPatternColor = 0xD9D9D9
    d.SaveAs2(str(OUT / "tabelle.docx"), 16)
    d.Close(0)


def listen(app) -> None:
    d = doc(app)
    d.Paragraphs(1).Range.Text = "Gliederung des Referats"
    d.Paragraphs(1).Style = -2
    items = [(1, "Einleitung"), (2, "Thema und Fragestellung"), (2, "Vorgehen"), (1, "Hauptteil"), (2, "Geschichte"), (3, "Antike"), (3, "Mittelalter"),
             (2, "Heute"), (1, "Schluss")]
    first = None
    for lvl, text in items:
        para = p(d, text)
        if first is None:
            first = para
    rng = d.Range(first.Range.Start, d.Content.End)
    rng.ListFormat.ApplyOutlineNumberDefault()
    for i, (lvl, _) in enumerate(items):
        for _ in range(lvl - 1):
            d.Paragraphs(i + 2).Range.ListFormat.ListIndent()
    p(d, "")
    for b in ["Material mitbringen", "Folien vorbereiten", "Handout drucken"]:
        para = p(d, b)
        para.Range.ListFormat.ApplyBulletDefault()
    d.SaveAs2(str(OUT / "listen.docx"), 16)
    d.Close(0)


def spalten(app) -> None:
    d = doc(app)
    d.Paragraphs(1).Range.Text = "Schülerzeitung – Ausgabe November"
    d.Paragraphs(1).Style = -1
    start = p(d, TEXT)
    for _ in range(8):
        p(d, TEXT + " " + TEXT)
    rng = d.Range(start.Range.Start, d.Content.End)
    rng.PageSetup.TextColumns.SetCount(2)
    d.SaveAs2(str(OUT / "spalten.docx"), 16)
    d.Close(0)


def bilder(app) -> None:
    png = OUT / "_bild.png"
    img = Image.new("RGB", (400, 260), (220, 235, 250))
    dr = ImageDraw.Draw(img)
    dr.ellipse((40, 30, 200, 190), fill=(255, 200, 0))
    dr.rectangle((220, 120, 380, 240), fill=(40, 140, 60))
    img.save(png)
    d = doc(app)
    d.Paragraphs(1).Range.Text = "Steckbrief Sonnenblume"
    d.Paragraphs(1).Style = -2
    para = p(d, "")
    s = d.InlineShapes.AddPicture(str(png), False, True, para.Range)
    s.Width, s.Height = 200, 130
    p(d, "Abbildung 1: Sonnenblume (Testbild)").Range.Font.Italic = True
    p(d, TEXT)
    p(d, TEXT)
    d.SaveAs2(str(OUT / "bilder.docx"), 16)
    d.Close(0)


def kopfzeile(app) -> None:
    d = doc(app)
    d.Paragraphs(1).Range.Text = "Protokoll der Projektwoche"
    d.Paragraphs(1).Style = -2
    for i in range(14):
        p(d, f"Tag {i + 1}: " + TEXT)
    sec = d.Sections(1)
    sec.Headers(1).Range.Text = "Raabeschule – Projektwoche 2026"
    sec.Footers(1).PageNumbers.Add(1)
    d.SaveAs2(str(OUT / "kopfzeile.docx"), 16)
    d.Close(0)


def fussnoten(app) -> None:
    d = doc(app)
    d.Paragraphs(1).Range.Text = "Facharbeit: Die Industrialisierung"
    d.Paragraphs(1).Style = -2
    for i in range(4):
        para = p(d, TEXT)
        d.Footnotes.Add(d.Range(para.Range.End - 1, para.Range.End - 1), Text=f"Quelle {i + 1}: Lehrbuch Geschichte, S. {10 + i}.")
    d.SaveAs2(str(OUT / "fussnoten.docx"), 16)
    d.Close(0)


def querformat(app) -> None:
    d = doc(app)
    d.PageSetup.Orientation = 1
    d.Paragraphs(1).Range.Text = "Zeitstrahl"
    d.Paragraphs(1).Style = -1
    t = d.Tables.Add(p(d).Range, 2, 8)
    t.Borders.Enable = True
    for i, y in enumerate(["1789", "1815", "1848", "1871", "1918", "1933", "1945", "1990"]):
        t.Cell(1, i + 1).Range.Text = y
        t.Cell(2, i + 1).Range.Text = "Ereignis " + y
    d.SaveAs2(str(OUT / "querformat.docx"), 16)
    d.Close(0)


def schriften(app) -> None:
    d = doc(app)
    d.Paragraphs(1).Range.Text = "Schriften im Vergleich"
    d.Paragraphs(1).Style = -2
    for f in ["Calibri", "Aptos", "Arial", "Times New Roman", "Cambria", "Georgia", "Verdana", "Comic Sans MS", "Segoe UI"]:
        para = p(d, f"{f}: Franz jagt im komplett verwahrlosten Taxi quer durch Bayern. 0123456789")
        para.Range.Font.Name = f
    para = p(d, "Fett, kursiv, unterstrichen, durchgestrichen, hochgestellt und tiefgestellt.")
    r = para.Range
    r.Words(1).Bold = True
    d.SaveAs2(str(OUT / "schriften.docx"), 16)
    d.Close(0)


def formen(app) -> None:
    d = doc(app)
    d.Paragraphs(1).Range.Text = "Wasserkreislauf"
    d.Paragraphs(1).Style = -2
    p(d, TEXT)
    anchor = p(d, "")
    shapes = [(1, 40, 100, 120, 50, "Verdunstung"), (9, 220, 80, 120, 80, "Wolken"), (5, 400, 100, 120, 50, "Niederschlag")]
    for kind, x, y, w, h, t in shapes:
        s = d.Shapes.AddShape(kind, x, y, w, h, anchor.Range)
        s.TextFrame.TextRange.Text = t
    d.Shapes.AddLine(160, 125, 220, 120, anchor.Range).Line.EndArrowheadStyle = 2
    d.Shapes.AddLine(340, 120, 400, 125, anchor.Range).Line.EndArrowheadStyle = 2
    for _ in range(6):
        p(d, "")
    p(d, "Erkläre die Stationen des Wasserkreislaufs.")
    d.SaveAs2(str(OUT / "formen.docx"), 16)
    d.Close(0)


def inhaltsverzeichnis(app) -> None:
    d = doc(app)
    d.Paragraphs(1).Range.Text = "Inhalt"
    d.Paragraphs(1).Style = -1
    toc_at = p(d, "")
    for i, h in enumerate(["Einleitung", "Material und Methoden", "Ergebnisse", "Diskussion", "Fazit"]):
        para = p(d, f"{i + 1} {h}", -2)
        para.Range.InsertParagraphBefore() if False else None
        for _ in range(3):
            p(d, TEXT, -1 if False else None)
            d.Paragraphs(d.Paragraphs.Count).Style = -67  # Standard
    d.TablesOfContents.Add(toc_at.Range, True, 1, 3)
    d.SaveAs2(str(OUT / "inhaltsverzeichnis.docx"), 16)
    d.Close(0)


def textfeld(app) -> None:
    d = doc(app)
    d.Paragraphs(1).Range.Text = "Merkkasten"
    d.Paragraphs(1).Style = -2
    anchor = p(d, TEXT)
    tb = d.Shapes.AddTextbox(1, 300, 120, 200, 90, anchor.Range)
    tb.TextFrame.TextRange.Text = "Merke: Pflanzen brauchen Licht, Wasser und Kohlenstoffdioxid."
    tb.Fill.ForeColor.RGB = 0xCCF2FF
    for _ in range(4):
        p(d, TEXT)
    d.SaveAs2(str(OUT / "textfeld.docx"), 16)
    d.Close(0)


BUILDERS = {f.__name__: f for f in [arbeitsblatt, klausur, elternbrief, tabelle, listen, spalten, bilder, kopfzeile, fussnoten, querformat,
                                    schriften, formen, inhaltsverzeichnis, textfeld]}


def main() -> None:
    import win32com.client
    OUT.mkdir(parents=True, exist_ok=True)
    only = sys.argv[1:] or list(BUILDERS)
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    done, failed = [], {}
    try:
        for name in only:
            try:
                BUILDERS[name](app)
                done.append(name)
            except Exception as e:  # noqa: BLE001
                failed[name] = str(e)[:160]
                while app.Documents.Count:
                    app.Documents(1).Close(0)
    finally:
        app.Quit(0)
    print(len(done), "built;", "failed:", failed)


if __name__ == "__main__":
    main()
