"""The look corpus, PowerPoint part (D-080 D): school-like presentations built by our own PowerPoint (only when no
other PowerPoint runs, feellab/powerpoint.py's rules) in TEMP/raabe-corpus/ppt, with PowerPoint's PNG of every slide
beside each file (<name>-<n>.png at 1280 px) for comparing RaabePresent's rendering. Never committed.

    .venv/Scripts/python.exe make_corpus_ppt.py [name ...]
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw

from feellab.powerpoint import THEMES, OwnerPowerPointRunning, PowerPointTarget

OUT = Path(tempfile.gettempdir()) / "raabe-corpus" / "ppt"
TEXT = "Pflanzen bilden mit Lichtenergie aus Kohlenstoffdioxid und Wasser Traubenzucker und Sauerstoff."


def title(s, t: str) -> None:
    for i in range(1, s.Shapes.Count + 1):
        sh = s.Shapes(i)
        if sh.Type == 14 and sh.PlaceholderFormat.Type in (1, 3):  # title, centre title
            sh.TextFrame.TextRange.Text = t
            return


def body(s, lines: list[tuple[int, str]]) -> None:
    for i in range(1, s.Shapes.Count + 1):
        sh = s.Shapes(i)
        if sh.Type == 14 and sh.PlaceholderFormat.Type in (2, 7):  # body, object
            tr = sh.TextFrame.TextRange
            tr.Text = "\r".join(t for _, t in lines)
            for k, (lvl, _) in enumerate(lines):
                tr.Paragraphs(k + 1).IndentLevel = lvl
            return


def referat(pp, pres) -> None:
    s = pp.add_slide(pres, 1)
    title(s, "Die Photosynthese")
    s.Shapes(2).TextFrame.TextRange.Text = "Referat Biologie, Klasse 8b"
    s = pp.add_slide(pres, 2)
    title(s, "Gliederung")
    body(s, [(1, "Was ist Photosynthese?"), (1, "Wo findet sie statt?"), (2, "Chloroplasten"), (2, "Blattgrün"), (1, "Bedeutung für uns")])
    s = pp.add_slide(pres, 2)
    title(s, "Was ist Photosynthese?")
    body(s, [(1, TEXT), (2, "Licht als Energiequelle"), (3, "rotes und blaues Licht")])


def zwei_inhalte(pp, pres) -> None:
    s = pp.add_slide(pres, 4)
    title(s, "Vergleich")
    n = 0
    for i in range(1, s.Shapes.Count + 1):
        sh = s.Shapes(i)
        if sh.Type == 14 and sh.PlaceholderFormat.Type in (2, 7):
            sh.TextFrame.TextRange.Text = ["Tag\rLicht\rSauerstoff", "Nacht\rKein Licht\rAtmung"][n]
            n += 1


def tabelle(pp, pres) -> None:
    s = pp.add_slide(pres, 6)
    title(s, "Notenspiegel")
    t = s.Shapes.AddTable(3, 7, 60, 150, 840, 120).Table
    for k, v in enumerate(["Note", "1", "2", "3", "4", "5", "6"]):
        t.Cell(1, k + 1).Shape.TextFrame.TextRange.Text = v
    for k, v in enumerate(["Anzahl", "2", "5", "8", "6", "3", "1"]):
        t.Cell(2, k + 1).Shape.TextFrame.TextRange.Text = v


def diagramm(pp, pres) -> None:
    s = pp.add_slide(pres, 6)
    title(s, "Ergebnisse")
    ch = s.Shapes.AddChart2(-1, 51, 80, 120, 800, 380).Chart
    try:
        ch.ChartData.Workbook.Close()
    except Exception:  # noqa: BLE001
        pass


def bild(pp, pres) -> None:
    png = OUT / "_bild.png"
    img = Image.new("RGB", (640, 400), (225, 240, 250))
    d = ImageDraw.Draw(img)
    d.ellipse((60, 40, 340, 320), fill=(255, 190, 0))
    d.rectangle((380, 180, 600, 380), fill=(30, 130, 70))
    img.save(png)
    s = pp.add_slide(pres, 6)
    title(s, "Ein Bild")
    s.Shapes.AddPicture(str(png), False, True, 200, 130, 480, 300)


def formen(pp, pres) -> None:
    s = pp.add_slide(pres, 6)
    title(s, "Ablauf")
    for i, t in enumerate(["Frage", "Hypothese", "Versuch", "Ergebnis"]):
        sh = s.Shapes.AddShape(5, 40 + i * 225, 220, 180, 90)
        sh.TextFrame.TextRange.Text = t
        if i < 3:
            ln = s.Shapes.AddConnector(1, 220 + i * 225, 265, 265 + i * 225, 265)
            ln.Line.EndArrowheadStyle = 2


def design(pp, name: str, thmx: str):
    def build(pp_, pres) -> None:
        s = pp_.add_slide(pres, 1)
        title(s, f"Design {name}")
        s.Shapes(2).TextFrame.TextRange.Text = "Untertitel der Präsentation"
        s = pp_.add_slide(pres, 2)
        title(s, "Inhalt")
        body(s, [(1, "Erster Punkt"), (2, "Unterpunkt"), (1, "Zweiter Punkt")])
    build.__name__ = f"design_{name.lower()}"
    build.theme = THEMES / thmx  # type: ignore[attr-defined]
    return build


def abschnitt(pp, pres) -> None:
    s = pp.add_slide(pres, 3)
    title(s, "Teil 2: Versuche")
    s = pp.add_slide(pres, 7)
    title(s, "Nur Titel")


def wordart(pp, pres) -> None:
    s = pp.add_slide(pres, 7)
    s.Shapes.AddTextEffect(0, "Willkommen!", "Calibri", 54, True, False, 200, 200)


def gruppe(pp, pres) -> None:
    s = pp.add_slide(pres, 7)
    a = s.Shapes.AddShape(9, 200, 150, 150, 150)
    b = s.Shapes.AddShape(1, 380, 180, 200, 90)
    b.TextFrame.TextRange.Text = "Gruppiert"
    s.Shapes.Range([a.Name, b.Name]).Group()


def notizen(pp, pres) -> None:
    s = pp.add_slide(pres, 2)
    title(s, "Mit Notizen")
    body(s, [(1, "Die Notizen erscheinen nicht auf der Folie.")])
    s.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = "Hier steht, was ich sagen will."


BUILDERS = {f.__name__: f for f in [referat, zwei_inhalte, tabelle, diagramm, bild, formen, abschnitt, wordart, gruppe, notizen,
                                    design(None, "Ion", "Ion.thmx"), design(None, "Facette", "Facet.thmx"), design(None, "Rueckblick", "Retrospect.thmx")]}


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    only = sys.argv[1:] or list(BUILDERS)
    try:
        pp = PowerPointTarget(visible=False)
    except OwnerPowerPointRunning as e:
        raise SystemExit(f"skipped: {e}")
    done, failed = [], {}
    try:
        for name in only:
            fn = BUILDERS[name]
            try:
                pres = pp.new_presentation(with_window=False, theme=getattr(fn, "theme", None), name=name)
                fn(pp, pres)
                pp.set_properties(pres)
                pp.save_copy(pres, OUT / f"{name}.pptx")
                for i in range(1, pres.Slides.Count + 1):
                    pp.export_png(pres, i, OUT / f"{name}-{i}.png")
                done.append(name)
            except Exception as e:  # noqa: BLE001
                failed[name] = str(e)[:160]
    finally:
        pp.close()
    print(len(done), "built;", "failed:", failed)


if __name__ == "__main__":
    main()
