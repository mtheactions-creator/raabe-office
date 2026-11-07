"""Builds test presentations with Microsoft PowerPoint itself into
packages/ooxml/test/fixtures/pptx-*.pptx, for RaabePresent's .pptx reader and writer (Step 8).

PowerPoint 365 starts new presentations with its current "Office" design (Aptos Display / Aptos);
the project wants the Office 2013-2022 look (Calibri Light / Calibri), which PowerPoint ships as
`Office 2013 - 2022 Theme.thmx`. Every fixture is built twice:

    pptx-<name>.pptx          the new presentation with the Office 2013-2022 design applied
    pptx-office-<name>.pptx   PowerPoint's own default (Office design, Aptos)

Next to each: <name>.json with PowerPoint's read-back of every slide and shape (feellab.powerpoint:
positions in pt and EMU, text, paragraphs, line positions, placeholders, z-order ...), and
pptx-png/<name>-<n>.png, every slide exported by PowerPoint at 1280 px width.

PowerPoint runs as one process per user (feellab/powerpoint.py): the script refuses to run while a
POWERPNT.EXE is running (the owner's). Most fixtures are built windowless; autofit and the equation
need a document window (PowerPoint shrinks text and builds equations only there), so a window of
our own PowerPoint appears briefly. No real input is sent. The contents are written for this
project; author and last author are "Raabe Office".

    .venv/Scripts/python.exe make_pptx_fixtures.py               # all
    .venv/Scripts/python.exe make_pptx_fixtures.py --shapes      # one
    .venv/Scripts/python.exe make_pptx_fixtures.py --chart       # opt-in: opens Excel visibly
"""

from __future__ import annotations

import datetime
import json
import re
import sys
import tempfile
import time
import zipfile
from pathlib import Path

from feellab.powerpoint import THEME_2013, OwnerPowerPointRunning, PowerPointTarget
from feellab.sanitize import private_names

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "packages" / "ooxml" / "test" / "fixtures"
PNG = OUT / "pptx-png"
NEUTRAL = "Raabe Office"
VARIANTS = {"pptx-": THEME_2013, "pptx-office-": None}

TITLE_TYPES = {1, 3, 5}           # Title, CenterTitle, VerticalTitle
TEXT_TYPES = {2, 4, 6, 7, 17}     # Body, Subtitle, VerticalBody, Object, VerticalObject

MINUS, PLUSMINUS, ROOT_SIGN = chr(0x2212), chr(0xB1), chr(0x221A)


def rgb(r: int, g: int, b: int) -> int:
    """PowerPoint's colour value (BGR)."""
    return r + (g << 8) + (b << 16)


def placeholders(slide):
    for i in range(1, slide.Shapes.Placeholders.Count + 1):
        ph = slide.Shapes.Placeholders(i)
        yield ph, int(ph.PlaceholderFormat.Type)


def set_paragraphs(shape, paras: list[tuple[int, str]]) -> None:
    """Text as (indent level, text) paragraphs."""
    tr = shape.TextFrame.TextRange
    tr.Text = "\r".join(t for _, t in paras)
    for i, (lvl, _) in enumerate(paras, start=1):
        tr.Paragraphs(i).IndentLevel = lvl


def title(slide, text: str) -> None:
    for ph, t in placeholders(slide):
        if t in TITLE_TYPES:
            ph.TextFrame.TextRange.Text = text
            return


def body(slide, n: int = 1):
    k = 0
    for ph, t in placeholders(slide):
        if t in TEXT_TYPES:
            k += 1
            if k == n:
                return ph
    raise LookupError("no body placeholder")


def show(t: PowerPointTarget, pres, slide) -> None:
    """Brings a slide into our document window (autofit and equations happen there)."""
    win = pres.Windows(1)
    win.Activate()
    win.View.GotoSlide(slide.SlideIndex)
    time.sleep(0.3)


# ----- the fixtures -----------------------------------------------------------------------------

def blank(t, pres, tmp):
    """A new presentation as Datei > Neu > Leere Präsentation gives it: one empty Titelfolie."""
    t.add_slide(pres, 1)


def layouts(t, pres, tmp):
    """One slide per layout of the design, text in every text placeholder."""
    lays = pres.SlideMaster.CustomLayouts
    for i in range(1, lays.Count + 1):
        name = lays(i).Name.replace(chr(10), "").replace(chr(13), "").replace(chr(11), "")
        s = t.add_slide(pres, i)
        n = 0
        for ph, typ in placeholders(s):
            if typ in TITLE_TYPES:
                ph.TextFrame.TextRange.Text = f"{name}: Titel"
            elif typ in TEXT_TYPES:
                n += 1
                if typ == 4:
                    ph.TextFrame.TextRange.Text = f"Untertitel der Folie {i}"
                else:
                    set_paragraphs(ph, [(1, f"Text {n} auf der Folie {i}"), (2, "Zweite Ebene")])


def bullets(t, pres, tmp):
    """Bullet levels 1-5 and a numbered list."""
    s = t.add_slide(pres, 2)
    title(s, "Aufzählungsebenen")
    set_paragraphs(body(s), [(1, "Ebene 1"), (2, "Ebene 2"), (3, "Ebene 3"), (4, "Ebene 4"), (5, "Ebene 5"), (1, "Wieder Ebene 1")])
    s = t.add_slide(pres, 4)
    title(s, "Nummerierung und ohne Aufzählungszeichen")
    left = body(s, 1)
    set_paragraphs(left, [(1, "Erster Schritt"), (1, "Zweiter Schritt"), (2, "Unterpunkt a"), (2, "Unterpunkt b"), (1, "Dritter Schritt")])
    for i in range(1, 6):
        p = left.TextFrame.TextRange.Paragraphs(i)
        p.ParagraphFormat.Bullet.Type = 2  # ppBulletNumbered
        if p.IndentLevel == 2:
            p.ParagraphFormat.Bullet.Style = 0  # ppBulletAlphaLCParenRight
    right = body(s, 2)
    set_paragraphs(right, [(1, "Absatz ohne Aufzählungszeichen"), (1, "Noch ein Absatz")])
    right.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = 0


LONG = [f"Absatz {i}: Dieser Text ist absichtlich lang, damit er nicht in den Platzhalter passt." for i in range(1, 15)]


def autofit(t, pres, tmp):
    """Overflowing text: shrunk by PowerPoint's autofit, not fitted, and a text box that grows."""
    s = t.add_slide(pres, 2)
    title(s, "Text verkleinern (automatisch anpassen)")
    set_paragraphs(body(s), [(1, x) for x in LONG])
    show(t, pres, s)
    s = t.add_slide(pres, 2)
    title(s, "Nicht automatisch anpassen")
    b = body(s)
    b.TextFrame2.AutoSize = 0  # msoAutoSizeNone
    set_paragraphs(b, [(1, x) for x in LONG[:9]])
    show(t, pres, s)
    s = t.add_slide(pres, 2)
    title(s, "Form an Text anpassen")
    b = body(s)
    b.TextFrame2.AutoSize = 1  # msoAutoSizeShapeToFitText
    set_paragraphs(b, [(1, x) for x in LONG[:8]])
    show(t, pres, s)
    time.sleep(1.0)


def textbox(t, pres, tmp):
    s = t.add_slide(pres, 7)
    a = s.Shapes.AddTextbox(1, 60, 40, 300, 40)
    a.TextFrame.TextRange.Text = "Ein Textfeld mit einer Zeile"
    b = s.Shapes.AddTextbox(1, 60, 120, 320, 40)
    b.TextFrame.TextRange.Text = "Ein Textfeld mit Zeilenumbruch: der Text ist länger als das Feld breit ist und bricht darum um."
    c = s.Shapes.AddTextbox(1, 460, 40, 420, 60)
    tr = c.TextFrame.TextRange
    tr.Text = "Fett, kursiv, farbig und groß"
    tr.Characters(1, 4).Font.Bold = -1
    tr.Characters(7, 6).Font.Italic = -1
    tr.Characters(15, 6).Font.Color.RGB = rgb(192, 0, 0)
    tr.Characters(25, 4).Font.Size = 32
    tr.ParagraphFormat.Alignment = 2  # ppAlignCenter
    d = s.Shapes.AddTextbox(2, 460, 160, 60, 300)  # msoTextOrientationUpward
    d.TextFrame.TextRange.Text = "Nach oben gedreht"
    e = s.Shapes.AddTextbox(1, 560, 200, 340, 200)
    e.TextFrame2.AutoSize = 0
    e.TextFrame2.WordWrap = -1
    e.TextFrame2.VerticalAnchor = 3  # msoAnchorMiddle
    e.TextFrame.TextRange.Text = "Feste Größe, vertikal zentriert\rZweiter Absatz"
    e.Line.Visible = -1
    e.Line.ForeColor.RGB = rgb(68, 114, 196)


def shapes(t, pres, tmp):
    s = t.add_slide(pres, 7)
    r = s.Shapes.AddShape(1, 40, 40, 200, 120)  # rectangle, theme style
    r.TextFrame.TextRange.Text = "Rechteck"
    rr = s.Shapes.AddShape(5, 280, 40, 200, 120)  # rounded rectangle
    rr.Adjustments.SetItem(1, 0.3)
    rr.Fill.ForeColor.RGB = rgb(112, 173, 71)
    rr.Line.ForeColor.RGB = rgb(56, 87, 35)
    rr.Line.Weight = 3
    ar = s.Shapes.AddShape(33, 520, 60, 200, 80)  # right arrow
    ar.Rotation = 30
    ar.Fill.ForeColor.RGB = rgb(255, 192, 0)
    tri = s.Shapes.AddShape(8, 760, 40, 140, 120)  # right triangle
    tri.Flip(0)  # msoFlipHorizontal
    st = s.Shapes.AddShape(92, 40, 220, 160, 160)  # 5-point star
    st.Flip(1)  # msoFlipVertical
    st.Fill.ForeColor.RGB = rgb(237, 125, 49)
    ov = s.Shapes.AddShape(9, 240, 220, 200, 140)  # oval without line
    ov.Line.Visible = 0
    ov.Fill.ForeColor.RGB = rgb(91, 155, 213)
    ov.TextFrame.TextRange.Text = "Ellipse"
    ln = s.Shapes.AddLine(480, 240, 700, 380)
    ln.Line.Weight = 2.25
    ln.Line.DashStyle = 4  # msoLineDash
    ln.Line.EndArrowheadStyle = 2  # msoArrowheadTriangle
    cn = s.Shapes.AddConnector(2, 740, 220, 900, 380)  # elbow connector
    cn.Line.ForeColor.RGB = rgb(0, 0, 0)
    bl = s.Shapes.AddShape(106, 520, 420, 220, 90)  # rounded rectangular callout
    bl.Rotation = 350


def group(t, pres, tmp):
    s = t.add_slide(pres, 7)
    a = s.Shapes.AddShape(1, 200, 150, 200, 120)
    a.Name = "Gruppe Rechteck"
    b = s.Shapes.AddShape(9, 450, 170, 160, 160)
    b.Name = "Gruppe Ellipse"
    c = s.Shapes.AddTextbox(1, 220, 300, 200, 40)
    c.Name = "Gruppe Text"
    c.TextFrame.TextRange.Text = "Text in der Gruppe"
    g = s.Shapes.Range([a.Name, b.Name, c.Name]).Group()
    g.Name = "Gruppe 1"
    g.Rotation = 15
    d = s.Shapes.AddShape(4, 680, 380, 120, 80)
    d.Name = "Außen"
    e = s.Shapes.AddShape(7, 820, 380, 100, 80)
    e.Name = "Außen Dreieck"
    inner = s.Shapes.Range([d.Name, e.Name]).Group()
    inner.Name = "Innere Gruppe"
    f = s.Shapes.AddShape(1, 680, 470, 240, 40)
    f.Name = "Außen Balken"
    outer = s.Shapes.Range([inner.Name, f.Name]).Group()
    outer.Name = "Äußere Gruppe"


def make_png(path: Path) -> Path:
    """A test picture of our own: colour bands and a grid, 400 x 300 px."""
    from PIL import Image, ImageDraw
    im = Image.new("RGB", (400, 300), (255, 255, 255))
    dr = ImageDraw.Draw(im)
    colors = [(68, 114, 196), (237, 125, 49), (165, 165, 165), (255, 192, 0), (91, 155, 213), (112, 173, 71)]
    for i, c in enumerate(colors):
        dr.rectangle([0, i * 50, 399, i * 50 + 49], fill=c)
    for x in range(0, 400, 40):
        dr.line([x, 0, x, 299], fill=(0, 0, 0), width=1)
    for y in range(0, 300, 30):
        dr.line([0, y, 399, y], fill=(0, 0, 0), width=1)
    dr.ellipse([150, 100, 250, 200], outline=(0, 0, 0), width=4)
    im.save(path)
    return path


def picture(t, pres, tmp):
    png = str(make_png(tmp / "raabe-testbild.png"))
    s = t.add_slide(pres, 7)
    a = s.Shapes.AddPicture(png, 0, -1, 60, 60)
    a.Name = "Bild ganz"
    b = s.Shapes.AddPicture(png, 0, -1, 520, 60, 400, 300)
    b.Name = "Bild beschnitten"
    pf = b.PictureFormat
    pf.CropLeft = 40
    pf.CropTop = 30
    pf.CropRight = 20
    pf.CropBottom = 10
    c = s.Shapes.AddPicture(png, 0, -1, 120, 380, 200, 150)
    c.Name = "Bild gedreht"
    c.Rotation = 20


def table(t, pres, tmp):
    s = t.add_slide(pres, 6)  # Nur Titel
    title(s, "Tabelle mit der Standardformatvorlage")
    sh = s.Shapes.AddTable(4, 3, 66, 150, 828, 200)
    tb = sh.Table
    for c in range(1, 4):
        tb.Cell(1, c).Shape.TextFrame.TextRange.Text = f"Spalte {c}"
    for r in range(2, 5):
        for c in range(1, 4):
            tb.Cell(r, c).Shape.TextFrame.TextRange.Text = f"Zeile {r - 1}, Zelle {c}"
    tb.Cell(4, 3).Shape.TextFrame.TextRange.Text = "Eine längere Zelle, die umbricht"
    s2 = t.add_slide(pres, 7)
    sh2 = s2.Shapes.AddTable(3, 2, 100, 100, 400, 120)
    sh2.Table.Cell(1, 1).Merge(sh2.Table.Cell(1, 2))
    sh2.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Verbundene Kopfzelle"
    sh2.Table.Cell(2, 1).Shape.TextFrame.TextRange.Text = "a"
    sh2.Table.Cell(2, 2).Shape.TextFrame.TextRange.Text = "b"


def equation(t, pres, tmp):
    """Equations built by PowerPoint from their linear form (Formel > Professionell)."""
    s = t.add_slide(pres, 6)
    title(s, "Formeln")
    forms = ["a^2+b^2=c^2", f"x=({MINUS}b{PLUSMINUS}{ROOT_SIGN}(b^2{MINUS}4ac))/(2a)"]
    for i, f in enumerate(forms):
        tb = s.Shapes.AddTextbox(1, 100, 160 + i * 120, 600, 60)
        tb.TextFrame.TextRange.Text = f
        show(t, pres, s)
        tb.TextFrame.TextRange.Select()
        t.app.CommandBars.ExecuteMso("EquationInsertNew")
        time.sleep(0.6)
        t.app.CommandBars.ExecuteMso("EquationProfessional")
        time.sleep(0.6)
    pres.Windows(1).Selection.Unselect()


def show_fixture(t, pres, tmp):
    """Notes, two sections, a transition and entrance animations."""
    texts = ["Einleitung", "Ein Übergang", "Eine Animation", "Noch eine Animation"]
    for i, x in enumerate(texts, start=1):
        s = t.add_slide(pres, 1 if i == 1 else 2)
        title(s, x)
        if i > 1:
            set_paragraphs(body(s), [(1, f"Punkt {i}.1"), (1, f"Punkt {i}.2")])
    for i in (1, 2):
        nb = None
        page = pres.Slides(i).NotesPage
        for k in range(1, page.Shapes.Placeholders.Count + 1):
            if int(page.Shapes.Placeholders(k).PlaceholderFormat.Type) == 2:
                nb = page.Shapes.Placeholders(k)
        nb.TextFrame.TextRange.Text = f"Notizen zur Folie {i}\rZweite Zeile der Notizen"
    sp = pres.SectionProperties
    sp.AddBeforeSlide(1, "Einleitung")
    sp.AddBeforeSlide(3, "Hauptteil")
    tr = pres.Slides(2).SlideShowTransition
    tr.EntryEffect = 3849  # ppEffectFadeSmoothly (Verblassen)
    tr.Duration = 0.7
    seq3 = pres.Slides(3).TimeLine.MainSequence
    seq3.AddEffect(body(pres.Slides(3)), 10, 0, 1)  # msoAnimEffectFade, on click
    seq4 = pres.Slides(4).TimeLine.MainSequence
    seq4.AddEffect(body(pres.Slides(4)), 2, 0, 1)  # msoAnimEffectFly (Hereinfliegen), on click


def chart(t, pres, tmp):
    """Opt-in: AddChart2 opens Excel's data window visibly (recorded 2026-11-06), so this is not built by default."""
    s = t.add_slide(pres, 6)
    title(s, "Diagramm")
    show(t, pres, s)
    ch = s.Shapes.AddChart2(-1, 51, 66, 140, 828, 360)  # xlColumnClustered
    time.sleep(2)
    ch.Chart.ChartData.Workbook.Close()


def texteffects(t, pres, tmp):
    """Texteffekte > Transformieren (all MsoWarpFormat values PowerPoint takes, four per slide) and a text
    Farbverlauf (D-076). Which warp each value gives is read from the saved XML by the render check."""
    k = 0
    while k < 40:
        s = t.add_slide(pres, 7)
        for j in range(4):
            x, y = 40 + (j % 2) * 460, 40 + (j // 2) * 250
            b = s.Shapes.AddTextbox(1, x, y, 400, 200)
            b.TextFrame2.AutoSize = 0
            b.TextFrame2.WordWrap = 0
            b.Width, b.Height = 400, 200
            b.TextFrame.TextRange.Text = "WordArt Text"
            f = b.TextFrame.TextRange.Font
            f.Size = 40
            f.Bold = -1
            f.Color.RGB = rgb(192, 0, 0)
            try:
                b.TextFrame2.WarpFormat = k
            except Exception:  # noqa: BLE001
                b.TextFrame.TextRange.Text = f"(kein {k})"
            k += 1
    s = t.add_slide(pres, 7)
    g = s.Shapes.AddTextbox(1, 60, 60, 800, 120)
    g.TextFrame.TextRange.Text = "Farbverlauf im Text"
    g.TextFrame.TextRange.Font.Size = 60
    g.TextFrame.TextRange.Font.Bold = -1
    fill = g.TextFrame2.TextRange.Font.Fill
    fill.TwoColorGradient(1, 1)  # msoGradientHorizontal, variant 1
    fill.ForeColor.RGB = rgb(0, 112, 192)
    fill.BackColor.RGB = rgb(255, 192, 0)
    h = s.Shapes.AddTextbox(1, 60, 260, 800, 120)
    h.TextFrame.TextRange.Text = "Vertikaler Verlauf"
    h.TextFrame.TextRange.Font.Size = 60
    fill = h.TextFrame2.TextRange.Font.Fill
    fill.TwoColorGradient(2, 1)  # msoGradientVertical
    fill.ForeColor.RGB = rgb(112, 48, 160)
    fill.BackColor.RGB = rgb(0, 176, 80)


FIXTURES = {
    "blank": (blank, False), "layouts": (layouts, False), "bullets": (bullets, False), "autofit": (autofit, True),
    "textbox": (textbox, False), "shapes": (shapes, False), "group": (group, False), "picture": (picture, False),
    "table": (table, False), "equation": (equation, True), "show": (show_fixture, False), "texteffects": (texteffects, False),
}
OPT_IN = {"chart": (chart, True)}


# ----- saving and checking ----------------------------------------------------------------------

def anonymize(path: Path) -> str:
    """PowerPoint writes the signed-in Office user as author and last author whatever the
    document properties say: set both to the neutral name (the only change to PowerPoint's file).
    Returns the name that was there (to check the rest of the file for it)."""
    with zipfile.ZipFile(path) as z:
        items = [(info, z.read(info)) for info in z.infolist()]
    found = ""
    out = []
    for info, data in items:
        if info.filename == "docProps/core.xml":
            text = data.decode("utf-8")
            for tag in ("dc:creator", "cp:lastModifiedBy"):
                m = re.search(f"<{tag}>([^<]*)</{tag}>", text)
                if m and m.group(1) != NEUTRAL:
                    found = found or m.group(1)
                text = re.sub(f"<{tag}>[^<]*</{tag}>", f"<{tag}>{NEUTRAL}</{tag}>", text)
            data = text.encode("utf-8")
        out.append((info, data))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for info, data in out:
            z.writestr(info, data)
    return found


def check_private(path: Path, users: set[str]) -> None:
    """No user names (Windows, Office, the signed-in identity) or private paths in a fixture;
    compared in lower case. PNGs are compressed and not searched."""
    texts = ["c:\\users", *private_names(), *(u.lower() for u in users)]
    needles = [t.encode(enc) for t in texts if t for enc in ("utf-8", "utf-16-le")]
    if path.suffix == ".pptx":
        with zipfile.ZipFile(path) as z:
            parts = [(n, z.read(n)) for n in z.namelist()]
    else:
        parts = [(path.name, path.read_bytes())]
    for n, data in parts:
        low = data.lower()
        for needle in needles:
            if needle in low:
                raise SystemExit(f"{path.name}: private text in {n}")


def build(t: PowerPointTarget, prefix: str, theme, name: str, fn, windowed: bool, tmp: Path, users: set[str]) -> None:
    pres = t.new_presentation(with_window=windowed, theme=theme, name=name)
    fn(t, pres, tmp)
    if windowed:
        # Autofit and equations settle while PowerPoint is idle.
        show(t, pres, pres.Slides(1))
        time.sleep(1.0)
    stem = f"{prefix}{name}"
    t.set_properties(pres, NEUTRAL)
    path = OUT / f"{stem}.pptx"
    t.save_copy(pres, path)
    old = anonymize(path)
    if old:
        users.add(old)
    PNG.mkdir(parents=True, exist_ok=True)
    for old_png in PNG.glob(f"{stem}-*.png"):
        old_png.unlink()
    for i in range(1, pres.Slides.Count + 1):
        t.export_png(pres, i, PNG / f"{stem}-{i}.png", 1280)
    info = {"built_by": f"PowerPoint {t.app.Version} build {t.app.Build}", "date": datetime.date.today().isoformat(),
            "design_file": Path(theme).name if theme else "(PowerPoint's default)", **t.presentation_info(pres)}
    (OUT / f"{stem}.json").write_text(json.dumps(info, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    pres.Saved = -1
    pres.Close()
    t.own.remove(pres)
    print("saved", stem, path.stat().st_size, "slides", len(info["slides"]), flush=True)


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    OUT.mkdir(parents=True, exist_ok=True)
    only = [a[2:] for a in sys.argv[1:] if a.startswith("--")]
    table_ = {**FIXTURES, **OPT_IN}
    names = only or list(FIXTURES)
    users: set[str] = set()
    try:
        t = PowerPointTarget(visible=False)
    except OwnerPowerPointRunning as e:
        raise SystemExit(f"skipped: {e}")
    try:
        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            for prefix, theme in VARIANTS.items():
                for name in names:
                    fn, windowed = table_[name]
                    try:
                        build(t, prefix, theme, name, fn, windowed, tmp, users)
                    except Exception as e:  # noqa: BLE001
                        print("FAILED", prefix + name, str(e)[:200], flush=True)
    finally:
        print(t.close())
    for f in sorted(OUT.glob("pptx-*.pptx")) + sorted(OUT.glob("pptx-*.json")):
        check_private(f, users)
    for f in sorted(PNG.glob("*.png")):
        # PowerPoint's PNG export carries no text chunks; make sure.
        if any(tag in f.read_bytes() for tag in (b"tEXt", b"iTXt", b"zTXt")):
            raise SystemExit(f"{f.name}: text chunk in the PNG")
    print("no private text in the fixtures")


if __name__ == "__main__":
    main()
