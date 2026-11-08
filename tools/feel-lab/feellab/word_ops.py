"""Word's own counterpart of RaabeWrite options for the by-hand audit's comparison (D-133). Keys are the sweep's
option paths ("Tab > Group / Label [kind] > Entry"), matched exactly; values apply the same option in Word to the base
document with the same selection (Word's object model, or the ribbon control itself through ExecuteMso).

Where RaabeWrite's option needs a value (a typed field), the sweep types the value given in sweep_write.py."""
from __future__ import annotations

CM = 28.3465


def _mso(name):
    def f(app, d, s):
        app.CommandBars.ExecuteMso(name)
    return f


def _font(attr, value):
    def f(app, d, s):
        setattr(s.Font, attr, value)
    return f


def _para(attr, value):
    def f(app, d, s):
        setattr(s.ParagraphFormat, attr, value)
    return f


def _case(v):
    def f(app, d, s):
        s.Range.Case = v
    return f


def _style(v):
    def f(app, d, s):
        s.Style = d.Styles(v)
    return f


def _spacing(rule, value=None):
    def f(app, d, s):
        pf = s.ParagraphFormat
        pf.LineSpacingRule = rule
        if value is not None:
            pf.LineSpacing = value
    return f


def _page(**kw):
    def f(app, d, s):
        ps = d.Sections(1).PageSetup
        for k, v in kw.items():
            setattr(ps, k, v)
    return f


def _break(t):
    def f(app, d, s):
        s.InsertBreak(t)
    return f


def _lines(active, mode=None):
    def f(app, d, s):
        ln = d.Sections(1).PageSetup.LineNumbering
        ln.Active = active
        if mode is not None:
            ln.RestartMode = mode
    return f


def _cols(n):
    def f(app, d, s):
        # Word applies Spalten to selected text only (it adds section breaks around it).
        s.PageSetup.TextColumns.SetCount(n)
    return f


def _call(path):
    def f(app, d, s):
        obj = s
        for part in path.split(".")[:-1]:
            obj = getattr(obj, part)
        getattr(obj, path.split(".")[-1])()
    return f


def _hyph(on):
    def f(app, d, s):
        d.AutoHyphenation = on
    return f


def _footnote(app, d, s):
    d.Footnotes.Add(s.Range)


def _endnote(app, d, s):
    d.Endnotes.Add(s.Range)


def _comment(app, d, s):
    d.Comments.Add(s.Range, "")


def _track(app, d, s):
    d.TrackRevisions = not d.TrackRevisions


def _table11(app, d, s):
    t = d.Tables.Add(s.Range, 1, 1)
    t.Style = "Tabellenraster"


def _blank_page(app, d, s):
    app.CommandBars.ExecuteMso("BlankPageInsert")


S = "Start > "
OPS = {
    # Start: Schriftart
    S + "Schriftart / Fett [button]": _font("Bold", True),
    S + "Schriftart / Kursiv [button]": _font("Italic", True),
    S + "Schriftart / Unterstreichen [split-main]": _font("Underline", 1),
    S + "Schriftart / Durchstreichen [button]": _font("StrikeThrough", True),
    S + "Schriftart / Tiefgestellt [button]": _font("Subscript", True),
    S + "Schriftart / Hochgestellt [button]": _font("Superscript", True),
    S + "Schriftart / Schriftart vergrößern [button]": _call("Font.Grow"),
    S + "Schriftart / Schriftart verkleinern [button]": _call("Font.Shrink"),
    S + "Schriftart / Alle Formatierungen löschen [button]": _call("ClearFormatting"),
    S + "Schriftart / Texthervorhebungsfarbe [color-main]": lambda app, d, s: setattr(s.Range, "HighlightColorIndex", 7),
    S + "Schriftart / Schriftfarbe [color-main]": _font("Color", 255),
    S + "Schriftart / Schriftgrad [combo-input] = 14": _font("Size", 14),
    S + "Schriftart / Schriftart [combo-input] = Arial": _font("Name", "Arial"),
    S + "Schriftart / Weitere Optionen [split-arrow] > Doppelt unterstreichen": _font("Underline", 3),
    S + "Schriftart / Weitere Optionen [split-arrow] > Unterstreichen": _font("Underline", 1),
    S + "Schriftart / Weitere Optionen [split-arrow] > Dick unterstreichen": _font("Underline", 6),
    S + "Schriftart / Weitere Optionen [split-arrow] > Gepunktet unterstreichen": _font("Underline", 4),
    S + "Schriftart / Weitere Optionen [split-arrow] > Gestrichelt unterstreichen": _font("Underline", 7),
    S + "Schriftart / Weitere Optionen [split-arrow] > Wellenlinie unterstreichen": _font("Underline", 11),
    S + "Schriftart / Groß-/Kleinschreibung [dropdown] > Ersten Buchstaben im Satz großschreiben.": _case(4),
    S + "Schriftart / Groß-/Kleinschreibung [dropdown] > kleinbuchstaben": _case(0),
    S + "Schriftart / Groß-/Kleinschreibung [dropdown] > GROSSBUCHSTABEN": _case(1),
    S + "Schriftart / Groß-/Kleinschreibung [dropdown] > Ersten Buchstaben Jedes Wortes Großschreiben": _case(2),
    S + "Schriftart / Groß-/Kleinschreibung [dropdown] > gROSS-/kLEINSCHREIBUNG umkehren": _case(5),
    S + "Schriftart / Texteffekte und Typografie [dropdown] > Ligaturen > Keine": _font("Ligatures", 0),
    S + "Schriftart / Texteffekte und Typografie [dropdown] > Ligaturen > Nur Standard": _font("Ligatures", 1),
    S + "Schriftart / Texteffekte und Typografie [dropdown] > Ligaturen > Historisch und frei": _font("Ligatures", 12),
    S + "Schriftart / Texteffekte und Typografie [dropdown] > Ligaturen > Alle Ligaturen": _font("Ligatures", 15),
    S + "Schriftart / Texteffekte und Typografie [dropdown] > Zahlenformatvorlagen > Proportionale Lineatur": lambda app, d, s: (setattr(s.Font, "NumberForm", 1), setattr(s.Font, "NumberSpacing", 1)),
    S + "Schriftart / Texteffekte und Typografie [dropdown] > Stil-Sets > Stil-Set 1": _font("StylisticSet", 1),
    S + "Schriftart / Texteffekte und Typografie [dropdown] > Stil-Sets > Stil-Set 20": _font("StylisticSet", 20),
    S + "Schriftart / Weitere Optionen [color-arrow] > Gelb": lambda app, d, s: setattr(s.Range, "HighlightColorIndex", 7),
    S + "Schriftart / Weitere Optionen [color-arrow] > Hellgrün": lambda app, d, s: setattr(s.Range, "HighlightColorIndex", 4),
    S + "Schriftart / Weitere Optionen [color-arrow] > Türkis": lambda app, d, s: setattr(s.Range, "HighlightColorIndex", 3),
    S + "Schriftart / Weitere Optionen [color-arrow] > Rosa": lambda app, d, s: setattr(s.Range, "HighlightColorIndex", 5),
    # Start: Absatz
    S + "Absatz / Aufzählungszeichen [split-main]": _call("Range.ListFormat.ApplyBulletDefault"),
    S + "Absatz / Nummerierung [split-main]": _call("Range.ListFormat.ApplyNumberDefault"),
    S + "Absatz / Einzug vergrößern [button]": _call("Paragraphs.Indent"),
    S + "Absatz / Einzug verkleinern [button]": _call("Paragraphs.Outdent"),
    S + "Absatz / Zentriert [button]": _para("Alignment", 1),
    S + "Absatz / Rechtsbündig [button]": _para("Alignment", 2),
    S + "Absatz / Blocksatz [button]": _para("Alignment", 3),
    S + "Absatz / Zeilen- und Absatzabstand [dropdown] > 1,0": _spacing(0),
    S + "Absatz / Zeilen- und Absatzabstand [dropdown] > 1,15": _spacing(5, 13.8),
    S + "Absatz / Zeilen- und Absatzabstand [dropdown] > 1,5": _spacing(1),
    S + "Absatz / Zeilen- und Absatzabstand [dropdown] > 2,0": _spacing(2),
    S + "Absatz / Zeilen- und Absatzabstand [dropdown] > 2,5": _spacing(5, 30),
    S + "Absatz / Zeilen- und Absatzabstand [dropdown] > 3,0": _spacing(5, 36),
    S + "Absatz / Zeilen- und Absatzabstand [dropdown] > Abstand vor Absatz hinzufügen": _para("SpaceBefore", 12),
    S + "Absatz / Zeilen- und Absatzabstand [dropdown] > Abstand nach Absatz entfernen": _para("SpaceAfter", 0),
    # Start: Formatvorlagen (gallery in the ribbon)
    S + "Formatvorlagen / Kein Leerraum [gallery-item]": _style("Kein Leerraum"),
    S + "Formatvorlagen / Überschrift 1 [gallery-item]": _style(-2),
    S + "Formatvorlagen / Überschrift 2 [gallery-item]": _style(-3),
    S + "Formatvorlagen / Titel [gallery-item]": _style(-63),
    S + "Formatvorlagen / Untertitel [gallery-item]": _style(-75),
    # Einfügen
    "Einfügen > Seiten / Seitenumbruch [button]": _break(7),
    "Einfügen > Seiten / Leere Seite [button]": _blank_page,
    "Einfügen > Kommentare / Kommentar [button]": _comment,
    # Layout
    "Layout > Seite einrichten / Ausrichtung [dropdown] > Querformat": _page(Orientation=1),
    "Layout > Seite einrichten / Spalten [dropdown] > Zwei": _cols(2),
    "Layout > Seite einrichten / Spalten [dropdown] > Drei": _cols(3),
    "Layout > Seite einrichten / Umbrüche [dropdown] > Seite": _break(7),
    "Layout > Seite einrichten / Umbrüche [dropdown] > Spalte": _break(8),
    "Layout > Seite einrichten / Umbrüche [dropdown] > Textumbruch": _break(11),
    "Layout > Seite einrichten / Umbrüche [dropdown] > Nächste Seite": _break(2),
    "Layout > Seite einrichten / Umbrüche [dropdown] > Fortlaufend": _break(3),
    "Layout > Seite einrichten / Umbrüche [dropdown] > Gerade Seite": _break(4),
    "Layout > Seite einrichten / Umbrüche [dropdown] > Ungerade Seite": _break(5),
    "Layout > Seite einrichten / Zeilennummern [dropdown] > Fortlaufend": _lines(True, 2),
    "Layout > Seite einrichten / Zeilennummern [dropdown] > Jede Seite neu beginnen": _lines(True, 0),
    "Layout > Seite einrichten / Zeilennummern [dropdown] > Jeden Abschnitt neu beginnen": _lines(True, 1),
    "Layout > Seite einrichten / Silbentrennung [dropdown] > Automatisch": _hyph(True),
    "Layout > Seite einrichten / Seitenränder [dropdown] > Schmal": _page(TopMargin=1.27 * CM, BottomMargin=1.27 * CM, LeftMargin=1.27 * CM, RightMargin=1.27 * CM),
    "Layout > Absatz / Links: [spinner-input] = 1": _para("LeftIndent", 1 * CM),
    "Layout > Absatz / Rechts: [spinner-input] = 1": _para("RightIndent", 1 * CM),
    "Layout > Absatz / Vor: [spinner-input] = 1": _para("SpaceBefore", 1),
    "Layout > Absatz / Nach: [spinner-input] = 1": _para("SpaceAfter", 1),
    # Referenzen
    "Referenzen > Fußnoten / Fußnote einfügen [button]": _footnote,
    "Referenzen > Fußnoten / Endnote einfügen [button]": _endnote,
    # Überprüfen
    "Überprüfen > Kommentare / Neuer Kommentar [button]": _comment,
    "Überprüfen > Nachverfolgung / Änderungen nachverfolgen [split-main]": _track,
}



def _tbl(fn):
    def f(app, d, s):
        fn(s, s.Tables(1))
    return f


def _cellalign(v, h):
    def f(app, d, s):
        s.Cells.VerticalAlignment = v
        s.ParagraphFormat.Alignment = h
    return f


# Keys without the group ("Tab > Label [kind] > Entry"): sweep_write.run_word tries these when the full path is not listed.
OPS_NO_GROUP = {
    "Tabellenlayout > Zeile Oberhalb [button]": _call("InsertRowsAbove"),
    "Tabellenlayout > Zeile unterhalb einfügen [button]": _call("InsertRowsBelow"),
    "Tabellenlayout > Spalte Links [button]": _call("InsertColumns"),
    "Tabellenlayout > Spalte rechts einfügen [button]": _call("InsertColumnsRight"),
    "Tabellenlayout > Löschen [dropdown] > Spalten löschen": _call("Columns.Delete"),
    "Tabellenlayout > Löschen [dropdown] > Zeilen löschen": _call("Rows.Delete"),
    "Tabellenlayout > Löschen [dropdown] > Tabelle löschen": _tbl(lambda s, t: t.Delete()),
    "Tabellenlayout > Tabelle teilen [button]": _tbl(lambda s, t: t.Split(s.Rows(1))),
    "Tabellenlayout > AutoAnpassen [dropdown] > Automatisch an Inhalt anpassen": _tbl(lambda s, t: t.AutoFitBehavior(1)),
    "Tabellenlayout > AutoAnpassen [dropdown] > Automatisch an Fenster anpassen": _tbl(lambda s, t: t.AutoFitBehavior(2)),
    "Tabellenlayout > Oben links ausrichten [button]": _cellalign(0, 0),
    "Tabellenlayout > Oben zentriert ausrichten [button]": _cellalign(0, 1),
    "Tabellenlayout > Oben rechts ausrichten [button]": _cellalign(0, 2),
    "Tabellenlayout > Mitte links ausrichten [button]": _cellalign(1, 0),
    "Tabellenlayout > Zentriert ausrichten [button]": _cellalign(1, 1),
    "Tabellenlayout > Mitte rechts ausrichten [button]": _cellalign(1, 2),
    "Tabellenlayout > Unten links ausrichten [button]": _cellalign(3, 0),
    "Tabellenlayout > Unten Mitte ausrichten [button]": _cellalign(3, 1),
    "Tabellenlayout > Unten rechts ausrichten [button]": _cellalign(3, 2),
    "Tabellenentwurf > Kopfzeile [check]": _tbl(lambda s, t: setattr(t, "ApplyStyleHeadingRows", not t.ApplyStyleHeadingRows)),
    "Tabellenentwurf > Ergebniszeile [check]": _tbl(lambda s, t: setattr(t, "ApplyStyleLastRow", not t.ApplyStyleLastRow)),
    "Tabellenentwurf > Gebänderte Zeilen [check]": _tbl(lambda s, t: setattr(t, "ApplyStyleRowBands", not t.ApplyStyleRowBands)),
    "Tabellenentwurf > Erste Spalte [check]": _tbl(lambda s, t: setattr(t, "ApplyStyleFirstColumn", not t.ApplyStyleFirstColumn)),
    "Tabellenentwurf > Letzte Spalte [check]": _tbl(lambda s, t: setattr(t, "ApplyStyleLastColumn", not t.ApplyStyleLastColumn)),
    "Tabellenentwurf > Gebänderte Spalten [check]": _tbl(lambda s, t: setattr(t, "ApplyStyleColumnBands", not t.ApplyStyleColumnBands)),
    "Layout > Ausrichtung [dropdown] > Querformat": _page(Orientation=1),
    "Layout > Seitenränder [dropdown] > Schmal": _page(TopMargin=1.27 * CM, BottomMargin=1.27 * CM, LeftMargin=1.27 * CM, RightMargin=1.27 * CM),
    "Layout > Seitenränder [dropdown] > Breit": _page(TopMargin=2.54 * CM, BottomMargin=2.54 * CM, LeftMargin=5.08 * CM, RightMargin=5.08 * CM),
    "Layout > Seitenränder [dropdown] > Mittel": _page(TopMargin=2.54 * CM, BottomMargin=2.54 * CM, LeftMargin=1.91 * CM, RightMargin=1.91 * CM),
}
