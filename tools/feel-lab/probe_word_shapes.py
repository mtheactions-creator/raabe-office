"""Probe: Word's shapes (Formen), invisible. Own Word instance (DispatchEx), closes only its
own documents.

    .venv/Scripts/python.exe probe_word_shapes.py names     # every AutoShape type: German name, prst, adjustments
    .venv/Scripts/python.exe probe_word_shapes.py defaults  # XML of a default shape and text box (2013 theme)
    .venv/Scripts/python.exe probe_word_shapes.py outline <docx> <json>
        # Word's PDF export of a document: the outline of every drawn path, per page

Names go to results/<date>-word-shapes.json, Word's XML (defaults) to captures/ (local only).
"""
from __future__ import annotations

import datetime
import json
import re
import sys
import tempfile
from pathlib import Path

import pythoncom
import win32com.client

from feellab import pdftext

HERE = Path(__file__).parent
RESULTS = HERE / "results"
THEMES = Path(r"C:\Program Files\Microsoft Office\root\Document Themes 16")


def word():
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    return app


def drawings(xml: str) -> list[str]:
    return re.findall(r"<w:drawing>.*?</w:drawing>", xml, re.S)


def names() -> dict:
    app = word()
    out: dict = {}
    try:
        d = app.Documents.Add()
        for t in range(1, 200):
            try:
                s = d.Shapes.AddShape(t, 50, 50, 100, 80)
            except Exception:  # noqa: BLE001
                continue
            try:
                adj = [round(s.Adjustments.Item(i + 1), 5) for i in range(s.Adjustments.Count)]
            except Exception:  # noqa: BLE001
                adj = []
            out[str(t)] = {"name": s.Name, "auto": s.AutoShapeType, "adj": adj}
        xml = d.Content.WordOpenXML
        # Shapes added at the same anchor come out in reverse order: match them by name.
        by_name = {}
        for dr in drawings(xml):
            n = re.search(r'<wp:docPr id="\d+" name="([^"]+)"', dr)
            if n:
                by_name[n.group(1)] = dr
        for k in out:
            dr = by_name.get(out[k]["name"], "")
            m = re.search(r'<a:prstGeom prst="([^"]+)"', dr)
            out[k]["prst"] = m.group(1) if m else None
            av = re.search(r"<a:avLst>(.*?)</a:avLst>", dr, re.S)
            out[k]["av"] = re.findall(r'<a:gd name="([^"]+)" fmla="([^"]+)"', av.group(1)) if av else []
        d.Close(0)
    finally:
        app.Quit(0)
    return out


def defaults() -> dict:
    app = word()
    out: dict = {}
    try:
        for theme in ("Office 2013 - 2022 Theme.thmx", None):
            d = app.Documents.Add()
            if theme:
                d.ApplyDocumentTheme(str(THEMES / theme))
            d.Shapes.AddShape(1, 100, 100, 72, 72)
            d.Shapes.AddTextbox(1, 100, 250, 144, 72)
            d.Shapes(2).TextFrame.TextRange.Text = "Text"
            d.Shapes.AddLine(100, 400, 200, 450)
            xml = d.Content.WordOpenXML
            key = theme or "default"
            out[key] = {"drawings": drawings(xml)}
            m = re.search(r"<a:theme .*?</a:theme>", xml, re.S)
            out[key]["theme"] = m.group(0) if m else None
            d.Close(0)
    finally:
        app.Quit(0)
    return out


def outline(docx: Path, dest: Path) -> None:
    app = word()
    pdf = Path(tempfile.gettempdir()) / (docx.stem + ".pdf")
    try:
        d = app.Documents.Open(str(docx), ReadOnly=True)
        d.ExportAsFixedFormat(str(pdf), 17)
        d.Close(0)
    finally:
        app.Quit(0)
    p = pdftext.Pdf(pdf.read_bytes())
    pages = []
    for num in p.pages():
        w, h = p.page_box(num)
        pages.append({"w": w, "h": h, "paths": pdftext.paths(p.contents(num), h)})
    dest.write_text(json.dumps(pages), encoding="utf-8", newline="\n")


def main() -> None:
    mode = sys.argv[1] if len(sys.argv) > 1 else "names"
    if mode == "outline":
        outline(Path(sys.argv[2]), Path(sys.argv[3]))
        return
    # Word's own XML (theme, drawings) stays local; the names table goes to results/.
    path = RESULTS / f"{datetime.date.today()}-word-shapes.json" if mode == "names" else HERE / "captures" / "word-shapes-defaults.json"
    data = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
    data[mode] = names() if mode == "names" else defaults()
    path.write_text(json.dumps(data, indent=1, ensure_ascii=False), encoding="utf-8", newline="\n")
    print(path)


if __name__ == "__main__":
    main()
