"""How Excel 365 stores ink (Zeichnen, D-093), recorded without real input. Excel's object model cannot draw ink
(Shapes has no AddInk), so this probe writes a workbook with an ink content part in the form Office documents use
(xdr:contentPart in mc:AlternateContent, an InkML part with the relationship type customXml), lets our own
invisible Excel open it, reads what Excel makes of it through COM (shape type 23 = msoInk, name, position,
size) and saves it again: the XML Excel writes is its own form of ink. Variants: a pen stroke, a red pen with
time stamps, a highlighter, two strokes in one part.

    .venv/Scripts/python.exe probe_excel_ink.py

Writes results/<date>-excel-ink.json and the fixture packages/ooxml/test/fixtures/excel-ink.json (our input and
Excel's saved parts: drawing, ink parts, relationships, content types)."""

from __future__ import annotations

import datetime
import json
import re
import sys
import tempfile
import zipfile
from pathlib import Path

from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
ROOT = HERE.parents[1]
TMP = Path(tempfile.gettempdir())
FIXTURE = ROOT / "packages" / "ooxml" / "test" / "fixtures" / "excel-ink.json"

NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
INK_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"


def to_fixture(src: Path, name: str) -> None:
    """Excel's saved workbook as a test fixture: the save folder (x15ac:absPath) left out, the user's name
    replaced (anonymize_office.py)."""
    import shutil
    import subprocess
    dst = ROOT / "packages" / "ooxml" / "test" / "fixtures" / name
    shutil.copyfile(src, dst)
    with zipfile.ZipFile(dst) as z:
        items = [(info, z.read(info)) for info in z.infolist()]
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for info, data in items:
            if info.filename == "xl/workbook.xml":
                data = re.sub(rb"<mc:AlternateContent\b(?:(?!</mc:AlternateContent>)[\s\S])*?x15ac:absPath[\s\S]*?</mc:AlternateContent>", b"", data)
            z.writestr(info, data)
    subprocess.run([sys.executable, str(HERE / "anonymize_office.py"), str(dst)], check=True, capture_output=True)


def inkml(traces: list[list[tuple[float, float, int]]], color: str, width_cm: float, highlighter: bool = False) -> str:
    """InkML as Office writes it: X/Y in 1/1000 cm (himetric), T in ms."""
    brush = (f'<inkml:brush xml:id="br0"><inkml:brushProperty name="width" value="{width_cm}" units="cm"/>'
             f'<inkml:brushProperty name="height" value="{width_cm}" units="cm"/><inkml:brushProperty name="color" value="{color}"/>'
             + ('<inkml:brushProperty name="tip" value="rectangle"/><inkml:brushProperty name="rasterOp" value="maskPen"/>' if highlighter else "")
             + "</inkml:brush>")
    ctx = ('<inkml:context xml:id="ctx0"><inkml:inkSource xml:id="inkSrc0"><inkml:traceFormat>'
           '<inkml:channel name="X" type="integer" max="32767" units="cm"/><inkml:channel name="Y" type="integer" max="32767" units="cm"/>'
           '<inkml:channel name="T" type="integer" max="2.14748E9" units="dev"/></inkml:traceFormat><inkml:channelProperties>'
           '<inkml:channelProperty channel="X" name="resolution" value="1000" units="1/cm"/><inkml:channelProperty channel="Y" name="resolution" value="1000" units="1/cm"/>'
           '<inkml:channelProperty channel="T" name="resolution" value="1" units="1/dev"/></inkml:channelProperties></inkml:inkSource>'
           '<inkml:timestamp xml:id="ts0" timeString="2026-11-08T10:00:00.000"/></inkml:context>')
    body = "".join(f'<inkml:trace contextRef="#ctx0" brushRef="#br0">{", ".join(f"{round(x)} {round(y)} {t}" for x, y, t in tr)}</inkml:trace>' for tr in traces)
    return f'<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML"><inkml:definitions>{ctx}{brush}</inkml:definitions>{body}</inkml:ink>'


def anchor(i: int, rid: str, col0: int, row0: int, col1: int, row1: int, emu: tuple[int, int, int, int]) -> str:
    x, y, cx, cy = emu
    return (f'<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
            f'<mc:Choice xmlns:xdr14="http://schemas.microsoft.com/office/excel/2010/spreadsheetDrawing" Requires="xdr14">'
            f'<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>{col0}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row0}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
            f'<xdr:to><xdr:col>{col1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
            f'<xdr:contentPart xmlns:r="{NS_R}" r:id="{rid}"><xdr14:nvContentPartPr><xdr14:cNvPr id="{i + 1}" name="Freihand {i}"/><xdr14:cNvContentPartPr/></xdr14:nvContentPartPr>'
            f'<xdr14:nvPr macro=""/><xdr14:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></xdr14:xfrm></xdr:contentPart>'
            f'<xdr:clientData/></xdr:twoCellAnchor></mc:Choice><mc:Fallback/></mc:AlternateContent>')


CASES = [
    # name, traces (cm*1000, ms), colour, width cm, highlighter, cells
    ("pen-black", [[(1000 + i * 100, 1000 + (i % 5) * 80, i * 8) for i in range(40)]], "#000000", 0.035, False),
    ("pen-red", [[(2000 + i * 60, 3000 + i * 40, i * 10) for i in range(30)]], "#FF0000", 0.035, False),
    ("highlighter", [[(1500 + i * 120, 5000, i * 12) for i in range(30)]], "#FFFF00", 0.6, True),
    ("two-strokes", [[(4000 + i * 50, 1000 + i * 50, i * 9) for i in range(20)], [(4000 + i * 50, 2000 - i * 50, 400 + i * 9) for i in range(20)]], "#0000FF", 0.05, False),
]


def build(base: Path, out: Path) -> list[dict]:
    """Our workbook: Excel's own blank file plus one drawing with an ink part per case."""
    sent = []
    with zipfile.ZipFile(base) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        names = zin.namelist()
        anchors, rels = [], []
        for i, (name, traces, color, width, hl) in enumerate(CASES, 1):
            xs = [p[0] for t in traces for p in t]
            ys = [p[1] for t in traces for p in t]
            # himetric (1/1000 cm) to EMU: 1 cm = 360000 EMU.
            emu = (int(min(xs) * 360), int(min(ys) * 360), int((max(xs) - min(xs)) * 360) or 3600, int((max(ys) - min(ys)) * 360) or 3600)
            col0, row0 = emu[0] // (609600), emu[1] // (190500)
            col1, row1 = (emu[0] + emu[2]) // 609600 + 1, (emu[1] + emu[3]) // 190500 + 1
            anchors.append(anchor(i, f"rId{i}", col0, row0, col1, row1, emu))
            rels.append(f'<Relationship Id="rId{i}" Type="{INK_REL}" Target="../ink/ink{i}.xml"/>')
            xml = inkml(traces, color, width, hl)
            z.writestr(f"xl/ink/ink{i}.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml)
            sent.append({"name": name, "emu": emu, "inkml": xml})
        drawing = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                   f'<xdr:wsDr xmlns:xdr="{NS_XDR}" xmlns:a="{NS_A}">' + "".join(anchors) + "</xdr:wsDr>")
        z.writestr("xl/drawings/drawing1.xml", drawing)
        z.writestr("xl/drawings/_rels/drawing1.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + "".join(rels) + "</Relationships>")
        for n in names:
            data = zin.read(n)
            if n == "[Content_Types].xml":
                s = data.decode("utf-8").replace("</Types>", '<Default Extension="xml_ink" ContentType="application/inkml+xml"/>'
                                                 '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
                                                 + "".join(f'<Override PartName="/xl/ink/ink{i}.xml" ContentType="application/inkml+xml"/>' for i in range(1, len(CASES) + 1))
                                                 + "</Types>").replace('<Default Extension="xml_ink" ContentType="application/inkml+xml"/>', "")
                data = s.encode("utf-8")
            elif n == "xl/worksheets/sheet1.xml":
                s = data.decode("utf-8")
                s = s.replace("</worksheet>", '<drawing r:id="rIdInk"/></worksheet>')
                if "xmlns:r=" not in s[:600]:
                    s = s.replace("<worksheet ", f'<worksheet xmlns:r="{NS_R}" ', 1)
                data = s.encode("utf-8")
            elif n == "xl/worksheets/_rels/sheet1.xml.rels":
                continue
            z.writestr(n, data)
        z.writestr("xl/worksheets/_rels/sheet1.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rIdInk" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>')
    return sent


def read_parts(path: Path) -> dict:
    out = {}
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if re.match(r"xl/(drawings|ink)/|xl/worksheets/_rels/|\[Content_Types\]", n):
                out[n] = z.read(n).decode("utf-8", "replace")
            elif n.startswith("xl/media/"):
                out[n] = f"<binary {len(z.read(n))} bytes>"
    return out


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    res: dict = {"date": datetime.date.today().isoformat()}
    base, ours, back = TMP / "raabe-ink-base.xlsx", TMP / "raabe-ink-ours.xlsx", TMP / "raabe-ink-excel.xlsx"
    x = ExcelTarget(visible=False)
    try:
        x.wb.Worksheets(1).Range("A1").Value2 = "Ink"
        x.wb.SaveAs(str(base), 51)
        x.wb.Close(False)
        res["sent"] = build(base, ours)
        x.wb = x.app.Workbooks.Add()
        wb = x.app.Workbooks.Open(str(ours))
        ws = wb.Worksheets(1)
        res["com"] = []
        for i in range(1, ws.Shapes.Count + 1):
            s = ws.Shapes(i)
            res["com"].append({"name": s.Name, "type": s.Type, "left": s.Left, "top": s.Top, "width": s.Width, "height": s.Height,
                               "tl": s.TopLeftCell.Address.replace("$", ""), "br": s.BottomRightCell.Address.replace("$", ""), "placement": s.Placement})
        print(res["com"], flush=True)
        # Excel saves the ink in its own form; a moved stroke shows how the position is kept.
        if ws.Shapes.Count:
            ws.Shapes(1).IncrementLeft(30)
        wb.SaveAs(str(back), 51)
        wb.Close(False)
        res["excel_saved"] = read_parts(back)
        # Excel reads its own file again.
        wb = x.app.Workbooks.Open(str(back))
        res["reopened"] = [{"name": wb.Worksheets(1).Shapes(i).Name, "type": wb.Worksheets(1).Shapes(i).Type} for i in range(1, wb.Worksheets(1).Shapes.Count + 1)]
        wb.Close(False)
    finally:
        x.close()
    to_fixture(TMP / "raabe-ink-excel.xlsx", "excel-ink.xlsx")
    out = HERE / "results" / f"{res['date']}-excel-ink.json"
    out.write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    FIXTURE.write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print("written", out.name, len(res["com"]), "shapes")


if __name__ == "__main__":
    main()
