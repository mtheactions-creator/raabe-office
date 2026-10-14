"""Probe: Word's arrowheads (a:tailEnd) — their size for line widths and sm/med/lg, from
Word's PDF export. Invisible, own Word instance.

    .venv/Scripts/python.exe probe_word_arrows.py
"""
from __future__ import annotations

import json
import math

import check_shapes as cs
from probe_word_shapes import outline

WIDTHS = [0.25, 0.5, 0.75, 1, 1.5, 2, 3, 6]
SIZES = ["sm", "med", "lg"]
TYPES = ["triangle", "stealth", "diamond", "oval", "arrow"]


def main() -> None:
    cases = []
    for t in TYPES:
        for lw in WIDTHS:
            for w in SIZES:
                for ln in SIZES:
                    if t != "triangle" and lw not in (0.75, 3):
                        continue
                    cases.append((t, lw, w, ln))
    spec = []
    for i, (t, lw, w, ln) in enumerate(cases):
        page, k = divmod(i, 40)
        r, c = divmod(k, 4)
        spec.append({"x": 40 + c * 140, "y": 40 + r * 19, "page": page, "t": t, "lw": lw, "w": w, "len": ln})
    pages: dict[int, list[str]] = {}
    for i, s in enumerate(spec):
        cx = 100 * 12700
        pages.setdefault(s["page"], []).append(
            f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="{i + 1}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="page"><wp:posOffset>{s["x"] * 12700}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>{s["y"] * 12700}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{cx}" cy="0"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="{i + 1}" name="L{i}"/><wp:cNvGraphicFramePr/>'
            f'<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvCnPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="0"/></a:xfrm>'
            f'<a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="{int(s["lw"] * 12700)}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:tailEnd type="{s["t"]}" w="{s["w"]}" len="{s["len"]}"/></a:ln></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')
    # Reuse check_shapes' document writer by handing it ready-made runs.
    path = cs.TMP / "arrows.docx"
    body = "".join(f"<w:p>{''.join(v)}{'<w:r><w:br w:type=\"page\"/></w:r>' if p < max(pages) else ''}</w:p>" for p, v in sorted(pages.items()))
    import zipfile
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
           'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
           'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:body>' + body + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>')
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        z.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr("word/document.xml", doc)
    out = cs.TMP / "arrows.json"
    outline(path, out)
    pdf = json.loads(out.read_text(encoding="utf-8"))
    res = []
    for s in spec:
        x0, y0 = s["x"], s["y"]
        pts = []
        for p in pdf[s["page"]]["paths"]:
            if p["paint"] == "stroke" and s["t"] != "arrow":
                continue
            for seg in p["d"]:
                for j in range(1, len(seg) - 1, 2):
                    px, py = seg[j], seg[j + 1]
                    if x0 + 60 <= px <= x0 + 115 and y0 - 9 <= py <= y0 + 9:
                        pts.append((px, py))
        if not pts:
            res.append({**s, "length": None})
            continue
        xs = [p[0] for p in pts]
        ys = [p[1] for p in pts]
        res.append({**s, "length": round(max(xs) - min(xs), 2), "width": round(max(ys) - min(ys), 2), "tip": round(max(xs) - (x0 + 100), 2)})
    for r in res:
        print(r["t"], r["lw"], r["w"], r["len"], "->", r.get("length"), r.get("width"), r.get("tip"))
    (cs.HERE / "results" / "2026-10-14-word-arrows.json").write_text(json.dumps(res, indent=0), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
