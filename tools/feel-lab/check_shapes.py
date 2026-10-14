"""Shape geometry against Word (D-038, package C): every preset of
packages/text-engine/src/geometry-presets.ts is drawn by Word (a generated document, exported
as PDF by an own invisible Word instance) and by our evaluator; the outlines are compared
point by point. Three variants: wide (80 x 50 pt), tall (50 x 80 pt) and wide with adjust
values moved off their defaults. Invisible; does not touch mouse, keyboard or screen.

    .venv/Scripts/python.exe check_shapes.py [prst ...]

Writes results/<date>-shapes.json (deviation per preset and variant, in points).
"""
from __future__ import annotations

import datetime
import json
import math
import re
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from probe_word_shapes import outline

HERE = Path(__file__).parent
ROOT = HERE.parent.parent
PRESETS = ROOT / "packages" / "text-engine" / "src" / "geometry-presets.ts"
TMP = Path(tempfile.gettempdir()) / "raabe-shapes"
TMP.mkdir(exist_ok=True)
EMU = 12700
COLS, ROWS, CW, CH, X0, Y0 = 4, 5, 140, 160, 20, 20


BASE = (0xCC, 0xE0, 0xFF)


def shade(mode: str) -> str:
    """The fill colour Word uses for a path's fill mode (measured on the bevel)."""
    f = {"darken": lambda c: c * 0.6, "darkenLess": lambda c: c * 0.8,
         "lighten": lambda c: c + (255 - c) * 0.4, "lightenLess": lambda c: c + (255 - c) * 0.2}.get(mode, lambda c: c)
    return "".join(f"{round(f(c)):02X}" for c in BASE)


def close(a: str, b: str) -> bool:
    return all(abs(int(a[i:i + 2], 16) - int(b[i:i + 2], 16)) <= 2 for i in (0, 2, 4))


def presets() -> list[str]:
    src = PRESETS.read_text(encoding="utf-8")
    body = src[src.index("PRESET_DEFS"):]
    return re.findall(r"^  (\w+): ", body, re.M)


def adjusted(defaults: dict[str, float]) -> dict[str, float]:
    """Adjust values off their defaults: angles turned by 45°, the rest scaled."""
    out = {}
    for k, v in defaults.items():
        if k in ("hf", "vf"):  # Word does not open an avLst with adj but without these
            out[k] = v
            continue
        if abs(v) >= 1_000_000:
            out[k] = (v + 2_700_000) % 21_600_000
        elif v == 0:
            out[k] = 10000
        else:
            out[k] = round(v * 0.6)
    return out


def layout(names: list[str], variant: str, defaults: dict[str, dict[str, float]]) -> list[dict]:
    w, h = (50, 80) if variant == "tall" else (80, 50)
    spec = []
    for i, n in enumerate(names):
        page, k = divmod(i, COLS * ROWS)
        r, c = divmod(k, COLS)
        adj = adjusted(defaults.get(n, {})) if variant == "adjusted" else {}
        spec.append({"prst": n, "page": page, "x": X0 + c * CW + 30, "y": Y0 + r * CH + 40, "w": w, "h": h, "adj": adj,
                     "cell": [X0 + c * CW, Y0 + r * CH, CW, CH]})
    return spec


def docx(spec: list[dict], path: Path) -> None:
    pages: dict[int, list[str]] = {}
    for i, s in enumerate(spec):
        av = "".join(f'<a:gd name="{k}" fmla="val {int(v)}"/>' for k, v in s["adj"].items())
        cx, cy = s["w"] * EMU, s["h"] * EMU
        pages.setdefault(s["page"], []).append(
            f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="{i + 1}" behindDoc="0" '
            f'locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="page"><wp:posOffset>{s["x"] * EMU}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="page"><wp:posOffset>{s["y"] * EMU}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
            f'<wp:docPr id="{i + 1}" name="Shape {i + 1}"/><wp:cNvGraphicFramePr/>'
            f'<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/>'
            f'<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="{s["prst"]}"><a:avLst>{av}</a:avLst></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="CCE0FF"/></a:solidFill><a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>'
            f'<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')
    paras = []
    n = max(pages) + 1 if pages else 1
    for p in range(n):
        brk = '<w:r><w:br w:type="page"/></w:r>' if p < n - 1 else ""
        paras.append(f"<w:p>{''.join(pages.get(p, []))}{brk}</w:p>")
    body = "".join(paras) + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="567" w:right="567" w:bottom="567" w:left="567" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>'
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
           'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
           'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
           'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
           f"<w:body>{body}</w:body></w:document>")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        z.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr("word/document.xml", doc)


def flatten(d: list, step: float = 0.5) -> list[tuple[float, float]]:
    """Dense points along a path of [cmd, ...] segments (from the PDF or ours)."""
    out: list[tuple[float, float]] = []
    cur = start = (0.0, 0.0)

    def line(a, b):
        n = max(1, int(math.dist(a, b) / step))
        for i in range(1, n + 1):
            t = i / n
            out.append((a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t))

    for seg in d:
        k = seg[0]
        if k == "M":
            cur = start = (seg[1], seg[2])
            out.append(cur)
        elif k == "L":
            line(cur, (seg[1], seg[2]))
            cur = (seg[1], seg[2])
        elif k == "Z":
            line(cur, start)
            cur = start
        elif k == "C":
            p0, p1, p2, p3 = cur, (seg[1], seg[2]), (seg[3], seg[4]), (seg[5], seg[6])
            n = max(4, int((math.dist(p0, p1) + math.dist(p1, p2) + math.dist(p2, p3)) / step))
            for i in range(1, n + 1):
                t = i / n
                u = 1 - t
                out.append((u ** 3 * p0[0] + 3 * u * u * t * p1[0] + 3 * u * t * t * p2[0] + t ** 3 * p3[0],
                            u ** 3 * p0[1] + 3 * u * u * t * p1[1] + 3 * u * t * t * p2[1] + t ** 3 * p3[1]))
            cur = p3
        elif k == "Q":
            p0, p1, p2 = cur, (seg[1], seg[2]), (seg[3], seg[4])
            n = max(4, int((math.dist(p0, p1) + math.dist(p1, p2)) / step))
            for i in range(1, n + 1):
                t = i / n
                u = 1 - t
                out.append((u * u * p0[0] + 2 * u * t * p1[0] + t * t * p2[0], u * u * p0[1] + 2 * u * t * p1[1] + t * t * p2[1]))
            cur = p2
    return out


def parse_svg(d: str, ox: float, oy: float) -> list:
    segs = []
    for m in re.finditer(r"([MLQCZ])([^MLQCZ]*)", d):
        n = [float(x) for x in m.group(2).split()]
        pts = [n[i] + (ox if i % 2 == 0 else oy) for i in range(len(n))]
        segs.append([m.group(1), *pts])
    return segs


def deviation(a: list[tuple[float, float]], b: list[tuple[float, float]]) -> float:
    """Largest distance from a point of one outline to the other outline (both ways)."""
    if not a or not b:
        return float("inf") if (a or b) else 0.0

    def one(p, q):
        grid: dict[tuple[int, int], list] = {}
        for x, y in q:
            grid.setdefault((int(x // 4), int(y // 4)), []).append((x, y))
        worst = 0.0
        for x, y in p:
            best = float("inf")
            gx, gy = int(x // 4), int(y // 4)
            for r in range(0, 40):
                for i in range(gx - r, gx + r + 1):
                    for j in range(gy - r, gy + r + 1):
                        if max(abs(i - gx), abs(j - gy)) != r:
                            continue
                        for qx, qy in grid.get((i, j), ()):
                            best = min(best, math.hypot(qx - x, qy - y))
                if best <= r * 4:
                    break
            worst = max(worst, best)
        return worst

    return max(one(a, b), one(b, a))


def overlay(name: str, word_pts: list, our_pts: list, s: dict) -> None:
    """Word's outline (red) over ours (blue), 4x, to captures/shapes/ (local only)."""
    from PIL import Image, ImageDraw
    cx, cy, cw, ch = s["cell"]
    k = 4
    img = Image.new("RGB", (cw * k, ch * k), "white")
    g = ImageDraw.Draw(img)
    for pts, col, r in ((word_pts, (230, 0, 0), 2), (our_pts, (0, 60, 255), 1)):
        for x, y in pts:
            px, py = (x - cx) * k, (y - cy) * k
            g.ellipse((px - r, py - r, px + r, py + r), fill=col)
    dest = HERE / "captures" / "shapes"
    dest.mkdir(parents=True, exist_ok=True)
    img.save(dest / f"{name}.png")


def ours(spec: list[dict]) -> list[dict]:
    inp, out = TMP / "spec.json", TMP / "ours.json"
    inp.write_text(json.dumps(spec), encoding="utf-8")
    env = {**__import__("os").environ, "SHAPES_IN": str(inp), "SHAPES_OUT": str(out)}
    subprocess.run("npx vitest run packages/text-engine/test/geometry.test.ts", cwd=ROOT, env=env, shell=True, check=True, capture_output=True)
    return json.loads(out.read_text(encoding="utf-8"))


def main() -> None:
    only = sys.argv[1:]
    names = [n for n in presets() if not only or n in only]
    first = ours([{"prst": n, "x": 0, "y": 0, "w": 80, "h": 50, "adj": {}, "page": 0} for n in names])
    defaults = {d["prst"]: d["adjust"] for d in first}
    result: dict = {}
    for variant in ("wide", "tall", "adjusted"):
        spec = layout(names, variant, defaults)
        path = TMP / f"shapes-{variant}.docx"
        docx(spec, path)
        pdf_json = TMP / f"shapes-{variant}.json"
        outline(path, pdf_json)
        pages = json.loads(pdf_json.read_text(encoding="utf-8"))
        mine = ours(spec)
        for s, m in zip(spec, mine):
            cx, cy, cw, ch = s["cell"]
            word_pts: list = []
            word_fills: set[str] = set()
            for p in pages[s["page"]]["paths"] if s["page"] < len(pages) else []:
                first = next((seg for seg in p["d"] if seg[0] == "M"), None)
                if first and cx <= first[1] < cx + cw and cy <= first[2] < cy + ch:
                    word_pts += flatten(p["d"])
                    if p.get("fill"):
                        word_fills.add(p["fill"])
            our_fills = {shade(p["fill"]) for p in m["paths"] if p["fill"] != "none"}
            same = all(any(close(a, b) for b in our_fills) for a in word_fills) and all(any(close(a, b) for b in word_fills) for a in our_fills)
            if not same:
                result.setdefault(s["prst"], {})["fills"] = f"word {sorted(word_fills)} ours {sorted(our_fills)}"
            our_pts: list = []
            for p in m["paths"]:
                if p["fill"] == "none" and not p["stroke"]:
                    continue
                our_pts += flatten(parse_svg(p["d"], s["x"], s["y"]))
            dev = round(deviation(word_pts, our_pts), 2)
            result.setdefault(s["prst"], {})[variant] = dev
            if only:
                overlay(f"{s['prst']}-{variant}", word_pts, our_pts, s)
    bad = {k: v for k, v in result.items() if "fills" in v or max(x for x in v.values() if isinstance(x, float)) > 0.5}
    for k, v in sorted(bad.items()):
        print(f"{k:28} {v}")
    print(f"{len(result) - len(bad)}/{len(result)} presets within 0.5 pt")
    if not only:
        dest = HERE / "results" / f"{datetime.date.today()}-shapes.json"
        dest.write_text(json.dumps(result, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
