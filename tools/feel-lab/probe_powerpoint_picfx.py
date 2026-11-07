"""Probe: how PowerPoint stores picture effects (Step 8, D-072), in a PowerPoint of our own without a
window (feellab/powerpoint.py refuses to run while the owner's PowerPoint is open).

A test picture (made here in TEMP) is inserted once per effect: Farbsättigung, Farbtemperatur,
Schärfen/Weichzeichnen and every Künstlerischer Effekt (MsoPictureEffectType), each with its default
parameters and, for the three adjustments, two more values. The presentation is saved, and for each
slide the probe keeps the a:blip element (as XML), the media parts it refers to (type and pixel size),
and the slide exported as PNG for the render comparison (in captures/, not committed).

    python probe_powerpoint_picfx.py

Result: results/<date>-powerpoint-picfx.json; PNGs and the .pptx in captures/picfx/.
"""
from __future__ import annotations

import datetime
import io
import json
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from feellab.powerpoint import OwnerPowerPointRunning, PowerPointTarget

HERE = Path(__file__).parent
CAPT = HERE / "captures" / ("picfx-temp" if "temps" in sys.argv[1:] else "picfx")

EFFECTS = {
    2: "Blur", 4: "Cement", 5: "CrisscrossEtching", 6: "ChalkSketch", 7: "ColorTemperature", 8: "Cutout",
    9: "FilmGrain", 10: "Glass", 11: "GlowDiffused", 12: "GlowEdges", 13: "LightScreen", 14: "LineDrawing",
    15: "Marker", 16: "MosaicBubbles", 17: "PaintBrush", 18: "PaintStrokes", 19: "PastelsSmooth",
    20: "PencilGrayscale", 21: "PencilSketch", 22: "Photocopy", 23: "PlasticWrap", 24: "Saturation",
    25: "SharpenSoften", 26: "Texturizer", 27: "WatercolorSponge",
}
# Extra values for the adjustments (first parameter): saturation 0..4, temperature 1500..11500, sharpen -1..1.
EXTRA = {24: [0.0, 2.0], 7: [4700.0, 11200.0], 25: [-0.5, 0.5]}


def test_picture(path: Path) -> None:
    """A colourful 320x240 picture: hue bands, a grey ramp and a sharp edge."""
    from PIL import Image, ImageDraw
    import colorsys
    img = Image.new("RGB", (320, 240), "white")
    d = ImageDraw.Draw(img)
    for x in range(320):
        r, g, b = colorsys.hsv_to_rgb(x / 320, 0.8, 0.95)
        d.line([(x, 0), (x, 119)], fill=(int(r * 255), int(g * 255), int(b * 255)))
    for x in range(320):
        v = int(x / 319 * 255)
        d.line([(x, 120), (x, 179)], fill=(v, v, v))
    d.rectangle([0, 180, 159, 239], fill=(30, 60, 120))
    d.rectangle([160, 180, 319, 239], fill=(240, 200, 60))
    d.ellipse([130, 190, 190, 235], fill=(200, 40, 40))
    img.save(path)


TEMPS = [1500, 2500, 3500, 4000, 4700, 5300, 5900, 6500, 7200, 8000, 8800, 10000, 11200, 11500] if "temps" in sys.argv[1:] else []


def fit_temps(z: zipfile.ZipFile, cases: list[dict]) -> list[dict]:
    """Least squares a, b per channel from the grey ramp row (y 150) of PowerPoint's rendered picture."""
    import numpy as np
    from PIL import Image
    fits = []
    ramp = np.array([int(x / 319 * 255) for x in range(320)], dtype=float)
    for i, c in enumerate(cases, start=1):
        rels = z.read(f"ppt/slides/_rels/slide{i}.xml.rels").decode("utf-8")
        target = re.search(r'Target="\.\./media/(image[^"]+)"', rels).group(1)
        px = np.asarray(Image.open(io.BytesIO(z.read("ppt/media/" + target))).convert("RGB")).astype(float)[150]
        row = {"temp": c["value"]}
        for ch, name in enumerate("rgb"):
            keep = (px[:, ch] > 2) & (px[:, ch] < 253)
            a, b = np.polyfit(ramp[keep], px[keep, ch], 1)
            err = float(np.abs(a * ramp[keep] + b - px[keep, ch]).max())
            row[name] = {"a": round(float(a), 4), "b": round(float(b), 2), "maxerr": round(err, 2)}
        fits.append(row)
    return fits


def main() -> int:
    try:
        ppt = PowerPointTarget(visible=False)
    except OwnerPowerPointRunning as e:
        print("skipped:", e)
        return 2
    out: dict = {"date": datetime.date.today().isoformat(), "what": "PowerPoint's XML for picture effects", "cases": []}
    tmp = Path(tempfile.mkdtemp(prefix="raabe-picfx-"))
    try:
        src = tmp / "test.png"
        test_picture(src)
        pres = ppt.new_presentation(name="picfx")
        cases: list[dict] = []
        if TEMPS:
            # Farbton: the gains PowerPoint applies per temperature (fitted from the grey ramp afterwards).
            cases = [{"kind": 7, "name": "ColorTemperature", "value": float(k)} for k in TEMPS]
        else:
            for kind, name in EFFECTS.items():
                for value in [None, *EXTRA.get(kind, [])]:
                    cases.append({"kind": kind, "name": name, "value": value})
        for i, c in enumerate(cases, start=1):
            slide = pres.Slides.Add(i, 12)  # ppLayoutBlank
            pic = slide.Shapes.AddPicture(str(src), 0, -1, 120, 90, 480, 360)
            try:
                fx = pic.Fill.PictureEffects.Insert(c["kind"])
                params = []
                for k in range(1, fx.EffectParameters.Count + 1):
                    p = fx.EffectParameters(k)
                    params.append({"name": str(p.Name), "value": float(p.Value)})
                if c["value"] is not None and fx.EffectParameters.Count:
                    fx.EffectParameters(1).Value = c["value"]
                    params[0]["value"] = float(fx.EffectParameters(1).Value)
                c["params"] = params
            except Exception as e:  # noqa: BLE001
                c["error"] = str(e)[:200]
        path = tmp / "picfx.pptx"
        ppt.save_copy(pres, path)
        CAPT.mkdir(parents=True, exist_ok=True)
        shutil.copy(path, CAPT / "picfx.pptx")
        for i in range(1, len(cases) + 1):
            ppt.export_png(pres, i, CAPT / f"slide{i}.png", width=1280)
        z = zipfile.ZipFile(path)
        names = z.namelist()
        from PIL import Image
        for i, c in enumerate(cases, start=1):
            xml = z.read(f"ppt/slides/slide{i}.xml").decode("utf-8")
            rels = z.read(f"ppt/slides/_rels/slide{i}.xml.rels").decode("utf-8")
            m = re.search(r"<a:blip\b.*?(</a:blip>|/>)", xml, re.S)
            c["blip"] = m.group(0) if m else None
            fx = re.search(r"<a14:imgEffect>(.*?)</a14:imgEffect>", xml, re.S)
            c["effect"] = fx.group(1) if fx else None
            media = {}
            for rid, target in re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels) + [(a, b) for b, a in re.findall(r'Target="([^"]+)"[^>]*Id="(rId\d+)"', rels)]:
                part = "ppt/" + target.replace("../", "")
                if part in names and "/media/" in part:
                    info = {"part": part.split("/")[-1], "bytes": len(z.read(part))}
                    try:
                        im = Image.open(io.BytesIO(z.read(part)))
                        info["size"] = list(im.size)
                        info["mode"] = im.mode
                    except Exception:  # noqa: BLE001
                        info["size"] = None
                    media[rid] = info
            c["media"] = media
            c["slide"] = i
        out["cases"] = cases
        if TEMPS:
            out["what"] = "PowerPoint's Farbton (colour temperature): per channel out = a * in + b, fitted on a grey ramp"
            out["fits"] = fit_temps(z, cases)
            for c in cases:
                c.pop("blip", None)
        res = HERE / "results" / f"{out['date']}-powerpoint-picfx{'-temp' if TEMPS else ''}.json"
        res.write_text(json.dumps(out, indent=1, ensure_ascii=False), encoding="utf-8", newline="\n")
        print("wrote", res.name, len(cases), "cases", sum(1 for c in cases if "error" in c), "errors")
    finally:
        print(ppt.close(kill_if_unreachable=True))
        shutil.rmtree(tmp, ignore_errors=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
