"""Slide renders against PowerPoint's own pictures (D-070): every slide of the fixtures PowerPoint built
(make_pptx_fixtures.py, exported by PowerPoint at 1280 px in fixtures/pptx-png) is drawn by RaabePresent
at the same width in a headless Edge (invisible, the dev server must run: npm run dev:present) and
compared pixel by pixel: the mean difference of all pixels (0-255) and the share of pixels that differ
by more than 48 (where text or shapes sit elsewhere).

    .venv/Scripts/python.exe check_present_render.py [fixture ...]

Result: results/<date>-present-render.json; side-by-side pictures in captures/present-render/ (not
committed)."""
from __future__ import annotations

import datetime
import io
import json
import sys
from pathlib import Path

from PIL import Image, ImageChops

from feellab.present_page import PresentPage

HERE = Path(__file__).parent
FIX = HERE.parent.parent / "packages" / "ooxml" / "test" / "fixtures"
PNG = FIX / "pptx-png"
CAP = HERE / "captures" / "present-render"
# The fixtures in the Office 2013-2022 design (our default); the pptx-office-* ones use Aptos.
DEFAULT = ["pptx-blank", "pptx-layouts", "pptx-bullets", "pptx-textbox", "pptx-shapes", "pptx-group", "pptx-picture", "pptx-table", "pptx-autofit", "pptx-show", "pptx-equation"]


def compare(a: Image.Image, b: Image.Image) -> dict:
    a = a.convert("RGB")
    b = b.convert("RGB").resize(a.size)
    d = ImageChops.difference(a, b).convert("L")
    hist = d.histogram()
    total = sum(hist)
    mean = sum(i * n for i, n in enumerate(hist)) / total
    big = sum(hist[49:]) / total
    return {"mean": round(mean, 2), "differ": round(big * 100, 2)}


def main(names: list[str]) -> None:
    CAP.mkdir(parents=True, exist_ok=True)
    page = PresentPage(1600, 1000, 1.0)
    out = {"date": datetime.date.today().isoformat(), "width": 1280, "slides": []}
    try:
        for name in names or DEFAULT:
            f = FIX / f"{name}.pptx"
            if not f.exists():
                continue
            n = page.open_file(f)
            for i in range(1, n + 1):
                ref = PNG / f"{name}-{i}.png"
                if not ref.exists():
                    continue
                theirs = Image.open(ref)
                r = page.js(f"window.__present.renderAlone({i}, {theirs.width})")
                page.js("new Promise(r => setTimeout(r, 250))")
                ours = Image.open(io.BytesIO(page.shot(0, 0, r["w"], r["h"])))
                m = compare(ours, theirs)
                out["slides"].append({"file": name, "slide": i, **m})
                side = Image.new("RGB", (theirs.width * 2 + 10, max(theirs.height, ours.height)), "white")
                side.paste(theirs.convert("RGB"), (0, 0))
                side.paste(ours.convert("RGB"), (theirs.width + 10, 0))
                side.save(CAP / f"{name}-{i}.png")
                print(f"{name} {i}: mean {m['mean']}, differ {m['differ']} %")
            page.js("window.__present.clearAlone()")
    finally:
        page.close()
    s = out["slides"]
    if s:
        out["summary"] = {"slides": len(s), "mean": round(sum(x["mean"] for x in s) / len(s), 2), "differ": round(sum(x["differ"] for x in s) / len(s), 2)}
        print("summary", out["summary"])
    (HERE / "results" / f"{out['date']}-present-render.json").write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))


if __name__ == "__main__":
    main(sys.argv[1:])
