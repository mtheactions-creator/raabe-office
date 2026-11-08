"""The look corpus for RaabePresent (D-080 D): school-like presentations PowerPoint built (make_corpus_ppt.py,
%TEMP%/raabe-corpus/ppt, copied to captures/corpus-ppt, not committed), each slide drawn by RaabePresent (headless
Edge on the dev server) at PowerPoint's picture width and compared pixel by pixel with PowerPoint's own picture:
the mean difference (0-255) and the share of pixels that differ by more than 48. Side-by-side pictures in
captures/corpus-ppt-compare.

    .venv/Scripts/python.exe check_present_corpus.py [name ...]      (npm run dev:present must be running)

Result: results/<date>-present-corpus.json.
"""
from __future__ import annotations

import datetime
import io
import json
import os
import shutil
import sys
import tempfile
from pathlib import Path

from PIL import Image

from check_present_render import compare
from feellab.present_page import PresentPage

HERE = Path(__file__).parent
SRC = Path(tempfile.gettempdir()) / "raabe-corpus" / "ppt"
CORPUS = HERE / "captures" / "corpus-ppt"
CAP = HERE / "captures" / "corpus-ppt-compare"


CLOUD = Path(os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts"))
FONTS = HERE / "captures" / "cloudfonts"


def cloud_faces() -> list[dict]:
    """Office's cloud fonts of this user (Aptos …), as the native command lists them (D-089): copied into captures
    (not committed) for the dev server, with families, weight and style read from each file's names. The headless
    page has the browser platform, which has no such command; this stands in for it."""
    from PIL import ImageFont
    out = []
    FONTS.mkdir(parents=True, exist_ok=True)
    for f in sorted(CLOUD.glob("Aptos*/*.ttf")) if CLOUD.exists() else []:
        fam, style = ImageFont.truetype(str(f), 10).getname()
        name = f"{fam} {style}"
        weight = 300 if "Light" in name else 600 if "SemiBold" in name else 800 if "ExtraBold" in name else 900 if "Black" in name else 700 if "Bold" in name else 400
        dst = FONTS / f.name
        if not dst.exists():
            shutil.copy(f, dst)
        # The family is the cache folder's: Aptos Display's files call themselves "Aptos" (style "Display …"), and
        # counted as Aptos they would replace Aptos Regular (narrower text, measured).
        out.append({"families": [f.parent.name], "weight": weight, "italic": "Italic" in name, "id": f.name})
    return out


def main(names: list[str]) -> None:
    CORPUS.mkdir(parents=True, exist_ok=True)
    CAP.mkdir(parents=True, exist_ok=True)
    if SRC.exists():
        for f in SRC.iterdir():
            shutil.copy(f, CORPUS / f.name)
    files = sorted(p.stem for p in CORPUS.glob("*.pptx") if not names or p.stem in names)
    out: dict = {"date": datetime.date.today().isoformat(), "slides": []}
    page = PresentPage(1600, 1000, 1.0)
    faces = cloud_faces()
    out["cloud_faces"] = len(faces)
    base = "/@fs/" + FONTS.resolve().as_posix() + "/"
    try:
        page.js(f"""(async () => {{ const t = await import('/@fs/{(HERE.parents[1] / 'packages/text-engine/src/index.ts').resolve().as_posix()}');
            const faces = {json.dumps(faces)};
            t.setCloudFonts({{ list: async () => faces, bytes: async (id) => (await fetch({json.dumps(base)} + id)).arrayBuffer() }}); }})()""")
        for name in files:
            n = page.open_file(CORPUS / f"{name}.pptx")
            for i in range(1, n + 1):
                ref = CORPUS / f"{name}-{i}.png"
                if not ref.exists():
                    continue
                theirs = Image.open(ref).convert("RGB")
                # Drawn once so the fonts are asked for, then again after they loaded (as the app does on onFontAlias).
                page.js(f"window.__present.renderAlone({i}, {theirs.width})")
                page.js("(async () => { await document.fonts.ready; await new Promise(r => setTimeout(r, 600)); await document.fonts.ready; })()")
                r = page.js(f"window.__present.renderAlone({i}, {theirs.width})")
                page.js("new Promise(r => setTimeout(r, 400))")
                ours = Image.open(io.BytesIO(page.shot(0, 0, r["w"], r["h"]))).convert("RGB")
                m = compare(ours, theirs)
                out["slides"].append({"file": name, "slide": i, **m})
                side = Image.new("RGB", (theirs.width * 2 + 10, max(theirs.height, ours.height)), "white")
                side.paste(theirs, (0, 0))
                side.paste(ours, (theirs.width + 10, 0))
                side.save(CAP / f"{name}-{i}.png")
                print(f"{name} {i}: mean {m['mean']}, differ {m['differ']} %", flush=True)
            page.js("window.__present.clearAlone()")
    finally:
        page.close()
    s = out["slides"]
    if s:
        out["summary"] = {"slides": len(s), "mean": round(sum(x["mean"] for x in s) / len(s), 2), "differ": round(sum(x["differ"] for x in s) / len(s), 2)}
        print("summary", out["summary"])
    p = HERE / "results" / f"{out['date']}-present-corpus.json"
    p.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))


if __name__ == "__main__":
    main(sys.argv[1:])
