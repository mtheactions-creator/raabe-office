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


def main(names: list[str]) -> None:
    CORPUS.mkdir(parents=True, exist_ok=True)
    CAP.mkdir(parents=True, exist_ok=True)
    if SRC.exists():
        for f in SRC.iterdir():
            shutil.copy(f, CORPUS / f.name)
    files = sorted(p.stem for p in CORPUS.glob("*.pptx") if not names or p.stem in names)
    out: dict = {"date": datetime.date.today().isoformat(), "slides": []}
    page = PresentPage(1600, 1000, 1.0)
    try:
        for name in files:
            n = page.open_file(CORPUS / f"{name}.pptx")
            for i in range(1, n + 1):
                ref = CORPUS / f"{name}-{i}.png"
                if not ref.exists():
                    continue
                theirs = Image.open(ref).convert("RGB")
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
