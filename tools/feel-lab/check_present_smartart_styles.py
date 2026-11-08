"""SmartArt colour sets and quick styles compared with PowerPoint (D-101). PowerPoint's file of
probe_powerpoint_smartart_styles.py (captures/smartart-styles, not committed: an "Einfache Blockliste" with
PowerPoint's five default items per colour set and per quick style) is opened in RaabePresent (headless Edge on
the dev server). Each slide is drawn twice and compared pixel by pixel with PowerPoint's own picture of it:
first PowerPoint's drawing as we draw it (how close our renderer gets), then the diagram laid out again by our
engine with the same colour set and quick style (how close our drawing gets).

    .venv/Scripts/python.exe check_present_smartart_styles.py      (npm run dev:present must be running)

Result: results/<date>-present-smartart-styles.json.
"""
from __future__ import annotations

import datetime
import io
import json
from pathlib import Path

from PIL import Image

from check_present_render import compare
from feellab.present_page import PresentPage

HERE = Path(__file__).parent
CAP = HERE / "captures" / "smartart-styles"


def main() -> None:
    rec = json.loads(sorted((HERE / "results").glob("*-powerpoint-smartart-styles.json"))[-1].read_text(encoding="utf-8"))
    names = {x["slide"]: (k, x["name"]) for k in ("colors", "styles") for x in rec[k] if "slide" in x}
    out: dict = {"date": datetime.date.today().isoformat(), "slides": []}
    page = PresentPage(1600, 1000, 1.0)
    try:
        n = page.open_file(CAP / "smartart-styles.pptx")
        for i in range(1, n + 1):
            ref = CAP / f"slide-{i}.png"
            if i not in names or not ref.exists():
                continue
            theirs = Image.open(ref).convert("RGB")
            row = {"slide": i, "kind": names[i][0], "name": names[i][1]}
            for which in ("powerpoint_drawing", "our_drawing"):
                if which == "our_drawing":
                    page.js(f"""(async () => {{ const a = window.__present.app; a.setCurrent({i - 1});
                        const m = await import('/src/smartart-create.ts');
                        const v = a.deck.view(a.deck.slides[{i - 1}].path).shapes.find((s) => s.frame === 'diagram');
                        const id = v.id; const d = m.diagramOf(a, id); m.relayout(a, id, d.items, d.layoutId); return [d.scheme, d.style]; }})()""")
                r = page.js(f"window.__present.renderAlone({i}, {theirs.width})")
                ours = Image.open(io.BytesIO(page.shot(0, 0, r["w"], r["h"]))).convert("RGB")
                row[which] = compare(theirs, ours)
                if which == "our_drawing":
                    side = Image.new("RGB", (theirs.width * 2 + 10, theirs.height), "white")
                    side.paste(theirs, (0, 0)); side.paste(ours, (theirs.width + 10, 0))
                    side.save(CAP / f"compare-{i}.png")
            out["slides"].append(row)
            print(i, row["name"], row["powerpoint_drawing"]["mean"], row["our_drawing"]["mean"], flush=True)
    finally:
        page.close()
    s = out["slides"]
    for which in ("powerpoint_drawing", "our_drawing"):
        out[f"mean_{which}"] = round(sum(x[which]["mean"] for x in s) / max(1, len(s)), 2)
    p = HERE / "results" / f"{out['date']}-present-smartart-styles.json"
    p.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))
    print("mean", out["mean_powerpoint_drawing"], out["mean_our_drawing"])


if __name__ == "__main__":
    main()
