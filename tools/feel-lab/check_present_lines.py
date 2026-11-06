"""Line breaks against PowerPoint's (D-070): for every text shape of the fixtures PowerPoint built, the
lines PowerPoint reports through COM (TextRange.Lines: where each starts) against where RaabePresent's
lines start, in a headless Edge on the dev server (invisible; npm run dev:present must run). Like
RaabeWrite's line-start measure (F4b).

    .venv/Scripts/python.exe check_present_lines.py

Result: results/<date>-present-lines.json.
"""
from __future__ import annotations

import datetime
import json
from pathlib import Path

from feellab.present_page import PresentPage

HERE = Path(__file__).parent
FIX = HERE.parent.parent / "packages" / "ooxml" / "test" / "fixtures"
FILES = ["pptx-layouts", "pptx-bullets", "pptx-textbox", "pptx-shapes", "pptx-group", "pptx-table", "pptx-autofit", "pptx-show",
         "pptx-office-layouts", "pptx-office-bullets", "pptx-office-textbox", "pptx-office-autofit"]


def flat(shapes):
    for s in shapes:
        yield s
        yield from flat(s.get("children", []))


def main() -> None:
    page = PresentPage(1600, 1000, 1.0)
    out = {"date": datetime.date.today().isoformat(), "shapes": []}
    total = same = 0
    # The pptx-office-* fixtures are in the current Office design with Aptos, which this PC's browser
    # may not have: counted apart.
    aptos = [0, 0]
    try:
        for name in FILES:
            f = FIX / f"{name}.pptx"
            ref = json.loads((FIX / f"{name}.json").read_text(encoding="utf-8"))
            page.open_file(f)
            for s in ref["slides"]:
                page.js(f"window.__present.renderAlone({s['index']}, 1280)")
                for sh in flat(s["shapes"]):
                    lines = (sh.get("text") or {}).get("lines")
                    if not lines or not sh.get("id"):
                        continue
                    ours = page.js(f"window.__present.lineStarts({sh['id']})") or []
                    theirs = [ln["start"] for ln in lines]
                    mine = [ln["start"] for ln in ours]
                    # A line counts when it starts at the same character in both; extra lines of ours count against.
                    n = max(len(theirs), len(mine))
                    hit = len(set(theirs) & set(mine))
                    if "office" not in name:
                        total += n
                        same += hit
                    else:
                        aptos[0] += hit
                        aptos[1] += n
                    tops = [round(ln["top"], 1) for ln in ours]
                    out["shapes"].append({"file": name, "slide": s["index"], "shape": sh["name"], "theirs": theirs, "ours": mine,
                                          "theirTops": [ln["top"] for ln in lines], "ourTops": tops, "same": hit, "of": n})
                    if theirs != mine:
                        print(f"{name} {s['index']} {sh['name']}: PowerPoint {theirs} ours {mine}")
    finally:
        page.close()
    out["summary"] = {"lineStarts": same, "of": total, "aptos": {"lineStarts": aptos[0], "of": aptos[1]}}
    print(f"line starts as PowerPoint: {same}/{total} (Office 2013-2022 design); Aptos design {aptos[0]}/{aptos[1]}")
    (HERE / "results" / f"{out['date']}-present-lines.json").write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))


if __name__ == "__main__":
    main()
