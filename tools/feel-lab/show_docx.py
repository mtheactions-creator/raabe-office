"""Prints the per-line table of the latest docx fidelity result.

    .venv/Scripts/python.exe show_docx.py [results/<date>-docx.json]
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

RESULTS = Path(__file__).parent / "results"
SUMMARY = ("lines_word", "lines_ours", "identical_line_starts", "percent", "pages_word", "pages_ours",
           "page_mismatches", "x", "baseline", "baseline_pdf", "render_vs_model")


def main(argv: list[str]) -> None:
    path = Path(argv[1]) if len(argv) > 1 else max(RESULTS.glob("*-docx.json"))
    data = json.loads(path.read_text(encoding="utf-8"))["results"]
    for name, r in data.items():
        if "error" in r:
            print(name, r)
            continue
        print(name, {k: r[k] for k in SUMMARY if k in r})
        hz = r.get("horizontal")
        if hz:
            print("horizontal:", {k: v for k, v in hz.items() if k not in ("worst", "rows")})
            for w in hz["worst"][:6]:
                print("   ", w)
        print("only in Word:", r["only_in_word"])
        print("only in RaabeWrite:", r["only_in_ours"])
        for d in r["lines"]:
            print(f"{d['start']:6} p{d['page'][0]}/{d['page'][1]} dx {d['dx']:6} base {d.get('dbase', ''):>6} "
                  f"pdf {d.get('dbase_pdf', ''):>6} render {d.get('render', ''):>6}  {d['context'][:34]!r}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main(sys.argv)
