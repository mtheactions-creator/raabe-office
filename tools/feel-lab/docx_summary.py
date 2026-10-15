"""Summary of a docx fidelity result (run.py docx): identical line starts and page breaks per
document, and in total. With a second file, the documents whose numbers differ between them.

    .venv/Scripts/python.exe docx_summary.py results/<date>-docx.json [earlier.json]
"""
from __future__ import annotations

import json
import sys
from pathlib import Path


def load(path: str) -> dict:
    d = json.loads(Path(path).read_text(encoding="utf-8"))
    return d.get("results", d)


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    now = load(sys.argv[1])
    before = load(sys.argv[2]) if len(sys.argv) > 2 else {}
    same = lines = pm = pages = 0
    for name, r in sorted(now.items()):
        if not isinstance(r, dict) or "lines_word" not in r:
            print(name, "error:", str(r)[:120])
            continue
        same += r["identical_line_starts"]
        lines += r["lines_word"]
        pm += r["page_mismatches"]
        pages += r["pages_word"]
        mark = ""
        b = before.get(name)
        if isinstance(b, dict) and "lines_word" in b and (b["identical_line_starts"], b["page_mismatches"]) != (r["identical_line_starts"], r["page_mismatches"]):
            mark = f"   (before: {b['identical_line_starts']}/{b['lines_word']}, {b['page_mismatches']} page differences)"
        print(f"{name:32} {r['identical_line_starts']:4}/{r['lines_word']:<4} pages {r['pages_ours']}/{r['pages_word']} page differences {r['page_mismatches']}{mark}")
    print(f"total: {same}/{lines} identical line starts, {pm} page differences in {pages} pages")


if __name__ == "__main__":
    main()
