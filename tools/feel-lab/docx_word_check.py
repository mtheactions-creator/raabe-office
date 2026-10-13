"""Opens original test documents and RaabeWrite's rewritten versions in Word (isolated,
invisible instance) and compares what Word sees in both. A file Word considers damaged fails
to open. Run after `ROUNDTRIP_OUT=tools/feel-lab/captures/roundtrip npx vitest run packages/ooxml`.

    .venv/Scripts/python.exe docx_word_check.py
"""

from __future__ import annotations

import sys
from pathlib import Path

import win32com.client

ROOT = Path(__file__).resolve().parents[2]
ORIG = ROOT / "packages" / "ooxml" / "test" / "fixtures"
OUT = Path(__file__).parent / "captures" / "roundtrip"


def inspect(app, path: Path) -> dict:
    d = app.Documents.Open(str(path), ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False, OpenAndRepair=False)
    try:
        return {
            "paragraphs": d.Paragraphs.Count, "tables": d.Tables.Count, "inline images": d.InlineShapes.Count,
            "shapes": d.Shapes.Count, "footnotes": d.Footnotes.Count, "comments": d.Comments.Count,
            "hyperlinks": d.Hyperlinks.Count, "revisions": d.Revisions.Count, "equations": d.OMaths.Count,
            "fields": d.Fields.Count, "bookmarks": d.Bookmarks.Count, "list paragraphs": d.ListParagraphs.Count,
            "sections": d.Sections.Count, "pages": d.ComputeStatistics(2), "text": d.Content.Text,
            "headers": d.Sections(1).Headers(1).Range.Text, "styles used": sorted({p.Style.NameLocal for p in d.Paragraphs}),
        }
    finally:
        d.Close(0)


def main() -> int:
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    failures = 0
    try:
        for f in sorted(OUT.glob("*.docx")):
            try:
                a = inspect(app, ORIG / f.name)
                b = inspect(app, f)
            except Exception as e:  # Word refused or repaired the file
                print(f"FAIL {f.name}: Word could not open it cleanly: {e}")
                failures += 1
                continue
            diffs = [k for k in a if a[k] != b[k]]
            if diffs:
                failures += 1
                print(f"FAIL {f.name}: differs in {diffs}")
                for k in diffs:
                    print(f"   {k}: {str(a[k])[:120]!r} -> {str(b[k])[:120]!r}")
            else:
                summary = {k: v for k, v in a.items() if k not in ("text", "headers", "styles used") and v}
                print(f"PASS {f.name}: identical in Word {summary}")
    finally:
        app.Quit(0)
    return failures


if __name__ == "__main__":
    sys.exit(main())
