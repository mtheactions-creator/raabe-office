"""Finds invisible and control characters in source files (they are easy to introduce by
accident and invisible in review) and, with --fix, replaces them with escapes.

    python tools/check-chars.py          # report
    python tools/check-chars.py --fix    # rewrite as \\uXXXX escapes

Only string contents are expected to hold such characters, and both TypeScript and Python
read \\uXXXX escapes there, so the fix keeps the meaning."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PATTERNS = ["packages/*/src/**/*.ts", "packages/*/test/**/*.ts", "apps/*/src/**/*.ts", "tools/**/*.py", "branding/*.py"]
# Visible-looking but special characters that must be written as escapes.
SPECIAL = {0xA0, 0xAD, 0x2011, 0x200B, 0x200C, 0x200D, 0x2060, 0xFEFF, 0xFFFC}


def bad(c: str) -> bool:
    o = ord(c)
    return o in SPECIAL or (o < 32 and c not in "\n\r\t")


def files() -> list[Path]:
    out: set[Path] = set()
    for pattern in PATTERNS:
        out.update(p for p in ROOT.glob(pattern) if ".venv" not in p.parts and "node_modules" not in p.parts)
    return sorted(out)


def main() -> int:
    fix = "--fix" in sys.argv[1:]
    problems = 0
    for f in files():
        text = f.read_text(encoding="utf-8")
        found = [i for i, c in enumerate(text) if bad(c)]
        if not found:
            continue
        problems += 1
        lines = sorted({text.count("\n", 0, i) + 1 for i in found})
        print(f"{f.relative_to(ROOT)}: {len(found)} character(s), lines {lines[:8]}")
        if fix:
            f.write_text("".join(f"\\u{ord(c):04x}" if bad(c) else c for c in text), encoding="utf-8", newline="\n")
            print("  fixed")
    return 1 if problems and not fix else 0


if __name__ == "__main__":
    sys.exit(main())
