"""Condenses captures/word-ribbon/*.json (capture_ribbon_widths.py) into the distinct layouts
each ribbon group of Word takes, with the widths where Word uses them. Prints a compact
description: per layout the columns, each a list of controls with size class (L large,
M medium = icon and label or icon and arrow, S small icon, G gallery).

    .venv/Scripts/python.exe ribbon_variants.py [Start …]
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

DIR = Path(__file__).parent / "captures" / "word-ribbon"
TABS = ["Start", "Einfügen", "Entwurf", "Layout", "Referenzen", "Überprüfen", "Ansicht"]
WIDTHS = [2560, 2300, 2100, 1900, 1700, 1500, 1333, 1200, 1050, 900, 800]


def columns(controls: list[dict]) -> list[list[dict]]:
    """Controls grouped into columns (large buttons, stacks, rows of small buttons)."""
    cs = sorted((c for c in controls if c["w"] > 18), key=lambda c: (c["x"], c["y"]))
    cols: list[list[dict]] = []
    for c in cs:
        for col in cols:
            if abs(col[0]["x"] - c["x"]) <= 3 and c["size"] in ("M", "S") and col[0]["size"] in ("M", "S") and c["y"] > max(o["y"] for o in col):
                col.append(c)
                break
        else:
            cols.append([c])
    return cols


def signature(g: dict) -> str:
    if g["collapsed"]:
        return "collapsed"
    rows = sorted({c["y"] for c in g["controls"] if c["w"] > 18})
    parts = []
    two_rows = len([y for y in rows if y < 70]) and all(abs(y - 10) < 4 or abs(y - 42) < 4 for y in rows if y < 70 and y > 3)
    if two_rows and len(rows) <= 2:
        # Rows of small buttons (Schriftart, Absatz): list each row.
        for y in rows:
            row = sorted((c for c in g["controls"] if c["y"] == y and c["w"] > 18), key=lambda c: c["x"])
            parts.append("row: " + ", ".join(f"{c['size']}:{c['name']}" for c in row))
        return " | ".join(parts)
    for col in columns(g["controls"]):
        if len(col) == 1:
            c = col[0]
            parts.append(f"{c['size']}:{c['name']}")
        else:
            parts.append("stack(" + ", ".join(f"{c['size']}:{c['name']}" for c in col) + ")")
    return " | ".join(parts)


def main() -> None:
    tabs = [a for a in sys.argv[1:]] or TABS
    for tab in tabs:
        print(f"\n######## {tab}")
        per_group: dict[str, list[tuple[int, str, int]]] = {}
        order: list[str] = []
        for w in WIDTHS:
            f = DIR / f"{tab}-{w}.json"
            if not f.exists():
                continue
            for g in json.loads(f.read_text(encoding="utf-8")):
                if g["name"] not in per_group:
                    order.append(g["name"])
                per_group.setdefault(g["name"], []).append((w, signature(g), g["w"]))
        for name in order:
            print(f"\n  [{name}]")
            last = None
            for w, sig, gw in per_group[name]:
                if sig != last:
                    print(f"    from {w}px (group {gw} px): {sig}")
                    last = sig


if __name__ == "__main__":
    main()
