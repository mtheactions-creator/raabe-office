"""What an inventory of Office's UI may keep in a committed results file (CLAUDE.md, D-058): labels and
names, not Microsoft's descriptions, and nothing of the owner's (recent files, paths, user names).
The full recording goes to captures/ (not committed)."""
from __future__ import annotations

import re

FILE = re.compile(r"\.(pdf|docx?|xlsx?|xlsm|pptx?|csv|txt|png|jpe?g)\b", re.I)
FUNC = re.compile(r"^([A-ZÄÖÜ][A-ZÄÖÜ0-9.]*), \S")


def clean_text(v: str) -> str | None:
    """A label as it may be committed; None: leave the entry out."""
    if FILE.search(v) or re.search(r"[A-Za-z]:\\|\\Users\\|/Users/", v):
        return None
    m = FUNC.match(v)
    if m:
        return m.group(1)
    if ": " in v and (". " in v or v.endswith(".")):
        return v.split(": ", 1)[0]
    if len(v) > 40 and (v.endswith(".") or ". " in v):
        return "[Beschreibung]"
    return v


def clean(o):
    if isinstance(o, dict):
        out = {}
        for k, v in o.items():
            if isinstance(v, str):
                c = clean_text(v)
                if c is None:
                    return None
                out[k] = c
            else:
                c = clean(v)
                if c is not None:
                    out[k] = c
        return out
    if isinstance(o, list):
        return [c for c in (clean(v) for v in o) if c is not None]
    return o
