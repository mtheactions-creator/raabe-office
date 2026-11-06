"""What an inventory of Office's UI may keep in a committed results file (CLAUDE.md, D-058): labels and
names, not Microsoft's descriptions, and nothing of the owner's (recent files, paths, user names).
The full recording goes to captures/ (not committed)."""
from __future__ import annotations

import os
import re

FILE = re.compile(r"\.(pdf|docx?|xlsx?|xlsm|pptx?|csv|txt|png|jpe?g)\b", re.I)
FUNC = re.compile(r"^([A-ZÄÖÜ][A-ZÄÖÜ0-9.]*), \S")
MAIL = re.compile(r"\S+@\S+\.\w+")
_NAMES: list[str] | None = None


def private_names() -> list[str]:
    """The owner's names as Windows and Office know them (Windows user, Office user name, signed-in
    identities), whole and by parts of 4+ letters, in lower case: Office shows them in the title bar,
    the backstage and greetings (PowerPoint inventory, Step 8). Read on this PC, never written out."""
    global _NAMES
    if _NAMES is not None:
        return _NAMES
    found: set[str] = {os.environ.get("USERNAME", ""), os.path.basename(os.environ.get("USERPROFILE", ""))}
    try:
        import winreg

        def value(key, name: str) -> str:
            try:
                return str(winreg.QueryValueEx(key, name)[0])
            except OSError:
                return ""

        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office\Common\UserInfo") as k:
                found.add(value(k, "UserName"))
        except OSError:
            pass
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office\16.0\Common\Identity\Identities") as k:
                i = 0
                while True:
                    try:
                        sub = winreg.EnumKey(k, i)
                    except OSError:
                        break
                    i += 1
                    with winreg.OpenKey(k, sub) as s:
                        for n in ("FriendlyName", "EmailAddress", "FirstName", "LastName"):
                            found.add(value(s, n))
        except OSError:
            pass
    except ImportError:
        pass
    names: set[str] = set()
    for f in found:
        f = f.strip()
        if len(f) >= 4:
            names.add(f.lower())
        for part in re.split(r"[\s@._-]+", f):
            if len(part) >= 4:
                names.add(part.lower())
    _NAMES = sorted(names)
    return _NAMES


def clean_text(v: str) -> str | None:
    """A label as it may be committed; None: leave the entry out."""
    # A button's name with its description on a second line (Office's picture source buttons): the label only.
    v = v.split(chr(10), 1)[0]
    if FILE.search(v) or re.search(r"[A-Za-z]:\\|\\Users\\|/Users/", v):
        return None
    if MAIL.search(v) or any(n in v.lower() for n in private_names()):
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
        # Static texts of six words or more are explanations, not labels (backstage, dialogs).
        if o.get("type") == "Text" and isinstance(o.get("name"), str) and len(o["name"].split()) >= 6:
            return None
        out = {}
        for k, v in o.items():
            # Keys can be names too (backstage pages named after recent files).
            if clean_text(k) is None:
                continue
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
