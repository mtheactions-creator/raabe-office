"""The repository is public (CLAUDE.md): Office writes the user's name into every file it saves (core
properties, comments, people, and the workbooks embedded in charts). This replaces it in .xlsx, .docx,
.pptx and their macro/template variants, also inside embedded packages, and reports what it found.

    python tools/feel-lab/anonymize_office.py <file or folder> ...     # rewrite
    python tools/feel-lab/anonymize_office.py --check <file or folder> ...  # only report (exit 1 if found)
"""
from __future__ import annotations

import io
import re
import sys
import zipfile
from pathlib import Path

TEST_AUTHOR = "Raabe Office Test"
# Names the harness itself writes (make_fixtures.py, RaabePresent's fixtures, RaabeX's own files).
ALLOWED = {b"Raabe Office Test", b"Raabe Office", b"Raabe Test", b"RaabeCalc", b"RaabeWrite", b"RaabePresent", b"0000000000000000"}
EXT = (".xlsx", ".xlsm", ".docx", ".docm", ".pptx", ".pptm", ".potx", ".ppsx", ".dotx", ".xltx", ".thmx")
NAME_TAGS = re.compile(rb"<(dc:creator|cp:lastModifiedBy)>([^<]*)</\1>")


def names_in(data: bytes) -> set[bytes]:
    out: set[bytes] = set()
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile:
        return out
    for n in z.namelist():
        t = z.read(n)
        if n.endswith("core.xml"):
            out |= {m.group(2) for m in NAME_TAGS.finditer(t) if m.group(2) and m.group(2) not in ALLOWED}
        for m in re.finditer(rb'(?:w:author|author|w15:userId|userId|displayName|name)="([^"]+)"', t) if n.endswith(("people.xml", "persons/person.xml")) else []:
            if m.group(1) not in ALLOWED:
                out.add(m.group(1))
        if n.lower().endswith(EXT):
            out |= names_in(t)
    return {x for x in out if x.strip()}


def rewrite(data: bytes, names: set[bytes]) -> bytes:
    z = zipfile.ZipFile(io.BytesIO(data))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as w:
        for info in z.infolist():
            t = z.read(info.filename)
            if info.filename.lower().endswith(EXT):
                t = rewrite(t, names)
            elif info.filename.endswith((".xml", ".rels", ".vml")):
                for n in names:
                    t = t.replace(n, TEST_AUTHOR.encode())
                t = re.sub(rb'w:initials="[^"]*"', b'w:initials="RT"', t)
                t = re.sub(rb'w15:userId="[^"]*"', b'w15:userId="0000000000000000"', t)
            w.writestr(info, t)
    return buf.getvalue()


def files(args: list[str]):
    for a in args:
        p = Path(a)
        if p.is_dir():
            yield from (f for f in p.rglob("*") if f.suffix.lower() in EXT)
        elif p.suffix.lower() in EXT:
            yield p


def main() -> None:
    check = "--check" in sys.argv
    args = [a for a in sys.argv[1:] if a != "--check"]
    found = False
    for f in files(args):
        data = f.read_bytes()
        n = names_in(data)
        if not n:
            continue
        found = True
        print(("found " if check else "anonymized ") + str(f), sorted(x.decode("utf-8", "replace") for x in n) if check else "")
        if not check:
            f.write_bytes(rewrite(data, n))
    sys.exit(1 if check and found else 0)


if __name__ == "__main__":
    main()
