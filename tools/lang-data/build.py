"""Builds RaabeWrite's thesaurus files from the MyThes data fetched by `fetch.py` (D-049).

Each language becomes one gzip file in apps/raabe-write/src-tauri/data/: the meanings once
each ("pos<TAB>term|term|…", antonyms marked with a leading "!"), then the headwords with the
numbers of their meanings in the source's order ("word<TAB>3,17,…"). Generic and related terms
(Oberbegriff, generic term, related term) are left out, as in Word's thesaurus.

    python -I tools/lang-data/build.py
"""
from __future__ import annotations

import gzip
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SRC = Path(__file__).parent / "src"
OUT = ROOT / "apps" / "raabe-write" / "src-tauri" / "data"
THESAURI = {"de": "de__th_de_DE_v2.dat", "en": "en__th_en_US_v2.dat", "fr": "fr_FR__dictionaries__thes_fr.dat", "es": "es__th_es_v2.dat"}
# Parts of speech as Word's German thesaurus names them.
POS = {
    "noun": "Subst.", "nom": "Subst.", "adj": "Adj.", "adjectif": "Adj.", "adv": "Adv.", "adverbe": "Adv.", "verb": "Verb", "verbe": "Verb",
    "adjectif nom": "Adj., Subst.", "préposition": "Präp.", "preposition": "Präp.", "interjection": "Interj.", "pronom": "Pron.", "conjonction": "Konj.",
}
DROP = re.compile(r"\((?:Oberbegriff|Unterbegriff|generic term|related term)\)")
ANTONYM = re.compile(r"\s*\((?:antonym|Antonym|Gegenwort|Gegensatz)\)")


def read(path: Path) -> list[str]:
    raw = path.read_bytes()
    first, _, rest = raw.partition(b"\n")
    enc = first.decode("ascii").strip().lower().replace("iso8859", "iso-8859")
    return rest.decode(enc, "replace").splitlines()


def build(lang: str, name: str) -> tuple[int, int, int]:
    lines = read(SRC / name)
    meanings: dict[tuple[str, tuple[str, ...]], int] = {}
    order: list[tuple[str, tuple[str, ...]]] = []
    index: dict[str, list[int]] = {}
    i = 0
    while i < len(lines):
        head = lines[i].rstrip("\r")
        i += 1
        if "|" not in head:
            continue
        word, _, count = head.rpartition("|")
        if not count.strip().isdigit():
            continue
        ids: list[int] = []
        for _ in range(int(count)):
            if i >= len(lines):
                break
            parts = lines[i].rstrip("\r").split("|")
            i += 1
            pos = parts[0].strip()
            pos = "" if pos == "-" else POS.get(pos.strip("()").lower(), "")
            terms = []
            for t in parts[1:]:
                t = t.strip()
                if not t or DROP.search(t):
                    continue
                if ANTONYM.search(t):
                    t = "!" + ANTONYM.sub("", t)
                terms.append(t)
            if not terms:
                continue
            key = (pos, tuple(terms))
            if key not in meanings:
                meanings[key] = len(order)
                order.append(key)
            if meanings[key] not in ids:
                ids.append(meanings[key])
        if ids:
            w = word.strip().lower()
            known = index.setdefault(w, [])
            known.extend(x for x in ids if x not in known)
    body = "\n".join(f"{p}\t{'|'.join(t)}" for p, t in order) + "\n\n" + "\n".join(f"{w}\t{','.join(map(str, ids))}" for w, ids in sorted(index.items()))
    data = gzip.compress(body.encode("utf-8"), compresslevel=9, mtime=0)
    (OUT / f"thes-{lang}.gz").write_bytes(data)
    return len(order), len(index), len(data)


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    for lang, name in THESAURI.items():
        m, w, size = build(lang, name)
        print(f"{lang}: {m} meanings, {w} words, {size / 1e6:.2f} MB")


if __name__ == "__main__":
    main()
