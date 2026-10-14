"""Builds RaabeWrite's thesaurus files and hyphenation patterns from the data fetched by
`fetch.py` (D-049).

Each language becomes one gzip file in apps/raabe-write/src-tauri/data/: the meanings once
each ("pos<TAB>term|term|…", antonyms marked with a leading "!"), then the headwords with the
numbers of their meanings in the source's order ("word<TAB>3,17,…"). Generic and related terms
(Oberbegriff, generic term, related term) are left out, as in Word's thesaurus.

Hyphenation: each Hyphen pattern file becomes a module in packages/text-engine/src/hyph/ with
its minimum lengths and its patterns (two levels for German: compound boundaries first, as
Hyphen's NEXTLEVEL; patterns with non-standard changes are left out).

    python -I tools/lang-data/build.py
"""
from __future__ import annotations

import gzip
import json
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


HYPH_OUT = ROOT / "packages" / "text-engine" / "src" / "hyph"
HYPHENATION = {
    "de": ("de__hyph_de_DE.dic", "German (reformed spelling): LibreOffice's hyph_de_DE.dic, based on TeX's dehyphn.tex, by Marco Huggenberger, Daniel Naber, Karl Zeiler; LGPL-2 or later (TeX patterns LPPL)."),
    "en-us": ("en__hyph_en_US.dic", "American English: LibreOffice's hyph_en_US.dic, from plain TeX's hyphen.tex, conversion by László Németh; BSD-style licence (unlimited copying, redistribution and modification with this notice)."),
    "en-gb": ("en__hyph_en_GB.dic", "British English: LibreOffice's hyph_en_GB.dic, from ukhyphen.tex, conversion by László Németh; BSD-style licence (unlimited copying, redistribution and modification with this notice)."),
    "fr": ("fr_FR__dictionaries__hyph_fr.dic", "French: LibreOffice's hyph_fr.dic 4.0, based on hyph-fr.tex (MIT), by Marc Lodewijck, Paul Pichaureau, Blaise Drayer; LGPL-2.1 or later."),
    "es": ("es__hyph_es.dic", "Spanish: LibreOffice's hyph_es.dic from tex-hyphen-spanish by Javier Bezos (CervanTeX), adapted by Edward Villegas; MIT/X11 (patterns), LGPL/GPL/MPL (adaptation)."),
}


def build_hyphenation(lang: str, name: str, note: str) -> tuple[int, int]:
    lines = read(SRC / name)
    levels: list[list[str]] = [[]]
    mins = {"LEFTHYPHENMIN": 2, "RIGHTHYPHENMIN": 2, "COMPOUNDLEFTHYPHENMIN": 2, "COMPOUNDRIGHTHYPHENMIN": 2}
    for raw in lines:
        line = raw.strip()
        if not line or line[0] in "%#":
            continue
        if line == "NEXTLEVEL":
            levels.append([])
            continue
        key, _, val = line.partition(" ")
        if key in mins:
            mins[key] = int(val)
            continue
        if key == "NOHYPHEN" or "/" in line:
            continue
        pat = line.split()[0]
        if re.fullmatch(r"[0-9.'\-\w]+", pat):
            levels[-1].append(pat)
    # Without NEXTLEVEL all patterns are the syllable level (Hyphen's default compound level only splits at hyphens).
    compound, syllables = (levels[0], levels[1]) if len(levels) > 1 else ([], levels[0])
    HYPH_OUT.mkdir(parents=True, exist_ok=True)
    nl = chr(10)
    body = (f"// Generated by tools/lang-data/build.py. {note}{nl}{nl}"
            f"export default {{ left: {mins['LEFTHYPHENMIN']}, right: {mins['RIGHTHYPHENMIN']}, compoundLeft: {mins['COMPOUNDLEFTHYPHENMIN']}, "
            f"compoundRight: {mins['COMPOUNDRIGHTHYPHENMIN']},{nl}  compound: {json.dumps(' '.join(compound), ensure_ascii=False)},{nl}"
            f"  patterns: {json.dumps(' '.join(syllables), ensure_ascii=False)} }};{nl}")
    (HYPH_OUT / f"{lang}.ts").write_text(body, encoding="utf-8", newline=nl)
    return len(compound), len(syllables)


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    for lang, name in THESAURI.items():
        m, w, size = build(lang, name)
        print(f"{lang}: {m} meanings, {w} words, {size / 1e6:.2f} MB")
    for lang, (name, note) in HYPHENATION.items():
        c, n = build_hyphenation(lang, name, note)
        print(f"hyph {lang}: {c} compound and {n} syllable patterns")


if __name__ == "__main__":
    main()
