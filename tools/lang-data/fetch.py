"""Fetches the open language data RaabeWrite builds its thesaurus and hyphenation from
(Step 4b F, D-049) into tools/lang-data/src/ (not in git; `build.py` turns it into the
compact files the app ships). All files come from LibreOffice's dictionaries repository:

- Thesaurus: German OpenThesaurus (LGPL-2.1+), English from WordNet (WordNet licence),
  French by Frédéric Labbé / Dicollecte (LGPL-2.1+), Spanish OpenThesaurus (LGPL-2.1).
- Hyphenation patterns (Hyphen format): German (LGPL-2+, from TeX's dehyphn), English US and
  GB (BSD-style), French (LGPL-2.1+, from hyph-fr, MIT), Spanish (LGPL/GPL/MPL).

    python tools/lang-data/fetch.py
"""
from __future__ import annotations

import sys
import urllib.request
from pathlib import Path

BASE = "https://raw.githubusercontent.com/LibreOffice/dictionaries/master/"
FILES = [
    "de/th_de_DE_v2.dat", "de/README_thesaurus.txt", "de/hyph_de_DE.dic", "de/README_hyph_de.txt",
    "en/th_en_US_v2.dat", "en/WordNet_license.txt", "en/hyph_en_US.dic", "en/README_hyph_en_US.txt",
    "en/hyph_en_GB.dic", "en/README_hyph_en_GB.txt",
    "fr_FR/dictionaries/thes_fr.dat", "fr_FR/dictionaries/README_thes_fr.txt",
    "fr_FR/dictionaries/hyph_fr.dic", "fr_FR/dictionaries/README_hyph_fr.txt",
    "es/th_es_v2.dat", "es/README_th_es.txt", "es/hyph_es.dic", "es/README_hyph_es.txt", "es/LICENSE.md",
]
SRC = Path(__file__).parent / "src"


def main() -> None:
    SRC.mkdir(exist_ok=True)
    for f in FILES:
        dest = SRC / f.replace("/", "__")
        if dest.exists() and dest.stat().st_size > 0:
            continue
        print("fetch", f, flush=True)
        with urllib.request.urlopen(BASE + f, timeout=120) as r:
            dest.write_bytes(r.read())
    print("ok", sum((SRC / f.replace("/", "__")).stat().st_size for f in FILES), "bytes")


if __name__ == "__main__":
    sys.exit(main())
