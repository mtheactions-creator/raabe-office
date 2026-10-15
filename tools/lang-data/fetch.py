"""Fetches the open language data RaabeWrite builds its thesaurus and hyphenation from
(Step 4b F, D-049) into tools/lang-data/src/ (not in git; `build.py` turns it into the
compact files the app ships). All files come from LibreOffice's dictionaries repository:

- Thesaurus: German OpenThesaurus (LGPL-2.1+), English from WordNet (WordNet licence),
  French by Frédéric Labbé / Dicollecte (LGPL-2.1+), Spanish OpenThesaurus (LGPL-2.1).
- Hyphenation patterns (Hyphen format): German (LGPL-2+, from TeX's dehyphn), English US and
  GB (BSD-style), French (LGPL-2.1+, from hyph-fr, MIT), Spanish (LGPL/GPL/MPL).

For the grammar checker (D-040, D-051): LanguageTool's rules (grammar.xml, disambiguation.xml
and word lists, LGPL-2.1+) from its GitHub repository and the POS lexicons from Maven Central
(german-pos-dict by Daniel Naber, CC-BY-SA-4.0; english-, french- and spanish-pos-dict).

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
LT = "https://raw.githubusercontent.com/languagetool-org/languagetool/master/languagetool-language-modules/"
LT_FILES = [
    "de/src/main/resources/org/languagetool/rules/de/grammar.xml",
    "de/src/main/resources/org/languagetool/resource/de/disambiguation.xml",
    "de/src/main/resources/org/languagetool/resource/de/tagset.txt",
    "de/src/main/resources/org/languagetool/resource/de/added.txt",
    "de/src/main/resources/org/languagetool/resource/de/removed.txt",
    "de/src/main/resources/org/languagetool/rules/de/replace.txt",
    "de/src/main/resources/org/languagetool/resource/de/case_rule_exceptions.txt",
    "en/src/main/resources/org/languagetool/rules/en/grammar.xml",
    "en/src/main/resources/org/languagetool/resource/en/disambiguation.xml",
    "en/src/main/resources/org/languagetool/resource/en/added.txt",
    "en/src/main/resources/org/languagetool/resource/en/removed.txt",
    "fr/src/main/resources/org/languagetool/rules/fr/grammar.xml",
    "fr/src/main/resources/org/languagetool/resource/fr/disambiguation.xml",
    "fr/src/main/resources/org/languagetool/resource/fr/added.txt",
    "fr/src/main/resources/org/languagetool/resource/fr/removed.txt",
    "es/src/main/resources/org/languagetool/rules/es/grammar.xml",
    "es/src/main/resources/org/languagetool/resource/es/disambiguation.xml",
    "es/src/main/resources/org/languagetool/resource/es/added.txt",
    "es/src/main/resources/org/languagetool/resource/es/removed.txt",
    "es/src/main/resources/org/languagetool/resource/es/entities.ent",
]
MAVEN = {
    "german-pos-dict-1.2.4.jar": "https://repo1.maven.org/maven2/de/danielnaber/german-pos-dict/1.2.4/german-pos-dict-1.2.4.jar",
    "english-pos-dict-0.6.jar": "https://repo1.maven.org/maven2/org/languagetool/english-pos-dict/0.6/english-pos-dict-0.6.jar",
    "french-pos-dict-0.7.jar": "https://repo1.maven.org/maven2/org/languagetool/french-pos-dict/0.7/french-pos-dict-0.7.jar",
    "spanish-pos-dict-2.5.jar": "https://repo1.maven.org/maven2/org/softcatala/spanish-pos-dict/2.5/spanish-pos-dict-2.5.jar",
}
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
    urls = [(BASE + f, f.replace("/", "__")) for f in FILES]
    urls += [(LT + f, "lt__" + f.split("/")[0] + "__" + f.rsplit("/", 1)[1]) for f in LT_FILES]
    urls += [(u, n) for n, u in MAVEN.items()]
    for url, name in urls[len(FILES):]:
        dest = SRC / name
        if dest.exists() and dest.stat().st_size > 0:
            continue
        print("fetch", url, flush=True)
        with urllib.request.urlopen(url, timeout=180) as r:
            dest.write_bytes(r.read())
    print("ok", sum((SRC / n).stat().st_size for _, n in urls), "bytes")


if __name__ == "__main__":
    sys.exit(main())
