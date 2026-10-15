# Language data compiled into RaabeWrite

Built by `tools/lang-data/build.py` from files fetched by `tools/lang-data/fetch.py`
(LibreOffice's dictionaries repository, https://github.com/LibreOffice/dictionaries).
Each thesaurus is one gzip file: its meanings once each, then the headwords with the numbers
of their meanings (see build.py). Generic and related terms are left out.

| File | Source | Licence |
|---|---|---|
| `thes-de.gz` | OpenThesaurus, Daniel Naber and contributors (de/th_de_DE_v2.dat) | LGPL-2.1 or later |
| `thes-en.gz` | WordNet-based English thesaurus (en/th_en_US_v2.dat), WordNet 2.1 © 2005 Princeton University | WordNet licence (permissive, see below) |
| `thes-fr.gz` | Dictionnaire des synonymes français, Frédéric Labbé / Dicollecte (fr_FR/thes_fr.dat) | LGPL-2.1 or later |
| `thes-es.gz` | OpenThesaurus Tesauro Español, Marcelo Garrone (es/th_es_v2.dat) | LGPL-2.1 |

WordNet licence: "Permission to use, copy, modify and distribute this software and database
and its documentation for any purpose and without fee or royalty is hereby granted, provided
that you agree to comply with the following copyright notice and statements, including the
disclaimer, and that the same appear on ALL copies of the software, database and
documentation, including modifications that you make for internal use or for distribution.
WordNet 2.1 Copyright 2005 by Princeton University. All rights reserved. THIS SOFTWARE AND
DATABASE IS PROVIDED "AS IS" AND PRINCETON UNIVERSITY MAKES NO REPRESENTATIONS OR WARRANTIES,
EXPRESS OR IMPLIED." The name of Princeton University may not be used in advertising or
publicity pertaining to distribution of the software and/or database.

The spelling dictionaries come from the npm packages `dictionary-de` (igerman98, GPL-2.0 or
GPL-3.0), `dictionary-en` and `dictionary-en-gb` (SCOWL, MIT and BSD), `dictionary-fr`
(Dicollecte, MPL-2.0) and `dictionary-es` (RLA-ES, GPL-3.0 or LGPL-3.0 or MPL-1.1); they are
compiled in from `node_modules` (`src/spell.rs`).

## Grammar checker (D-051)

| File | Source | Licence |
|---|---|---|
| `german.dict`, `german_synth.dict` | german-pos-dict 1.2.4 by Daniel Naber (Morfologik automata, from Morphy data) | CC-BY-SA-4.0 |
| `english.dict`, `english_synth.dict` | english-pos-dict 0.6 (LanguageTool, Morfologik automata) | LGPL-2.1 |
| `french.dict`, `french_synth.dict` | french-pos-dict 0.7 (LanguageTool, from Dicollecte's lexicon) | CC-BY-SA-4.0 |
| `spanish.dict`, `spanish_synth.dict` | spanish-pos-dict 2.5 (Softcatalà) | LGPL-2.1 |
| `*-added.txt.gz`, `*-removed.txt.gz` | LanguageTool's additions to and removals from these lexicons | LGPL-2.1 or later |
| `grammar-de.json.gz`, `grammar-en.json.gz`, `grammar-fr.json.gz`, `grammar-es.json.gz` | LanguageTool's rules (grammar.xml) and disambiguation (disambiguation.xml) per language, compiled by `tools/lang-data/build_grammar.py` | LGPL-2.1 or later |

LanguageTool: © LanguageTool contributors, https://languagetool.org, https://github.com/languagetool-org/languagetool.
german-pos-dict: https://github.com/languagetool-org/german-pos-dict. english-pos-dict and
french-pos-dict: https://github.com/languagetool-org (Maven org.languagetool). spanish-pos-dict:
https://github.com/Softcatala/spanish-dict-tools (Maven org.softcatala).

