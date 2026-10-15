"""Measures RaabeWrite's grammar checker against Word's offline grammar check (D-040, D-051).

Corpora (sentence level):
- errors: tools/lang-data/school-errors-de.txt (typical school errors, written for this) and a
  fixed sample of LanguageTool's incorrect example sentences;
- clean: paragraphs of German Wikipedia articles on school topics (CC-BY-SA, fetched once into
  tools/lang-data/work/, not in git) and a sample of LanguageTool's correct examples.
Word checks them through COM in its own invisible instance (Document.GrammaticalErrors; Word's
options are not changed), ours through the Rust checker (cargo test check_file). Reports how many
error sentences each finds, how many clean paragraphs each flags, and our rules behind false
alarms. Results: tools/feel-lab/results/<date>-grammar-de.json.

    python tools/lang-data/grammar_measure.py [de|en|fr|es] [--no-word]

Word has no Spanish proofing tools on the development PC, so Spanish is measured without it.
"""
from __future__ import annotations

import datetime
import json
import os
import random
import subprocess
import sys
import tempfile
import time
import urllib.parse
from collections import Counter
from pathlib import Path

HERE = Path(__file__).parent
ROOT = HERE.parents[1]
WORK = HERE / "work"
TAURI = ROOT / "apps" / "raabe-write" / "src-tauri"
RESULTS = ROOT / "tools" / "feel-lab" / "results"
TITLES_EN = ["Photosynthesis", "Cellular respiration", "French Revolution", "Industrial Revolution", "Water cycle", "Separation of powers",
             "Middle Ages", "Atmosphere of Earth", "Electric current", "Probability", "Ecosystem", "Forest", "Roman Empire", "Reformation",
             "Johann Wolfgang von Goethe", "Poetry", "Democracy", "European Union", "Volcano", "Bee", "Atom", "Chemical reaction",
             "Pythagorean theorem", "Alps", "Rhine", "Braunschweig", "Thirty Years' War", "Cell (biology)", "World War I", "Weather"]
TITLES_FR = ["Photosynthèse", "Respiration cellulaire", "Révolution française", "Révolution industrielle", "Cycle de l'eau", "Séparation des pouvoirs",
             "Moyen Âge", "Atmosphère terrestre", "Courant électrique", "Probabilité", "Écosystème", "Forêt", "Empire romain", "Réforme protestante",
             "Johann Wolfgang von Goethe", "Poésie", "Démocratie", "Union européenne", "Volcan", "Abeille", "Atome", "Réaction chimique",
             "Théorème de Pythagore", "Alpes", "Rhin", "Brunswick (Basse-Saxe)", "Guerre de Trente Ans", "Cellule (biologie)", "Première Guerre mondiale", "Météorologie"]
TITLES_ES = ["Fotosíntesis", "Respiración celular", "Revolución francesa", "Revolución Industrial", "Ciclo hidrológico", "Separación de poderes",
             "Edad Media", "Atmósfera terrestre", "Corriente eléctrica", "Probabilidad", "Ecosistema", "Bosque", "Imperio romano", "Reforma protestante",
             "Johann Wolfgang von Goethe", "Poesía", "Democracia", "Unión Europea", "Volcán", "Apis mellifera", "Átomo", "Reacción química",
             "Teorema de Pitágoras", "Alpes", "Rin", "Brunswick", "Guerra de los Treinta Años", "Célula", "Primera Guerra Mundial", "Tiempo atmosférico"]
WORD_LANG = {"de": 1031, "en": 1033, "fr": 1036, "es": None}
TITLES = ["Fotosynthese", "Zellatmung", "Französische Revolution", "Industrialisierung", "Wasserkreislauf", "Gewaltenteilung",
          "Mittelalter", "Erdatmosphäre", "Elektrischer Strom", "Wahrscheinlichkeitsrechnung", "Ökosystem", "Wald", "Römisches Reich",
          "Reformation", "Johann Wolfgang von Goethe", "Gedicht", "Demokratie", "Europäische Union", "Vulkan", "Biene", "Atom",
          "Chemische Reaktion", "Satz des Pythagoras", "Alpen", "Rhein", "Braunschweig", "Dreißigjähriger Krieg", "Zelle (Biologie)",
          "Erster Weltkrieg", "Wetter"]


def wiki(lang: str = "de", n_max: int = 500) -> list[str]:
    cache = WORK / f"wiki-{lang}.json"
    if cache.exists():
        return json.loads(cache.read_text(encoding="utf-8"))
    paras: list[str] = []
    titles = {"de": TITLES, "en": TITLES_EN, "fr": TITLES_FR, "es": TITLES_ES}[lang]
    for title in titles:
        url = f"https://{lang}.wikipedia.org/w/api.php?" + urllib.parse.urlencode(
            {"action": "query", "prop": "extracts", "explaintext": 1, "titles": title, "format": "json", "redirects": 1})
        # curl uses Windows' certificate store (Python's own may be out of date).
        raw = subprocess.run(["curl", "-s", "-A", "RaabeOffice/0.1 (https://github.com/mtheactions-creator/raabe-office; grammar measurement)", url],
                             capture_output=True, check=True).stdout
        time.sleep(2)
        try:
            pages = json.loads(raw)["query"]["pages"]
        except ValueError:
            print("not available:", title, raw[:80])
            continue
        text = next(iter(pages.values())).get("extract", "")
        for p in text.split("\n"):
            p = p.strip()
            if len(p) >= 80 and not p.startswith("=") and p.count(".") >= 1 and "\\" not in p and "{" not in p:
                paras.append(p)
    random.Random(1).shuffle(paras)
    paras = paras[:n_max]
    WORK.mkdir(exist_ok=True)
    cache.write_text(json.dumps(paras, ensure_ascii=False), encoding="utf-8")
    return paras


def corpora(lang: str) -> dict[str, list[dict]]:
    school = []
    for line in (HERE / f"school-errors-{lang}.txt").read_text(encoding="utf-8").splitlines():
        if line and not line.startswith("#"):
            kind, text = line.split("\t", 1)
            school.append({"kind": kind, "text": text})
    examples = json.loads((WORK / f"grammar-{lang}-examples.json").read_text(encoding="utf-8"))
    rnd = random.Random(7)
    bad = [e for e in examples if e.get("fix") is not None]
    good = [e for e in examples if e.get("fix") is None and "a" not in e and e.get("type") != "triggers_error"]
    return {
        "school errors": school,
        "LT incorrect examples": [{"kind": e["rule"], "text": e["text"]} for e in rnd.sample(bad, min(400, len(bad)))],
        "Wikipedia": [{"kind": "clean", "text": t} for t in wiki(lang)],
        "LT correct examples": [{"kind": e["rule"], "text": e["text"]} for e in rnd.sample(good, min(400, len(good)))],
    }


def ours(texts: list[str], lang: str) -> list[list[dict]]:
    with tempfile.TemporaryDirectory() as tmp:
        inp, out = Path(tmp) / "in.json", Path(tmp) / "out.json"
        inp.write_text(json.dumps(texts, ensure_ascii=False), encoding="utf-8")
        env = {**os.environ, "GRAMMAR_IN": str(inp), "GRAMMAR_OUT": str(out), "GRAMMAR_LANG": lang, "PATH": str(Path.home() / ".cargo" / "bin") + os.pathsep + os.environ["PATH"]}
        subprocess.run(["cargo", "test", "--release", "--lib", "grammar::tests::check_file", "--", "--ignored", "--nocapture"],
                       cwd=TAURI, env=env, check=True, capture_output=True)
        return json.loads(out.read_text(encoding="utf-8"))


def word(texts: list[str], lang: str) -> list[bool]:
    """Whether Word's grammar check flags each text (cached in work/word-grammar-cache[-lang].json)."""
    cache_path = WORK / ("word-grammar-cache.json" if lang == "de" else f"word-grammar-cache-{lang}.json")
    cache = json.loads(cache_path.read_text(encoding="utf-8")) if cache_path.exists() else {}
    todo = [t for t in texts if t not in cache]
    if todo:
        for t, v in zip(todo, word_check(todo, WORD_LANG[lang])):
            cache[t] = v
        cache_path.write_text(json.dumps(cache, ensure_ascii=False), encoding="utf-8")
    return [cache[t] for t in texts]


def word_check(texts: list[str], language_id: int) -> list[bool]:
    """Whether Word's grammar check flags each text (its own invisible instance; options unchanged)."""
    import win32com.client
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    try:
        d = app.Documents.Add()
        d.Content.Text = "\r".join(t.replace("\r", " ").replace("\n", " ") for t in texts)
        d.Content.LanguageID = language_id
        # Per paragraph: reading the document's error collection item by item is quadratic.
        flagged = [p.Range.GrammaticalErrors.Count > 0 for p in d.Paragraphs][:len(texts)]
        d.Close(0)
        return flagged
    finally:
        app.Quit(0)


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    lang = next((a for a in sys.argv[1:] if a in WORD_LANG), "de")
    use_word = "--no-word" not in sys.argv and WORD_LANG[lang] is not None
    data = corpora(lang)
    report: dict = {"date": datetime.date.today().isoformat(), "language": lang, "corpora": {}}
    for name, items in data.items():
        texts = [x["text"] for x in items]
        found = ours(texts, lang)
        w = word(texts, lang) if use_word else None
        n = len(texts)
        r = {"items": n, "ours": sum(1 for f in found if f), "word": sum(w) if w else None}
        r["both"] = sum(1 for i in range(n) if found[i] and w and w[i]) if w else None
        rules = Counter(f["rule"] for fs in found for f in fs)
        r["our rules"] = rules.most_common(40)
        if name == "school errors":
            kinds = Counter(x["kind"] for x in items)
            r["by kind"] = {k: {"of": kinds[k], "ours": sum(1 for i, x in enumerate(items) if x["kind"] == k and found[i]),
                                "word": sum(1 for i, x in enumerate(items) if x["kind"] == k and w and w[i]) if w else None} for k in kinds}
            r["missed by us"] = [x["text"] for i, x in enumerate(items) if not found[i]]
            r["word found, we did not"] = [x["text"] for i, x in enumerate(items) if w and w[i] and not found[i]]
        if name in ("Wikipedia", "LT correct examples"):
            r["our false alarms"] = [{"text": texts[i][max(0, f["start"] - 40):f["end"] + 40], "rule": f["rule"], "message": f["message"]}
                                     for i, fs in enumerate(found) for f in fs][:200]
        report["corpora"][name] = r
        print(f"{name}: {n} items; ours {r['ours']}, Word {r['word']}, both {r['both']}", flush=True)
    # The full report with text excerpts stays local (Wikipedia's texts are CC BY-SA); the
    # repository keeps counts, rule ids and messages.
    WORK.mkdir(exist_ok=True)
    (WORK / f"{report['date']}-grammar-{lang}-full.json").write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    for r in report["corpora"].values():
        for f in r.get("our false alarms", []):
            f.pop("text", None)
    RESULTS.mkdir(parents=True, exist_ok=True)
    out = RESULTS / f"{report['date']}-grammar-{lang}.json"
    out.write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print("saved to", out)


if __name__ == "__main__":
    main()
