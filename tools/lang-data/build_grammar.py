"""Compiles LanguageTool's XML grammar rules (LGPL-2.1+) into the compact JSON that
RaabeWrite's own grammar checker reads (apps/raabe-write/src-tauri/src/grammar, D-051), and
their example sentences into a test file (tools/lang-data/work, not shipped).

Rules that need LanguageTool's Java code (filters), unification or chunks, and rules that are
off by default are left out; the Rust engine drops rules whose regular expressions it cannot
compile. Each kept rule keeps LanguageTool's id, so findings can be compared with its examples.

    python -I tools/lang-data/build_grammar.py de
"""
from __future__ import annotations

import gzip
import json
import re
import sys
import xml.etree.ElementTree as ET
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SRC = Path(__file__).parent / "src"
WORK = Path(__file__).parent / "work"
OUT = ROOT / "apps" / "raabe-write" / "src-tauri" / "data"

skipped: Counter = Counter()

# Typographic advice (protected spaces in "z. B.", typographic apostrophes and ellipses) that
# Word never gives and that fires on most correct texts (measured on Wikipedia, D-051).
DISABLED_BY_LANG = {
    "de": ("ABKUERZUNG_LEERZEICHEN", "AUSLASSUNGSPUNKTE", "AUSLASSUNGSPUNKTE_LEERZEICHEN", "TYPOGRAFISCHE_APOSTROPHE", "LEERZEICHEN_HINTER_DOPPELPUNKT",
           # Years read as quantities ("1891 kam die Volksinitiative"): LanguageTool tells them apart in Java.
           "VIELZAHL_PLUS_SINGULAR[2]"),
    "en": ("EN_QUOTES", "EN_ELLIPSIS"),
    # Dashes, protected spaces before units and the 1990 reform's spellings without hyphen are style.
    "fr": ("TIRET", "ESPACE_UNITES", "PAS_DE_TRAIT_UNION", "POINTS_SUSPENSION", "POINTS_SUSPENSIONS_SPACE",
           # A comma before "mais" is advice, not a rule; "l'Est de la France" is a region's name.
           "PLACE_DE_LA_VIRGULE[5]", "NORD_SUD[9]",
           # Needs LanguageTool's chunker: fires on most correct sentences with a long subject.
           "VIRG_NON_TROUVEE",
           # Postal address formats: an 18 MB expression of place names, nothing for school.
           "ADRESSES_FRANCE", "ADRESSES_BELGIQUE"),
    "es": ("COMILLAS_TIPOGRAFICAS", "GUION_LARGO", "PUNTOS_SUSPENSIVOS",
           # Commas before "pero" and "sino que" are advice; the agreement rules below need
           # LanguageTool's chunker ("dos o más palabras", "positivos o negativos", "los demás").
           "COMMA_PERO", "COMMA_SINO[1]", "AGREEMENT_NUMERAL_PLURAL[1]", "AGREEMENT_ADJ_NOUN[4]", "AGREEMENT_ADJ_NOUN[5]", "DEMAS[0]"),
}
DISABLED: tuple = ()


def yes(v: str | None) -> bool:
    return v == "yes"


def token_text(el: ET.Element) -> str:
    """A token's text: before its exceptions or after them (<token><exception/>I</token>)."""
    text = el.text or ""
    for c in el:
        if c.tag in ("exception", "match"):
            text += c.tail or ""
    return text.strip()


def match_ref(el: ET.Element) -> dict:
    m = {"no": int(el.get("no", "1"))}
    for a, k in (("case_conversion", "cc"), ("regexp_match", "rm"), ("regexp_replace", "rr"), ("postag", "pos"),
                 ("postag_replace", "pr"), ("include_skipped", "inc")):
        if el.get(a) is not None:
            m[k] = el.get(a)
    if yes(el.get("postag_regexp")):
        m["pre"] = True
    return m


def pieces(el: ET.Element) -> list:
    """Text with <match> references and inline <suggestion>s, as a list of pieces."""
    out: list = []
    if el.text:
        out.append(el.text)
    for c in el:
        if c.tag == "match":
            out.append({"m": match_ref(c)})
        elif c.tag == "suggestion":
            out.append({"s": pieces(c)})
        elif c.tag == "marker":
            out.extend(pieces(c))
        if c.tail:
            out.append(c.tail)
    return out


def token(el: ET.Element, case_default: bool) -> dict | None:
    if el.find("unify") is not None or el.get("chunk") or el.get("chunk_re"):
        return None
    t: dict = {}
    text = token_text(el)
    if text:
        t["t"] = text
    # <match no="0"/> inside a token: the same word as pattern token 0 (counted from 0 here; lemma when inflected).
    ref = el.find("match")
    if ref is not None:
        if ref.get("postag") or ref.get("regexp_match") or ref.get("postag_replace"):
            return None
        t["ref"] = int(ref.get("no", "0"))
    if yes(el.get("regexp")):
        t["re"] = True
    cs = el.get("case_sensitive")
    if (cs == "yes") or (cs is None and case_default):
        t["cs"] = True
    if yes(el.get("inflected")):
        t["inf"] = True
    if yes(el.get("negate")):
        t["neg"] = True
    if el.get("postag"):
        t["pos"] = el.get("postag")
        if yes(el.get("postag_regexp")):
            t["pre"] = True
    if yes(el.get("negate_pos")):
        t["npos"] = True
    if el.get("skip"):
        t["skip"] = int(el.get("skip"))
    if el.get("min"):
        t["min"] = int(el.get("min"))
    if el.get("max"):
        t["max"] = int(el.get("max"))
    if el.get("spacebefore") in ("yes", "no"):
        t["sb"] = el.get("spacebefore") == "yes"
    exc = []
    for e in el.findall("exception"):
        x: dict = {"scope": e.get("scope", "current")}
        if token_text(e):
            x["t"] = token_text(e)
        if yes(e.get("regexp")):
            x["re"] = True
        if yes(e.get("case_sensitive")):
            x["cs"] = True
        if yes(e.get("inflected")):
            x["inf"] = True
        if yes(e.get("negate")):
            x["neg"] = True
        if e.get("postag"):
            x["pos"] = e.get("postag")
            if yes(e.get("postag_regexp")):
                x["pre"] = True
        if yes(e.get("negate_pos")):
            x["npos"] = True
        if e.get("spacebefore") in ("yes", "no"):
            x["sb"] = e.get("spacebefore") == "yes"
        exc.append(x)
    if exc:
        t["exc"] = exc
    for kind in ("and", "or"):
        sub = el.find(kind)
        if sub is not None:
            parts = [token(s, case_default) for s in sub.findall("token")]
            if any(p is None for p in parts):
                return None
            t[kind] = parts
    return t


def pattern(el: ET.Element) -> tuple[list, int, int, list] | None:
    """The pattern's tokens and the marker's first and last token (indices). Tokens inside
    <unify> get "u": the index of their group in the returned list of unify groups; tokens in
    <unify-ignore> take part in the pattern only."""
    case_default = yes(el.get("case_sensitive"))
    toks: list = []
    mark: list[int] = []
    groups: list = []

    def walk(parent: ET.Element, marked: bool, group: int | None) -> bool:
        for c in parent:
            if c.tag == "token":
                t = token(c, case_default)
                if t is None:
                    return False
                if marked:
                    mark.append(len(toks))
                if group is not None:
                    t["u"] = group
                toks.append(t)
            elif c.tag == "marker":
                if not walk(c, True, group):
                    return False
            elif c.tag == "unify":
                if group is not None:
                    return False
                feats = [f.get("id") for f in c.findall("feature")]
                groups.append({"features": feats, **({"neg": True} if yes(c.get("negate")) else {})})
                if not walk(c, marked, len(groups) - 1):
                    return False
            elif c.tag == "unify-ignore":
                if not walk(c, marked, None):
                    return False
            elif c.tag in ("and", "or", "phrase", "phraseref"):
                return False
        return True

    if not walk(el, False, None) or not toks:
        return None
    a, b = (mark[0], mark[-1]) if mark else (0, len(toks) - 1)
    return toks, a, b, groups


def unification(lang: str, file: str = "disambiguation") -> dict:
    """The features unification compares (number, case, gender): their types as tag patterns,
    from grammar.xml or disambiguation.xml (each file defines its own)."""
    path = SRC / f"lt__{lang}__{file}.xml"
    out: dict = {}
    if path.exists():
        for u in parse(path).findall("unification"):
            out[u.get("feature")] = [[e.get("type"), e.find("token").get("postag")] for e in u.findall("equivalence")]
    return out


def example(e: ET.Element) -> dict:
    """An example sentence with the marked part (start, end) and its correction."""
    text, a, b = "", -1, -1
    if e.text:
        text += e.text
    for c in e:
        if c.tag == "marker":
            a = len(text)
            text += "".join(c.itertext())
            b = len(text)
        if c.tail:
            text += c.tail
    out: dict = {"text": text}
    if a >= 0:
        out["a"], out["b"] = a, b
    if e.get("correction") is not None:
        out["fix"] = e.get("correction")
    if e.get("type"):
        out["type"] = e.get("type")
    return out


def rule(el: ET.Element, group: ET.Element | None, cat: ET.Element, index: int) -> tuple[dict | None, list]:
    gid = group.get("id") if group is not None else el.get("id")
    rid = f"{gid}[{index}]" if group is not None else gid
    if gid in DISABLED or rid in DISABLED:
        skipped["typography"] += 1
        return None, []
    examples = [example(e) for e in el.findall("example")]
    off = (el.get("default") or (group.get("default") if group is not None else None)) in ("off", "temp_off")
    if off:
        skipped["off"] += 1
        return None, []
    if el.find("filter") is not None:
        skipped["filter"] += 1
        return None, []
    r: dict = {"id": rid, "cat": cat.get("id"), "name": el.get("name") or (group.get("name") if group is not None else "")}
    pat = el.find("pattern")
    rex = el.find("regexp")
    if pat is not None:
        p = pattern(pat)
        if p is None:
            skipped["unsupported pattern"] += 1
            return None, []
        r["pat"], r["ma"], r["mb"], groups = p
        if groups:
            r["unify"] = groups
    elif rex is not None:
        r["regex"] = {"re": (rex.text or "").strip(), "cs": yes(rex.get("case_sensitive")), "mark": int(rex.get("mark", "0"))}
    else:
        skipped["no pattern"] += 1
        return None, []
    anti, antiu = [], []
    for a in ([*group.findall("antipattern")] if group is not None else []) + el.findall("antipattern"):
        p = pattern(a)
        if p is None:
            skipped["unsupported antipattern"] += 1
            return None, []
        anti.append(p[0])
        antiu.append(p[3])
    if anti:
        r["anti"] = anti
        if any(antiu):
            r["antiu"] = antiu
    msg = el.find("message")
    if msg is None:
        skipped["no message"] += 1
        return None, []
    r["msg"] = pieces(msg)
    sugg = [pieces(s) for s in el.findall("suggestion")]
    if sugg:
        r["sugg"] = sugg
    short = el.find("short") if el.find("short") is not None else (group.find("short") if group is not None else None)
    if short is not None and short.text:
        r["short"] = short.text.strip()
    return r, [{"rule": rid, **x} for x in examples]


LEXICONS = {
    # language: (jar, folder in the jar, file stem, output stem)
    "de": ("german-pos-dict-1.2.4.jar", "org/languagetool/resource/de", "german", "german"),
    "en": ("english-pos-dict-0.6.jar", "org/languagetool/resource/en", "english", "english"),
    "fr": ("french-pos-dict-0.7.jar", "org/languagetool/resource/fr", "french", "french"),
    "es": ("spanish-pos-dict-2.5.jar", "org/languagetool/resource/es", "es-ES", "spanish"),
}


def lexicon(lang: str) -> None:
    """The POS and synthesis lexicons (Morfologik automata: german-pos-dict CC-BY-SA-4.0,
    english-pos-dict LGPL-2.1, french-pos-dict CC-BY-SA-4.0, spanish-pos-dict LGPL-2.1) and
    LanguageTool's additions and removals (LGPL-2.1+), as the Rust tagger reads them."""
    import zipfile
    jar, folder, stem, out = LEXICONS[lang]
    with zipfile.ZipFile(SRC / jar) as z:
        (OUT / f"{out}.dict").write_bytes(z.read(f"{folder}/{stem}.dict"))
        (OUT / f"{out}_synth.dict").write_bytes(z.read(f"{folder}/{stem}_synth.dict"))
    for name in ("added", "removed"):
        src = SRC / f"lt__{lang}__{name}.txt"
        data = src.read_bytes() if src.exists() else b""
        (OUT / f"{out}-{name}.txt.gz").write_bytes(gzip.compress(data, 9, mtime=0))


def parse(path: Path) -> ET.Element:
    """Parses a LanguageTool XML file; an external entity file (Spanish: entities.ent) is inlined."""
    text = path.read_text(encoding="utf-8")
    m = re.search(r'<!ENTITY\s+%\s+(\w+)\s+SYSTEM\s+"([^"]+)"\s*>', text)
    if m:
        ent = SRC / f"{path.name.split('__')[0]}__{path.name.split('__')[1]}__{Path(m.group(2)).name}"
        body = ent.read_text(encoding="utf-8") if ent.exists() else ""
        text = text.replace(m.group(0), "").replace(f"%{m.group(1)};", body)
    return ET.fromstring(text.encode("utf-8"))


def disambiguation(lang: str) -> list[dict]:
    """LanguageTool's disambiguation rules that change readings before the rules run (remove,
    add, filter, replace) or protect tokens (immunize); spelling-only ones are left out."""
    path = SRC / f"lt__{lang}__disambiguation.xml"
    if not path.exists():
        return []
    root = parse(path)
    out = []
    items = []
    for el in root:
        if el.tag == "rulegroup":
            items += [(el, r, i) for i, r in enumerate(el.findall("rule"))]
        elif el.tag == "rule":
            items.append((None, el, 0))
    for group, r, i in items:
        d = r.find("disambig")
        if d is None:
            continue
        action = d.get("action", "replace")
        if action not in ("remove", "add", "filter", "replace", "immunize", "unify", "filterall"):
            skipped["disambig " + action] += 1
            continue
        pat = r.find("pattern")
        p = pattern(pat) if pat is not None else None
        if p is None:
            skipped["disambig pattern"] += 1
            continue
        anti, antiu = [], []
        for a in ([*group.findall("antipattern")] if group is not None else []) + r.findall("antipattern"):
            ap = pattern(a)
            if ap is not None:
                anti.append(ap[0])
                antiu.append(ap[3])
        x = {"id": (group.get("id") if group is not None else r.get("id")) or "", "pat": p[0], "ma": p[1], "mb": p[2], "action": action}
        if p[3]:
            x["unify"] = p[3]
        if any(antiu):
            x["antiu"] = antiu
        if d.get("postag"):
            x["pos"] = d.get("postag")
        wds = [{k: v for k, v in (("lemma", w.get("lemma")), ("pos", w.get("pos")), ("text", (w.text or "").strip() or None)) if v} for w in d.findall("wd")]
        if wds:
            x["wd"] = wds
        if anti:
            x["anti"] = anti
        out.append(x)
    return out


def main() -> None:
    lang = sys.argv[1] if len(sys.argv) > 1 else "de"
    global DISABLED
    DISABLED = DISABLED_BY_LANG.get(lang, ())
    OUT.mkdir(parents=True, exist_ok=True)
    lexicon(lang)
    root = parse(SRC / f"lt__{lang}__grammar.xml")
    rules, examples = [], []
    for cat in root.iter("category"):
        if cat.get("default") == "off":
            continue
        for el in cat:
            if el.tag == "rulegroup":
                for i, r in enumerate(el.findall("rule")):
                    x, ex = rule(r, el, cat, i)
                    if x:
                        rules.append(x)
                        examples.extend(ex)
            elif el.tag == "rule":
                x, ex = rule(el, None, cat, 0)
                if x:
                    rules.append(x)
                    examples.extend(ex)
    disambig = disambiguation(lang)
    unify_rules = unification(lang, "grammar") or unification(lang)
    data = json.dumps({"rules": rules, "disambig": disambig, "unification": unify_rules, "unification_disambig": unification(lang)},
                      ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    OUT.mkdir(parents=True, exist_ok=True)
    (OUT / f"grammar-{lang}.json.gz").write_bytes(gzip.compress(data, 9, mtime=0))
    WORK.mkdir(exist_ok=True)
    (WORK / f"grammar-{lang}-examples.json").write_text(json.dumps(examples, ensure_ascii=False), encoding="utf-8", newline="\n")
    print(f"{lang}: {len(disambig)} disambiguation rules")
    print(f"{lang}: {len(rules)} rules ({len(data) / 1e6:.1f} MB, {len(gzip.compress(data, 9)) / 1e6:.2f} MB gz), {len(examples)} examples; skipped {dict(skipped)}")


if __name__ == "__main__":
    main()
