// RaabeWrite's own grammar checker (D-040, D-051): sentences and tokens (text.rs), German
// part-of-speech tags from LanguageTool's open lexicon (tagger.rs, fsa.rs) and LanguageTool's
// XML rules compiled into JSON and run by our own engine (rules.rs). Everything is compiled into
// the program and runs offline; each language loads on first use. Findings are returned with
// UTF-16 offsets, as the frontend counts.

mod de_rules;
mod en_rules;
mod es_rules;
mod fr_rules;
mod fsa;
mod minire;
mod rules;
mod tagger;
mod text;

use rules::{ATok, Disambig, Finding, Rule, RuleFile};
use std::collections::{HashMap, HashSet};
use std::io::Read;
use std::sync::{Mutex, OnceLock};
use tagger::{Reading, Tagger};

/// A language's compiled data: lexicon (separator, frequency letter), additions, removals, rules, synthesis.
struct LangData {
    key: &'static str,
    dict: &'static [u8],
    sep: u8,
    freq: bool,
    added: &'static [u8],
    removed: &'static [u8],
    rules: &'static [u8],
    synth: &'static [u8],
}

const LANGS: [LangData; 4] = [
    LangData { key: "de", dict: include_bytes!("../../data/german.dict"), sep: b'_', freq: false, added: include_bytes!("../../data/german-added.txt.gz"),
        removed: include_bytes!("../../data/german-removed.txt.gz"), rules: include_bytes!("../../data/grammar-de.json.gz"), synth: include_bytes!("../../data/german_synth.dict") },
    LangData { key: "en", dict: include_bytes!("../../data/english.dict"), sep: b'+', freq: false, added: include_bytes!("../../data/english-added.txt.gz"),
        removed: include_bytes!("../../data/english-removed.txt.gz"), rules: include_bytes!("../../data/grammar-en.json.gz"), synth: include_bytes!("../../data/english_synth.dict") },
    LangData { key: "fr", dict: include_bytes!("../../data/french.dict"), sep: b'_', freq: true, added: include_bytes!("../../data/french-added.txt.gz"),
        removed: include_bytes!("../../data/french-removed.txt.gz"), rules: include_bytes!("../../data/grammar-fr.json.gz"), synth: include_bytes!("../../data/french_synth.dict") },
    LangData { key: "es", dict: include_bytes!("../../data/spanish.dict"), sep: b'_', freq: true, added: include_bytes!("../../data/spanish-added.txt.gz"),
        removed: include_bytes!("../../data/spanish-removed.txt.gz"), rules: include_bytes!("../../data/grammar-es.json.gz"), synth: include_bytes!("../../data/spanish_synth.dict") },
];

pub struct Checker {
    lang: &'static str,
    tagger: Tagger,
    disambig: Vec<Disambig>,
    rules: Vec<Rule>,
    pub dropped: usize,
    disabled: HashSet<String>,
    synth: de_rules::Synth,
    unifier: std::sync::Arc<rules::Unifier>,
}

impl Checker {
    fn german() -> Option<Checker> {
        Checker::new("de")
    }

    fn new(key: &str) -> Option<Checker> {
        let d = LANGS.iter().find(|d| d.key == key)?;
        let tagger = Tagger::new(d.dict, d.added, d.removed, d.sep, d.freq, d.key)?;
        let mut json = String::new();
        flate2::read::GzDecoder::new(d.rules).read_to_string(&mut json).ok()?;
        let file: RuleFile = serde_json::from_str(&json).ok()?;
        let unifier = std::sync::Arc::new(rules::Unifier::new(&file.unification));
        let unifier_d = std::sync::Arc::new(rules::Unifier::new(if file.unification_disambig.is_empty() { &file.unification } else { &file.unification_disambig }));
        let total = file.rules.len();
        let disambig: Vec<Disambig> = file.disambig.into_iter().filter_map(|j| Disambig::new(j, &unifier_d)).collect();
        let rules: Vec<Rule> = file.rules.into_iter().filter_map(|j| Rule::new(j, &unifier)).collect();
        let dropped = total - rules.len();
        Some(Checker { lang: d.key, tagger, disambig, rules, dropped, disabled: HashSet::new(), synth: de_rules::Synth::new(d.synth, d.sep)?, unifier })
    }

    /// The analysed tokens of a sentence: SENT_START first, SENT_END on the last token.
    fn analyse(&self, toks: &[text::Token]) -> Vec<ATok> {
        let mut out = Vec::with_capacity(toks.len() + 1);
        out.push(ATok { text: String::new(), lower: String::new(), start: toks.first().map_or(0, |t| t.start), end: toks.first().map_or(0, |t| t.start),
            space_before: false, readings: vec![Reading { lemma: String::new(), tag: "SENT_START".into(), guessed: false }], immune: false });
        let mut first = true;
        for t in toks {
            let mut readings = self.tagger.tag(&t.text, first);
            if t.text.chars().any(char::is_alphanumeric) {
                first = false;
            }
            if readings.is_empty() {
                readings.push(Reading { lemma: t.text.clone(), tag: "UNKNOWN".into(), guessed: false });
            }
            // Rules write the plain apostrophe (n'est, don't); the text may have the typographic one.
            let text = t.text.replace('\u{2019}', "'");
            out.push(ATok { lower: text.to_lowercase(), text, start: t.start, end: t.end, space_before: t.space_before, readings, immune: false });
        }
        if out.len() > 1 {
            let last = out.last_mut().unwrap();
            let lemma = last.text.clone();
            last.readings.push(Reading { lemma, tag: "SENT_END".into(), guessed: false });
        }
        out
    }

    /// Findings in a paragraph (char offsets).
    pub fn check_chars(&self, chars: &[char]) -> Vec<Finding> {
        let toks = text::split_clitics(text::tokens(chars, 0, chars.len()), self.lang, |w| self.tagger.has(w));
        let mut out = Vec::new();
        for range in text::sentences(&toks) {
            let sent = &toks[range.clone()];
            if sent.is_empty() {
                continue;
            }
            let mut atoks = self.analyse(sent);
            // Disambiguation first: it narrows the readings the rules see.
            let lowers: Vec<String> = atoks.iter().map(|t| t.lower.clone()).collect();
            let present: HashSet<&str> = lowers.iter().map(String::as_str).collect();
            for d in &self.disambig {
                if d.needs(&present) {
                    d.apply(&mut atoks);
                }
            }
            let words: HashSet<&str> = atoks.iter().map(|t| t.lower.as_str()).collect();
            let span = (sent[0].start, sent[sent.len() - 1].end);
            for r in &self.rules {
                if self.disabled.contains(&r.id) || !r.literals.iter().all(|l| words.contains(l.as_str())) {
                    continue;
                }
                r.check(&atoks, chars, span, &mut out);
            }
            if self.lang == "de" {
                de_rules::check(&atoks, chars, &self.synth, &self.unifier, &mut out);
            }
            if self.lang == "en" {
                en_rules::check(&atoks, chars, &self.synth, &mut out);
            }
            if self.lang == "fr" {
                fr_rules::check(&atoks, chars, &self.synth, &mut out);
            }
            if self.lang == "es" {
                es_rules::check(&atoks, chars, &self.synth, &mut out);
            }
        }
        out.sort_by_key(|f| (f.start, f.end));
        out
    }
}

fn checkers() -> &'static Mutex<HashMap<&'static str, Option<Checker>>> {
    static C: OnceLock<Mutex<HashMap<&'static str, Option<Checker>>>> = OnceLock::new();
    C.get_or_init(|| Mutex::new(HashMap::new()))
}

/// A finding for the frontend: UTF-16 offsets.
#[derive(serde::Serialize)]
pub struct Issue {
    start: usize,
    end: usize,
    rule: String,
    category: String,
    message: String,
    short: String,
    suggestions: Vec<String>,
}

fn utf16(chars: &[char], f: Finding) -> Issue {
    let at = |i: usize| chars[..i.min(chars.len())].iter().map(|c| c.len_utf16()).sum();
    Issue { start: at(f.start), end: at(f.end), rule: f.rule, category: f.category, message: f.message, short: f.short, suggestions: f.suggestions }
}

/// Checks paragraphs in a language (German for now; other languages have no findings yet).
#[tauri::command]
pub async fn grammar_check(lang: String, paragraphs: Vec<String>) -> Vec<Vec<Issue>> {
    tauri::async_runtime::spawn_blocking(move || {
        let l = lang.to_ascii_lowercase();
        let Some(key) = LANGS.iter().map(|d| d.key).find(|k| l.starts_with(k)) else {
            return paragraphs.iter().map(|_| Vec::new()).collect();
        };
        let mut map = match checkers().lock() { Ok(m) => m, Err(_) => return Vec::new() };
        let c = map.entry(key).or_insert_with(|| Checker::new(key));
        let Some(c) = c.as_ref() else { return paragraphs.iter().map(|_| Vec::new()).collect() };
        paragraphs.iter().map(|p| {
            let chars: Vec<char> = p.chars().collect();
            c.check_chars(&chars).into_iter().map(|f| utf16(&chars, f)).collect()
        }).collect()
    })
    .await
    .unwrap_or_default()
}

/// Loads the German checker in the background, so the first check is quick.
pub fn preload() {
    std::thread::spawn(|| {
        if let Ok(mut map) = checkers().lock() {
            map.entry("de").or_insert_with(Checker::german);
        }
    });
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The language under test: $GRAMMAR_LANG (de by default).
    fn lang() -> String {
        std::env::var("GRAMMAR_LANG").unwrap_or_else(|_| "de".into())
    }

    fn checker() -> Checker {
        Checker::new(&lang()).expect("checker")
    }

    fn rules_hit(c: &Checker, s: &str) -> Vec<String> {
        let chars: Vec<char> = s.chars().collect();
        c.check_chars(&chars).into_iter().map(|f| f.rule).collect()
    }

    #[test]
    fn tags_german_words() {
        let c = Checker::german().unwrap();
        let tags: Vec<String> = c.tagger.tag("Kinder", false).into_iter().map(|r| r.tag).collect();
        assert!(tags.contains(&"SUB:NOM:PLU:NEU".to_string()));
        assert!(c.tagger.tag("Fotosynthese", false).iter().any(|r| r.tag.starts_with("SUB") && r.lemma == "Fotosynthese"));
        assert_eq!(c.tagger.tag(",", false)[0].tag, "PKT");
        eprintln!("rules: {} kept, {} dropped", c.rules.len(), c.dropped);
    }

    #[test]
    fn finds_typical_errors() {
        let c = Checker::german().unwrap();
        for s in ["Gut, das dass heute besser klappt.", "Alle sind am essen.", "Das ist ein schöner Tag."] {
            let chars: Vec<char> = s.chars().collect();
            let toks = text::tokens(&chars, 0, chars.len());
            let at = c.analyse(&toks);
            eprintln!("{s}: {:?}", at.iter().map(|t| format!("{}/{}", t.text, t.readings.iter().map(|r| r.tag.as_str()).collect::<Vec<_>>().join("|"))).collect::<Vec<_>>());
            eprintln!("  -> {:?}", c.check_chars(&chars).iter().map(|f| (f.rule.clone(), f.start, f.end, f.message.clone())).collect::<Vec<_>>());
        }
        assert!(rules_hit(&c, "Gut, das dass heute besser klappt.").iter().any(|r| r.starts_with("DAS_DASS")));
        assert!(rules_hit(&c, "Alle sind am essen.").iter().any(|r| r == "AM_VERB_KLEIN"));
        assert!(rules_hit(&c, "Das ist ein schöner Tag.").is_empty(), "{:?}", rules_hit(&c, "Das ist ein schöner Tag."));
    }

    /// Our own English, French and Spanish rules: a typical pupil's error each, with the
    /// suggestion, and a correct sentence that stays clean.
    #[test]
    fn finds_errors_in_other_languages() {
        let cases = [
            ("en", "She don't like apples.", "RAABE_EN_3RD_PERSON", "does"),
            ("en", "Does he likes music?", "RAABE_EN_BASE_FORM", "like"),
            ("en", "I have a apple.", "RAABE_EN_A_AN", "an"),
            ("fr", "Le maison est grande.", "RAABE_FR_DET_NOM", "La maison"),
            ("fr", "Il faut que tu viens.", "RAABE_FR_SUBJONCTIF", "viennes"),
            ("es", "Los niño juegan en el parque.", "RAABE_ES_DET_NOMBRE", "Los niños"),
            ("es", "Madrid es en España.", "RAABE_ES_SER_ESTAR", "está"),
        ];
        for (lang, s, rule, fix) in cases {
            let c = Checker::new(lang).unwrap();
            let chars: Vec<char> = s.chars().collect();
            let found = c.check_chars(&chars);
            assert!(found.iter().any(|f| f.rule == rule && f.suggestions.first().map(String::as_str) == Some(fix)), "{s}: {:?}",
                found.iter().map(|f| (&f.rule, &f.suggestions)).collect::<Vec<_>>());
        }
        for (lang, s) in [("en", "He made it work, and the dog wagged its tail."), ("fr", "Je pense que tu viens demain."), ("es", "La fiesta es en mi casa.")] {
            let c = Checker::new(lang).unwrap();
            assert!(rules_hit(&c, s).is_empty(), "{s}: {:?}", rules_hit(&c, s));
        }
    }

    /// Counts the bytes in use, for `memory_use`.
    struct Counting;
    static LIVE: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);
    unsafe impl std::alloc::GlobalAlloc for Counting {
        unsafe fn alloc(&self, l: std::alloc::Layout) -> *mut u8 {
            LIVE.fetch_add(l.size(), std::sync::atomic::Ordering::Relaxed);
            unsafe { std::alloc::System.alloc(l) }
        }
        unsafe fn dealloc(&self, p: *mut u8, l: std::alloc::Layout) {
            LIVE.fetch_sub(l.size(), std::sync::atomic::Ordering::Relaxed);
            unsafe { std::alloc::System.dealloc(p, l) }
        }
    }
    #[global_allocator]
    static ALLOC: Counting = Counting;

    /// Memory in use after loading each language's checker (and its parts with
    /// $GRAMMAR_MEM_PARTS=xx).
    #[test]
    #[ignore]
    fn memory_use() {
        let private = || LIVE.load(std::sync::atomic::Ordering::Relaxed) / (1 << 20);
        eprintln!("start {} MB", private());
        if let Ok(l) = std::env::var("GRAMMAR_MEM_PARTS") {
            let d = LANGS.iter().find(|d| d.key == l).unwrap();
            let _tagger = Tagger::new(d.dict, d.added, d.removed, d.sep, d.freq, d.key).unwrap();
            eprintln!("tagger: {} MB", private());
            let mut json = String::new();
            flate2::read::GzDecoder::new(d.rules).read_to_string(&mut json).unwrap();
            let file: RuleFile = serde_json::from_str(&json).unwrap();
            drop(json);
            let unifier = std::sync::Arc::new(rules::Unifier::new(&file.unification));
            let _dis: Vec<Disambig> = file.disambig.into_iter().filter_map(|j| Disambig::new(j, &unifier)).collect();
            let mut rs: Vec<Rule> = Vec::new();
            let mut heavy: Vec<(usize, String)> = Vec::new();
            for j in file.rules {
                let id = j.id.clone();
                let before = LIVE.load(std::sync::atomic::Ordering::Relaxed);
                rs.extend(Rule::new(j, &unifier));
                heavy.push((LIVE.load(std::sync::atomic::Ordering::Relaxed).saturating_sub(before), id));
            }
            eprintln!("all: {} MB ({} rules, {} expressions for fancy-regex)", private(), rs.len(), rules::fancy_count());
            heavy.sort_by(|a, b| b.0.cmp(&a.0));
            for (b, id) in heavy.iter().take(12) {
                eprintln!("  {} KB {id}", b / 1024);
            }
            return;
        }
        let mut keep = Vec::new();
        for l in ["de", "en", "fr", "es"] {
            keep.push(Checker::new(l).unwrap());
            eprintln!("after {l}: {} MB", private());
        }
    }

    /// How long loading a language takes, by phase, and the slowest rules to compile.
    #[test]
    #[ignore]
    fn load_time() {
        let d = LANGS.iter().find(|d| d.key == lang()).unwrap();
        let t = std::time::Instant::now();
        let _tagger = Tagger::new(d.dict, d.added, d.removed, d.sep, d.freq, d.key).unwrap();
        eprintln!("tagger {:?}", t.elapsed());
        let t = std::time::Instant::now();
        let mut json = String::new();
        flate2::read::GzDecoder::new(d.rules).read_to_string(&mut json).unwrap();
        let file: RuleFile = serde_json::from_str(&json).unwrap();
        eprintln!("json {:?}", t.elapsed());
        let unifier = std::sync::Arc::new(rules::Unifier::new(&file.unification));
        let t = std::time::Instant::now();
        let n = file.disambig.len();
        let _d: Vec<Disambig> = file.disambig.into_iter().filter_map(|j| Disambig::new(j, &unifier)).collect();
        eprintln!("{n} disambiguation rules {:?}", t.elapsed());
        let t = std::time::Instant::now();
        let mut slow: Vec<(std::time::Duration, String)> = Vec::new();
        for j in file.rules {
            let id = j.id.clone();
            let s = std::time::Instant::now();
            let _ = Rule::new(j, &unifier);
            slow.push((s.elapsed(), id));
        }
        eprintln!("rules {:?}", t.elapsed());
        slow.sort_by(|a, b| b.0.cmp(&a.0));
        for (d, id) in slow.iter().take(15) {
            eprintln!("  {d:?} {id}");
        }
    }

    /// Prints the tokens, readings and findings of $GRAMMAR_DEBUG (for finding engine differences).
    #[test]
    #[ignore]
    fn debug_text() {
        let Ok(s) = std::env::var("GRAMMAR_DEBUG") else { return };
        let c = checker();
        let chars: Vec<char> = s.chars().collect();
        let toks = text::split_clitics(text::tokens(&chars, 0, chars.len()), &lang(), |w| c.tagger.has(w));
        for range in text::sentences(&toks) {
            let mut at = c.analyse(&toks[range.clone()]);
            let lowers: Vec<String> = at.iter().map(|t| t.lower.clone()).collect();
            let present: HashSet<&str> = lowers.iter().map(String::as_str).collect();
            for d in &c.disambig {
                if d.needs(&present) {
                    d.apply(&mut at);
                }
            }
            for t in &at {
                eprintln!("  {:>3}-{:<3} {:20} {}{}", t.start, t.end, t.text, t.readings.iter().map(|r| format!("{}/{}", r.lemma, r.tag)).collect::<Vec<_>>().join(" "), if t.immune { " IMMUNE" } else { "" });
            }
        }
        for f in c.check_chars(&chars) {
            let seen: String = chars[f.start..f.end].iter().collect();
            eprintln!("{} [{}..{}] '{}': {} {:?}", f.rule, f.start, f.end, seen, f.message, f.suggestions);
        }
    }

    /// For tools/lang-data/grammar_measure.py: checks the paragraphs of $GRAMMAR_IN (a JSON list)
    /// and writes the findings to $GRAMMAR_OUT.
    #[test]
    #[ignore]
    fn check_file() {
        let (Ok(input), Ok(output)) = (std::env::var("GRAMMAR_IN"), std::env::var("GRAMMAR_OUT")) else { return };
        let paras: Vec<String> = serde_json::from_str(&std::fs::read_to_string(input).unwrap()).unwrap();
        let c = checker();
        let t0 = std::time::Instant::now();
        let out: Vec<Vec<serde_json::Value>> = paras.iter().map(|p| {
            let chars: Vec<char> = p.chars().collect();
            c.check_chars(&chars).into_iter().map(|f| serde_json::json!({"start": f.start, "end": f.end, "rule": f.rule, "category": f.category,
                "message": f.message, "suggestions": f.suggestions})).collect()
        }).collect();
        eprintln!("checked {} paragraphs in {:?}", paras.len(), t0.elapsed());
        std::fs::write(output, serde_json::to_string(&out).unwrap()).unwrap();
    }

    /// LanguageTool's example sentences (tools/lang-data/work/grammar-de-examples.json, built by
    /// build_grammar.py): each rule should find its incorrect examples and leave the correct
    /// ones alone. Writes a report per rule to tools/lang-data/work/grammar-de-report.json.
    #[test]
    #[ignore]
    fn language_tool_examples() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join(format!("../../../tools/lang-data/work/grammar-{}-examples.json", lang()));
        let Ok(text) = std::fs::read_to_string(&path) else { eprintln!("no examples at {path:?}"); return };
        let examples: Vec<serde_json::Value> = serde_json::from_str(&text).unwrap();
        let c = checker();
        if std::env::var_os("GRAMMAR_PROFILE").is_some() {
            *minire::PROFILE.lock().unwrap() = Some(HashMap::new());
        }
        let known: HashSet<String> = c.rules.iter().map(|r| r.id.clone()).collect();
        let mut stats: HashMap<String, [u32; 4]> = HashMap::new();
        let t0 = std::time::Instant::now();
        let (mut sentences, mut chars_total) = (0usize, 0usize);
        for e in &examples {
            let rule = e["rule"].as_str().unwrap().to_string();
            if !known.contains(&rule) {
                continue;
            }
            let s = e["text"].as_str().unwrap();
            let chars: Vec<char> = s.chars().collect();
            sentences += 1;
            chars_total += chars.len();
            let found = c.check_chars(&chars);
            let st = stats.entry(rule.clone()).or_default();
            let incorrect = e.get("fix").is_some() || e.get("type").and_then(|t| t.as_str()) == Some("incorrect");
            if incorrect {
                // The marked part, as char offsets (Python counted code points too).
                let (a, b) = (e["a"].as_u64().unwrap_or(0) as usize, e["b"].as_u64().unwrap_or(chars.len() as u64) as usize);
                let hit = found.iter().any(|f| f.rule == rule && f.start < b && a < f.end);
                st[0] += 1;
                if hit { st[1] += 1; }
            } else if e.get("type").and_then(|t| t.as_str()) != Some("triggers_error") {
                st[2] += 1;
                if found.iter().any(|f| f.rule == rule) { st[3] += 1; }
            }
        }
        let elapsed = t0.elapsed();
        let (mut inc, mut hit, mut cor, mut fa) = (0, 0, 0, 0);
        for s in stats.values() {
            inc += s[0]; hit += s[1]; cor += s[2]; fa += s[3];
        }
        eprintln!("incorrect examples found: {hit}/{inc}; correct examples flagged: {fa}/{cor}; {sentences} sentences, {chars_total} chars in {elapsed:?}; {} MB in use",
            LIVE.load(std::sync::atomic::Ordering::Relaxed) >> 20);
        if let Some(m) = minire::PROFILE.lock().unwrap().take() {
            let mut v: Vec<_> = m.into_iter().collect();
            v.sort_by(|a, b| b.1 .1.cmp(&a.1 .1));
            let total: u128 = v.iter().map(|x| x.1 .1).sum();
            eprintln!("matcher time {} ms", total / 1_000_000);
            for (src, (n, ns)) in v.iter().take(25) {
                eprintln!("  {} ms {n} calls {src}", ns / 1_000_000);
            }
        }
        let report: serde_json::Map<String, serde_json::Value> = stats.into_iter().map(|(k, v)| (k, serde_json::json!(v))).collect();
        let out = path.with_file_name(format!("grammar-{}-report.json", lang()));
        std::fs::write(out, serde_json::to_string(&report).unwrap()).unwrap();
    }
}
