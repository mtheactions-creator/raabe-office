// RaabeWrite's own grammar checker (D-040, D-051): sentences and tokens (text.rs), German
// part-of-speech tags from LanguageTool's open lexicon (tagger.rs, fsa.rs) and LanguageTool's
// XML rules compiled into JSON and run by our own engine (rules.rs). Everything is compiled into
// the program and runs offline; each language loads on first use. Findings are returned with
// UTF-16 offsets, as the frontend counts.

mod de_rules;
mod fsa;
mod rules;
mod tagger;
mod text;

use rules::{ATok, Disambig, Finding, Rule, RuleFile};
use std::collections::{HashMap, HashSet};
use std::io::Read;
use std::sync::{Mutex, OnceLock};
use tagger::{Reading, Tagger};

const DE_DICT: &[u8] = include_bytes!("../../data/german.dict");
const DE_ADDED: &[u8] = include_bytes!("../../data/german-added.txt.gz");
const DE_REMOVED: &[u8] = include_bytes!("../../data/german-removed.txt.gz");
const DE_RULES: &[u8] = include_bytes!("../../data/grammar-de.json.gz");
const DE_SYNTH: &[u8] = include_bytes!("../../data/german_synth.dict");

pub struct Checker {
    tagger: Tagger,
    disambig: Vec<Disambig>,
    rules: Vec<Rule>,
    pub dropped: usize,
    disabled: HashSet<String>,
    synth: de_rules::Synth,
}

impl Checker {
    fn german() -> Option<Checker> {
        let tagger = Tagger::new(DE_DICT, DE_ADDED, DE_REMOVED)?;
        let mut json = String::new();
        flate2::read::GzDecoder::new(DE_RULES).read_to_string(&mut json).ok()?;
        let file: RuleFile = serde_json::from_str(&json).ok()?;
        let _ = rules::UNIFIER.set(rules::Unifier::new(&file.unification));
        let total = file.rules.len();
        let disambig: Vec<Disambig> = file.disambig.into_iter().filter_map(Disambig::new).collect();
        let rules: Vec<Rule> = file.rules.into_iter().filter_map(Rule::new).collect();
        let dropped = total - rules.len();
        Some(Checker { tagger, disambig, rules, dropped, disabled: HashSet::new(), synth: de_rules::Synth::new(DE_SYNTH)? })
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
            out.push(ATok { lower: t.text.to_lowercase(), text: t.text.clone(), start: t.start, end: t.end, space_before: t.space_before, readings, immune: false });
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
        let toks = text::tokens(chars, 0, chars.len());
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
            if let Some(u) = rules::UNIFIER.get() {
                de_rules::check(&atoks, chars, &self.synth, u, &mut out);
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
        if !lang.to_ascii_lowercase().starts_with("de") {
            return paragraphs.iter().map(|_| Vec::new()).collect();
        }
        let mut map = match checkers().lock() { Ok(m) => m, Err(_) => return Vec::new() };
        let c = map.entry("de").or_insert_with(Checker::german);
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

    fn checker() -> Checker {
        Checker::german().expect("German checker")
    }

    fn rules_hit(c: &Checker, s: &str) -> Vec<String> {
        let chars: Vec<char> = s.chars().collect();
        c.check_chars(&chars).into_iter().map(|f| f.rule).collect()
    }

    #[test]
    fn tags_german_words() {
        let c = checker();
        let tags: Vec<String> = c.tagger.tag("Kinder", false).into_iter().map(|r| r.tag).collect();
        assert!(tags.contains(&"SUB:NOM:PLU:NEU".to_string()));
        assert!(c.tagger.tag("Fotosynthese", false).iter().any(|r| r.tag.starts_with("SUB") && r.lemma == "Fotosynthese"));
        assert_eq!(c.tagger.tag(",", false)[0].tag, "PKT");
        eprintln!("rules: {} kept, {} dropped", c.rules.len(), c.dropped);
    }

    #[test]
    fn finds_typical_errors() {
        let c = checker();
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

    /// Prints the tokens, readings and findings of $GRAMMAR_DEBUG (for finding engine differences).
    #[test]
    #[ignore]
    fn debug_text() {
        let Ok(s) = std::env::var("GRAMMAR_DEBUG") else { return };
        let c = checker();
        let chars: Vec<char> = s.chars().collect();
        let toks = text::tokens(&chars, 0, chars.len());
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
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../../tools/lang-data/work/grammar-de-examples.json");
        let Ok(text) = std::fs::read_to_string(&path) else { eprintln!("no examples at {path:?}"); return };
        let examples: Vec<serde_json::Value> = serde_json::from_str(&text).unwrap();
        let c = checker();
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
        eprintln!("incorrect examples found: {hit}/{inc}; correct examples flagged: {fa}/{cor}; {sentences} sentences, {chars_total} chars in {elapsed:?}");
        let report: serde_json::Map<String, serde_json::Value> = stats.into_iter().map(|(k, v)| (k, serde_json::json!(v))).collect();
        let out = path.with_file_name("grammar-de-report.json");
        std::fs::write(out, serde_json::to_string(&report).unwrap()).unwrap();
    }
}
