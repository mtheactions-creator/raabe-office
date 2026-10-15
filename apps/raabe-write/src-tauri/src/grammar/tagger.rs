// Part-of-speech tags of German words, as LanguageTool tags them (Morphy tags such as
// SUB:NOM:SIN:MAS, VER:3:SIN:PRÄ:SFT): Daniel Naber's german-pos-dict (CC-BY-SA-4.0) with
// LanguageTool's added.txt and removed.txt (LGPL-2.1+). Words the lexicon lacks are analysed
// as compounds by their last part (Fotosynthese → Synthese), as LanguageTool's GermanTagger
// does; numbers are ZAL, punctuation PKT, words without any reading UNKNOWN.

use super::fsa::Fsa;
use std::collections::{HashMap, HashSet};
use std::io::Read;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Reading {
    pub lemma: String,
    pub tag: String,
    /// Guessed from a compound's last part (Fotosynthese → Synthese), not in the lexicon.
    pub guessed: bool,
}

pub struct Tagger {
    fsa: Fsa,
    added: HashMap<String, Vec<Reading>>,
    removed: HashSet<(String, String, String)>,
    /// The lexicon's separator byte and whether its tags end with a frequency letter.
    sep: u8,
    freq: bool,
    lang: &'static str,
}

fn lines_of(gz: &[u8]) -> String {
    let mut s = String::new();
    let _ = flate2::read::GzDecoder::new(gz).read_to_string(&mut s);
    s
}

fn upper_first(s: &str) -> String {
    let mut c = s.chars();
    match c.next() {
        Some(f) => f.to_uppercase().collect::<String>() + c.as_str(),
        None => String::new(),
    }
}

fn lower_first(s: &str) -> String {
    let mut c = s.chars();
    match c.next() {
        Some(f) => f.to_lowercase().collect::<String>() + c.as_str(),
        None => String::new(),
    }
}

impl Tagger {
    pub fn new(dict: &[u8], added_gz: &[u8], removed_gz: &[u8], sep: u8, freq: bool, lang: &'static str) -> Option<Tagger> {
        let fsa = Fsa::new(dict)?;
        let mut added: HashMap<String, Vec<Reading>> = HashMap::new();
        for line in lines_of(added_gz).lines() {
            if line.starts_with('#') {
                continue;
            }
            // Tab-separated; French uses semicolons (Spontex;Spontex;Z e sp).
            let mut p = line.split(if line.contains('\t') { '\t' } else { ';' });
            if let (Some(form), Some(lemma), Some(tag)) = (p.next(), p.next(), p.next()) {
                added.entry(form.to_string()).or_default().push(Reading { lemma: lemma.to_string(), tag: tag.trim().to_string(), guessed: false });
            }
        }
        let mut removed = HashSet::new();
        for line in lines_of(removed_gz).lines() {
            if line.starts_with('#') {
                continue;
            }
            // Tab-separated; French uses semicolons (Spontex;Spontex;Z e sp).
            let mut p = line.split(if line.contains('\t') { '\t' } else { ';' });
            if let (Some(form), Some(lemma), Some(tag)) = (p.next(), p.next(), p.next()) {
                removed.insert((form.to_string(), lemma.to_string(), tag.trim().to_string()));
            }
        }
        Some(Tagger { fsa, added, removed, sep, freq, lang })
    }

    /// Whether the lexicon knows a word (typographic apostrophes as plain ones).
    pub fn has(&self, word: &str) -> bool {
        !self.known(word).is_empty()
    }

    /// The lexicon's readings of a word as written.
    pub fn known(&self, word: &str) -> Vec<Reading> {
        if word.contains('\u{2019}') {
            let plain = word.replace('\u{2019}', "'");
            let r = self.known_exact(&plain);
            if !r.is_empty() {
                return r;
            }
        }
        self.known_exact(word)
    }

    fn known_exact(&self, word: &str) -> Vec<Reading> {
        let freq = self.freq;
        let mut out: Vec<Reading> = self.fsa.lookup(word, self.sep).into_iter()
            .map(|(l, mut t)| { if freq { t.pop(); } (l, t) })
            .filter(|(l, t)| !self.removed.contains(&(word.to_string(), l.clone(), t.clone())))
            .map(|(lemma, tag)| Reading { lemma, tag, guessed: false }).collect();
        if let Some(a) = self.added.get(word) {
            for r in a {
                if !out.contains(r) {
                    out.push(r.clone());
                }
            }
        }
        out
    }

    /// The readings of a token (`first`: the first word of a sentence, which may be capitalized).
    pub fn tag(&self, word: &str, first: bool) -> Vec<Reading> {
        if word.chars().all(|c| !c.is_alphanumeric()) {
            let known = self.known(word);
            let tag = match self.lang {
                "de" => "PKT".to_string(),
                "es" => es_punct(word).to_string(),
                _ if !known.is_empty() => return known,
                "en" => word.to_string(),
                _ => "UNKNOWN".to_string(),
            };
            return vec![Reading { lemma: word.to_string(), tag, guessed: false }];
        }
        if word.chars().all(|c| c.is_ascii_digit() || c == '.' || c == ',') {
            let tag = match self.lang { "de" => "ZAL", "en" => "CD", "fr" => "UNKNOWN", _ => "Z" };
            return vec![Reading { lemma: word.to_string(), tag: tag.into(), guessed: false }];
        }
        let w = word.replace('\u{ad}', "");
        let mut out = self.known(&w);
        let upper = w.chars().next().is_some_and(char::is_uppercase);
        // The first word of a sentence, and words in capitals, are looked up in lower case too.
        if upper && (first || out.is_empty()) {
            let lower = lower_first(&w);
            for r in self.known(&lower) {
                if !out.contains(&r) {
                    out.push(r);
                }
            }
            if out.is_empty() && w.chars().count() > 1 && w.chars().all(|c| !c.is_lowercase()) {
                let cap = upper_first(&w.to_lowercase());
                out = self.known(&cap);
                if out.is_empty() {
                    out = self.known(&w.to_lowercase());
                }
            }
        }
        if out.is_empty() && self.lang == "de" {
            out = self.compound(&w);
            for r in &mut out {
                r.guessed = true;
            }
        }
        out
    }

    /// Compounds the lexicon lacks, by their last part: nouns (Fotosynthese, Kohlenstoff-Atom),
    /// adjectives and participles (hellblaue).
    fn compound(&self, w: &str) -> Vec<Reading> {
        if let Some(i) = w.rfind('-') {
            let last = &w[i + 1..];
            if last.chars().count() >= 2 {
                let r = self.known(last);
                if !r.is_empty() {
                    return r.into_iter().map(|x| Reading { lemma: format!("{}{}", &w[..i + 1], x.lemma), tag: x.tag, guessed: true }).collect();
                }
            }
        }
        let chars: Vec<char> = w.chars().collect();
        if chars.len() < 6 {
            return Vec::new();
        }
        let noun = chars[0].is_uppercase();
        // The first part must be a word too (Foto|synthese, Wald|boden, Arbeit-s|blatt), so that
        // names like "Etruskern" are not split into "Etrus|kern".
        let known_part = |p: &str| {
            let cap = upper_first(p);
            let low = p.to_lowercase();
            let any = |x: &str| !self.known(x).is_empty();
            any(&cap) || any(&low) || ["s", "es", "n", "en", "e", "er"].iter().any(|l| p.strip_suffix(l).is_some_and(|b| b.chars().count() >= 3 && (any(&upper_first(b)) || any(&b.to_lowercase()))))
        };
        for k in 3..chars.len().saturating_sub(2) {
            let prefix: String = chars[..k].iter().collect();
            let rest: String = chars[k..].iter().collect();
            if rest.chars().count() < 3 {
                break;
            }
            if !known_part(&prefix) {
                continue;
            }
            let part = if noun { upper_first(&rest) } else { rest.clone() };
            let found: Vec<Reading> = self.known(&part).into_iter()
                .filter(|r| if noun { r.tag.starts_with("SUB") } else { r.tag.starts_with("ADJ") || r.tag.starts_with("PA1") || r.tag.starts_with("PA2") })
                .collect();
            if !found.is_empty() {
                return found.into_iter().map(|r| Reading {
                    lemma: format!("{}{}", prefix, if noun { lower_first(&r.lemma) } else { r.lemma.clone() }),
                    tag: r.tag,
                    guessed: true,
                }).collect();
            }
        }
        Vec::new()
    }
}

/// Spanish punctuation tags (EAGLES, as LanguageTool's Spanish tagger gives them).
fn es_punct(p: &str) -> &'static str {
    match p {
        "," => "Fc", "." => "Fp", ";" => "Fx", ":" => "Fd", "?" => "Fit", "¿" => "Fia", "!" => "Fat", "¡" => "Faa", "(" => "Fpa", ")" => "Fpt",
        "\"" | "“" | "”" => "Fe", "«" => "Fra", "»" => "Frc", "-" | "–" | "—" => "Fg", "…" => "Fs", "/" => "Fh", "%" => "Ft", "[" => "Fca", "]" => "Fct",
        _ => "Fz",
    }
}
