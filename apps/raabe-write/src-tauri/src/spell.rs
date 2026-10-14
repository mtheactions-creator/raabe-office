// Spell checking with Hunspell dictionaries (spellbook, MPL-2.0): German (igerman98, GPL-2.0
// or GPL-3.0), US and British English (SCOWL, MIT and BSD), French (Dicollecte, MPL-2.0) and
// Spanish (RLA-ES, GPL-3.0 or LGPL-3.0 or MPL-1.1). The dictionaries are compiled into the
// program, so checking works offline; each is loaded on first use.

use spellbook::Dictionary;
use std::collections::HashMap;
use std::sync::{Mutex, OnceLock};

const DE_AFF: &str = include_str!("../../../../node_modules/dictionary-de/index.aff");
const DE_DIC: &str = include_str!("../../../../node_modules/dictionary-de/index.dic");
const EN_AFF: &str = include_str!("../../../../node_modules/dictionary-en/index.aff");
const EN_DIC: &str = include_str!("../../../../node_modules/dictionary-en/index.dic");
const GB_AFF: &str = include_str!("../../../../node_modules/dictionary-en-gb/index.aff");
const GB_DIC: &str = include_str!("../../../../node_modules/dictionary-en-gb/index.dic");
const FR_AFF: &str = include_str!("../../../../node_modules/dictionary-fr/index.aff");
const FR_DIC: &str = include_str!("../../../../node_modules/dictionary-fr/index.dic");
const ES_AFF: &str = include_str!("../../../../node_modules/dictionary-es/index.aff");
const ES_DIC: &str = include_str!("../../../../node_modules/dictionary-es/index.dic");

fn dictionaries() -> &'static Mutex<HashMap<&'static str, Dictionary>> {
    static D: OnceLock<Mutex<HashMap<&'static str, Dictionary>>> = OnceLock::new();
    D.get_or_init(|| Mutex::new(HashMap::new()))
}

/// The dictionary key for a language tag ("de-DE" → "de", "en-GB" → "en-GB", "en-US" → "en",
/// "fr-CA" → "fr", "es-MX" → "es").
fn key(lang: &str) -> Option<&'static str> {
    let l = lang.to_ascii_lowercase();
    if l.starts_with("de") { Some("de") }
    else if ["en-gb", "en-au", "en-nz", "en-ie", "en-za", "en-in"].contains(&l.as_str()) { Some("en-GB") }
    else if l.starts_with("en") { Some("en") }
    else if l.starts_with("fr") { Some("fr") }
    else if l.starts_with("es") { Some("es") }
    else { None }
}

fn with_dict<T>(lang: &str, f: impl FnOnce(&Dictionary) -> T) -> Option<T> {
    let k = key(lang)?;
    let mut map = dictionaries().lock().ok()?;
    if !map.contains_key(k) {
        let (aff, dic) = match k { "de" => (DE_AFF, DE_DIC), "en-GB" => (GB_AFF, GB_DIC), "fr" => (FR_AFF, FR_DIC), "es" => (ES_AFF, ES_DIC), _ => (EN_AFF, EN_DIC) };
        let d = Dictionary::new(aff, dic).ok()?;
        map.insert(k, d);
    }
    map.get(k).map(f)
}

fn capitalize(s: &str) -> String {
    let mut c = s.chars();
    match c.next() {
        Some(f) => f.to_uppercase().collect::<String>() + c.as_str(),
        None => String::new(),
    }
}

/// German compounds the dictionary does not list (Krautschicht, Waldbodenlebewesen…): a word
/// is right when it splits into known words, each later part a noun, with the usual linking
/// letters (Arbeit-s-blatt, Sonne-n-licht) on the parts before.
fn compound_ok(d: &Dictionary, word: &str, depth: usize) -> bool {
    if depth > 3 {
        return false;
    }
    let chars: Vec<char> = word.chars().collect();
    if chars.len() < 6 {
        return false;
    }
    let known_head = |h: &str| {
        if h.chars().count() < 3 {
            return false;
        }
        if d.check(h) {
            return true;
        }
        for link in ["s", "es", "n", "en", "e", "er"] {
            if let Some(base) = h.strip_suffix(link) {
                if base.chars().count() >= 3 && d.check(base) {
                    return true;
                }
            }
        }
        false
    };
    for i in (3..=chars.len() - 3).rev() {
        let head: String = chars[..i].iter().collect();
        if !known_head(&head) {
            continue;
        }
        // Nouns continue with a noun (Kraut-schicht → Schicht), adjectives with a word as written (gras-grün).
        let tail: String = chars[i..].iter().collect();
        let rest = if chars[0].is_uppercase() { capitalize(&tail) } else { tail };
        if d.check(&rest) || compound_ok(d, &rest, depth + 1) {
            return true;
        }
    }
    false
}

/// For each word: true when it is spelled correctly (or the language has no dictionary).
#[tauri::command]
pub async fn spell_check(lang: String, words: Vec<String>) -> Vec<bool> {
    let german = key(&lang) == Some("de");
    tauri::async_runtime::spawn_blocking(move || {
        with_dict(&lang, |d| {
            words.iter().map(|w| d.check(w) || (german && w.chars().count() >= 7 && compound_ok(d, w, 0))).collect::<Vec<_>>()
        })
        .unwrap_or_else(|| vec![true; words.len()])
    })
    .await
    .unwrap_or_default()
}

/// Suggestions for a misspelled word (best first).
#[tauri::command]
pub async fn spell_suggest(lang: String, word: String) -> Vec<String> {
    tauri::async_runtime::spawn_blocking(move || {
        with_dict(&lang, |d| {
            let mut out = Vec::new();
            d.suggest(&word, &mut out);
            out.truncate(8);
            out
        })
        .unwrap_or_default()
    })
    .await
    .unwrap_or_default()
}

/// Loads the German dictionary in the background at start-up, so the first check is quick.
pub fn preload() {
    std::thread::spawn(|| { let _ = with_dict("de", |_| ()); });
}

#[cfg(test)]
mod tests {
    use super::with_dict;

    #[test]
    fn checks_french_and_spanish() {
        let fr = with_dict("fr-FR", |d| (d.check("maison"), d.check("aujourd'hui"), d.check("légumes"), d.check("maisonn"))).unwrap();
        assert_eq!(fr, (true, true, true, false));
        let es = with_dict("es-ES", |d| (d.check("mañana"), d.check("playa"), d.check("calor"), d.check("plaia"))).unwrap();
        assert_eq!(es, (true, true, true, false));
    }
}
