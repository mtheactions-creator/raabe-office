// Spell checking with Hunspell dictionaries (spellbook, MPL-2.0): German (igerman98, GPL-2.0
// or GPL-3.0), US and British English (SCOWL, MIT and BSD). The dictionaries are compiled
// into the program, so checking works offline; each is loaded on first use.

use spellbook::Dictionary;
use std::collections::HashMap;
use std::sync::{Mutex, OnceLock};

const DE_AFF: &str = include_str!("../../../../node_modules/dictionary-de/index.aff");
const DE_DIC: &str = include_str!("../../../../node_modules/dictionary-de/index.dic");
const EN_AFF: &str = include_str!("../../../../node_modules/dictionary-en/index.aff");
const EN_DIC: &str = include_str!("../../../../node_modules/dictionary-en/index.dic");
const GB_AFF: &str = include_str!("../../../../node_modules/dictionary-en-gb/index.aff");
const GB_DIC: &str = include_str!("../../../../node_modules/dictionary-en-gb/index.dic");

fn dictionaries() -> &'static Mutex<HashMap<&'static str, Dictionary>> {
    static D: OnceLock<Mutex<HashMap<&'static str, Dictionary>>> = OnceLock::new();
    D.get_or_init(|| Mutex::new(HashMap::new()))
}

/// The dictionary key for a language tag ("de-DE" → "de", "en-GB" → "en-GB", "en-US" → "en").
fn key(lang: &str) -> Option<&'static str> {
    let l = lang.to_ascii_lowercase();
    if l.starts_with("de") { Some("de") }
    else if l == "en-gb" || l == "en-au" || l == "en-nz" || l == "en-ie" { Some("en-GB") }
    else if l.starts_with("en") { Some("en") }
    else { None }
}

fn with_dict<T>(lang: &str, f: impl FnOnce(&Dictionary) -> T) -> Option<T> {
    let k = key(lang)?;
    let mut map = dictionaries().lock().ok()?;
    if !map.contains_key(k) {
        let (aff, dic) = match k { "de" => (DE_AFF, DE_DIC), "en-GB" => (GB_AFF, GB_DIC), _ => (EN_AFF, EN_DIC) };
        let d = Dictionary::new(aff, dic).ok()?;
        map.insert(k, d);
    }
    map.get(k).map(f)
}

/// For each word: true when it is spelled correctly (or the language has no dictionary).
#[tauri::command]
pub async fn spell_check(lang: String, words: Vec<String>) -> Vec<bool> {
    tauri::async_runtime::spawn_blocking(move || {
        with_dict(&lang, |d| words.iter().map(|w| d.check(w)).collect::<Vec<_>>()).unwrap_or_else(|| vec![true; words.len()])
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
