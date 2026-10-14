// Thesaurus (Überprüfen > Thesaurus, Synonyme in the context menu): open thesauri from
// LibreOffice's dictionaries, built into compact gzip files by tools/lang-data/build.py (D-049):
// German OpenThesaurus (LGPL-2.1+), English from WordNet (WordNet licence), French by
// Frédéric Labbé / Dicollecte (LGPL-2.1+), Spanish OpenThesaurus (LGPL-2.1). Compiled into
// the program; each language is unpacked on first use.

use serde::Serialize;
use std::collections::HashMap;
use std::io::Read;
use std::sync::{Mutex, OnceLock};

const DE: &[u8] = include_bytes!("../data/thes-de.gz");
const EN: &[u8] = include_bytes!("../data/thes-en.gz");
const FR: &[u8] = include_bytes!("../data/thes-fr.gz");
const ES: &[u8] = include_bytes!("../data/thes-es.gz");

struct Thesaurus {
    meanings: Vec<(String, Vec<String>)>,
    index: HashMap<String, Vec<u32>>,
}

#[derive(Serialize)]
pub struct Meaning {
    /// Part of speech as Word's German thesaurus writes it ("Adj.", "Subst." …) or "".
    pos: String,
    /// The terms; antonyms start with "!".
    terms: Vec<String>,
}

#[derive(Serialize)]
pub struct Lookup {
    /// The word the meanings belong to (the base form found, e.g. "schön" for "schönen").
    word: String,
    meanings: Vec<Meaning>,
}

fn key(lang: &str) -> Option<&'static str> {
    let l = lang.to_ascii_lowercase();
    ["de", "en", "fr", "es"].into_iter().find(|k| l.starts_with(k))
}

fn load(k: &str) -> Option<Thesaurus> {
    let data = match k { "de" => DE, "en" => EN, "fr" => FR, _ => ES };
    let mut text = String::new();
    flate2::read::GzDecoder::new(data).read_to_string(&mut text).ok()?;
    let (meanings, index) = text.split_once("\n\n")?;
    let meanings = meanings.lines().map(|l| {
        let (pos, terms) = l.split_once('\t').unwrap_or(("", l));
        (pos.to_string(), terms.split('|').map(str::to_string).collect())
    }).collect();
    let index = index.lines().filter_map(|l| {
        let (w, ids) = l.split_once('\t')?;
        Some((w.to_string(), ids.split(',').filter_map(|x| x.parse().ok()).collect()))
    }).collect();
    Some(Thesaurus { meanings, index })
}

fn thesauri() -> &'static Mutex<HashMap<&'static str, Thesaurus>> {
    static T: OnceLock<Mutex<HashMap<&'static str, Thesaurus>>> = OnceLock::new();
    T.get_or_init(|| Mutex::new(HashMap::new()))
}

/// Base forms to try for an inflected word (schönen → schön, houses → house, belles → belle).
fn candidates(k: &str, word: &str) -> Vec<String> {
    let w = word.trim().to_lowercase();
    let mut out = vec![w.clone()];
    let endings: &[&str] = match k {
        "de" => &["sten", "ste", "ern", "en", "em", "er", "es", "st", "te", "e", "n", "s", "t"],
        "en" => &["ies", "es", "s", "ed", "d", "ing", "ly", "er", "est"],
        "fr" => &["aux", "es", "s", "x", "e", "ée", "ées", "és"],
        _ => &["es", "s", "as", "os", "a", "o"],
    };
    for e in endings {
        let Some(base) = w.strip_suffix(e) else { continue };
        if base.chars().count() < 3 {
            continue;
        }
        out.push(base.to_string());
        match (k, *e) {
            ("en", "ies") => out.push(format!("{base}y")),
            ("en", "ing" | "ed") => out.push(format!("{base}e")),
            ("es", "a" | "as" | "os") => out.push(format!("{base}o")),
            ("fr", "aux") => out.push(format!("{base}al")),
            _ => {}
        }
    }
    out
}

fn lookup(lang: &str, word: &str) -> Option<Lookup> {
    let k = key(lang)?;
    let mut map = thesauri().lock().ok()?;
    if !map.contains_key(k) {
        map.insert(k, load(k)?);
    }
    let t = map.get(k)?;
    candidates(k, word).into_iter().find_map(|c| {
        let ids = t.index.get(&c)?;
        let meanings = ids.iter().filter_map(|&i| t.meanings.get(i as usize))
            .map(|(pos, terms)| Meaning { pos: pos.clone(), terms: terms.clone() }).collect();
        Some(Lookup { word: c, meanings })
    })
}

/// The meanings of a word in a language (none when the thesaurus has none).
#[tauri::command]
pub async fn thesaurus_lookup(lang: String, word: String) -> Option<Lookup> {
    tauri::async_runtime::spawn_blocking(move || lookup(&lang, &word)).await.ok().flatten()
}

#[cfg(test)]
mod tests {
    use super::lookup;

    #[test]
    fn finds_meanings_and_base_forms() {
        let de = lookup("de-DE", "schönen").expect("de");
        assert_eq!(de.word, "schön");
        assert!(de.meanings.iter().any(|m| m.terms.iter().any(|t| t == "hübsch")));
        let en = lookup("en-US", "houses").expect("en");
        assert_eq!(en.word, "house");
        let fr = lookup("fr-FR", "maison").expect("fr");
        assert!(fr.meanings.iter().any(|m| m.pos == "Subst."));
        assert!(lookup("es-ES", "casa").is_some());
        assert!(lookup("it-IT", "casa").is_none());
    }
}
