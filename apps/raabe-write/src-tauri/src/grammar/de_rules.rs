// RaabeWrite's own German rules (D-040, D-051), for what LanguageTool checks with Java code and
// for typical school errors: spaces around punctuation, comparative + "wie", subject and verb,
// article, adjective and noun, the case a preposition needs, capitals (nominalised infinitives,
// capitalised verbs), the genitive apostrophe, the comma before "aber"/"denn"/"sondern", and
// "das" for "dass" after a comma. Every rule is cautious: it only reports what the tags make
// certain. Suggestions are built with LanguageTool's synthesis lexicon (german_synth.dict).

use super::fsa::Fsa;
use super::rules::{ATok, Finding, Unifier};
use super::tagger::Reading;

pub struct Synth {
    fsa: Fsa,
    sep: u8,
}

impl Synth {
    pub fn new(data: &[u8], sep: u8) -> Option<Synth> {
        Some(Synth { fsa: Fsa::new(data)?, sep })
    }

    /// The forms of a lemma with a tag ("ein", "ART:IND:AKK:SIN:NEU" → ["ein"]).
    pub fn forms(&self, lemma: &str, tag: &str) -> Vec<String> {
        self.fsa.lookup(&format!("{lemma}|{tag}"), self.sep).into_iter().map(|(form, _)| form).collect()
    }
}

fn has(t: &ATok, pre: &str) -> bool {
    t.readings.iter().any(|r| r.tag.starts_with(pre))
}

fn only(t: &ATok, pres: &[&str]) -> bool {
    t.readings.iter().all(|r| r.tag == "SENT_END" || pres.iter().any(|p| r.tag.starts_with(p)))
}

fn is_punct(t: &ATok) -> bool {
    !t.text.chars().any(char::is_alphanumeric)
}

/// Person and number of a finite verb reading (VER:3:SIN:PRÄ:SFT, VER:AUX:1:PLU:PRT …).
fn person(tag: &str) -> Option<(char, &str, String)> {
    let parts: Vec<&str> = tag.split(':').collect();
    if parts.first() != Some(&"VER") {
        return None;
    }
    let i = if matches!(parts.get(1), Some(&"AUX") | Some(&"MOD")) { 2 } else { 1 };
    let p = parts.get(i)?.chars().next()?;
    if !"123".contains(p) {
        return None;
    }
    Some((p, parts.get(i + 1)?, parts[i + 2..].join(":")))
}

fn finite(t: &ATok) -> bool {
    t.readings.iter().any(|r| person(&r.tag).is_some())
}

/// What a subject pronoun needs from its verb.
fn pronoun_needs(text: &str) -> Option<&'static [(char, &'static str)]> {
    Some(match text {
        "ich" | "Ich" => &[('1', "SIN")],
        "du" | "Du" => &[('2', "SIN")],
        "er" | "Er" | "man" | "Man" => &[('3', "SIN")],
        "sie" | "Sie" => &[('3', "SIN"), ('3', "PLU")],
        "wir" | "Wir" => &[('1', "PLU")],
        "ihr" | "Ihr" => &[('2', "PLU")],
        _ => return None,
    })
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

struct Ctx<'a> {
    toks: &'a [ATok],
    chars: &'a [char],
    synth: &'a Synth,
    unifier: &'a Unifier,
    out: &'a mut Vec<Finding>,
}

impl Ctx<'_> {
    fn text(&self, a: usize, b: usize) -> String {
        self.chars[a..b].iter().collect()
    }

    fn report(&mut self, a: usize, b: usize, rule: &str, category: &str, message: &str, suggestions: Vec<String>) {
        let seen = self.text(a, b);
        let mut sugg: Vec<String> = Vec::new();
        for s in suggestions {
            if s != seen && !sugg.contains(&s) {
                sugg.push(s);
            }
        }
        self.out.push(Finding { start: a, end: b, rule: rule.into(), category: category.into(), message: message.into(), short: String::new(), suggestions: sugg });
    }
}

const SENT_PUNCT: &[&str] = &[",", ".", ";", ":", "!", "?", ")"];

/// Checks a sentence (tokens with SENT_START at 0).
pub fn check(toks: &[ATok], chars: &[char], synth: &Synth, unifier: &Unifier, out: &mut Vec<Finding>) {
    let mut c = Ctx { toks, chars, synth, unifier, out };
    spaces(&mut c);
    comparative(&mut c);
    subject_verb(&mut c);
    agreement(&mut c);
    preposition_case(&mut c);
    capitals(&mut c);
    apostrophe(&mut c);
    comma_conjunction(&mut c);
    das_dass(&mut c);
}

// ----- spaces ----------------------------------------------------------------------------------------

fn spaces(c: &mut Ctx) {
    let t = c.toks;
    for i in 2..t.len() {
        let (prev, cur) = (&t[i - 1], &t[i]);
        // "Das ist ein Haus ." — a space before a punctuation mark.
        // Not emoticons (:-o, :( ), domains (.com) or ratios (1 : 6).
        let glued = t.get(i + 1).is_some_and(|n| !n.space_before);
        let ratio = cur.text == ":" && prev.text.chars().all(|ch| ch.is_ascii_digit()) && t.get(i + 1).is_some_and(|n| n.text.chars().all(|ch| ch.is_ascii_digit() || ch == ','));
        if cur.space_before && SENT_PUNCT.contains(&cur.text.as_str()) && !is_punct(prev) && !glued && !ratio {
            let sugg = vec![cur.text.clone()];
            c.report(prev.end, cur.end, "RAABE_LEERZEICHEN_VOR_SATZZEICHEN", "TYPOGRAPHY", "Vor einem Satzzeichen steht kein Leerzeichen.", sugg);
        }
        // "Er kam,sah und siegte." — no space after a comma or semicolon.
        if prev.text == "," && !cur.space_before && cur.text.chars().next().is_some_and(char::is_alphabetic) && i >= 2 && !is_punct(&t[i - 2]) {
            let sugg = vec![format!("{} {}", prev.text, cur.text)];
            c.report(prev.start, cur.end, "RAABE_LEERZEICHEN_NACH_KOMMA", "TYPOGRAPHY", "Nach einem Komma steht ein Leerzeichen.", sugg);
        }
        // Two spaces between words.
        if cur.space_before && prev.end < cur.start {
            let gap: String = c.text(prev.end, cur.start);
            if gap.chars().filter(|&ch| ch == ' ').count() >= 2 && !gap.contains('\t') && !gap.contains('\u{a0}') {
                c.report(prev.end, cur.start, "RAABE_DOPPELTES_LEERZEICHEN", "TYPOGRAPHY", "Hier stehen zwei Leerzeichen hintereinander.", vec![" ".into()]);
            }
        }
    }
}

// ----- "größer wie" -----------------------------------------------------------------------------------

fn comparative(c: &mut Ctx) {
    let t = c.toks;
    for i in 1..t.len().saturating_sub(1) {
        let w = &t[i + 1];
        if w.lower != "wie" {
            continue;
        }
        let kom = t[i].readings.iter().any(|r| r.tag.starts_with("ADJ") && r.tag.contains(":KOM")) || t[i].lower == "anders";
        // "so … wie", "eher wie" and the like are fine.
        if !kom || t[i].lower == "eher" || t[i].lower == "mehr" || t[i].lower == "weniger" {
            continue;
        }
        let (a, b) = (w.start, w.end);
        let sugg = vec![if w.text.starts_with('W') { "Als".into() } else { "als".into() }];
        c.report(a, b, "RAABE_KOMPARATIV_ALS", "GRAMMAR", "Nach einem Komparativ steht „als“ (größer als, nicht größer wie).", sugg);
    }
}

// ----- subject and verb -----------------------------------------------------------------------------------

/// The verb forms that agree with a subject (from the verb's own readings: same tense and mood).
fn verb_forms(c: &Ctx, verb: &ATok, needs: &[(char, &str)]) -> Vec<String> {
    let mut out = Vec::new();
    for r in &verb.readings {
        let Some((_, _, rest)) = person(&r.tag) else { continue };
        let aux = if r.tag.starts_with("VER:AUX") { "AUX:" } else if r.tag.starts_with("VER:MOD") { "MOD:" } else { "" };
        for (p, n) in needs {
            for f in c.synth.forms(&r.lemma, &format!("VER:{aux}{p}:{n}:{rest}")) {
                let f = if verb.text.chars().next().is_some_and(char::is_uppercase) { upper_first(&f) } else { f };
                if !out.contains(&f) {
                    out.push(f);
                }
            }
        }
    }
    out
}

fn agrees(verb: &ATok, needs: &[(char, &str)]) -> bool {
    verb.readings.iter().filter_map(|r| person(&r.tag)).any(|(p, n, _)| needs.iter().any(|(q, m)| *q == p && *m == n))
}

/// A noun phrase from `i`: determiner, adjectives, noun (its end, exclusive). None when it is not one.
fn noun_phrase(t: &[ATok], i: usize) -> Option<usize> {
    let det = t.get(i)?;
    if !(has(det, "ART") || det.readings.iter().any(|r| r.tag.starts_with("PRO:POS") || r.tag.starts_with("PRO:DEM") || (r.tag.starts_with("PRO:IND") && !r.tag.ends_with(":STV")))) {
        return None;
    }
    let mut k = i + 1;
    while k < t.len() && (has(&t[k], "ADJ") || has(&t[k], "PA1") || has(&t[k], "PA2")) && !has(&t[k], "SUB") {
        k += 1;
    }
    let n = t.get(k)?;
    if !has(n, "SUB") || n.readings.iter().any(|r| r.tag.starts_with("SUB") && r.tag.ends_with(":ADJ")) {
        return None;
    }
    Some(k + 1)
}

fn subject_verb(c: &mut Ctx) {
    let t = c.toks;
    if t.len() < 3 {
        return;
    }
    // "Wir hat keine Zeit." — a subject pronoun first, the verb right after it.
    if let Some(needs) = pronoun_needs(&t[1].text) {
        let v = &t[2];
        // "Es sind viele": an expletive "es" goes with the noun after the verb.
        let expletive = t[1].lower == "es" && v.readings.iter().any(|r| matches!(r.lemma.as_str(), "sein" | "werden" | "bleiben" | "geben"));
        if !expletive && only(v, &["VER"]) && finite(v) && !agrees(v, needs) && !has(&t[1], "PRO:POS") {
            let sugg = verb_forms(c, v, needs);
            c.report(v.start, v.end, "RAABE_SUBJEKT_VERB", "GRAMMAR", "Subjekt und Verb passen nicht zusammen (Person oder Zahl).", sugg);
        }
    }
    // "Die Kinder spielt im Garten." — a noun phrase first, then the verb.
    let Some(mut end) = noun_phrase(t, 1) else { return };
    let mut needs: Vec<(char, &str)> = Vec::new();
    let first: Vec<&ATok> = t[1..end].iter().collect();
    let unified = c.unifier.unify(&first, &[]);
    let Some(keep) = unified else { return };
    // The noun phrase's number when it is a nominative.
    let noun = &t[end - 1];
    let nom: Vec<&Reading> = noun.readings.iter().zip(keep.last().unwrap()).filter(|(r, ok)| **ok && r.tag.contains(":NOM:")).map(|(r, _)| r).collect();
    if nom.is_empty() {
        return;
    }
    let sin = nom.iter().any(|r| r.tag.contains(":SIN"));
    let plu = nom.iter().any(|r| r.tag.contains(":PLU"));
    if sin && plu {
        return;
    }
    // "Meine Freundin und ich …", "Der Hund und die Katze …".
    let mut coordinated = false;
    if t.get(end).is_some_and(|x| x.lower == "und") {
        if let Some(n2) = pronoun_needs(&t.get(end + 1).map_or(String::new(), |x| x.text.clone())) {
            needs = if n2.contains(&('1', "SIN")) { vec![('1', "PLU")] } else if n2.contains(&('2', "SIN")) { vec![('2', "PLU")] } else { vec![('3', "PLU")] };
            end += 2;
            coordinated = true;
        } else if let Some(e2) = noun_phrase(t, end + 1) {
            needs = vec![('3', "PLU")];
            end = e2;
            coordinated = true;
        } else {
            return;
        }
    }
    if !coordinated {
        needs = vec![('3', if plu { "PLU" } else { "SIN" })];
    }
    let Some(v) = t.get(end) else { return };
    if !(only(v, &["VER"]) && finite(v)) || agrees(v, &needs) {
        return;
    }
    // "Die Kinder sieht er.", "Das größte Wachstum haben tropische Regenwälder": when a noun
    // phrase or pronoun follows the verb, it may be the subject.
    if let Some(after) = t.get(end + 1) {
        if after.readings.iter().any(|r| r.tag.starts_with("PRO") || r.tag.starts_with("EIG") || r.tag.starts_with("ART") || r.tag.starts_with("ADJ") || r.tag.starts_with("SUB") || r.tag.starts_with("ZAL")) {
            return;
        }
    }
    if v.readings.iter().any(|r| matches!(r.lemma.as_str(), "sein" | "werden" | "bleiben")) {
        return;
    }
    let sugg = verb_forms(c, v, &needs);
    c.report(v.start, v.end, "RAABE_SUBJEKT_VERB", "GRAMMAR", "Subjekt und Verb passen nicht zusammen (Person oder Zahl).", sugg);
}

// ----- article, adjective, noun --------------------------------------------------------------------------

fn determiner_forms(c: &Ctx, det: &ATok, rest: &[&ATok]) -> Vec<String> {
    // The combinations the adjectives and the noun allow, as determiner forms.
    let Some(keep) = c.unifier.unify(rest, &[]) else { return Vec::new() };
    let noun = rest.last().unwrap();
    let combos: Vec<(String, String, String)> = noun.readings.iter().zip(keep.last().unwrap()).filter(|(_, ok)| **ok).filter_map(|(r, _)| {
        let p: Vec<&str> = r.tag.split(':').collect();
        Some((p.get(1)?.to_string(), p.get(2)?.to_string(), p.get(3)?.to_string()))
    }).collect();
    let det_cases: Vec<String> = det.readings.iter().filter_map(|r| r.tag.split(':').find(|x| matches!(*x, "NOM" | "AKK" | "DAT" | "GEN")).map(str::to_string)).collect();
    let mut out = Vec::new();
    for prefer in [true, false] {
        for r in &det.readings {
            let p: Vec<&str> = r.tag.split(':').collect();
            let kind = match p.first() { Some(&"ART") => format!("ART:{}", p.get(1).unwrap_or(&"DEF")), Some(&"PRO") => format!("PRO:{}", p.get(1).unwrap_or(&"POS")), _ => continue };
            for (case, num, gender) in &combos {
                if prefer != det_cases.contains(case) {
                    continue;
                }
                let suffix = if kind.starts_with("PRO") { ":BEG" } else { "" };
                for f in c.synth.forms(&r.lemma, &format!("{kind}:{case}:{num}:{gender}{suffix}")) {
                    let f = if det.text.chars().next().is_some_and(char::is_uppercase) { upper_first(&f) } else { f };
                    if !out.contains(&f) {
                        out.push(f);
                    }
                }
            }
        }
        if !out.is_empty() {
            break;
        }
    }
    out.truncate(3);
    out
}

const SKIP_NOUNS: &[&str] = &["herr", "herrn", "frau", "paar", "bisschen", "wenig", "menge", "prozent", "euro", "meter", "jahr", "mal", "uhr"];

fn agreement(c: &mut Ctx) {
    let t = c.toks;
    let mut i = 1;
    while i < t.len() {
        let det = &t[i];
        let is_det = det.readings.iter().any(|r| r.tag.starts_with("ART") || (r.tag.starts_with("PRO:POS") && r.tag.ends_with(":BEG")));
        // A capital article in the middle of a sentence starts a title (in die Gedichte Die Metamorphose).
        let title = i > 1 && det.text.chars().next().is_some_and(char::is_uppercase) && !t[1..i].iter().all(is_punct);
        if !is_det || det.immune || title {
            i += 1;
            continue;
        }
        // ", die zunehmend …": after a comma der/die/das are relative pronouns.
        if t[i - 1].text == "," && matches!(det.lower.as_str(), "der" | "die" | "das" | "dem" | "den" | "deren" | "dessen") {
            i += 1;
            continue;
        }
        let mut k = i + 1;
        let mut words: Vec<&ATok> = vec![det];
        while k < t.len() && (has(&t[k], "ADJ") || has(&t[k], "PA1") || has(&t[k], "PA2") || has(&t[k], "ADV") || (t[k].lower == "zu" && t.get(k + 1).is_some_and(|x| has(x, "PA1")))) && !has(&t[k], "SUB") {
            if t[k].text.chars().any(|ch| ch.is_ascii_digit()) {
                break;
            }
            // Adjectives used as adverbs (ganz, deutlich, ursprünglich) are not inflected.
            let adverbial = t[k].lower == "zu" || t[k].readings.iter().any(|r| r.tag.starts_with("ADV") || r.tag.contains(":PRD:"));
            if !adverbial {
                words.push(&t[k]);
            }
            k += 1;
        }
        let Some(n) = t.get(k) else { break };
        let foreign_after = t.get(k + 1).is_some_and(|x| x.readings.iter().any(|r| r.tag == "UNKNOWN") && x.text.chars().next().is_some_and(char::is_uppercase));
        let foreign_before = t[i + 1..k].iter().any(|x| x.readings.iter().any(|r| r.tag == "UNKNOWN"));
        let noun_ok = !foreign_after && !foreign_before && n.text.chars().count() > 1 && only(n, &["SUB"]) && !n.readings.iter().any(|r| r.tag.ends_with(":ADJ") || r.guessed) && n.text.chars().next().is_some_and(char::is_uppercase)
            && !SKIP_NOUNS.contains(&n.lower.as_str()) && !n.immune && !t.get(k + 1).is_some_and(|x| x.text == "-" || x.text.starts_with('-'))
            && !t.get(k + 1).is_some_and(|x| x.text.chars().next().is_some_and(char::is_uppercase) && !has(x, "SUB") && !has(x, "VER") && !has(x, "ART"));
        if !noun_ok || k > i + 4 {
            i += 1;
            continue;
        }
        words.push(n);
        // "hatte das fatale Konsequenzen": der/die/das can stand alone when the rest agrees.
        let standalone = matches!(det.lower.as_str(), "das" | "die" | "der" | "dies" | "diese" | "dieses") && words.len() > 2 && c.unifier.unify(&words[1..], &[]).is_some();
        if !standalone && c.unifier.unify(&words, &[]).is_none() {
            let rest: Vec<&ATok> = words[1..].to_vec();
            let forms = determiner_forms(c, det, &rest);
            let tail = c.text(t[i + 1].start, n.end);
            let sugg: Vec<String> = forms.into_iter().map(|f| format!("{f} {tail}")).collect();
            c.report(det.start, n.end, "RAABE_KONGRUENZ", "GRAMMAR", "Artikel, Adjektiv und Nomen passen nicht zusammen (Fall, Zahl oder Geschlecht).", sugg);
            i = k + 1;
            continue;
        }
        i += 1;
    }
}

// ----- prepositions --------------------------------------------------------------------------------------

fn preposition_case(c: &mut Ctx) {
    let t = c.toks;
    for i in 1..t.len().saturating_sub(2) {
        let p = &t[i];
        let case = match p.lower.as_str() {
            "für" | "gegen" | "ohne" => "AKK",
            "mit" | "nach" | "bei" | "von" | "zu" | "gegenüber" | "seit" => "DAT",
            _ => continue,
        };
        // "was für einem", "von hier aus den Berg" (a particle), "mit" as an adverb.
        if (p.lower == "für" && t[i - 1].lower == "was") || (i >= 2 && matches!(t[i - 1].lower.as_str(), "hier" | "da" | "dort" | "oben" | "unten")) {
            continue;
        }
        // "seit" and "um" can start a clause or an infinitive.
        if p.readings.iter().any(|r| r.tag.starts_with("KON")) || (p.lower == "um" && t[i + 1..].iter().any(|x| x.lower == "zu")) {
            continue;
        }
        let Some(end) = noun_phrase(t, i + 1) else { continue };
        let det = &t[i + 1];
        let foreign = t.get(end).is_some_and(|x| x.text.chars().next().is_some_and(char::is_uppercase) && !has(x, "VER"));
        if foreign || !det.readings.iter().any(|r| r.tag.starts_with("ART")) || t[end - 1].readings.iter().any(|r| r.guessed) {
            continue;
        }
        let np: Vec<&ATok> = t[i + 1..end].iter().collect();
        let Some(keep) = c.unifier.unify(&np, &[]) else { continue };
        let fits = det.readings.iter().zip(&keep[0]).any(|(r, ok)| *ok && r.tag.contains(&format!(":{case}:")));
        if fits {
            continue;
        }
        // The determiner in the case the preposition needs, for the noun's gender and number.
        let noun = &t[end - 1];
        let mut sugg = Vec::new();
        for r in &noun.readings {
            let q: Vec<&str> = r.tag.split(':').collect();
            if q.first() != Some(&"SUB") || q.len() < 4 {
                continue;
            }
            for d in &det.readings {
                let dp: Vec<&str> = d.tag.split(':').collect();
                if dp.first() != Some(&"ART") {
                    continue;
                }
                for f in c.synth.forms(&d.lemma, &format!("ART:{}:{case}:{}:{}", dp.get(1).unwrap_or(&"DEF"), q[2], q[3])) {
                    let tail = c.text(det.end, noun.end);
                    let s = format!("{f}{tail}");
                    if !sugg.contains(&s) {
                        sugg.push(s);
                    }
                }
            }
        }
        sugg.truncate(3);
        let msg = if case == "AKK" { format!("Nach „{}“ steht der Akkusativ.", p.text) } else { format!("Nach „{}“ steht der Dativ.", p.text) };
        c.report(det.start, noun.end, "RAABE_PRAEPOSITION_FALL", "GRAMMAR", &msg, sugg);
    }
}

// ----- capitals --------------------------------------------------------------------------------------

fn capitals(c: &mut Ctx) {
    let t = c.toks;
    // "Das lesen macht Spaß." — an infinitive used as a noun.
    for i in 1..t.len().saturating_sub(2) {
        let d = &t[i];
        if !matches!(d.lower.as_str(), "das" | "beim" | "zum" | "vom") {
            continue;
        }
        let w = &t[i + 1];
        let next = &t[i + 2];
        let starts = i == 1 || matches!(t[i - 1].text.as_str(), "," | ";" | ":");
        if starts && w.text.chars().next().is_some_and(char::is_lowercase) && w.readings.iter().any(|r| r.tag.starts_with("VER:INF")) && only(w, &["VER"])
            && (only(next, &["VER"]) && finite(next) && d.lower == "das" || (d.lower != "das" && (is_punct(next) || has(next, "VER"))))
        {
            c.report(w.start, w.end, "RAABE_NOMINALISIERUNG", "CASING", "Ein Verb, das als Nomen gebraucht wird, schreibt man groß.", vec![upper_first(&w.text)]);
        }
    }
    // "Am besten Fragst du …" — a capital where no noun or name is.
    for i in 2..t.len() {
        let w = &t[i];
        let prev = &t[i - 1];
        if !w.text.chars().next().is_some_and(char::is_uppercase) || w.text.chars().count() < 3 || w.text.chars().skip(1).any(char::is_uppercase) {
            continue;
        }
        if is_punct(prev) && !matches!(prev.text.as_str(), ",") {
            continue;
        }
        if t[1..i].iter().all(is_punct) {
            continue;
        }
        // Names come in groups (Wilhelm Raabe, Werner Flechsig); compounds with a hyphen (Stark-).
        let prev_lower = prev.text.chars().next().is_some_and(char::is_lowercase) && !has(prev, "PRP") && !matches!(prev.lower.as_str(), "von" | "zu" | "de" | "van" | "der" | "und");
        let next = t.get(i + 1);
        if !prev_lower || next.is_some_and(|x| x.text.chars().next().is_some_and(char::is_uppercase) || x.text.starts_with('-')) {
            continue;
        }
        let capital_known = w.readings.iter().any(|r| r.tag != "UNKNOWN" && r.tag != "SENT_END" && r.lemma.chars().next().is_some_and(char::is_uppercase));
        if capital_known || w.readings.iter().any(|r| r.tag == "UNKNOWN") {
            continue;
        }
        let lower_ok = w.readings.iter().all(|r| r.tag == "SENT_END" || r.tag.starts_with("VER") || r.tag.starts_with("ADV") || r.tag.starts_with("PRP") || r.tag.starts_with("KON") || r.tag.starts_with("ADJ:PRD"));
        if lower_ok && finite(w) && !w.readings.iter().any(|r| r.guessed) && next.is_some_and(|x| !is_punct(x)) {
            c.report(w.start, w.end, "RAABE_GROSS_MITTEN", "CASING", "Dieses Wort ist kein Nomen und wird kleingeschrieben.", vec![lower_first(&w.text)]);
        }
    }
}

// ----- apostrophe -------------------------------------------------------------------------------------

fn apostrophe(c: &mut Ctx) {
    let t = c.toks;
    for w in &t[1..] {
        let text = &w.text;
        let lower = text.to_lowercase();
        let short = match lower.as_str() {
            "in's" | "in’s" => Some("ins"), "an's" | "an’s" => Some("ans"), "auf's" | "auf’s" => Some("aufs"), "für's" | "für’s" => Some("fürs"),
            "um's" | "um’s" => Some("ums"), "durch's" | "durch’s" => Some("durchs"), "über's" | "über’s" => Some("übers"), "vor's" | "vor’s" => Some("vors"),
            _ => None,
        };
        if let Some(s) = short {
            let s = if text.chars().next().is_some_and(char::is_uppercase) { upper_first(s) } else { s.to_string() };
            c.report(w.start, w.end, "RAABE_APOSTROPH", "TYPOGRAPHY", "Verschmelzungen von Präposition und Artikel schreibt man ohne Apostroph.", vec![s]);
            continue;
        }
        // "Peter's Fahrrad" — the genitive -s of names has no apostrophe.
        let chars: Vec<char> = text.chars().collect();
        let n = chars.len();
        let english = c.toks.iter().position(|x| std::ptr::eq(x, w)).and_then(|k| c.toks.get(k + 1)).is_some_and(|x| x.readings.iter().any(|r| r.tag == "UNKNOWN"));
        if !english && n >= 4 && chars[0].is_uppercase() && (chars[n - 2] == '\'' || chars[n - 2] == '’') && chars[n - 1] == 's' && chars[..n - 2].iter().all(|ch| ch.is_alphabetic()) {
            let name: String = chars[..n - 2].iter().collect();
            let fixed = if name.ends_with('s') { name.clone() } else { format!("{name}s") };
            c.report(w.start, w.end, "RAABE_APOSTROPH", "TYPOGRAPHY", "Der Genitiv von Namen wird ohne Apostroph geschrieben (Peters Fahrrad).", vec![fixed]);
        }
    }
}

// ----- commas ------------------------------------------------------------------------------------------

fn comma_conjunction(c: &mut Ctx) {
    let t = c.toks;
    for i in 2..t.len().saturating_sub(2) {
        let k = &t[i];
        if !matches!(k.lower.as_str(), "aber" | "denn" | "sondern") {
            continue;
        }
        let prev = &t[i - 1];
        if is_punct(prev) || matches!(prev.lower.as_str(), "und" | "oder" | "nicht" | "es") {
            continue;
        }
        let subj = &t[i + 1];
        let verb = &t[i + 2];
        if pronoun_needs(&subj.text).is_some() && only(verb, &["VER"]) && finite(verb) {
            let sugg = vec![format!("{}, {}", prev.text, k.text)];
            c.report(prev.start, k.end, "RAABE_KOMMA_KONJUNKTION", "PUNCTUATION", &format!("Vor „{}“ steht ein Komma, wenn ein Hauptsatz folgt.", k.text), sugg);
        }
    }
}

fn das_dass(c: &mut Ctx) {
    let t = c.toks;
    for i in 2..t.len().saturating_sub(2) {
        if t[i].text != "das" || t[i - 1].text != "," {
            continue;
        }
        let mut b = i - 2;
        while b > 1 && is_punct(&t[b]) {
            b -= 1;
        }
        let before = &t[b];
        if is_punct(&t[i - 2]) && t[i - 2].text != ")" {
            continue;
        }
        // "Das Kind, das …", "etwas, das …": a relative pronoun.
        // A dass-clause follows verbs (glauben, sagen, wissen …) and some adjectives; after
        // nouns and pronouns "das" is a relative pronoun (das erste, das …; etwas anderes, das …).
        let verb = before.readings.iter().all(|r| r.tag.starts_with("VER") || r.tag == "SENT_END") && finite(before);
        let adjective = matches!(before.lower.as_str(), "wichtig" | "klar" | "schade" | "gut" | "schön" | "möglich" | "sicher" | "wahrscheinlich" | "richtig"
            | "falsch" | "toll" | "besser" | "logisch" | "unglaublich" | "erstaunlich" | "seltsam" | "komisch" | "traurig" | "froh" | "stolz" | "egal" | "fraglich");
        if !verb && !adjective {
            continue;
        }
        let next = &t[i + 1];
        if pronoun_needs(&next.text).is_some() && !has(next, "PRO:POS") {
            c.report(t[i].start, t[i].end, "RAABE_DAS_DASS", "GRAMMAR", "Leitet „das“ hier einen Nebensatz ein, schreibt man „dass“.", vec!["dass".into()]);
        }
    }
}
