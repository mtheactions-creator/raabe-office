// RaabeWrite's own French rules (D-051): errors pupils make that LanguageTool finds in Java or
// with its chunker: determiner and noun (le maison, les enfant, mon sœur), the subjunctive after
// "il faut que" (que tu viens), the participle after avoir/être (je suis aller, j'ai manger),
// "Ca" for "Ça" and the space French typography wants before ! ? ; :.
// Tags from LanguageTool's french-pos-dict ("D f s", "N e p", "V sub pres 2 s"); forms from its
// synthesis lexicon.

use super::de_rules::Synth;
use super::rules::{ATok, Finding};

const NBSP: char = '\u{a0}';

fn has_pre(t: &ATok, pre: &str) -> bool {
    t.readings.iter().any(|r| r.tag.starts_with(pre))
}

/// Every reading starts with one of these prefixes (ignoring SENT_END).
fn only(t: &ATok, pres: &[&str]) -> bool {
    t.readings.iter().all(|r| r.tag == "SENT_END" || pres.iter().any(|p| r.tag.starts_with(p)))
}

fn is_punct(t: &ATok) -> bool {
    !t.text.chars().any(char::is_alphanumeric)
}

fn cap(s: &str, like: &str) -> String {
    if like.chars().next().is_some_and(char::is_uppercase) {
        let mut c = s.chars();
        match c.next() {
            Some(f) => f.to_uppercase().collect::<String>() + c.as_str(),
            None => String::new(),
        }
    } else {
        s.to_string()
    }
}

/// Gender and number of a tag's last two fields ("D f s" → ('f', "s"), "N m sp" → ('m', "sp")).
fn gn(tag: &str) -> Option<(char, &str)> {
    let mut parts = tag.rsplit(' ');
    let n = parts.next()?;
    let g = parts.next()?.chars().next()?;
    (matches!(g, 'm' | 'f' | 'e') && matches!(n, "s" | "p" | "sp")).then_some((g, n))
}

fn agree(a: (char, &str), b: (char, &str)) -> bool {
    (a.0 == 'e' || b.0 == 'e' || a.0 == b.0) && (a.1 == "sp" || b.1 == "sp" || a.1 == b.1)
}

struct Ctx<'a> {
    toks: &'a [ATok],
    synth: &'a Synth,
    out: &'a mut Vec<Finding>,
    first: usize,
}

impl Ctx<'_> {
    fn report(&mut self, a: usize, b: usize, rule: &str, message: &str, sugg: Vec<String>) {
        let mut s: Vec<String> = Vec::new();
        for x in sugg {
            if !s.contains(&x) {
                s.push(x);
            }
        }
        if let Some(k) = (self.first..self.out.len()).find(|&k| self.out[k].start < b && a < self.out[k].end) {
            let o = &self.out[k];
            if o.start == a && o.end == b && o.suggestions.is_empty() && !s.is_empty() {
                self.out.remove(k);
            } else {
                return;
            }
        }
        self.out.push(Finding { start: a, end: b, rule: rule.into(), category: "GRAMMAR".into(), message: message.into(), short: String::new(), suggestions: s });
    }

    fn lower(&self, i: usize) -> &str {
        self.toks.get(i).map_or("", |t| t.lower.as_str())
    }

    fn form(&self, lemma: &str, tag: &str, like: &str) -> Option<String> {
        self.synth.forms(lemma, tag).into_iter().next().map(|f| cap(&f, like))
    }
}

pub fn check(toks: &[ATok], _chars: &[char], synth: &Synth, out: &mut Vec<Finding>) {
    let first = out.iter().rposition(|f| toks.first().is_some_and(|t| f.start < t.start)).map_or(0, |k| k + 1);
    let mut c = Ctx { toks, synth, out, first };
    determiners(&mut c);
    subjunctive(&mut c);
    participle(&mut c);
    words(&mut c);
    spaces(&mut c);
}

// ----- determiner and noun --------------------------------------------------------------------------

/// The determiner of a gender and number, by family (le → la, les; ce → cette, ces …).
fn determiner(lower: &str, g: char, n: &str, vowel: bool) -> Option<&'static str> {
    let plural = n == "p";
    let fem = g == 'f';
    Some(match lower {
        "le" | "la" | "les" | "l'" => if plural { "les" } else if vowel { "l'" } else if fem { "la" } else { "le" },
        "un" | "une" => if plural { "des" } else if fem { "une" } else { "un" },
        "ce" | "cet" | "cette" | "ces" => if plural { "ces" } else if fem { "cette" } else if vowel { "cet" } else { "ce" },
        "mon" | "ma" | "mes" => if plural { "mes" } else if fem && !vowel { "ma" } else { "mon" },
        "ton" | "ta" | "tes" => if plural { "tes" } else if fem && !vowel { "ta" } else { "ton" },
        "son" | "sa" | "ses" => if plural { "ses" } else if fem && !vowel { "sa" } else { "son" },
        "notre" | "nos" => if plural { "nos" } else { "notre" },
        "votre" | "vos" => if plural { "vos" } else { "votre" },
        "du" => if plural { "des" } else if vowel { "de l'" } else if fem { "de la" } else { "du" },
        "au" | "aux" => if plural { "aux" } else if vowel { "à l'" } else if fem { "à la" } else { "au" },
        _ => return None,
    })
}

fn determiners(c: &mut Ctx) {
    let t = c.toks;
    for i in 1..t.len().saturating_sub(1) {
        let (d, n) = (&t[i], &t[i + 1]);
        let lower = d.lower.as_str();
        // le, la, les and leur are also pronouns (je le livre demain): the noun must be only a noun.
        let pronoun_too = matches!(lower, "le" | "la" | "les" | "l'");
        if determiner(lower, 'm', "s", false).is_none() || !n.space_before && lower != "l'" {
            continue;
        }
        let noun_ok = has_pre(n, "N ") && if pronoun_too { only(n, &["N "]) } else { only(n, &["N ", "V "]) };
        if !noun_ok || n.text.chars().next().is_some_and(char::is_uppercase) || t.get(i + 2).is_some_and(|x| x.text == "-" && !x.space_before) {
            continue;
        }
        let dets: Vec<(char, &str)> = d.readings.iter().filter(|r| r.tag.starts_with("D ") || r.tag.starts_with("P+D ")).filter_map(|r| gn(&r.tag)).collect();
        let nouns: Vec<(char, &str, &str)> = n.readings.iter().filter(|r| r.tag.starts_with("N ")).filter_map(|r| gn(&r.tag).map(|(g, x)| (g, x, r.lemma.as_str()))).collect();
        if dets.is_empty() || nouns.is_empty() {
            continue;
        }
        let vowel = n.lower.starts_with(['a', 'e', 'i', 'o', 'u', 'y', 'é', 'è', 'ê', 'â', 'î', 'ô', 'û', 'h']);
        // mon amie is right; mon sœur is not (a feminine noun with a consonant takes ma).
        let mon_fem = matches!(lower, "mon" | "ton" | "son") && !vowel && nouns.iter().all(|x| x.0 == 'f');
        if !mon_fem && dets.iter().any(|&a| nouns.iter().any(|&(g, x, _)| agree(a, (g, x)))) {
            continue;
        }
        let (g, num, lemma) = nouns[0];
        let mut sugg = Vec::new();
        // The determiner for the noun (la maison), and for a number mismatch also the noun for
        // the determiner (les enfants).
        let det_num = dets[0].1;
        let number_clash = !(det_num == "sp" || num == "sp" || det_num == num);
        if number_clash && det_num != "sp" {
            if let Some(f) = c.form(lemma, &format!("N {g} {det_num}"), &n.text) {
                sugg.push(format!("{} {f}", d.text));
            }
        }
        if let Some(nd) = determiner(lower, g, num, vowel) {
            let nd = cap(nd, &d.text);
            sugg.push(if nd.ends_with('\'') { format!("{nd}{}", n.text) } else { format!("{nd} {}", n.text) });
        }
        let msg = if number_clash { "Le déterminant et le nom ne s’accordent pas en nombre." } else { "Le déterminant et le nom ne s’accordent pas en genre." };
        c.report(d.start, n.end, "RAABE_FR_DET_NOM", msg, sugg);
    }
}

// ----- the subjunctive --------------------------------------------------------------------------------

fn subjunctive(c: &mut Ctx) {
    let t = c.toks;
    for i in 1..t.len().saturating_sub(2) {
        if !matches!(c.lower(i), "que" | "qu'") {
            continue;
        }
        let before = |k: usize| if i >= k { c.lower(i - k) } else { "" };
        let trigger = matches!(before(1), "faut" | "pour" | "afin" | "avant" | "sans" | "veux" | "veut" | "voulons" | "voulez" | "veulent" | "voudrais" | "voudrait" | "aimerais" | "souhaite" | "souhaites" | "faudrait")
            // bien que (but si bien que, aussi bien que: indicative or comparison)
            || before(1) == "bien" && !matches!(before(2), "si" | "aussi" | "tant" | "eh" | "ou")
            || before(1) == "condition" && before(2) == "à"
            // il est (très) important que, but not "plus important que"
            || matches!(before(1), "important" | "nécessaire" | "essentiel" | "indispensable" | "possible") && matches!(before(2), "est" | "très" | "vraiment" | "particulièrement");
        if !trigger {
            continue;
        }
        let person = match c.lower(i + 1) {
            "je" | "j'" => "1 s", "tu" => "2 s", "il" | "elle" | "on" => "3 s", "nous" => "1 p", "vous" => "2 p", "ils" | "elles" => "3 p",
            _ => continue,
        };
        // que tu ne viennes, que tu le fasses: skip ne and object pronouns.
        let mut k = i + 2;
        while matches!(c.lower(k), "ne" | "n'" | "me" | "m'" | "te" | "t'" | "se" | "s'" | "le" | "la" | "l'" | "les" | "lui" | "leur" | "en" | "y" | "nous" | "vous") && k < i + 6 {
            k += 1;
        }
        let Some(v) = t.get(k) else { continue };
        let ind = v.readings.iter().find(|r| r.tag == format!("V ind pres {person}") || r.tag.starts_with("V ") && r.tag.contains(" ind pres ") && r.tag.ends_with(person));
        let sub = v.readings.iter().any(|r| r.tag.contains(" sub ") && r.tag.ends_with(person));
        if let (Some(r), false) = (ind, sub) {
            let lemma = r.lemma.clone();
            let tag = r.tag.replace(" ind ", " sub ");
            if let Some(f) = c.form(&lemma, &tag, &v.text) {
                c.report(v.start, v.end, "RAABE_FR_SUBJONCTIF", "Après cette expression, le verbe se met au subjonctif.", vec![f]);
            }
        }
    }
}

// ----- the participle after avoir and être ---------------------------------------------------------------

const MOVEMENT: &[&str] = &["aller", "venir", "partir", "arriver", "rentrer", "sortir", "entrer", "tomber", "rester", "monter", "descendre", "retourner", "revenir", "devenir", "passer", "naître", "mourir"];

fn participle(c: &mut Ctx) {
    let t = c.toks;
    for i in 1..t.len().saturating_sub(2) {
        let subj = c.lower(i);
        let (gender, number) = match subj {
            "je" | "j'" | "tu" | "il" | "on" => ('m', 's'),
            "elle" => ('f', 's'),
            "nous" | "vous" | "ils" => ('m', 'p'),
            "elles" => ('f', 'p'),
            _ => continue,
        };
        let aux = c.lower(i + 1);
        let etre = matches!(aux, "suis" | "es" | "est" | "sommes" | "êtes" | "sont");
        let avoir = matches!(aux, "ai" | "as" | "a" | "avons" | "avez" | "ont");
        if !etre && !avoir {
            continue;
        }
        let v = &t[i + 2];
        // "Je suis aller", "j'ai manger": an infinitive in -er where the participle belongs.
        let Some(r) = v.readings.iter().find(|r| r.tag == "V inf") else { continue };
        // ("aller" is also a noun: un aller simple.)
        if !v.lower.ends_with("er") || etre && !MOVEMENT.contains(&r.lemma.as_str()) || avoir && has_pre(v, "N ") || has_pre(v, "J ") {
            continue;
        }
        // "j'ai à faire", "il a beau dire" have something in between; "nous avons manger" does not.
        let tag = if etre { format!("V ppa {gender} {number}") } else { "V ppa m s".to_string() };
        if let Some(f) = c.form(&r.lemma, &tag, &v.text) {
            c.report(v.start, v.end, "RAABE_FR_PARTICIPE", "Après « avoir » ou « être », le participe passé : « allé », « mangé ».", vec![f]);
        }
    }
}

// ----- words ------------------------------------------------------------------------------------------------

fn words(c: &mut Ctx) {
    let t = c.toks;
    for i in 1..t.len() {
        let w = &t[i];
        // "Ca va bien": ça with a cedilla (ca. is circa).
        if w.lower == "ca" && t.get(i + 1).is_some_and(|x| x.text != "." && !is_punct(x)) {
            c.report(w.start, w.end, "RAABE_FR_CA", "« Ça » s’écrit avec une cédille.", vec![cap("ça", &w.text)]);
        }
    }
}

/// French typography wants a (non-breaking) space before ! ? ; and :.
fn spaces(c: &mut Ctx) {
    let t = c.toks;
    for i in 2..t.len() {
        let (prev, cur) = (&t[i - 1], &t[i]);
        if cur.space_before || !matches!(cur.text.as_str(), "!" | "?" | ";" | ":") || !prev.text.chars().last().is_some_and(char::is_alphanumeric) {
            continue;
        }
        // 10:30, http://, ?! and ;-) stay as they are.
        if t.get(i + 1).is_some_and(|n| !n.space_before) {
            continue;
        }
        c.report(cur.start, cur.end, "RAABE_FR_ESPACE", "En français, une espace (insécable) précède ce signe.", vec![format!("{NBSP}{}", cur.text)]);
    }
}
