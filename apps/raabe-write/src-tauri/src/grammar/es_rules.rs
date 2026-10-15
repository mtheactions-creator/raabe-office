// RaabeWrite's own Spanish rules (D-051): errors pupils make that LanguageTool finds in Java or
// with its chunker: determiner and noun (los niño, la problema, mi hermanos), "ser" for a place
// (Madrid es en España), wrong participles (ha escribido), the opening ¿ and ¡, and spaces before
// punctuation. EAGLES tags from LanguageTool's spanish-pos-dict (DA0MS0, NCFP000); forms from its
// synthesis lexicon.

use super::de_rules::Synth;
use super::rules::{ATok, Finding};

fn has_pre(t: &ATok, pre: &str) -> bool {
    t.readings.iter().any(|r| r.tag.starts_with(pre))
}

fn only(t: &ATok, pres: &[&str]) -> bool {
    t.readings.iter().all(|r| r.tag == "SENT_END" || r.tag.starts_with('_') || pres.iter().any(|p| r.tag.starts_with(p)))
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

/// Gender and number of a determiner (DA0MS0: M, S) or noun (NCFP000: F, P) tag.
fn gn(tag: &str) -> Option<(char, char)> {
    let c: Vec<char> = tag.chars().collect();
    let (g, n) = if tag.starts_with('D') { (*c.get(3)?, *c.get(4)?) } else if tag.starts_with("NC") { (*c.get(2)?, *c.get(3)?) } else { return None };
    (matches!(g, 'M' | 'F' | 'C' | 'N' | '0') && matches!(n, 'S' | 'P' | 'N' | '0')).then_some((g, n))
}

fn agree(a: (char, char), b: (char, char)) -> bool {
    let g = |x: char| matches!(x, 'C' | 'N' | '0');
    let n = |x: char| matches!(x, 'N' | '0');
    (g(a.0) || g(b.0) || a.0 == b.0) && (n(a.1) || n(b.1) || a.1 == b.1)
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
    ser_estar(&mut c);
    participles(&mut c);
    opening_marks(&mut c);
    spaces(&mut c);
}

// ----- determiner and noun --------------------------------------------------------------------------

fn finite(t: &ATok) -> bool {
    t.readings.iter().any(|r| r.tag.starts_with('V') && matches!(r.tag.chars().nth(2), Some('I' | 'S')))
}

fn determiners(c: &mut Ctx) {
    let t = c.toks;
    for i in 1..t.len().saturating_sub(1) {
        let (d, n) = (&t[i], &t[i + 1]);
        let dets: Vec<(&str, &str)> = d.readings.iter().filter(|r| matches!(r.tag.get(..2), Some("DA" | "DI" | "DD" | "DP"))).map(|r| (r.lemma.as_str(), r.tag.as_str())).collect();
        if dets.is_empty() || !n.space_before || n.text.chars().next().is_some_and(char::is_uppercase) {
            continue;
        }
        // la, las, lo, los are also pronouns (las casa: he marries them off); mi is also a note.
        // The noun must be only a noun, or a verb as well only when a finite verb follows.
        if !has_pre(n, "NC") || !only(n, &["NC", "AQ", "V"]) || n.lower.ends_with("mente") {
            continue;
        }
        if has_pre(n, "V") && !t.get(i + 2).is_some_and(finite) {
            continue;
        }
        if has_pre(n, "AQ") && t.get(i + 2).is_some_and(|x| has_pre(x, "NC")) {
            continue;
        }
        let nouns: Vec<(&str, (char, char))> = n.readings.iter().filter(|r| r.tag.starts_with("NC")).filter_map(|r| gn(&r.tag).map(|g| (r.lemma.as_str(), g))).collect();
        if nouns.is_empty() {
            continue;
        }
        // el agua, un alma: feminine nouns with a stressed a take el and un.
        let stressed_a = n.lower.starts_with(['a', 'á']) || n.lower.starts_with("ha") || n.lower.starts_with("há");
        if stressed_a && matches!(d.lower.as_str(), "el" | "un" | "algún" | "ningún") && nouns.iter().any(|x| x.1 .0 == 'F' && x.1 .1 == 'S') {
            continue;
        }
        if dets.iter().any(|&(_, dt)| gn(dt).is_some_and(|a| nouns.iter().any(|&(_, b)| agree(a, b)))) {
            continue;
        }
        // "la lactato deshidrogenasa": a noun used as a modifier; the determiner goes with the next.
        let next_agrees = t.get(i + 2).is_some_and(|x| !finite(x) && x.readings.iter().filter(|r| r.tag.starts_with("NC")).filter_map(|r| gn(&r.tag))
            .any(|b| dets.iter().any(|&(_, dt)| gn(dt).is_some_and(|a| agree(a, b)))));
        if next_agrees {
            continue;
        }
        let (dl, dt) = dets[0];
        let Some((dg, dn)) = gn(dt) else { continue };
        let (nl, (ng, nn)) = nouns[0];
        let mut sugg = Vec::new();
        let number_clash = !(matches!(dn, 'N' | '0') || matches!(nn, 'N' | '0') || dn == nn);
        // The noun for the determiner's number (los niños), then the determiner for the noun.
        if number_clash {
            let g = if matches!(ng, 'C' | 'N' | '0') { 'C' } else { ng };
            if let Some(f) = c.form(nl, &format!("NC{g}{dn}000"), &n.text).or_else(|| c.form(nl, &format!("NC{ng}{dn}000"), &n.text)) {
                sugg.push(format!("{} {f}", d.text));
            }
        }
        let want_g = if matches!(dg, 'C' | 'N' | '0') || matches!(ng, 'C' | 'N' | '0') { dg } else { ng };
        let want_n = if matches!(nn, 'N' | '0') { dn } else { nn };
        let mut tag: Vec<char> = dt.chars().collect();
        tag[3] = want_g;
        tag[4] = want_n;
        let tag: String = tag.into_iter().collect();
        if let Some(f) = c.form(dl, &tag, &d.text) {
            sugg.push(format!("{f} {}", n.text));
        }
        let msg = if number_clash { "El determinante y el sustantivo no concuerdan en número." } else { "El determinante y el sustantivo no concuerdan en género." };
        c.report(d.start, n.end, "RAABE_ES_DET_NOMBRE", msg, sugg);
    }
}

// ----- ser and estar ---------------------------------------------------------------------------------

fn ser_estar(c: &mut Ctx) {
    let t = c.toks;
    for i in 1..t.len().saturating_sub(3) {
        let (s, v) = (&t[i], &t[i + 1]);
        // A person or a place (yo, ella, Madrid) "es en …": where something is takes estar.
        // Events take ser (la fiesta es en mi casa), so only pronouns and names count.
        let subject = matches!(s.lower.as_str(), "yo" | "tú" | "él" | "ella" | "nosotros" | "nosotras" | "vosotros" | "vosotras" | "ellos" | "ellas")
            || i == 1 && has_pre(s, "NP") && s.text.chars().next().is_some_and(char::is_uppercase);
        let fix = match v.lower.as_str() {
            "soy" => "estoy", "eres" => "estás", "es" => "está", "somos" => "estamos", "sois" => "estáis", "son" => "están",
            _ => continue,
        };
        if !subject || c.lower(i + 2) != "en" {
            continue;
        }
        // "… es en mayo", "es en 1990": times take ser.
        let after = c.lower(i + 3);
        if after.chars().next().is_some_and(|ch| ch.is_ascii_digit()) || matches!(after, "enero" | "febrero" | "marzo" | "abril" | "mayo" | "junio" | "julio" | "agosto" | "septiembre" | "octubre" | "noviembre" | "diciembre" | "verano" | "invierno" | "otoño" | "primavera") {
            continue;
        }
        c.report(v.start, v.end, "RAABE_ES_SER_ESTAR", "Para indicar dónde está algo o alguien se usa «estar».", vec![cap(fix, &v.text)]);
    }
}

// ----- irregular participles -----------------------------------------------------------------------------

fn participles(c: &mut Ctx) {
    let t = c.toks;
    for i in 1..t.len() {
        let w = &t[i];
        let fix = match w.lower.as_str() {
            "escribido" => "escrito", "abrido" => "abierto", "hacido" => "hecho", "ponido" => "puesto", "volvido" => "vuelto", "rompido" => "roto",
            "vido" => "visto", "morido" => "muerto", "cubrido" => "cubierto", "resolvido" => "resuelto", "descubrido" => "descubierto", "devolvido" => "devuelto",
            "decido" if i > 1 && matches!(c.lower(i - 1), "he" | "has" | "ha" | "hemos" | "habéis" | "han" | "había" | "habían") => "dicho",
            "satisfacido" => "satisfecho", "deshacido" => "deshecho", "imponido" => "impuesto", "componido" => "compuesto", "proponido" => "propuesto",
            _ => continue,
        };
        c.report(w.start, w.end, "RAABE_ES_PARTICIPIO", "Este participio es irregular.", vec![cap(fix, &w.text)]);
    }
}

// ----- ¿ and ¡ ---------------------------------------------------------------------------------------------

fn opening_marks(c: &mut Ctx) {
    let t = c.toks;
    let Some(last) = t.iter().rposition(|x| !is_punct(x) || matches!(x.text.as_str(), "?" | "!")) else { return };
    let end = t[last].text.as_str();
    let (open, msg) = match end {
        "?" => ("¿", "Las preguntas empiezan con «¿»."),
        "!" => ("¡", "Las exclamaciones empiezan con «¡»."),
        _ => return,
    };
    if t.iter().any(|x| x.text.contains(open)) {
        return;
    }
    // Titles in other languages ("Who Started the Great War?") are mostly words the Spanish
    // lexicon does not know or names.
    let words: Vec<&ATok> = t.iter().filter(|x| !is_punct(x)).collect();
    let foreign = words.iter().filter(|x| x.readings.iter().all(|r| r.tag == "UNKNOWN" || r.tag.starts_with("NP") || r.tag == "SENT_END")).count();
    if foreign * 3 > words.len() {
        return;
    }
    // The first word of the sentence: "Cómo estás?" → "¿Cómo".
    let Some(w) = t.iter().skip(1).find(|x| !is_punct(x)) else { return };
    c.report(w.start, w.end, "RAABE_ES_APERTURA", msg, vec![format!("{open}{}", w.text)]);
}

fn spaces(c: &mut Ctx) {
    let t = c.toks;
    for i in 2..t.len() {
        let (prev, cur) = (&t[i - 1], &t[i]);
        let glued = t.get(i + 1).is_some_and(|n| !n.space_before);
        if cur.space_before && matches!(cur.text.as_str(), "," | "." | ";") && prev.text.chars().any(char::is_alphanumeric) && !glued {
            c.report(prev.end, cur.end, "RAABE_ES_ESPACIO", "Sobra un espacio antes de este signo.", vec![cur.text.clone()]);
        }
    }
}
