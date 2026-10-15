// RaabeWrite's own English rules (D-051): the errors German-speaking pupils make that
// LanguageTool finds with its chunker or in Java: subject and verb (he go, I has, we was, there is
// many, do he), the base form after do/can/to (didn't saw, can sings, does he likes, to went),
// "more taller", a/an, its/it's, your/you're, their/there, whose, "I am agree", "look forward to
// see", "much friends", "borrow me", uncountable plurals and spaces before punctuation.
// Penn Treebank tags from LanguageTool's english-pos-dict; suggestions from its synthesis lexicon.
// The tokenizer splits contractions as LanguageTool does: do|n't, ca|n't, it|'s, I|'m.

use super::de_rules::Synth;
use super::rules::{ATok, Finding};

fn has(t: &ATok, tag: &str) -> bool {
    t.readings.iter().any(|r| r.tag == tag)
}

fn has_pre(t: &ATok, pre: &str) -> bool {
    t.readings.iter().any(|r| r.tag.starts_with(pre))
}

/// Every reading has one of these tags (ignoring SENT_END).
fn only(t: &ATok, tags: &[&str]) -> bool {
    t.readings.iter().all(|r| r.tag == "SENT_END" || tags.iter().any(|x| r.tag == *x))
}

fn lemma_of(t: &ATok, tags: &[&str]) -> Option<String> {
    t.readings.iter().find(|r| tags.contains(&r.tag.as_str())).map(|r| r.lemma.clone())
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

struct Ctx<'a> {
    toks: &'a [ATok],
    synth: &'a Synth,
    out: &'a mut Vec<Finding>,
    /// LanguageTool's findings in this sentence came first: where one covers the same words ours
    /// gives way, unless it has no suggestion and ours has.
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

    /// A verb in another form ("like" + VBZ → "likes"), capitalised like the word it replaces.
    fn form(&self, lemma: &str, tag: &str, like: &str) -> Vec<String> {
        self.synth.forms(lemma, tag).into_iter().filter(|f| !f.starts_with('\'')).map(|f| cap(&f, like)).collect()
    }

    fn lower(&self, i: usize) -> &str {
        self.toks.get(i).map_or("", |t| t.lower.as_str())
    }

    /// "n't" or "not" at i.
    fn negation(&self, i: usize) -> bool {
        matches!(self.lower(i), "n't" | "not")
    }

    /// Whether a clause starts before token i (sentence start, punctuation, a conjunction).
    fn clause_start(&self, i: usize) -> bool {
        i == 1 || is_punct(&self.toks[i - 1]) || matches!(self.lower(i - 1), "and" | "but" | "because" | "when" | "if" | "so" | "while" | "although" | "since" | "before" | "after" | "then" | "now" | "today" | "yesterday" | "tomorrow")
    }
}

/// A subject pronoun: its person and whether the verb is plural ("you" takes plural verbs).
/// `you` and `it` are also objects (I gave you kisses, call it love): they count only at a
/// clause start.
fn pronoun(c: &Ctx, i: usize) -> Option<(u8, bool)> {
    let p = match c.lower(i) {
        "i" => (1, false),
        "he" | "she" => (3, false),
        "we" | "they" => (2, true),
        "it" if c.clause_start(i) => (3, false),
        "you" if c.clause_start(i) => (2, true),
        _ => return None,
    };
    // "I" is upper case; a lower-case "i" is a letter or a typing slip the speller reports.
    if p.0 == 1 && c.toks[i].text != "I" {
        return None;
    }
    Some(p)
}

const WH: &[&str] = &["what", "where", "when", "why", "how", "who", "which"];
/// After these the pronoun is not the verb's subject or the verb is a base form anyway
/// (let him go, make it work, does he like, that she be).
const NOT_SUBJECT_AFTER: &[&str] = &["do", "does", "did", "can", "could", "will", "would", "shall", "should", "must", "may", "might", "let", "lets", "make", "makes", "made",
    "help", "helps", "see", "saw", "watch", "hear", "heard", "feel", "felt", "if", "that", "lest", "and", "or", "nor", "than", "as", "'d", "'ll", "to"];
const COMMON_VERBS: &[&str] = &["go", "play", "live", "love", "want", "need", "eat", "drink", "sleep", "swim", "cook", "dance", "speak", "come", "hate",
    "read", "write", "visit", "watch", "help", "walk", "ride", "drive", "learn", "teach", "buy", "sell", "make", "take", "get", "give", "say", "tell"];
const FREQUENCY: &[&str] = &["always", "never", "often", "usually", "sometimes", "also", "really", "only", "still", "just", "rarely", "seldom", "normally"];

pub fn check(toks: &[ATok], _chars: &[char], synth: &Synth, out: &mut Vec<Finding>) {
    let first = out.iter().rposition(|f| toks.first().is_some_and(|t| f.start < t.start)).map_or(0, |k| k + 1);
    let mut c = Ctx { toks, synth, out, first };
    agreement(&mut c);
    questions(&mut c);
    base_form(&mut c);
    comparatives(&mut c);
    articles(&mut c);
    confusions(&mut c);
    spaces(&mut c);
}

// ----- subject and verb ----------------------------------------------------------------------------

fn agreement(c: &mut Ctx) {
    let t = c.toks;
    for i in 1..t.len().saturating_sub(1) {
        let Some((person, plural)) = pronoun(c, i) else { continue };
        if i > 1 && NOT_SUBJECT_AFTER.contains(&c.lower(i - 1)) {
            continue;
        }
        let mut k = i + 1;
        if FREQUENCY.contains(&c.lower(k)) {
            k += 1;
        }
        let Some(v) = t.get(k) else { continue };
        let lower = v.lower.as_str();
        if person == 1 && matches!(lower, "is" | "was") && i > 1 {
            continue;
        }
        // be: I am, he is, we are; we were.
        let be = match (lower, person, plural) {
            ("are", 1, false) => Some("am"), ("are", 3, false) => Some("is"), ("is", 1, false) => Some("am"),
            ("is" | "am", _, true) => Some("are"), ("am", 3, false) => Some("is"), ("was", _, true) => Some("were"),
            _ => None,
        };
        if let Some(fix) = be {
            c.report(v.start, v.end, "RAABE_EN_BE", "The form of \"be\" does not match the subject.", vec![cap(fix, &v.text)]);
            continue;
        }
        match lower {
            "has" if plural || person == 1 => {
                c.report(v.start, v.end, "RAABE_EN_HAVE", "With I, you, we or they: \"have\".", vec![cap("have", &v.text)]);
                continue;
            }
            "have" if person == 3 && !plural => {
                c.report(v.start, v.end, "RAABE_EN_HAVE", "With he, she or it: \"has\".", vec![cap("has", &v.text)]);
                continue;
            }
            _ => {}
        }
        // he go, she don't → goes, does (a present-tense verb without -s).
        let it = c.lower(i) == "it";
        if person == 3 && !plural && has(v, "VBP") && !has(v, "VBZ") && !has(v, "VBD") && !has(v, "MD") && only(v, &["VB", "VBP", "NN", "NN:UN", "NN:U", "NNS", "JJ"])
            && !(it && has_pre(v, "NN")) {
            if let Some(lemma) = lemma_of(v, &["VBP", "VB"]) {
                let sugg = c.form(&lemma, "VBZ", &v.text);
                if !sugg.is_empty() {
                    c.report(v.start, v.end, "RAABE_EN_3RD_PERSON", "With he, she or it the verb takes -s.", sugg);
                }
            }
        }
        // I/we/they goes, they doesn't → go, don't.
        if (plural || person == 1) && has(v, "VBZ") && only(v, &["VBZ", "NNS"]) && !(matches!(c.lower(i), "you" | "we") && has(v, "NNS")) {
            if let Some(lemma) = lemma_of(v, &["VBZ"]) {
                let sugg = c.form(&lemma, "VBP", &v.text);
                c.report(v.start, v.end, "RAABE_EN_3RD_PERSON", "Only he, she and it take the verb with -s.", sugg);
            }
        }
    }
    // "My brother like football", "Everybody have a book": a singular subject at the start of the
    // sentence and a verb without -s. Collective nouns (the team, the police) are left alone.
    for i in 1..t.len().saturating_sub(2) {
        let lower = c.lower(i);
        let k = if matches!(lower, "everybody" | "everyone" | "somebody" | "someone" | "nobody" | "anybody" | "anyone" | "everything" | "nothing" | "something") && i == 1 {
            i + 1
        } else if i == 1 && matches!(lower, "my" | "his" | "her" | "our" | "their" | "your" | "the") {
            let n = &t[i + 1];
            let collective = matches!(n.lower.as_str(), "police" | "people" | "family" | "team" | "staff" | "government" | "class" | "audience" | "public" | "crew" | "majority" | "couple" | "group" | "band" | "jury" | "committee" | "data" | "media");
            if collective || !(has(n, "NN") || has(n, "NN:UN")) || !only(n, &["NN", "NN:UN", "NN:U"]) || n.text.chars().next().is_some_and(char::is_uppercase) {
                continue;
            }
            i + 2
        } else {
            continue;
        };
        let v = &t[k];
        let next_ok = t.get(k + 1).is_some_and(|x| has(x, "DT") || has_pre(x, "NN") || has_pre(x, "PRP") || has(x, "TO") || has_pre(x, "RB") || has(x, "IN"));
        if !next_ok {
            continue;
        }
        let fix = match v.lower.as_str() {
            "have" => Some(vec!["has".to_string()]),
            "are" => Some(vec!["is".into()]),
            "were" => Some(vec!["was".into()]),
            "do" if c.negation(k + 1) => Some(vec!["does".into()]),
            _ if has(v, "VBP") && (only(v, &["VB", "VBP"]) || v.lower == "like") => lemma_of(v, &["VBP"]).map(|l| c.form(&l, "VBZ", &v.text)),
            // go, play, love … are nouns too ("the school play in May"): only before an object,
            // "to" or a pronoun (My brother go to school, My sister play the piano).
            _ if has(v, "VBP") && COMMON_VERBS.contains(&v.lower.as_str()) && t.get(k + 1).is_some_and(|x| has(x, "DT") || has(x, "TO") || has_pre(x, "PRP")) => {
                lemma_of(v, &["VBP"]).map(|l| c.form(&l, "VBZ", &v.text))
            }
            _ => None,
        };
        if let Some(sugg) = fix.filter(|s| !s.is_empty()) {
            c.report(v.start, v.end, "RAABE_EN_3RD_PERSON", "With a singular subject the verb takes -s.", sugg);
        }
    }
    // "One of the boys are late": the verb goes with "one".
    for i in 1..t.len().saturating_sub(4) {
        if c.lower(i) != "one" || c.lower(i + 1) != "of" || !matches!(c.lower(i + 2), "the" | "my" | "his" | "her" | "our" | "their" | "your" | "these" | "those") {
            continue;
        }
        let (n, v) = (&t[i + 3], &t[i + 4]);
        if !(has(n, "NNS") && !has(n, "NN")) {
            continue;
        }
        let fix = match v.lower.as_str() {
            "are" => "is",
            "were" => "was",
            "have" => "has",
            _ => continue,
        };
        c.report(v.start, v.end, "RAABE_EN_ONE_OF", "The verb goes with \"one\": singular.", vec![fix.into()]);
    }
    // "My sister she is very nice": the pronoun repeats the subject.
    for i in 1..t.len().saturating_sub(3) {
        let (n, p, v) = (&t[i + 1], &t[i + 2], &t[i + 3]);
        let family = matches!(n.lower.as_str(), "sister" | "brother" | "mother" | "father" | "mum" | "mom" | "dad" | "friend" | "teacher" | "grandma" | "grandpa"
            | "grandmother" | "grandfather" | "cousin" | "aunt" | "uncle" | "parents" | "friends" | "dog" | "cat");
        if i == 1 && matches!(c.lower(i), "my" | "his" | "her" | "our" | "their" | "your") && family && matches!(p.lower.as_str(), "he" | "she" | "they" | "it")
            && (matches!(v.lower.as_str(), "is" | "was" | "are" | "were" | "has" | "have") || has(v, "VBZ") || has(v, "VBD")) {
            c.report(n.end, p.end, "RAABE_EN_DOUBLE_SUBJECT", "The pronoun repeats the subject.", vec![String::new()]);
        }
    }
    // "There is many people", "These is", "This are".
    for i in 1..t.len().saturating_sub(2) {
        let (a, v, n) = (&t[i], &t[i + 1], &t[i + 2]);
        if a.lower == "there" && matches!(v.lower.as_str(), "is" | "was" | "'s") {
            let many = matches!(n.lower.as_str(), "many" | "several" | "two" | "three" | "four" | "five" | "six" | "seven" | "eight" | "nine" | "ten");
            let noun = if many || n.lower == "some" { t.get(i + 3) } else { Some(n) };
            let plural_noun = noun.is_some_and(|x| has(x, "NNS") && only(x, &["NNS"]) && !x.lower.ends_with("ics") && x.lower != "news");
            if many || plural_noun {
                let fix = if v.lower == "was" { "were" } else { "are" };
                let a0 = if v.lower == "'s" { a.start } else { v.start };
                let s = if v.lower == "'s" { format!("{} are", a.text) } else { fix.to_string() };
                c.report(a0, v.end, "RAABE_EN_THERE_IS", "With a plural noun: \"there are\" / \"there were\".", vec![s]);
            }
        }
        if matches!(a.lower.as_str(), "these" | "those") && matches!(v.lower.as_str(), "is" | "was") {
            c.report(v.start, v.end, "RAABE_EN_THESE", "\"These\" and \"those\" take a plural verb.", vec![if v.lower == "is" { "are".into() } else { "were".into() }]);
        }
        if i == 1 && a.lower == "this" && matches!(v.lower.as_str(), "are" | "were") {
            c.report(v.start, v.end, "RAABE_EN_THESE", "\"This\" takes a singular verb.", vec![if v.lower == "are" { "is".into() } else { "was".into() }]);
        }
    }
}

/// Questions: "Do he like", "Does they", "Is they", "Has you".
fn questions(c: &mut Ctx) {
    let t = c.toks;
    if t.last().is_none_or(|x| x.text != "?") {
        return;
    }
    for i in 1..t.len().saturating_sub(2) {
        if !(i == 1 || WH.contains(&c.lower(i - 1))) {
            continue;
        }
        let aux = &t[i];
        let k = if c.negation(i + 1) { i + 2 } else { i + 1 };
        let Some((person, plural)) = (match c.lower(k) {
            "i" => Some((1, false)), "he" | "she" | "it" => Some((3, false)), "we" | "they" | "you" => Some((2, true)), _ => None,
        }) else { continue };
        let fix = match (aux.lower.as_str(), person, plural) {
            ("do", 3, false) => "does", ("does", _, true) | ("does", 1, _) => "do", ("is", _, true) => "are", ("are", 3, false) => "is",
            ("was", _, true) => "were", ("has", _, true) | ("has", 1, _) => "have", ("have", 3, false) => "has",
            _ => continue,
        };
        c.report(aux.start, aux.end, "RAABE_EN_QUESTION", "The verb does not match the subject.", vec![cap(fix, &aux.text)]);
    }
}

// ----- the base form after auxiliaries -------------------------------------------------------------

fn base_form(c: &mut Ctx) {
    let t = c.toks;
    for i in 1..t.len().saturating_sub(1) {
        let a = &t[i];
        // May, Will and Can with a capital inside a sentence are names; "the can", "a will" nouns.
        if i > 1 && a.text.chars().next().is_some_and(char::is_uppercase) || i > 1 && (has(&t[i - 1], "DT") || has_pre(&t[i - 1], "PRP$") || has_pre(&t[i - 1], "JJ")) {
            continue;
        }
        let lower = a.lower.as_str();
        let modal = matches!(lower, "can" | "ca" | "cannot" | "could" | "will" | "wo" | "would" | "should" | "must" | "may" | "might" | "shall");
        let doaux = matches!(lower, "do" | "does" | "did");
        let mut k = i + 1;
        let neg = c.negation(k);
        if neg {
            k += 1;
        }
        // "Does he like", "What did you see": the verb comes after the subject.
        if doaux && !neg && (i == 1 || WH.contains(&c.lower(i - 1))) && matches!(c.lower(k), "i" | "you" | "he" | "she" | "it" | "we" | "they") {
            k += 1;
        } else if doaux && !neg {
            continue;
        }
        let to = lower == "to";
        if !modal && !doaux && !to {
            continue;
        }
        let Some(v) = t.get(k) else { continue };
        let wrong = if to {
            // "to went", "to has": only forms that cannot be anything else.
            (has(v, "VBD") && only(v, &["VBD"])) || (has(v, "VBZ") && only(v, &["VBZ"]) && v.lower != "'s")
        } else {
            (has(v, "VBD") && only(v, &["VBD", "VBN", "JJ"])) || (has(v, "VBZ") && only(v, &["VBZ", "NNS"]))
        };
        if wrong {
            if let Some(lemma) = lemma_of(v, &["VBD", "VBZ"]) {
                let sugg = c.form(&lemma, "VB", &v.text);
                c.report(v.start, v.end, "RAABE_EN_BASE_FORM", "After do, did, can, will, to and the like the verb stays in its base form.", sugg);
            }
        }
    }
    // "I am agree", "she is agree", "I'm live in Hanover".
    for i in 1..t.len().saturating_sub(2) {
        let Some((person, plural)) = pronoun(c, i) else { continue };
        let (be, v) = (&t[i + 1], &t[i + 2]);
        if !matches!(be.lower.as_str(), "am" | "is" | "are" | "'m" | "'re" | "'s") {
            continue;
        }
        let live = v.lower == "live" && matches!(c.lower(i + 3), "in" | "with" | "near");
        if live || matches!(v.lower.as_str(), "agree" | "disagree" | "come" | "go" | "want" | "need" | "know" | "think" | "believe") && !c.lower(i + 3).starts_with('-') {
            let sugg = if person == 3 && !plural { c.form(&v.lower, "VBZ", &v.text) } else { vec![v.text.clone()] };
            let a = if be.space_before { be.start } else { t[i].start };
            let sugg = if be.space_before { sugg } else { sugg.into_iter().map(|s| format!("{} {s}", t[i].text)).collect() };
            c.report(a, v.end, "RAABE_EN_AM_AGREE", "\"Agree\", \"know\" and the like are verbs: no \"am\" or \"is\" before them.", sugg);
        }
    }
    // "look forward to see you".
    for i in 1..t.len().saturating_sub(3) {
        if t[i].lower.starts_with("look") && t[i + 1].lower == "forward" && t[i + 2].lower == "to" {
            let v = &t[i + 3];
            if has(v, "VB") && only(v, &["VB", "VBP", "NN", "NN:UN", "NN:U"]) {
                if let Some(lemma) = lemma_of(v, &["VB"]) {
                    let sugg = c.form(&lemma, "VBG", &v.text);
                    c.report(v.start, v.end, "RAABE_EN_LOOK_FORWARD", "After \"look forward to\" comes the -ing form.", sugg);
                }
            }
        }
    }
}

// ----- comparisons -------------------------------------------------------------------------------------

fn comparatives(c: &mut Ctx) {
    let t = c.toks;
    for i in 1..t.len().saturating_sub(1) {
        let (a, b) = (&t[i], &t[i + 1]);
        if a.lower == "more" && has(b, "JJR") && only(b, &["JJR", "RBR"]) {
            c.report(a.start, b.end, "RAABE_EN_MORE_TALLER", "Use either \"more\" or -er, not both.", vec![b.text.clone()]);
        }
        if a.lower == "most" && has(b, "JJS") && only(b, &["JJS", "RBS"]) {
            c.report(a.start, b.end, "RAABE_EN_MORE_TALLER", "Use either \"most\" or -est, not both.", vec![b.text.clone()]);
        }
    }
}

// ----- a/an --------------------------------------------------------------------------------------------

/// Whether a word starts with a vowel sound (an apple, an hour; a university, a one-way street).
fn vowel_sound(w: &str) -> bool {
    let l = w.to_lowercase();
    if ["hour", "honest", "honour", "honor", "heir"].iter().any(|x| l.starts_with(x)) {
        return true;
    }
    let you = ["unique", "unit", "unif", "union", "unis", "univ", "unicorn", "unicy", "unila", "unimol", "unicell", "use", "usu", "uti", "uto", "ura", "ure", "uri", "uro", "ubi", "uku", "eu", "ewe"];
    if you.iter().any(|x| l.starts_with(x)) || l == "one" || l.starts_with("one-") || l == "once" {
        return false;
    }
    l.starts_with(['a', 'e', 'i', 'o', 'u'])
}

fn articles(c: &mut Ctx) {
    let t = c.toks;
    for i in 1..t.len().saturating_sub(1) {
        let (a, w) = (&t[i], &t[i + 1]);
        // "Vitamin A is": a capital A inside a sentence is a letter.
        if a.text == "A" && i > 1 || !w.space_before || !matches!(a.lower.as_str(), "a" | "an") {
            continue;
        }
        if !(has_pre(w, "NN") || has_pre(w, "JJ") || has(w, "CD") || has_pre(w, "RB") || has(w, "VBG") || has(w, "VBN") || has(w, "UNKNOWN")) || has(w, "CC") {
            continue;
        }
        let letters: Vec<char> = w.text.chars().collect();
        // Abbreviations (an MP, a UFO), single letters and "X-ray" are read letter by letter.
        if !letters[0].is_alphabetic() || letters.len() < 2 || letters[1] == '-' || letters.iter().filter(|ch| ch.is_uppercase()).count() > 1 {
            continue;
        }
        if a.lower == "a" && vowel_sound(&w.text) {
            c.report(a.start, a.end, "RAABE_EN_A_AN", "Before a vowel sound: \"an\".", vec![cap("an", &a.text)]);
        } else if a.lower == "an" && !vowel_sound(&w.text) && !w.lower.starts_with('h') {
            c.report(a.start, a.end, "RAABE_EN_A_AN", "Before a consonant sound: \"a\".", vec![cap("a", &a.text)]);
        }
    }
}

// ----- easily confused words -----------------------------------------------------------------------------

/// A noun followed by a verb ("it's tail is", "who's turn was").
fn noun_then_verb(c: &Ctx, k: usize) -> bool {
    let (Some(n), Some(v)) = (c.toks.get(k), c.toks.get(k + 1)) else { return false };
    has_pre(n, "NN") && !has(n, "VBG") && !has_pre(n, "JJ") && !has_pre(n, "PRP") && !has(n, "DT") && !has_pre(n, "RB") && !has(n, "IN")
        && !matches!(n.lower.as_str(), "time" | "fun" | "home")
        && (has(v, "VBZ") || has(v, "VBD") || has(v, "MD")) && !has(v, "NNS")
}

/// Things an animal or a thing has ("The dog wagged it's tail.").
const POSSESSED: &[&str] = &["tail", "tails", "name", "head", "eyes", "ears", "colour", "color", "wings", "leaves", "shape", "size", "owner", "legs", "fur", "skin",
    "nest", "capital", "population", "purpose", "mouth", "body", "roof", "teeth", "feathers", "paws", "babies", "young", "prey", "food", "history", "origin", "meaning"];

fn confusions(c: &mut Ctx) {
    let t = c.toks;
    for i in 1..t.len() {
        let w = &t[i];
        let next = t.get(i + 1);
        let n = c.lower(i + 1);
        let glued = next.is_some_and(|x| !x.space_before);
        match w.lower.as_str() {
            // "Its a nice day", "its not fair".
            "its" if matches!(n, "a" | "an" | "the" | "not" | "so" | "too" | "been" | "raining" | "snowing") || n == "going" && c.lower(i + 2) == "to" => {
                c.report(w.start, w.end, "RAABE_EN_ITS", "Did you mean \"it's\" (it is)?", vec![cap("it's", &w.text)]);
            }
            // "The dog wagged it's tail", "who's turn is it".
            "it" | "who" if n == "'s" && glued && (noun_then_verb(c, i + 2)
                || w.lower == "it" && (c.lower(i + 2) == "own" || POSSESSED.contains(&c.lower(i + 2)) && i > 1 && has(&t[i - 1], "VBD"))) => {
                let (fix, rule, msg) = if w.lower == "it" {
                    ("its", "RAABE_EN_ITS", "Did you mean \"its\" (belonging to it)?")
                } else {
                    ("whose", "RAABE_EN_WHOSE", "Did you mean \"whose\" (belonging to whom)?")
                };
                c.report(w.start, t[i + 1].end, rule, msg, vec![cap(fix, &w.text)]);
            }
            // "Your welcome", "your going to".
            "your" if n == "welcome" || next.is_some_and(|x| has(x, "VBG") && only(x, &["VBG"])) || n == "going" && c.lower(i + 2) == "to" => {
                c.report(w.start, w.end, "RAABE_EN_YOUR", "Did you mean \"you're\" (you are)?", vec![cap("you're", &w.text)]);
            }
            "their" if next.is_some_and(|x| has(x, "VBG") && only(x, &["VBG"])) || n == "going" && c.lower(i + 2) == "to" => {
                c.report(w.start, w.end, "RAABE_EN_THEIR", "Did you mean \"they're\" (they are)?", vec![cap("they're", &w.text)]);
            }
            // "to there house", "with there friends".
            "there" if i > 1 && matches!(c.lower(i - 1), "to" | "with" | "for" | "of" | "about" | "at") && next.is_some_and(|x| has_pre(x, "NN") && only(x, &["NN", "NN:UN", "NN:U", "NNS"])) => {
                c.report(w.start, w.end, "RAABE_EN_THEIR", "Did you mean \"their\" (belonging to them)?", vec![cap("their", &w.text)]);
            }
            "lets" if i == 1 && next.is_some_and(|x| has(x, "VB") && !has_pre(x, "PRP") && !has(x, "DT")) => {
                c.report(w.start, w.end, "RAABE_EN_LETS", "Did you mean \"let's\" (let us)?", vec![cap("let's", &w.text)]);
            }
            "loose" if i > 1 && c.lower(i - 1) == "to" => {
                c.report(w.start, w.end, "RAABE_EN_LOSE", "Did you mean \"lose\"? \"Loose\" means not tight.", vec!["lose".into()]);
            }
            // "Can you borrow me your pen?"
            "borrow" if matches!(n, "me" | "us" | "him") && t.get(i + 2).is_some_and(|x| has(x, "DT") || has_pre(x, "PRP$") || has(x, "CD")) => {
                c.report(w.start, w.end, "RAABE_EN_LEND", "You lend something to someone, and borrow it from someone.", vec![cap("lend", &w.text)]);
            }
            "much" if next.is_some_and(|x| has(x, "NNS") && !has_pre(x, "NN:") && !has(x, "NN") && !has(x, "JJ") && !matches!(x.lower.as_str(), "thanks" | "news") && !x.lower.ends_with("ics"))
                && t.get(i + 2).is_none_or(|y| is_punct(y) || matches!(y.lower.as_str(), "in" | "at" | "and" | "but" | "here" | "there" | "for" | "with")) => {
                c.report(w.start, w.end, "RAABE_EN_MUCH", "With countable nouns: \"many\".", vec![cap("many", &w.text)]);
            }
            "informations" | "advices" | "furnitures" | "homeworks" | "knowledges" | "luggages" | "feedbacks" => {
                let s = w.text[..w.text.len() - 1].to_string();
                c.report(w.start, w.end, "RAABE_EN_UNCOUNTABLE", "This noun has no plural.", vec![s]);
            }
            // "I go to home": home without "to" (but "go to home page").
            "home" if i > 2 && c.lower(i - 1) == "to" && matches!(c.lower(i - 2), "go" | "goes" | "going" | "went" | "gone" | "come" | "comes" | "coming" | "came" | "get" | "got")
                && !next.is_some_and(|x| has_pre(x, "NN") || has_pre(x, "JJ")) => {
                c.report(t[i - 1].start, w.end, "RAABE_EN_GO_HOME", "\"Go home\", without \"to\".", vec!["home".into()]);
            }
            _ => {}
        }
    }
}

fn spaces(c: &mut Ctx) {
    let t = c.toks;
    for i in 2..t.len() {
        let (prev, cur) = (&t[i - 1], &t[i]);
        // " ,", " ." but not " .NET" or " :)".
        let glued = t.get(i + 1).is_some_and(|n| !n.space_before);
        if cur.space_before && matches!(cur.text.as_str(), "," | "." | ";" | "!" | "?") && prev.text.chars().any(char::is_alphanumeric) && !glued {
            c.report(prev.end, cur.end, "RAABE_EN_SPACE", "No space before a punctuation mark.", vec![cur.text.clone()]);
        }
    }
}
