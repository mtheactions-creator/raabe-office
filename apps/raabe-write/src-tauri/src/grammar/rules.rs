// LanguageTool-style pattern rules (compiled from LanguageTool's XML by
// tools/lang-data/build_grammar.py, D-051): token conditions on the text, the lemma and the
// part-of-speech tag (literal or regular expression, negated or not), exceptions for the token,
// the next or the previous token, repetitions (min/max), skipped tokens, markers for what is
// reported, antipatterns that cancel a finding, and messages with suggestions built from the
// matched tokens (case conversion, regular-expression replacement).

use super::tagger::Reading;
use fancy_regex::Regex;
use serde::Deserialize;

// ----- the compiled JSON ---------------------------------------------------------------------------

#[derive(Deserialize, Default)]
pub struct TokJson {
    t: Option<String>,
    #[serde(default)]
    re: bool,
    #[serde(default)]
    cs: bool,
    #[serde(default)]
    inf: bool,
    #[serde(default)]
    neg: bool,
    pos: Option<String>,
    #[serde(default)]
    pre: bool,
    #[serde(default)]
    npos: bool,
    skip: Option<i32>,
    min: Option<i32>,
    max: Option<i32>,
    sb: Option<bool>,
    /// Inside <unify>: the group's index in the pattern's unify list.
    u: Option<usize>,
    /// The same word as pattern token `ref` (0-based; its lemma when `inf`).
    #[serde(rename = "ref")]
    backref: Option<usize>,
    #[serde(default)]
    exc: Vec<ExcJson>,
    #[serde(default, rename = "and")]
    and_: Vec<TokJson>,
    #[serde(default, rename = "or")]
    or_: Vec<TokJson>,
}

#[derive(Deserialize, Default)]
pub struct ExcJson {
    scope: Option<String>,
    t: Option<String>,
    #[serde(default)]
    re: bool,
    #[serde(default)]
    cs: bool,
    #[serde(default)]
    inf: bool,
    #[serde(default)]
    neg: bool,
    pos: Option<String>,
    #[serde(default)]
    pre: bool,
    #[serde(default)]
    npos: bool,
    sb: Option<bool>,
}

#[derive(Deserialize, Clone)]
pub struct MatchRef {
    no: usize,
    cc: Option<String>,
    rm: Option<String>,
    rr: Option<String>,
    pos: Option<String>,
    pr: Option<String>,
}

#[derive(Deserialize, Clone)]
#[serde(untagged)]
pub enum Piece {
    Text(String),
    Match { m: MatchRef },
    Sugg { s: Vec<Piece> },
}

#[derive(Deserialize, Clone)]
pub struct UnifyJson {
    #[serde(default)]
    features: Vec<String>,
    #[serde(default)]
    neg: bool,
}

#[derive(Deserialize)]
pub struct RegexJson {
    re: String,
    #[serde(default)]
    cs: bool,
    #[serde(default)]
    mark: usize,
}

#[derive(Deserialize)]
pub struct RuleJson {
    pub id: String,
    pub cat: Option<String>,
    #[serde(default)]
    pub name: String,
    pat: Option<Vec<TokJson>>,
    ma: Option<usize>,
    mb: Option<usize>,
    regex: Option<RegexJson>,
    #[serde(default)]
    anti: Vec<Vec<TokJson>>,
    #[serde(default)]
    antiu: Vec<Vec<UnifyJson>>,
    #[serde(default)]
    unify: Vec<UnifyJson>,
    msg: Vec<Piece>,
    #[serde(default)]
    sugg: Vec<Vec<Piece>>,
    short: Option<String>,
}

#[derive(Deserialize)]
pub struct WdJson {
    lemma: Option<String>,
    pos: Option<String>,
}

#[derive(Deserialize)]
pub struct DisambigJson {
    pat: Vec<TokJson>,
    ma: usize,
    mb: usize,
    action: String,
    pos: Option<String>,
    #[serde(default)]
    wd: Vec<WdJson>,
    #[serde(default)]
    anti: Vec<Vec<TokJson>>,
    #[serde(default)]
    antiu: Vec<Vec<UnifyJson>>,
    #[serde(default)]
    unify: Vec<UnifyJson>,
}

#[derive(Deserialize)]
pub struct RuleFile {
    pub rules: Vec<RuleJson>,
    #[serde(default)]
    pub disambig: Vec<DisambigJson>,
    #[serde(default)]
    pub unification: std::collections::HashMap<String, Vec<(String, String)>>,
}

// ----- unification -----------------------------------------------------------------------------------

/// LanguageTool's unification: features (number, case, gender) whose values are tag patterns;
/// tokens unify when one choice of readings shares a value of every feature.
pub struct Unifier {
    features: Vec<(String, Vec<Regex>)>,
}

impl Unifier {
    pub fn new(defs: &std::collections::HashMap<String, Vec<(String, String)>>) -> Unifier {
        let mut features: Vec<(String, Vec<Regex>)> = defs.iter()
            .map(|(f, types)| (f.clone(), types.iter().filter_map(|(_, re)| compile(re, true)).collect())).collect();
        features.sort_by(|a, b| a.0.cmp(&b.0));
        Unifier { features }
    }

    /// For the readings of each token: whether it takes part in a unification over `wanted`
    /// features; None when the tokens do not unify.
    pub fn unify(&self, toks: &[&ATok], wanted: &[String]) -> Option<Vec<Vec<bool>>> {
        let feats: Vec<&Vec<Regex>> = self.features.iter().filter(|(f, _)| wanted.is_empty() || wanted.contains(f)).map(|(_, t)| t).collect();
        if feats.is_empty() || toks.is_empty() {
            return None;
        }
        // Each reading's values: one bit set per feature.
        let values = |r: &Reading| -> Vec<u32> { feats.iter().map(|types| types.iter().enumerate().fold(0u32, |m, (i, re)| if re.is_match(&r.tag).unwrap_or(false) { m | 1 << i } else { m })).collect() };
        let per: Vec<Vec<Vec<u32>>> = toks.iter().map(|t| t.readings.iter().map(values).collect()).collect();
        // Every combination of one value per feature (at most 2 × 4 × 3).
        let mut combos: Vec<Vec<usize>> = vec![vec![]];
        for types in &feats {
            combos = combos.into_iter().flat_map(|c| (0..types.len()).map(move |i| { let mut n = c.clone(); n.push(i); n })).collect();
        }
        let supports = |r: &Vec<u32>, c: &Vec<usize>| c.iter().enumerate().all(|(f, &i)| r[f] & (1 << i) != 0);
        let unified: Vec<&Vec<usize>> = combos.iter().filter(|c| per.iter().all(|readings| readings.iter().any(|r| supports(r, c)))).collect();
        if unified.is_empty() {
            return None;
        }
        Some(per.iter().map(|readings| readings.iter().map(|r| unified.iter().any(|c| supports(r, c))).collect()).collect())
    }
}

// ----- compiled --------------------------------------------------------------------------------------

/// A token of a sentence with its readings (never empty: UNKNOWN when the lexicon has none).
pub struct ATok {
    pub text: String,
    pub lower: String,
    pub start: usize,
    pub end: usize,
    pub space_before: bool,
    pub readings: Vec<Reading>,
    /// Protected by a disambiguation rule (immunize): no rule reports it.
    pub immune: bool,
}

/// Java's regular expressions as fancy-regex reads them (anchored: the whole token must match).
pub fn compile(re: &str, case_sensitive: bool) -> Option<Regex> {
    let mut s = re.to_string();
    for (a, b) in [("\\p{Punct}", "[\\p{P}\\p{S}]"), ("\\p{Upper}", "\\p{Lu}"), ("\\p{Lower}", "\\p{Ll}"), ("\\p{Alpha}", "\\p{L}"),
        ("\\p{Digit}", "\\d"), ("\\p{IsAlphabetic}", "\\p{Alphabetic}"), ("\\p{IsLetter}", "\\p{L}"), ("\\p{javaLowerCase}", "\\p{Ll}"),
        ("\\p{javaUpperCase}", "\\p{Lu}"), ("\\p{Lu}+", "\\p{Lu}+"), ("(?iu)", "(?i)"), ("(?ui)", "(?i)"), ("(?u)", ""), ("(?-iu)", "(?-i)"),
        ("*+", "*"), ("++", "+"), ("?+", "?")] {
        s = s.replace(a, b);
    }
    let flags = if case_sensitive { "" } else { "(?i)" };
    Regex::new(&format!("^{flags}(?:{s})$")).ok()
}

enum TextCond {
    Any,
    Lit(String, bool),
    Re(Regex),
}

impl TextCond {
    fn new(t: &Option<String>, re: bool, cs: bool) -> Option<TextCond> {
        match t {
            None => Some(TextCond::Any),
            Some(t) if re => compile(t, cs).map(TextCond::Re),
            Some(t) => Some(TextCond::Lit(if cs { t.clone() } else { t.to_lowercase() }, cs)),
        }
    }

    fn ok(&self, s: &str, lower: &str) -> bool {
        match self {
            TextCond::Any => true,
            TextCond::Lit(l, cs) => if *cs { s == l } else { lower == l },
            TextCond::Re(r) => r.is_match(s).unwrap_or(false),
        }
    }
}

enum PosCond {
    Any,
    Lit(String),
    Re(Regex),
}

impl PosCond {
    fn new(pos: &Option<String>, pre: bool) -> Option<PosCond> {
        match pos {
            None => Some(PosCond::Any),
            Some(p) if pre => compile(p, true).map(PosCond::Re),
            Some(p) => Some(PosCond::Lit(p.clone())),
        }
    }

    fn ok(&self, tag: &str) -> bool {
        match self {
            PosCond::Any => true,
            PosCond::Lit(l) => tag == l,
            PosCond::Re(r) => r.is_match(tag).unwrap_or(false),
        }
    }
}

struct Cond {
    text: TextCond,
    inf: bool,
    neg: bool,
    pos: PosCond,
    has_pos: bool,
    npos: bool,
    sb: Option<bool>,
}

impl Cond {
    fn new(t: &Option<String>, re: bool, cs: bool, inf: bool, neg: bool, pos: &Option<String>, pre: bool, npos: bool, sb: Option<bool>) -> Option<Cond> {
        Some(Cond { text: TextCond::new(t, re, cs)?, inf, neg, pos: PosCond::new(pos, pre)?, has_pos: pos.is_some(), npos, sb })
    }

    /// LanguageTool's PatternToken.isMatched for one reading: (text or lemma, negated) and
    /// (tag, negated); a token matches when any of its readings does.
    fn reading(&self, tok: &ATok, r: &Reading) -> bool {
        let text = match self.text {
            TextCond::Any => !self.neg,
            _ => {
                let ok = if self.inf { self.text.ok(&r.lemma, &r.lemma.to_lowercase()) } else { self.text.ok(&tok.text, &tok.lower) };
                ok != self.neg
            }
        };
        if !text {
            return false;
        }
        let pos = if self.has_pos { self.pos.ok(&r.tag) } else { true };
        pos != self.npos
    }

    fn matches(&self, tok: &ATok) -> bool {
        if let Some(sb) = self.sb {
            if tok.space_before != sb {
                return false;
            }
        }
        tok.readings.iter().any(|r| self.reading(tok, r))
    }

    /// Only the first reading (LanguageTool checks "next" exceptions of a token without skip so).
    fn matches_first(&self, tok: &ATok) -> bool {
        if let Some(sb) = self.sb {
            if tok.space_before != sb {
                return false;
            }
        }
        tok.readings.first().is_some_and(|r| self.reading(tok, r))
    }
}

struct PTok {
    cond: Cond,
    exc_cur: Vec<Cond>,
    exc_next: Vec<Cond>,
    exc_prev: Vec<Cond>,
    and_: Vec<PTok>,
    or_: Vec<PTok>,
    skip: i32,
    min: usize,
    max: usize,
    /// A literal the token needs (lower case), for skipping rules quickly.
    literal: Option<String>,
    /// The same word as an earlier pattern token (0-based) and whether lemmas count.
    backref: Option<(usize, bool, bool)>,
    /// The unify group the token belongs to.
    group: Option<usize>,
}

impl PTok {
    fn new(j: &TokJson) -> Option<PTok> {
        let cond = Cond::new(&j.t, j.re, j.cs, j.inf, j.neg, &j.pos, j.pre, j.npos, j.sb)?;
        let (mut cur, mut next, mut prev) = (Vec::new(), Vec::new(), Vec::new());
        for e in &j.exc {
            let c = Cond::new(&e.t, e.re, e.cs, e.inf, e.neg, &e.pos, e.pre, e.npos, e.sb)?;
            match e.scope.as_deref() {
                Some("next") => next.push(c),
                Some("previous") => prev.push(c),
                _ => cur.push(c),
            }
        }
        let and_ = j.and_.iter().map(PTok::new).collect::<Option<Vec<_>>>()?;
        let or_ = j.or_.iter().map(PTok::new).collect::<Option<Vec<_>>>()?;
        let min = j.min.unwrap_or(1).max(0) as usize;
        let max = match j.max { Some(m) if m < 0 => usize::MAX, Some(m) => m.max(0) as usize, None => min.max(1) };
        let literal = match (&j.t, j.re, j.neg, j.inf) {
            (Some(t), false, false, false) if min > 0 => Some(t.to_lowercase()),
            _ => None,
        };
        // Inside patterns LanguageTool counts from 0 (in messages from 1).
        let backref = j.backref.map(|n| (n, j.inf, j.cs));
        Some(PTok { cond, exc_cur: cur, exc_next: next, exc_prev: prev, and_, or_, skip: j.skip.unwrap_or(0), min, max, literal, backref, group: j.u })
    }

    fn matches(&self, toks: &[ATok], i: usize) -> bool {
        let Some(t) = toks.get(i) else { return false };
        if !self.cond.matches(t) || self.exc_cur.iter().any(|c| c.matches(t)) {
            return false;
        }
        if !self.and_.iter().all(|a| a.matches(toks, i)) {
            return false;
        }
        if !self.or_.is_empty() && !self.or_.iter().any(|a| a.matches(toks, i)) {
            return false;
        }
        if !self.exc_prev.is_empty() && i > 0 && self.exc_prev.iter().any(|c| c.matches(&toks[i - 1])) {
            return false;
        }
        true
    }

    fn next_excepted(&self, toks: &[ATok], i: usize) -> bool {
        toks.get(i).is_some_and(|t| self.exc_next.iter().any(|c| c.matches(t)))
    }

    fn next_excepted_first(&self, toks: &[ATok], i: usize) -> bool {
        toks.get(i).is_some_and(|t| self.exc_next.iter().any(|c| c.matches_first(t)))
    }
}

struct Pattern {
    toks: Vec<PTok>,
    unify: Vec<UnifyJson>,
}

/// The unifier of the language being checked (set when its checker loads).
pub static UNIFIER: std::sync::OnceLock<Unifier> = std::sync::OnceLock::new();

impl Pattern {
    fn new(j: &[TokJson], unify: &[UnifyJson]) -> Option<Pattern> {
        Some(Pattern { toks: j.iter().map(PTok::new).collect::<Option<Vec<_>>>()?, unify: unify.to_vec() })
    }

    /// Matches starting at token `ti`: the tokens each pattern token took (start, count).
    fn at(&self, toks: &[ATok], ti: usize) -> Option<Vec<(usize, usize)>> {
        let mut spans = vec![(0, 0); self.toks.len()];
        let mut budget = 20_000u32;
        if !self.from(toks, 0, ti, &mut spans, &mut budget) {
            return None;
        }
        for g in 0..self.unify.len() {
            let ok = self.unified(toks, &spans, g).is_some();
            if ok == self.unify[g].neg {
                return None;
            }
        }
        Some(spans)
    }

    /// The tokens of unify group `g` and, per token, which readings unify (None: they do not).
    fn unified(&self, toks: &[ATok], spans: &[(usize, usize)], g: usize) -> Option<(Vec<usize>, Vec<Vec<bool>>)> {
        let members: Vec<usize> = self.toks.iter().enumerate().filter(|(_, t)| t.group == Some(g))
            .flat_map(|(i, _)| { let (s, n) = spans[i]; s..s + n }).collect();
        let u = UNIFIER.get()?;
        let refs: Vec<&ATok> = members.iter().map(|&k| &toks[k]).collect();
        let keep = u.unify(&refs, &self.unify[g].features)?;
        Some((members, keep))
    }

    fn from(&self, toks: &[ATok], pi: usize, ti: usize, spans: &mut Vec<(usize, usize)>, budget: &mut u32) -> bool {
        if pi == self.toks.len() {
            return true;
        }
        if *budget == 0 {
            return false;
        }
        *budget -= 1;
        let p = &self.toks[pi];
        let room = toks.len().saturating_sub(ti);
        let most = p.max.min(room);
        let mut k = 0;
        while k < most && p.matches(toks, ti + k) && self.same_as(p, toks, ti + k, spans) {
            k += 1;
        }
        if k < p.min {
            return false;
        }
        for c in (p.min..=k).rev() {
            if c > 0 && p.skip == 0 && p.next_excepted_first(toks, ti + c) {
                continue;
            }
            spans[pi] = (ti, c);
            let next = ti + c;
            let limit = if p.skip < 0 { toks.len() } else { next + p.skip as usize };
            let mut s = next;
            loop {
                if self.from(toks, pi + 1, s, spans, budget) {
                    return true;
                }
                if s >= limit || s >= toks.len() || pi + 1 == self.toks.len() {
                    break;
                }
                // A skipping token's "next" exceptions stop the skip.
                if p.skip != 0 && p.next_excepted(toks, s) {
                    break;
                }
                s += 1;
            }
        }
        false
    }

    /// A back-reference: the token is the word an earlier pattern token matched.
    fn same_as(&self, p: &PTok, toks: &[ATok], i: usize, spans: &[(usize, usize)]) -> bool {
        let Some((n, inf, cs)) = p.backref else { return true };
        let Some(&(at, count)) = spans.get(n) else { return true };
        if count == 0 {
            return false;
        }
        let (a, b) = (&toks[at], &toks[i]);
        if inf {
            a.readings.iter().any(|x| b.readings.iter().any(|y| x.lemma.to_lowercase() == y.lemma.to_lowercase()))
        } else if cs {
            a.text == b.text
        } else {
            a.lower == b.lower
        }
    }

    /// Literals the pattern needs (all of them must be in the sentence).
    fn literals(&self) -> Vec<String> {
        self.toks.iter().filter_map(|t| t.literal.clone()).collect()
    }
}

pub struct Rule {
    pub id: String,
    pub category: String,
    pub name: String,
    pattern: Option<Pattern>,
    regex: Option<(Regex, usize)>,
    ma: usize,
    mb: usize,
    anti: Vec<Pattern>,
    msg: Vec<Piece>,
    sugg: Vec<Vec<Piece>>,
    pub short: Option<String>,
    pub literals: Vec<String>,
}

/// A finding: char offsets in the paragraph.
#[derive(Debug, Clone, serde::Serialize)]
pub struct Finding {
    pub start: usize,
    pub end: usize,
    pub rule: String,
    pub category: String,
    pub message: String,
    pub short: String,
    pub suggestions: Vec<String>,
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

impl Rule {
    pub fn new(j: RuleJson) -> Option<Rule> {
        let pattern = match &j.pat { Some(p) => Some(Pattern::new(p, &j.unify)?), None => None };
        let regex = match &j.regex {
            Some(r) => {
                let flags = if r.cs { "" } else { "(?i)" };
                let mut s = r.re.clone();
                for (a, b) in [("\\p{Punct}", "[\\p{P}\\p{S}]"), ("(?iu)", "(?i)"), ("(?u)", "")] {
                    s = s.replace(a, b);
                }
                Some((Regex::new(&format!("{flags}{s}")).ok()?, r.mark))
            }
            None => None,
        };
        if pattern.is_none() && regex.is_none() {
            return None;
        }
        let anti = j.anti.iter().enumerate().map(|(i, a)| Pattern::new(a, j.antiu.get(i).map_or(&[][..], |v| &v[..]))).collect::<Option<Vec<_>>>()?;
        let literals = pattern.as_ref().map(|p| p.literals()).unwrap_or_default();
        Some(Rule {
            id: j.id, category: j.cat.unwrap_or_default(), name: j.name, pattern, regex, ma: j.ma.unwrap_or(0), mb: j.mb.unwrap_or(0),
            anti, msg: j.msg, sugg: j.sugg, short: j.short, literals,
        })
    }

    /// The text a pattern token matched (several tokens keep their spaces).
    fn span_text(toks: &[ATok], span: (usize, usize), chars: &[char]) -> String {
        if span.1 == 0 {
            return String::new();
        }
        let a = toks[span.0].start;
        let b = toks[span.0 + span.1 - 1].end;
        chars[a..b].iter().collect()
    }

    /// A <match>: the token's text, converted; None when it needs word forms we cannot build.
    fn render_match(m: &MatchRef, toks: &[ATok], spans: &[(usize, usize)], chars: &[char]) -> Option<String> {
        if m.pos.is_some() || m.pr.is_some() {
            return None;
        }
        let span = *spans.get(m.no.checked_sub(1)?)?;
        let mut s = Rule::span_text(toks, span, chars);
        if let (Some(rm), Some(rr)) = (&m.rm, &m.rr) {
            let re = Regex::new(&rm.replace("(?iu)", "(?i)").replace("(?u)", "")).ok()?;
            s = re.replace_all(&s, rr.as_str()).into_owned();
        }
        Some(match m.cc.as_deref() {
            Some("startupper") => upper_first(&s),
            Some("startlower") => lower_first(&s),
            Some("allupper") => s.to_uppercase(),
            Some("alllower") => s.to_lowercase(),
            _ => s,
        })
    }

    /// Text with \1 references and <match> pieces; suggestions found inside go to `sugg`.
    fn render(pieces: &[Piece], toks: &[ATok], spans: &[(usize, usize)], chars: &[char], sugg: &mut Vec<String>, ok: &mut bool) -> String {
        let mut out = String::new();
        for p in pieces {
            match p {
                Piece::Text(t) => {
                    let mut rest = t.as_str();
                    while let Some(i) = rest.find('\\') {
                        out.push_str(&rest[..i]);
                        let digits: String = rest[i + 1..].chars().take_while(char::is_ascii_digit).collect();
                        if digits.is_empty() {
                            out.push('\\');
                            rest = &rest[i + 1..];
                            continue;
                        }
                        let n: usize = digits.parse().unwrap_or(0);
                        if let Some(span) = n.checked_sub(1).and_then(|k| spans.get(k)) {
                            out.push_str(&Rule::span_text(toks, *span, chars));
                        }
                        rest = &rest[i + 1 + digits.len()..];
                    }
                    out.push_str(rest);
                }
                Piece::Match { m } => match Rule::render_match(m, toks, spans, chars) {
                    Some(s) => out.push_str(&s),
                    None => *ok = false,
                },
                Piece::Sugg { s } => {
                    let mut good = true;
                    let mut inner = Vec::new();
                    let text = Rule::render(s, toks, spans, chars, &mut inner, &mut good);
                    if good {
                        sugg.push(text.clone());
                        out.push('„');
                        out.push_str(&text);
                        out.push('“');
                    } else {
                        out.push_str(&text);
                    }
                }
            }
        }
        out
    }

    fn finding(&self, toks: &[ATok], spans: &[(usize, usize)], chars: &[char]) -> Option<Finding> {
        let (ma, mb) = (self.ma.min(spans.len() - 1), self.mb.min(spans.len() - 1));
        // The marked tokens (pattern tokens that matched nothing are left out).
        let marked: Vec<(usize, usize)> = spans[ma..=mb].iter().copied().filter(|s| s.1 > 0).collect();
        let (first, last) = (marked.first()?, marked.last()?);
        let start = toks[first.0].start;
        let end = toks[last.0 + last.1 - 1].end;
        if end <= start {
            return None;
        }
        let mut sugg = Vec::new();
        let mut ok = true;
        let message = Rule::render(&self.msg, toks, spans, chars, &mut sugg, &mut ok);
        for s in &self.sugg {
            let mut good = true;
            let text = Rule::render(s, toks, spans, chars, &mut Vec::new(), &mut good);
            if good {
                sugg.push(text);
            }
        }
        // At the start of a sentence suggestions start with a capital, as the text does.
        let original: String = chars[start..end].iter().collect();
        if first.0 == 1 && original.chars().next().is_some_and(char::is_uppercase) {
            for s in &mut sugg {
                *s = upper_first(s);
            }
        }
        sugg.dedup();
        sugg.retain(|s| *s != original && !s.trim().is_empty());
        Some(Finding { start, end, rule: self.id.clone(), category: self.category.clone(), message: message.trim().to_string(),
            short: self.short.clone().unwrap_or_default(), suggestions: sugg })
    }

    /// The rule's findings in a sentence (tokens with SENT_START at 0; `text` is the sentence).
    pub fn check(&self, toks: &[ATok], chars: &[char], sentence: (usize, usize), out: &mut Vec<Finding>) {
        if let Some((re, mark)) = &self.regex {
            let text: String = chars[sentence.0..sentence.1].iter().collect();
            let offsets: Vec<usize> = text.char_indices().map(|(b, _)| b).chain(std::iter::once(text.len())).collect();
            let char_at = |byte: usize| offsets.iter().position(|&b| b >= byte).unwrap_or(offsets.len() - 1);
            for caps in re.captures_iter(&text).flatten() {
                let Some(g) = caps.get(*mark) else { continue };
                if g.start() == g.end() {
                    continue;
                }
                let (a, b) = (sentence.0 + char_at(g.start()), sentence.0 + char_at(g.end()));
                let mut sugg = Vec::new();
                let mut ok = true;
                let message = Rule::render(&self.msg, toks, &[], chars, &mut sugg, &mut ok);
                out.push(Finding { start: a, end: b, rule: self.id.clone(), category: self.category.clone(), message: message.trim().to_string(),
                    short: self.short.clone().unwrap_or_default(), suggestions: sugg });
            }
            return;
        }
        let Some(p) = &self.pattern else { return };
        let mut anti: Option<Vec<(usize, usize)>> = None;
        let mut ti = 0;
        while ti < toks.len() {
            let Some(spans) = p.at(toks, ti) else { ti += 1; continue };
            let used: Vec<&(usize, usize)> = spans.iter().filter(|s| s.1 > 0).collect();
            let (a, b) = match (used.first(), used.last()) { (Some(f), Some(l)) => (f.0, l.0 + l.1), _ => { ti += 1; continue } };
            // Antipatterns: a finding inside text an antipattern matches is no finding.
            let covers = anti.get_or_insert_with(|| {
                let mut v = Vec::new();
                for ap in &self.anti {
                    let mut i = 0;
                    while i < toks.len() {
                        if let Some(s) = ap.at(toks, i) {
                            let u: Vec<&(usize, usize)> = s.iter().filter(|x| x.1 > 0).collect();
                            if let (Some(f), Some(l)) = (u.first(), u.last()) {
                                v.push((f.0, l.0 + l.1));
                            }
                        }
                        i += 1;
                    }
                }
                v
            });
            let immune = spans[self.ma.min(spans.len() - 1)..=self.mb.min(spans.len() - 1)].iter()
                .any(|&(s, n)| (s..s + n).any(|k| toks[k].immune));
            if !immune && !covers.iter().any(|&(x, y)| x < b && a < y) {
                if let Some(f) = self.finding(toks, &spans, chars) {
                    out.push(f);
                }
            }
            ti = b.max(ti + 1);
        }
    }
}

// ----- disambiguation ----------------------------------------------------------------------------

enum Action {
    Remove,
    Filter,
    Add,
    Replace,
    Immunize,
    Unify,
}

/// A disambiguation rule (LanguageTool's disambiguation.xml): it changes the readings of the
/// marked tokens before the rules run, or protects them.
pub struct Disambig {
    pattern: Pattern,
    ma: usize,
    mb: usize,
    action: Action,
    pos: Option<Regex>,
    wd: Vec<(Option<String>, Option<String>)>,
    anti: Vec<Pattern>,
    literals: Vec<String>,
}

impl Disambig {
    pub fn new(j: DisambigJson) -> Option<Disambig> {
        let pattern = Pattern::new(&j.pat, &j.unify)?;
        let action = match j.action.as_str() {
            "unify" => Action::Unify,
            "remove" => Action::Remove,
            "filter" => Action::Filter,
            "add" => Action::Add,
            "replace" => Action::Replace,
            "immunize" => Action::Immunize,
            _ => return None,
        };
        let pos = match &j.pos { Some(p) => Some(compile(p, true)?), None => None };
        let anti = j.anti.iter().enumerate().map(|(i, a)| Pattern::new(a, j.antiu.get(i).map_or(&[][..], |v| &v[..]))).collect::<Option<Vec<_>>>()?;
        let literals = pattern.literals();
        Some(Disambig { pattern, ma: j.ma, mb: j.mb, action, pos, wd: j.wd.into_iter().map(|w| (w.lemma, w.pos)).collect(), anti, literals })
    }

    pub fn needs(&self, words: &std::collections::HashSet<&str>) -> bool {
        self.literals.iter().all(|l| words.contains(l.as_str()))
    }

    pub fn apply(&self, toks: &mut [ATok]) {
        let mut ti = 0;
        while ti < toks.len() {
            let Some(spans) = self.pattern.at(toks, ti) else { ti += 1; continue };
            let used: Vec<&(usize, usize)> = spans.iter().filter(|s| s.1 > 0).collect();
            let end = used.last().map_or(ti + 1, |l| l.0 + l.1);
            let (a, b) = (used.first().map_or(ti, |f| f.0), end);
            let blocked = self.anti.iter().any(|ap| (0..toks.len()).any(|i| ap.at(toks, i).is_some_and(|s| {
                let u: Vec<&(usize, usize)> = s.iter().filter(|x| x.1 > 0).collect();
                matches!((u.first(), u.last()), (Some(f), Some(l)) if f.0 < b && a < l.0 + l.1)
            })));
            if !blocked && matches!(self.action, Action::Unify) {
                // Only the readings that agree stay (der Mann: ART and SUB in the same case …).
                for g in 0..self.pattern.unify.len() {
                    if let Some((members, keep)) = self.pattern.unified(toks, &spans, g) {
                        for (k, ok) in members.into_iter().zip(keep) {
                            let t = &mut toks[k];
                            let readings: Vec<Reading> = t.readings.iter().zip(&ok).filter(|(r, o)| **o || r.tag == "SENT_END" || r.tag == "SENT_START").map(|(r, _)| r.clone()).collect();
                            if readings.iter().any(|r| r.tag != "SENT_END") {
                                t.readings = readings;
                            }
                        }
                    }
                }
            } else if !blocked {
                let (ma, mb) = (self.ma.min(spans.len() - 1), self.mb.min(spans.len() - 1));
                let targets: Vec<usize> = spans[ma..=mb].iter().flat_map(|&(s, n)| s..s + n).collect();
                for k in targets {
                    self.change(&mut toks[k]);
                }
            }
            ti = end.max(ti + 1);
        }
    }

    fn change(&self, t: &mut ATok) {
        let tag_ok = |tag: &str| self.pos.as_ref().is_some_and(|p| p.is_match(tag).unwrap_or(false));
        match self.action {
            Action::Immunize => t.immune = true,
            Action::Unify => {}
            Action::Remove => {
                let keep: Vec<Reading> = if self.wd.is_empty() {
                    t.readings.iter().filter(|r| !tag_ok(&r.tag)).cloned().collect()
                } else {
                    t.readings.iter().filter(|r| !self.wd.iter().any(|(l, p)|
                        l.as_ref().is_none_or(|l| *l == r.lemma) && p.as_ref().is_none_or(|p| *p == r.tag))).cloned().collect()
                };
                if !keep.is_empty() {
                    t.readings = keep;
                }
            }
            Action::Filter => {
                let keep: Vec<Reading> = t.readings.iter().filter(|r| tag_ok(&r.tag) || r.tag == "SENT_END").cloned().collect();
                if keep.iter().any(|r| r.tag != "SENT_END") {
                    t.readings = keep;
                }
            }
            Action::Add | Action::Replace => {
                let new: Vec<Reading> = self.wd.iter().filter_map(|(l, p)| Some(Reading { lemma: l.clone().unwrap_or_else(|| t.text.clone()), tag: p.clone()?, guessed: false })).collect();
                if new.is_empty() {
                    return;
                }
                if matches!(self.action, Action::Replace) {
                    let end = t.readings.iter().any(|r| r.tag == "SENT_END");
                    t.readings = new;
                    if end {
                        t.readings.push(Reading { lemma: t.text.clone(), tag: "SENT_END".into(), guessed: false });
                    }
                } else {
                    t.readings.retain(|r| r.tag != "UNKNOWN");
                    for r in new {
                        if !t.readings.contains(&r) {
                            t.readings.push(r);
                        }
                    }
                }
            }
        }
    }
}
