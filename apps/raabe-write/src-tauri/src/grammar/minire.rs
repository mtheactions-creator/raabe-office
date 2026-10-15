// A small regular-expression matcher for the rules' token and tag conditions (D-052). The rules
// hold tens of thousands of short expressions ("SUB:NOM:.*", "[NJ].* p", "(?i)wiede?r"); compiled
// with the regex crate each takes 14-20 KB, which made a language's checker 100-300 MB. This one
// parses Java's syntax into a Pike VM program of a few dozen instructions and matches whole
// tokens (anchored), case-insensitively by lower-casing like Java's UNICODE_CASE. What it does
// not know (lookaround, back-references, word boundaries, \p{P}, nested classes) is left to
// fancy-regex.

use std::cell::RefCell;

#[derive(Clone, Copy)]
enum Prop {
    Digit,
    Word,
    Space,
    Letter,
    Upper,
    Lower,
    Numeric,
}

impl Prop {
    fn test(self, c: char) -> bool {
        match self {
            Prop::Digit => c.is_ascii_digit(),
            Prop::Word => c.is_alphanumeric() || c == '_',
            Prop::Space => c.is_whitespace(),
            Prop::Letter => c.is_alphabetic(),
            Prop::Upper => c.is_uppercase(),
            Prop::Lower => c.is_lowercase(),
            Prop::Numeric => c.is_numeric(),
        }
    }
}

#[derive(Clone)]
enum Item {
    Range(char, char),
    Prop(Prop, bool),
}

#[derive(Clone)]
struct Class {
    items: Vec<Item>,
    neg: bool,
    /// Ranges compare case-insensitively; properties keep their case (Java leaves \p{Lu} alone).
    ci: bool,
}

impl Class {
    fn matches(&self, c: char) -> bool {
        let inside = self.items.iter().any(|it| match *it {
            Item::Range(a, b) => {
                let hit = |x: char| a <= x && x <= b;
                hit(c) || self.ci && (hit(lower(c)) || hit(upper(c)))
            }
            Item::Prop(p, neg) => p.test(c) != neg,
        });
        inside != self.neg
    }
}

#[derive(Clone)]
enum Node {
    Char(char, bool),
    Class(Class),
    Any,
    Cat(Vec<Node>),
    Alt(Vec<Node>),
    Rep(Box<Node>, u32, Option<u32>),
}

enum Inst {
    Char(char),
    /// A lower-cased letter, compared with the lower-cased input.
    CharI(char),
    Class(Box<Class>),
    Any,
    Split(u32, u32),
    Jmp(u32),
    Match,
}

pub struct MiniRe {
    prog: Vec<Inst>,
    /// "SUB:NOM:.*", "ADJ.*|PA2.*": literal beginnings, tested without the VM.
    prefixes: Option<Vec<Vec<(char, bool)>>>,
    /// What every match starts with ("SUB:" in "SUB:(NOM|AKK):.*"), or the first letters of the
    /// branches: most tests fail there, before the VM runs.
    lead: Vec<(char, bool)>,
    firsts: Option<Vec<(char, bool)>>,
    #[cfg(test)]
    pub src: String,
}

#[cfg(test)]
pub static PROFILE: std::sync::Mutex<Option<std::collections::HashMap<String, (u64, u128)>>> = std::sync::Mutex::new(None);

/// The letters a branch starts with, up to its first non-literal part.
fn leading(n: &Node) -> Vec<(char, bool)> {
    let mut out = Vec::new();
    match n {
        Node::Cat(items) => {
            for x in items {
                match x {
                    Node::Char(c, ci) => out.push(if *ci { (lower(*c), true) } else { (*c, false) }),
                    _ => break,
                }
            }
        }
        Node::Char(c, ci) => out.push(if *ci { (lower(*c), true) } else { (*c, false) }),
        _ => {}
    }
    out
}

/// The literal beginning of a branch that ends in ".*" (case-insensitive letters marked).
fn prefix(n: &Node) -> Option<Vec<(char, bool)>> {
    let Node::Cat(items) = n else { return None };
    let (last, init) = items.split_last()?;
    if !matches!(last, Node::Rep(b, 0, None) if matches!(**b, Node::Any)) {
        return None;
    }
    init.iter().map(|x| match x {
        Node::Char(c, ci) => Some(if *ci { (lower(*c), true) } else { (*c, false) }),
        _ => None,
    }).collect()
}

fn starts(s: &str, p: &[(char, bool)]) -> bool {
    let mut it = s.chars();
    p.iter().all(|&(c, ci)| it.next().is_some_and(|x| if ci { lower(x) == c } else { x == c }))
}

fn lower(c: char) -> char {
    let mut l = c.to_lowercase();
    match (l.next(), l.next()) {
        (Some(x), None) => x,
        _ => c,
    }
}

fn upper(c: char) -> char {
    let mut u = c.to_uppercase();
    match (u.next(), u.next()) {
        (Some(x), None) => x,
        _ => c,
    }
}

struct Parser<'a> {
    s: &'a [char],
    i: usize,
}

impl Parser<'_> {
    fn peek(&self) -> Option<char> {
        self.s.get(self.i).copied()
    }

    fn eat(&mut self, c: char) -> bool {
        if self.peek() == Some(c) {
            self.i += 1;
            true
        } else {
            false
        }
    }

    fn at(&self, s: &str) -> bool {
        s.chars().enumerate().all(|(k, c)| self.s.get(self.i + k) == Some(&c))
    }

    fn alt(&mut self, ci: bool) -> Option<Node> {
        // An inline (?i) or (?-i) holds for the rest of its group, later alternatives included.
        let mut ci = ci;
        let mut alts = vec![self.cat(&mut ci)?];
        while self.eat('|') {
            alts.push(self.cat(&mut ci)?);
        }
        Some(if alts.len() == 1 { alts.pop()? } else { Node::Alt(alts) })
    }

    fn cat(&mut self, ci: &mut bool) -> Option<Node> {
        let mut items = Vec::new();
        while let Some(c) = self.peek() {
            if c == '|' || c == ')' {
                break;
            }
            if self.at("(?i)") {
                self.i += 4;
                *ci = true;
                continue;
            }
            if self.at("(?-i)") {
                self.i += 5;
                *ci = false;
                continue;
            }
            // ^ and $ at a branch's ends say nothing new: the whole token is matched anyway.
            if c == '^' && items.is_empty() || c == '$' && matches!(self.s.get(self.i + 1), None | Some('|' | ')')) {
                self.i += 1;
                continue;
            }
            let atom = self.atom(*ci)?;
            items.push(self.repeat(atom)?);
        }
        Some(Node::Cat(items))
    }

    fn number(&mut self) -> Option<u32> {
        let start = self.i;
        while self.peek().is_some_and(|c| c.is_ascii_digit()) {
            self.i += 1;
        }
        if start == self.i {
            return None;
        }
        self.s[start..self.i].iter().collect::<String>().parse().ok()
    }

    fn repeat(&mut self, atom: Node) -> Option<Node> {
        let (min, max) = match self.peek() {
            Some('*') => { self.i += 1; (0, None) }
            Some('+') => { self.i += 1; (1, None) }
            Some('?') => { self.i += 1; (0, Some(1)) }
            Some('{') => {
                self.i += 1;
                let m = self.number()?;
                let n = if self.eat(',') { if self.peek() == Some('}') { None } else { Some(self.number()?) } } else { Some(m) };
                if !self.eat('}') || n.is_some_and(|n| n < m) || m > 20 || n.is_some_and(|n| n > 20) {
                    return None;
                }
                (m, n)
            }
            _ => return Some(atom),
        };
        // Lazy and possessive forms match the same whole tokens.
        if !self.eat('?') {
            self.eat('+');
        }
        if matches!(self.peek(), Some('*' | '+' | '?' | '{')) {
            return None;
        }
        Some(Node::Rep(Box::new(atom), min, max))
    }

    fn atom(&mut self, ci: bool) -> Option<Node> {
        let c = self.peek()?;
        self.i += 1;
        match c {
            '.' => Some(Node::Any),
            '[' => self.class(ci).map(Node::Class),
            '(' => {
                let mut ci = ci;
                if self.eat('?') {
                    if self.eat(':') {
                    } else if self.peek() == Some('i') && self.s.get(self.i + 1) == Some(&':') {
                        self.i += 2;
                        ci = true;
                    } else if self.peek() == Some('-') && self.s.get(self.i + 1) == Some(&'i') && self.s.get(self.i + 2) == Some(&':') {
                        self.i += 3;
                        ci = false;
                    } else {
                        return None;
                    }
                }
                let inner = self.alt(ci)?;
                if !self.eat(')') {
                    return None;
                }
                Some(inner)
            }
            '\\' => match self.escape(false)? {
                Esc::Char(x) => Some(Node::Char(x, ci)),
                Esc::Prop(p, neg) => Some(Node::Class(Class { items: vec![Item::Prop(p, neg)], neg: false, ci })),
            },
            '^' | '$' | ')' | '*' | '+' | '?' | '{' | '|' => None,
            ']' | '}' => Some(Node::Char(c, ci)),
            _ => Some(Node::Char(c, ci)),
        }
    }

    fn escape(&mut self, in_class: bool) -> Option<Esc> {
        let c = self.peek()?;
        self.i += 1;
        Some(match c {
            'd' => Esc::Prop(Prop::Digit, false),
            'D' => Esc::Prop(Prop::Digit, true),
            'w' => Esc::Prop(Prop::Word, false),
            'W' => Esc::Prop(Prop::Word, true),
            's' => Esc::Prop(Prop::Space, false),
            'S' => Esc::Prop(Prop::Space, true),
            'p' | 'P' => {
                let neg = c == 'P';
                let name: String = if self.eat('{') {
                    let start = self.i;
                    while self.peek().is_some_and(|x| x != '}') {
                        self.i += 1;
                    }
                    let n = self.s[start..self.i].iter().collect();
                    if !self.eat('}') {
                        return None;
                    }
                    n
                } else {
                    self.peek().map(|x| { self.i += 1; x.to_string() })?
                };
                let p = match name.as_str() {
                    "L" | "IsL" | "Alpha" | "IsAlphabetic" | "Alphabetic" | "IsLetter" | "L&" => Prop::Letter,
                    "Lu" | "IsLu" | "Upper" | "javaUpperCase" | "IsUppercase" => Prop::Upper,
                    "Ll" | "IsLl" | "Lower" | "javaLowerCase" | "IsLowercase" => Prop::Lower,
                    "N" | "Nd" | "IsDigit" | "Digit" => Prop::Numeric,
                    _ => return None,
                };
                Esc::Prop(p, neg)
            }
            't' => Esc::Char('\t'),
            'n' => Esc::Char('\n'),
            'r' => Esc::Char('\r'),
            'u' => {
                let hex: String = self.s.get(self.i..self.i + 4)?.iter().collect();
                self.i += 4;
                Esc::Char(char::from_u32(u32::from_str_radix(&hex, 16).ok()?)?)
            }
            'x' => {
                let hex: String = if self.eat('{') {
                    let start = self.i;
                    while self.peek().is_some_and(|x| x != '}') {
                        self.i += 1;
                    }
                    let h = self.s[start..self.i].iter().collect();
                    self.eat('}');
                    h
                } else {
                    let h = self.s.get(self.i..self.i + 2)?.iter().collect();
                    self.i += 2;
                    h
                };
                Esc::Char(char::from_u32(u32::from_str_radix(&hex, 16).ok()?)?)
            }
            // \b, \B, \A, \z, \Q, back-references: not here.
            x if x.is_alphanumeric() => return None,
            x => {
                let _ = in_class;
                Esc::Char(x)
            }
        })
    }

    fn class(&mut self, ci: bool) -> Option<Class> {
        let neg = self.eat('^');
        let mut items = Vec::new();
        let mut first = true;
        loop {
            let c = self.peek()?;
            if c == ']' && !first {
                self.i += 1;
                break;
            }
            first = false;
            // Java's unions, intersections and nested classes: not here.
            if c == '[' || (c == '&' && self.s.get(self.i + 1) == Some(&'&')) {
                return None;
            }
            self.i += 1;
            let lo = if c == '\\' {
                match self.escape(true)? {
                    Esc::Char(x) => x,
                    Esc::Prop(p, n) => {
                        items.push(Item::Prop(p, n));
                        continue;
                    }
                }
            } else {
                c
            };
            if self.peek() == Some('-') && self.s.get(self.i + 1).is_some_and(|&x| x != ']') {
                self.i += 1;
                let h = self.peek()?;
                self.i += 1;
                let hi = if h == '\\' {
                    match self.escape(true)? {
                        Esc::Char(x) => x,
                        Esc::Prop(..) => return None,
                    }
                } else {
                    h
                };
                if hi < lo {
                    return None;
                }
                items.push(Item::Range(lo, hi));
            } else {
                items.push(Item::Range(lo, lo));
            }
        }
        Some(Class { items, neg, ci })
    }
}

enum Esc {
    Char(char),
    Prop(Prop, bool),
}

fn emit(n: &Node, prog: &mut Vec<Inst>) -> Option<()> {
    if prog.len() > 60_000 {
        return None;
    }
    match n {
        Node::Char(c, ci) => {
            let l = lower(*c);
            prog.push(if *ci && (l != *c || upper(*c) != *c) { Inst::CharI(l) } else { Inst::Char(*c) });
        }
        Node::Class(c) => prog.push(Inst::Class(Box::new(c.clone()))),
        Node::Any => prog.push(Inst::Any),
        Node::Cat(items) => {
            for it in items {
                emit(it, prog)?;
            }
        }
        Node::Alt(alts) => {
            // Split(a, next split) … each alternative jumps to the end.
            let mut jumps = Vec::new();
            for (k, a) in alts.iter().enumerate() {
                if k + 1 < alts.len() {
                    let split = prog.len();
                    prog.push(Inst::Split(0, 0));
                    emit(a, prog)?;
                    jumps.push(prog.len());
                    prog.push(Inst::Jmp(0));
                    let next = prog.len();
                    prog[split] = Inst::Split(split as u32 + 1, next as u32);
                } else {
                    emit(a, prog)?;
                }
            }
            let end = prog.len();
            for j in jumps {
                prog[j] = Inst::Jmp(end as u32);
            }
        }
        Node::Rep(body, min, max) => {
            for _ in 0..*min {
                emit(body, prog)?;
            }
            match max {
                None => {
                    let split = prog.len();
                    prog.push(Inst::Split(0, 0));
                    emit(body, prog)?;
                    prog.push(Inst::Jmp(split as u32));
                    let end = prog.len();
                    prog[split] = Inst::Split(split as u32 + 1, end as u32);
                }
                Some(m) => {
                    let mut splits = Vec::new();
                    for _ in *min..*m {
                        splits.push(prog.len());
                        prog.push(Inst::Split(0, 0));
                        emit(body, prog)?;
                    }
                    let end = prog.len();
                    for s in splits {
                        prog[s] = Inst::Split(s as u32 + 1, end as u32);
                    }
                }
            }
        }
    }
    Some(())
}

thread_local! {
    static SCRATCH: RefCell<(Vec<usize>, Vec<usize>, Vec<u32>, Vec<usize>, u32)> = const { RefCell::new((Vec::new(), Vec::new(), Vec::new(), Vec::new(), 0)) };
}

/// The words an expression matches when they are few ("para[iî]tre|viser|ferrer", "chats?"),
/// lower-cased when it is case-insensitive, and whether it is (a leading (?-i) or (?i) decides
/// over the token's flag); None for anything open-ended.
pub fn finite(re: &str, case_insensitive: bool, limit: usize) -> Option<(Vec<String>, bool)> {
    let (node, ci) = parse(re, case_insensitive)?;
    // (?-i:…) inside a case-insensitive expression (or the other way round): not as a set.
    if !uniform(&node, ci) {
        return None;
    }
    let case_insensitive = ci;
    let mut out = expand(&node, limit)?;
    if case_insensitive {
        for w in &mut out {
            *w = w.to_lowercase();
        }
    }
    out.sort();
    out.dedup();
    Some((out, case_insensitive))
}

fn uniform(n: &Node, ci: bool) -> bool {
    match n {
        Node::Char(c, f) => *f == ci || lower(*c) == *c && upper(*c) == *c,
        Node::Class(c) => c.ci == ci,
        Node::Any => true,
        Node::Cat(v) | Node::Alt(v) => v.iter().all(|x| uniform(x, ci)),
        Node::Rep(b, ..) => uniform(b, ci),
    }
}

fn expand(n: &Node, limit: usize) -> Option<Vec<String>> {
    let cross = |a: Vec<String>, b: Vec<String>| -> Option<Vec<String>> {
        if a.len() * b.len() > limit {
            return None;
        }
        Some(a.iter().flat_map(|x| b.iter().map(move |y| format!("{x}{y}"))).collect())
    };
    match n {
        Node::Char(c, _) => Some(vec![c.to_string()]),
        Node::Class(c) => {
            if c.neg {
                return None;
            }
            let mut v = Vec::new();
            for it in &c.items {
                match *it {
                    Item::Range(a, b) if (b as u32) - (a as u32) < 40 => {
                        for x in a..=b {
                            v.push(x.to_string());
                        }
                    }
                    _ => return None,
                }
            }
            Some(v)
        }
        Node::Any => None,
        Node::Cat(items) => {
            let mut acc = vec![String::new()];
            for it in items {
                acc = cross(acc, expand(it, limit)?)?;
            }
            Some(acc)
        }
        Node::Alt(alts) => {
            let mut v = Vec::new();
            for a in alts {
                v.extend(expand(a, limit)?);
                if v.len() > limit {
                    return None;
                }
            }
            Some(v)
        }
        Node::Rep(body, min, max) => {
            let max = (*max)?;
            if max > 4 {
                return None;
            }
            let one = expand(body, limit)?;
            let mut v = Vec::new();
            let mut cur = vec![String::new()];
            for k in 0..=max {
                if k >= *min {
                    v.extend(cur.iter().cloned());
                }
                if k < max {
                    cur = cross(cur, one.clone())?;
                }
                if v.len() > limit {
                    return None;
                }
            }
            Some(v)
        }
    }
}

/// Parses an expression; the flag is the one in force at its start.
fn parse(re: &str, case_insensitive: bool) -> Option<(Node, bool)> {
    let mut chars: Vec<char> = re.chars().collect();
    let mut ci = case_insensitive;
    loop {
        let s: String = chars.iter().take(5).collect();
        if s.starts_with("(?i)") {
            chars.drain(..4);
            ci = true;
        } else if s.starts_with("(?-i)") {
            chars.drain(..5);
            ci = false;
        } else {
            break;
        }
    }
    let mut p = Parser { s: &chars, i: 0 };
    let node = p.alt(ci)?;
    (p.i == chars.len()).then_some((node, ci))
}

impl MiniRe {
    /// The matcher for a Java expression (already rewritten by `compile`'s replacements), or
    /// None when it uses something this matcher does not know.
    pub fn new(re: &str, case_insensitive: bool) -> Option<MiniRe> {
        let mut chars: Vec<char> = re.chars().collect();
        let mut ci = case_insensitive;
        // A leading (?i) or (?-i) applies to the whole expression.
        loop {
            let s: String = chars.iter().take(5).collect();
            if s.starts_with("(?i)") {
                chars.drain(..4);
                ci = true;
            } else if s.starts_with("(?-i)") {
                chars.drain(..5);
                ci = false;
            } else {
                break;
            }
        }
        let mut p = Parser { s: &chars, i: 0 };
        let node = p.alt(ci)?;
        if p.i != chars.len() {
            return None;
        }
        let prefixes = match &node {
            Node::Alt(alts) => alts.iter().map(prefix).collect(),
            n => prefix(n).map(|x| vec![x]),
        };
        let lead = if matches!(node, Node::Alt(_)) { Vec::new() } else { leading(&node) };
        let firsts = match &node {
            Node::Alt(alts) => alts.iter().map(|a| leading(a).first().copied()).collect(),
            _ => None,
        };
        let mut prog = Vec::new();
        if prefixes.is_none() {
            emit(&node, &mut prog)?;
        }
        prog.push(Inst::Match);
        Some(MiniRe { prog, prefixes, lead, firsts, #[cfg(test)] src: re.to_string() })
    }

    /// Whether the whole of `s` matches.
    pub fn is_match(&self, s: &str) -> bool {
        #[cfg(test)]
        {
            let on = PROFILE.lock().unwrap().is_some();
            if on {
                let t = std::time::Instant::now();
                let r = self.is_match_inner(s);
                let d = t.elapsed().as_nanos();
                if let Some(m) = PROFILE.lock().unwrap().as_mut() {
                    let e = m.entry(self.src.chars().take(100).collect()).or_default();
                    e.0 += 1;
                    e.1 += d;
                }
                return r;
            }
        }
        self.is_match_inner(s)
    }

    fn is_match_inner(&self, s: &str) -> bool {
        if let Some(ps) = &self.prefixes {
            return ps.iter().any(|p| starts(s, p));
        }
        if !starts(s, &self.lead) {
            return false;
        }
        if let Some(f) = &self.firsts {
            let Some(c) = s.chars().next() else { return false };
            let lc = lower(c);
            if !f.iter().any(|&(x, ci)| if ci { lc == x } else { c == x }) {
                return false;
            }
        }
        SCRATCH.with(|cell| {
            let mut guard = cell.borrow_mut();
            let (cur, next, marks, stack, generation) = &mut *guard;
            if marks.len() < self.prog.len() {
                marks.resize(self.prog.len(), 0);
            }
            cur.clear();
            *generation = generation.wrapping_add(1);
            if *generation == 0 {
                marks.iter_mut().for_each(|m| *m = 0);
                *generation = 1;
            }
            self.add(cur, marks, stack, *generation, 0);
            for c in s.chars() {
                if cur.is_empty() {
                    return false;
                }
                *generation = generation.wrapping_add(1);
                if *generation == 0 {
                    marks.iter_mut().for_each(|m| *m = 0);
                    *generation = 1;
                }
                next.clear();
                let lc = lower(c);
                for k in 0..cur.len() {
                    let pc = cur[k];
                    let ok = match &self.prog[pc] {
                        Inst::Char(x) => c == *x,
                        Inst::CharI(x) => lc == *x,
                        Inst::Class(cl) => cl.matches(c),
                        Inst::Any => true,
                        _ => false,
                    };
                    if ok {
                        self.add(next, marks, stack, *generation, pc + 1);
                    }
                }
                std::mem::swap(cur, next);
            }
            cur.iter().any(|&pc| matches!(self.prog[pc], Inst::Match))
        })
    }

    /// Adds a thread at `pc`, following splits and jumps.
    fn add(&self, list: &mut Vec<usize>, marks: &mut [u32], stack: &mut Vec<usize>, generation: u32, pc: usize) {
        stack.clear();
        stack.push(pc);
        while let Some(pc) = stack.pop() {
            if marks[pc] == generation {
                continue;
            }
            marks[pc] = generation;
            match self.prog[pc] {
                Inst::Jmp(x) => stack.push(x as usize),
                Inst::Split(a, b) => {
                    stack.push(b as usize);
                    stack.push(a as usize);
                }
                _ => list.push(pc),
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::MiniRe;

    fn m(re: &str, ci: bool, s: &str) -> bool {
        MiniRe::new(re, ci).unwrap_or_else(|| panic!("{re} not supported")).is_match(s)
    }

    #[test]
    fn matches_like_java() {
        assert!(m("SUB:NOM:.*", false, "SUB:NOM:SIN:MAS"));
        assert!(!m("SUB:NOM:.*", false, "SUB:AKK:SIN:MAS"));
        assert!(m("[NJ].* p", false, "N m p"));
        assert!(m("V.* (ind|con|sub).*", false, "V etre ind pres 3 s"));
        assert!(m("ADJ.*|PA[12].*", false, "PA2:PRD:GRU:VER"));
        assert!(m("wiede?r", true, "Wieder"));
        assert!(m("(?i)über", false, "Über"));
        assert!(!m("über", false, "Über"));
        assert!(m("(?-i:\\p{Lu}).*", true, "Haus"));
        assert!(!m("(?-i:\\p{Lu}).*", true, "haus"));
        assert!(m("[^aeiou]+", false, "xyz"));
        assert!(!m("[^aeiou]+", false, "xaz"));
        assert!(m("\\d{4}", false, "1891"));
        assert!(!m("\\d{4}", false, "189"));
        assert!(m("a{2,3}b?", false, "aaab"));
        assert!(m("(der|die|das)", true, "DAS"));
        assert!(m("[a-zäöüß]+", true, "Straße"));
        assert!(m("\\.\\.\\.", false, "..."));
        assert!(m("x*", false, ""));
        assert!(m("[,;]|und|oder", false, ";"));
        assert!(m(".*es$", false, "pommes"));
        assert!(m("deux|tous|(?-i)MES", true, "DEUX"));
        assert!(m("deux|tous|(?-i)MES", true, "MES"));
        assert!(!m("deux|tous|(?-i)MES", true, "mes"));
        assert!(m("(?-i)^[àâa].*|hash", true, "hash"));
        assert!(!m("(?-i)^[àâa].*|hash", true, "HASH"));
        assert!(!m("(?-i)^[àâa].*|hash", true, "Haus"));
        assert!(MiniRe::new("(?!ab).*", false).is_none());
        assert!(MiniRe::new("\\bab", false).is_none());
        assert!(MiniRe::new("(a)\\1", false).is_none());
        assert!(MiniRe::new("[a-z&&[^e]]", false).is_none());
    }
}
