// Sentences and tokens of a paragraph, as LanguageTool's German tokenizer cuts them: words
// (with inner hyphens, apostrophes and the comma or point of numbers), every other character
// its own token, whitespace only remembered as "space before". Sentences end at . ! ? … when
// what follows starts a new sentence (not after abbreviations, initials or ordinal numbers).
// Offsets are char indices into the paragraph.

#[derive(Debug, Clone)]
pub struct Token {
    pub text: String,
    pub start: usize,
    pub end: usize,
    pub space_before: bool,
}

fn word_char(c: char) -> bool {
    c.is_alphanumeric() || c == '\u{ad}' || matches!(c, '\u{300}'..='\u{36f}')
}

/// The tokens of chars[from..to].
pub fn tokens(chars: &[char], from: usize, to: usize) -> Vec<Token> {
    let mut out = Vec::new();
    let mut i = from;
    let mut space = false;
    while i < to {
        let c = chars[i];
        if c.is_whitespace() || c == '\u{a0}' || c == '\u{fffc}' {
            space = true;
            i += 1;
            continue;
        }
        let start = i;
        if word_char(c) {
            i += 1;
            while i < to {
                let d = chars[i];
                if word_char(d) {
                    i += 1;
                    continue;
                }
                // Inner hyphens, apostrophes and the separators of numbers belong to the word.
                let next = chars.get(i + 1).copied().unwrap_or(' ');
                let prev = chars[i - 1];
                let inner = match d {
                    '-' | '\u{2011}' => word_char(next),
                    '\'' | '\u{2019}' => prev.is_alphabetic() && next.is_alphabetic(),
                    '.' | ',' => prev.is_ascii_digit() && next.is_ascii_digit(),
                    _ => false,
                };
                if inner && i + 1 < to {
                    i += 2;
                } else {
                    break;
                }
            }
        } else {
            i += 1;
        }
        out.push(Token { text: chars[start..i].iter().collect(), start, end: i, space_before: space });
        space = false;
    }
    out
}

const ABBREVIATIONS: &[&str] = &[
    "bzw", "usw", "ca", "Nr", "Dr", "Prof", "St", "vgl", "Vgl", "evtl", "ggf", "inkl", "bzgl", "etc", "Abb", "Tab", "Kap", "Jh", "Mio", "Mrd",
    "Hr", "Hrn", "Fr", "Str", "Std", "Min", "min", "Sek", "sog", "usf", "zzgl", "Bd", "Hrsg", "Aufl", "Jan", "Feb", "Mär", "Apr", "Jun", "Jul",
    "Aug", "Sep", "Sept", "Okt", "Nov", "Dez", "Mo", "Di", "Mi", "Do", "Fr", "Sa", "So", "Ing", "Dipl", "Mag", "geb", "gest", "Tel", "Anm",
    "Ausg", "Art", "Abs", "Bsp", "bspw", "eigtl", "einschl", "entspr", "gem", "jew", "lt", "max", "mind", "Mrd", "Pkt", "rd", "s", "u", "z",
    "zz", "zzt", "Ziff", "mtl", "Mio", "Co", "Jr", "Mr", "Mrs", "Ms", "vs", "e.V", "Inc", "Ltd",
];

const MONTHS: &[&str] = &["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember", "Jahrhundert", "Jahrhunderts"];

/// The sentences of a paragraph: ranges of its tokens.
pub fn sentences(toks: &[Token]) -> Vec<std::ops::Range<usize>> {
    let mut out = Vec::new();
    let mut start = 0;
    let mut i = 0;
    while i < toks.len() {
        let t = &toks[i].text;
        if matches!(t.as_str(), "." | "!" | "?" | "…") {
            // Closing quotes and brackets stay with the sentence.
            let mut end = i + 1;
            while end < toks.len() && matches!(toks[end].text.as_str(), "." | "!" | "?" | "…" | "“" | "\"" | "”" | "«" | "»" | "'" | "’" | ")" | "]") && !toks[end].space_before {
                end += 1;
            }
            let next = toks.get(end);
            let prev = if i > 0 { toks[i - 1].text.as_str() } else { "" };
            let abbrev = t == "." && i > 0 && !toks[i].space_before && (ABBREVIATIONS.contains(&prev) || (prev.chars().count() == 1 && prev.chars().all(char::is_alphabetic)));
            let ordinal = t == "." && i > 0 && !toks[i].space_before && prev.chars().all(|c| c.is_ascii_digit())
                && next.is_some_and(|n| MONTHS.contains(&n.text.as_str()) || n.text.chars().next().is_some_and(char::is_lowercase));
            let starts_sentence = next.is_none_or(|n| n.space_before && n.text.chars().next().is_some_and(|c| c.is_uppercase() || c.is_ascii_digit() || "„\"“‚'(«»".contains(c)));
            if !abbrev && !ordinal && starts_sentence {
                out.push(start..end);
                start = end;
                i = end;
                continue;
            }
        }
        i += 1;
    }
    if start < toks.len() {
        out.push(start..toks.len());
    }
    out
}
