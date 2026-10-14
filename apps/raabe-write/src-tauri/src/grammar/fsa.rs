// Morfologik CFSA2 automata, the format of LanguageTool's POS dictionaries (german.dict, D-051).
// A dictionary maps "word_encodedLemma_tag" sequences; the lemma is stored as Morfologik's
// SUFFIX encoding (the first byte says how many bytes to cut from the word's end, the rest is
// appended). tools/lang-data/morfologik.py is the same reader in Python.

const BIT_TARGET_NEXT: u8 = 1 << 7;
const BIT_LAST_ARC: u8 = 1 << 6;
const BIT_FINAL_ARC: u8 = 1 << 5;
const LABEL_MASK: u8 = (1 << 5) - 1;
const FLAG_NUMBERS: u16 = 1 << 8;

pub struct Fsa {
    arcs: Vec<u8>,
    labels: Vec<u8>,
    numbers: bool,
}

impl Fsa {
    pub fn new(data: &[u8]) -> Option<Fsa> {
        if data.len() < 8 || &data[..4] != b"\\fsa" || data[4] != 0xC6 {
            return None;
        }
        let flags = u16::from_be_bytes([data[5], data[6]]);
        let n = data[7] as usize;
        Some(Fsa { labels: data[8..8 + n].to_vec(), arcs: data[8 + n..].to_vec(), numbers: flags & FLAG_NUMBERS != 0 })
    }

    fn vint(&self, mut off: usize) -> usize {
        let mut b = self.arcs[off];
        let mut v = (b & 0x7F) as usize;
        let mut shift = 7;
        while b & 0x80 != 0 {
            off += 1;
            b = self.arcs[off];
            v |= ((b & 0x7F) as usize) << shift;
            shift += 7;
        }
        v
    }

    fn skip_vint(&self, mut off: usize) -> usize {
        while self.arcs[off] & 0x80 != 0 {
            off += 1;
        }
        off + 1
    }

    fn first_arc(&self, node: usize) -> usize {
        if self.numbers { self.skip_vint(node) } else { node }
    }

    fn is_last(&self, arc: usize) -> bool {
        self.arcs[arc] & BIT_LAST_ARC != 0
    }

    fn is_final(&self, arc: usize) -> bool {
        self.arcs[arc] & BIT_FINAL_ARC != 0
    }

    fn label(&self, arc: usize) -> u8 {
        let i = (self.arcs[arc] & LABEL_MASK) as usize;
        if i > 0 { self.labels[i] } else { self.arcs[arc + 1] }
    }

    fn skip_arc(&self, mut off: usize) -> usize {
        let flag = self.arcs[off];
        off += 1;
        if flag & LABEL_MASK == 0 {
            off += 1;
        }
        if flag & BIT_TARGET_NEXT == 0 {
            off = self.skip_vint(off);
        }
        off
    }

    fn next_arc(&self, arc: usize) -> usize {
        if self.is_last(arc) { 0 } else { self.skip_arc(arc) }
    }

    fn target(&self, mut arc: usize) -> usize {
        if self.arcs[arc] & BIT_TARGET_NEXT != 0 {
            while !self.is_last(arc) {
                arc = self.next_arc(arc);
            }
            self.skip_arc(arc)
        } else {
            self.vint(arc + if self.arcs[arc] & LABEL_MASK == 0 { 2 } else { 1 })
        }
    }

    fn root(&self) -> usize {
        self.target(self.first_arc(0))
    }

    fn find(&self, node: usize, label: u8) -> usize {
        let mut arc = self.first_arc(node);
        while arc != 0 {
            if self.label(arc) == label {
                return arc;
            }
            arc = self.next_arc(arc);
        }
        0
    }

    /// The node reached after `prefix` (None when no sequence starts with it).
    fn walk(&self, prefix: &[u8]) -> Option<usize> {
        let mut node = self.root();
        for &b in prefix {
            let arc = self.find(node, b);
            if arc == 0 {
                return None;
            }
            node = self.target(arc);
            if node == 0 {
                return None;
            }
        }
        Some(node)
    }

    fn collect(&self, node: usize, acc: &mut Vec<u8>, out: &mut Vec<Vec<u8>>, depth: usize) {
        if depth > 200 {
            return;
        }
        let mut arc = self.first_arc(node);
        while arc != 0 {
            acc.push(self.label(arc));
            if self.is_final(arc) {
                out.push(acc.clone());
            }
            let t = self.target(arc);
            if t != 0 {
                self.collect(t, acc, out, depth + 1);
            }
            acc.pop();
            arc = self.next_arc(arc);
        }
    }

    /// Every (lemma, tag) of a word in a dictionary built with separator `sep` and SUFFIX encoding.
    pub fn lookup(&self, word: &str, sep: u8) -> Vec<(String, String)> {
        let w = word.as_bytes();
        let mut key = w.to_vec();
        key.push(sep);
        let Some(node) = self.walk(&key) else { return Vec::new() };
        let mut seqs = Vec::new();
        self.collect(node, &mut Vec::new(), &mut seqs, 0);
        seqs.into_iter().filter_map(|s| {
            let i = s.iter().position(|&b| b == sep)?;
            let (enc, tag) = (&s[..i], &s[i + 1..]);
            let lemma = decode_suffix(w, enc);
            Some((String::from_utf8_lossy(&lemma).into_owned(), String::from_utf8_lossy(tag).into_owned()))
        }).collect()
    }
}

fn decode_suffix(word: &[u8], enc: &[u8]) -> Vec<u8> {
    if enc.is_empty() {
        return word.to_vec();
    }
    let k = enc[0] as i32 - b'A' as i32;
    if k < 0 || k as usize > word.len() {
        return enc[1..].to_vec();
    }
    let mut out = word[..word.len() - k as usize].to_vec();
    out.extend_from_slice(&enc[1..]);
    out
}
