"""Reads Morfologik CFSA2 automata (LanguageTool's POS dictionaries, e.g. german.dict) and
looks up words: their lemmas and tags. A reference for the Rust reader (grammar/fsa.rs) and
for building test data.

    python -I tools/lang-data/morfologik.py <file.dict> <word> …
"""
from __future__ import annotations

import sys
from pathlib import Path

BIT_TARGET_NEXT, BIT_LAST_ARC, BIT_FINAL_ARC = 1 << 7, 1 << 6, 1 << 5
LABEL_MASK = (1 << 5) - 1
NUMBERS = 1 << 8


class CFSA2:
    def __init__(self, data: bytes):
        assert data[:4] == b"\\fsa" and data[4] == 0xC6, "not a CFSA2 automaton"
        self.flags = int.from_bytes(data[5:7], "big")
        n = data[7]
        self.labels = data[8:8 + n]
        self.arcs = data[8 + n:]
        self.numbers = bool(self.flags & NUMBERS)

    def vint(self, off: int) -> int:
        a = self.arcs
        b = a[off]
        v = b & 0x7F
        shift = 7
        while b & 0x80:
            off += 1
            b = a[off]
            v |= (b & 0x7F) << shift
            shift += 7
        return v

    def skip_vint(self, off: int) -> int:
        while self.arcs[off] & 0x80:
            off += 1
        return off + 1

    def first_arc(self, node: int) -> int:
        return self.skip_vint(node) if self.numbers else node

    def is_last(self, arc: int) -> bool:
        return bool(self.arcs[arc] & BIT_LAST_ARC)

    def is_final(self, arc: int) -> bool:
        return bool(self.arcs[arc] & BIT_FINAL_ARC)

    def label(self, arc: int) -> int:
        i = self.arcs[arc] & LABEL_MASK
        return self.labels[i] if i else self.arcs[arc + 1]

    def skip_arc(self, off: int) -> int:
        flag = self.arcs[off]
        off += 1
        if not flag & LABEL_MASK:
            off += 1
        if not flag & BIT_TARGET_NEXT:
            off = self.skip_vint(off)
        return off

    def next_arc(self, arc: int) -> int:
        return 0 if self.is_last(arc) else self.skip_arc(arc)

    def target(self, arc: int) -> int:
        if self.arcs[arc] & BIT_TARGET_NEXT:
            while not self.is_last(arc):
                arc = self.next_arc(arc)
            return self.skip_arc(arc)
        return self.vint(arc + (2 if not self.arcs[arc] & LABEL_MASK else 1))

    def root(self) -> int:
        return self.target(self.first_arc(0))

    def arcs_of(self, node: int):
        arc = self.first_arc(node)
        while arc:
            yield arc
            arc = self.next_arc(arc)

    def find(self, node: int, label: int) -> int:
        for arc in self.arcs_of(node):
            if self.label(arc) == label:
                return arc
        return 0

    def walk(self, prefix: bytes) -> int | None:
        """The node after `prefix`, or None."""
        node = self.root()
        for i, b in enumerate(prefix):
            arc = self.find(node, b)
            if not arc:
                return None
            if i == len(prefix) - 1:
                return self.target(arc) if self.target(arc) else None
            node = self.target(arc)
            if not node:
                return None
        return node

    def sequences(self, node: int, acc: bytes = b""):
        for arc in self.arcs_of(node):
            seq = acc + bytes([self.label(arc)])
            if self.is_final(arc):
                yield seq
            t = self.target(arc)
            if t:
                yield from self.sequences(t, seq)


def decode_suffix(word: bytes, enc: bytes) -> bytes:
    """Morfologik's TrimSuffixEncoder: the first byte says how many bytes to cut from the end."""
    k = enc[0] - ord("A")
    if k < 0 or k > len(word):
        return enc[1:]
    return word[:len(word) - k] + enc[1:]


def lookup(fsa: CFSA2, word: str, sep: bytes = b"_") -> list[tuple[str, str]]:
    w = word.encode("utf-8")
    node = fsa.walk(w + sep)
    out = []
    if node is None:
        return out
    for seq in fsa.sequences(node):
        enc, _, tag = seq.partition(sep)
        out.append((decode_suffix(w, enc).decode("utf-8", "replace"), tag.decode("utf-8", "replace")))
    return out


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    fsa = CFSA2(Path(sys.argv[1]).read_bytes())
    for word in sys.argv[2:]:
        print(word, lookup(fsa, word))
