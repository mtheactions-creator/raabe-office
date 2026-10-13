"""Minimal PDF reader (standard library only): where each piece of text is drawn.

Enough of PDF for the files Word and WebView2 export: classic and compressed cross-reference
data, object streams, Flate-compressed content streams, the graphics and text matrices.
Glyphs are not decoded; only positions and sizes are read, which is what layout comparison
needs. Coordinates are returned in points from the top left of each page."""

from __future__ import annotations

import re
import zlib
from dataclasses import dataclass


@dataclass
class TextRun:
    x: float
    y: float  # baseline, from the top of the page
    size: float  # effective font size (text matrix × CTM × Tf)
    chars: int  # bytes shown (a rough length)


@dataclass
class Page:
    width: float
    height: float
    runs: list[TextRun]


# ----- objects ---------------------------------------------------------------------------

_OBJ = re.compile(rb"(\d+)\s+(\d+)\s+obj\b")
_REF = re.compile(rb"(\d+)\s+0\s+R")
WS = b" \t\r\n\x0c\x00"
DELIM = b"()<>[]{}/%"


def _dict_value(d: bytes, key: bytes) -> bytes | None:
    m = re.search(rb"/" + key + rb"(?![A-Za-z0-9])\s*", d)
    if not m:
        return None
    rest = d[m.end():]
    if rest.startswith(b"["):
        depth = 0
        for i, c in enumerate(rest):
            if c == 0x5B:
                depth += 1
            elif c == 0x5D:
                depth -= 1
                if depth == 0:
                    return rest[: i + 1]
    r = _REF.match(rest)
    if r:
        return r.group(0)
    m2 = re.match(rb"[^/\s>\]]+", rest)
    return m2.group(0) if m2 else rest[:1]


class Pdf:
    def __init__(self, data: bytes):
        self.data = data
        self.objects: dict[int, bytes] = {}
        self.streams: dict[int, bytes] = {}
        self._scan()

    def _scan(self) -> None:
        d = self.data
        pending: list[tuple[int, bytes, int]] = []
        for m in _OBJ.finditer(d):
            num = int(m.group(1))
            start = m.end()
            s = d.find(b"stream", start)
            e = d.find(b"endobj", start)
            if s != -1 and (e == -1 or s < e):
                head = d[start:s]
                p = s + 6
                if d[p:p + 2] == b"\r\n":
                    p += 2
                elif d[p:p + 1] in (b"\n", b"\r"):
                    p += 1
                pending.append((num, head, p))
                self.objects[num] = head
            elif e != -1:
                self.objects[num] = d[start:e]
        for num, head, p in pending:
            length = _dict_value(head, b"Length")
            n = self._int(length)
            raw = self.data[p:p + n] if n is not None else self.data[p:self.data.find(b"endstream", p)]
            self.streams[num] = self._decode(head, raw)
        # Object streams hold further objects.
        for num, head in list(self.objects.items()):
            if b"/ObjStm" in head and num in self.streams:
                body = self.streams[num]
                n, first = self._int(_dict_value(head, b"N")), self._int(_dict_value(head, b"First"))
                nums = [int(x) for x in body[:first].split()]
                for i in range(n):
                    onum, off = nums[2 * i], nums[2 * i + 1]
                    end = nums[2 * i + 3] if i + 1 < n else len(body) - first
                    self.objects.setdefault(onum, body[first + off:first + end])

    def _int(self, v: bytes | None) -> int | None:
        if v is None:
            return None
        r = _REF.match(v)
        if r:
            v = self.objects.get(int(r.group(1)), b"").strip()
        try:
            return int(v)
        except ValueError:
            return None

    @staticmethod
    def _decode(head: bytes, raw: bytes) -> bytes:
        if b"FlateDecode" in head:
            try:
                return zlib.decompress(raw)
            except zlib.error:
                return zlib.decompressobj().decompress(raw)
        return raw

    def ref(self, v: bytes | None) -> int | None:
        r = _REF.match(v or b"")
        return int(r.group(1)) if r else None

    def pages(self) -> list[int]:
        root = None
        for num, body in self.objects.items():
            if re.search(rb"/Type\s*/Catalog", body):
                root = num
                break
        if root is None:
            raise ValueError("no catalog")
        out: list[int] = []

        def walk(num: int) -> None:
            body = self.objects[num]
            if re.search(rb"/Type\s*/Pages", body):
                kids = _dict_value(body, b"Kids") or b"[]"
                for r in _REF.finditer(kids):
                    walk(int(r.group(1)))
            else:
                out.append(num)

        walk(self.ref(_dict_value(self.objects[root], b"Pages")))
        return out

    def page_box(self, num: int) -> tuple[float, float]:
        body = self.objects[num]
        box = _dict_value(body, b"MediaBox")
        while box is None:  # inherited from the parent
            parent = self.ref(_dict_value(body, b"Parent"))
            if parent is None:
                return 595.3, 841.9
            body = self.objects[parent]
            box = _dict_value(body, b"MediaBox")
        r = self.ref(box)
        if r is not None:
            box = self.objects[r]
        x0, y0, x1, y1 = (float(v) for v in re.findall(rb"-?[\d.]+", box)[:4])
        return x1 - x0, y1 - y0

    def contents(self, num: int) -> bytes:
        v = _dict_value(self.objects[num], b"Contents")
        if v is None:
            return b""
        refs = [int(r.group(1)) for r in _REF.finditer(v)]
        if len(refs) == 1 and refs[0] not in self.streams:  # a reference to an array
            refs = [int(r.group(1)) for r in _REF.finditer(self.objects[refs[0]])]
        return b"\n".join(self.streams.get(r, b"") for r in refs)


# ----- content streams ---------------------------------------------------------------------

def _tokens(s: bytes):
    i, n = 0, len(s)
    while i < n:
        c = s[i]
        if c in WS:
            i += 1
        elif c == 0x25:  # % comment
            while i < n and s[i] not in b"\r\n":
                i += 1
        elif c == 0x28:  # (string)
            depth, j = 1, i + 1
            while j < n and depth:
                if s[j] == 0x5C:
                    j += 2
                    continue
                if s[j] == 0x28:
                    depth += 1
                elif s[j] == 0x29:
                    depth -= 1
                j += 1
            yield ("str", j - i - 2)
            i = j
        elif c == 0x3C and s[i + 1:i + 2] == b"<":
            yield ("op", b"<<")
            i += 2
        elif c == 0x3E and s[i + 1:i + 2] == b">":
            yield ("op", b">>")
            i += 2
        elif c == 0x3C:  # <hex>
            j = s.find(b">", i)
            yield ("str", max(0, (j - i - 1) // 2))
            i = j + 1
        elif c in b"[]":
            yield ("op", bytes([c]))
            i += 1
        elif c == 0x2F:  # /name
            j = i + 1
            while j < n and s[j] not in WS and s[j] not in DELIM:
                j += 1
            yield ("name", s[i + 1:j])
            i = j
        else:
            j = i
            while j < n and s[j] not in WS and s[j] not in DELIM:
                j += 1
            word = s[i:j]
            i = j if j > i else i + 1
            try:
                yield ("num", float(word))
            except ValueError:
                yield ("op", word)
                if word == b"ID":  # inline image data up to EI
                    k = s.find(b"EI", i)
                    i = n if k == -1 else k + 2


def _mul(a, b):
    return (a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3],
            a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3],
            a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5])


def text_runs(content: bytes, height: float) -> list[TextRun]:
    I = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
    ctm, stack = I, []
    tm = tlm = I
    size, leading, rise = 0.0, 0.0, 0.0
    ops: list = []
    arr: list | None = None
    out: list[TextRun] = []

    def show(nchars: int) -> None:
        m = _mul((1, 0, 0, 1, 0, rise), _mul(tm, ctm))
        scale = (abs(m[3]) or abs(m[1])) * size if size else 0
        out.append(TextRun(round(m[4], 3), round(height - m[5], 3), round(scale, 3), nchars))

    for kind, v in _tokens(content):
        if kind == "op" and v == b"[":
            arr = []
            continue
        if kind == "op" and v == b"]":
            ops.append(("arr", arr or []))
            arr = None
            continue
        if arr is not None:
            arr.append((kind, v))
            continue
        if kind != "op":
            ops.append((kind, v))
            continue
        nums = [x for k, x in ops if k == "num"]
        if v == b"q":
            stack.append(ctm)
        elif v == b"Q":
            ctm = stack.pop() if stack else I
        elif v == b"cm" and len(nums) >= 6:
            ctm = _mul(tuple(nums[-6:]), ctm)
        elif v == b"BT":
            tm = tlm = I
        elif v == b"Tf" and nums:
            size = nums[-1]
        elif v == b"TL" and nums:
            leading = nums[-1]
        elif v == b"Ts" and nums:
            rise = nums[-1]
        elif v in (b"Td", b"TD") and len(nums) >= 2:
            if v == b"TD":
                leading = -nums[-1]
            tlm = _mul((1, 0, 0, 1, nums[-2], nums[-1]), tlm)
            tm = tlm
        elif v == b"Tm" and len(nums) >= 6:
            tm = tlm = tuple(nums[-6:])
        elif v == b"T*":
            tlm = _mul((1, 0, 0, 1, 0, -leading), tlm)
            tm = tlm
        elif v in (b"Tj", b"'", b'"'):
            if v != b"Tj":
                tlm = _mul((1, 0, 0, 1, 0, -leading), tlm)
                tm = tlm
            strs = [x for k, x in ops if k == "str"]
            show(strs[-1] if strs else 0)
        elif v == b"TJ":
            a = next((x for k, x in reversed(ops) if k == "arr"), [])
            show(sum(x for k, x in a if k == "str"))
        ops = []
    return out


def read(path) -> list[Page]:
    from pathlib import Path
    pdf = Pdf(Path(path).read_bytes())
    pages = []
    for num in pdf.pages():
        w, h = pdf.page_box(num)
        pages.append(Page(w, h, text_runs(pdf.contents(num), h)))
    return pages


def baselines(page: Page, tolerance: float = 4.0) -> list[dict]:
    """Groups the runs of a page into lines (top to bottom). A line's baseline is the one its
    largest text sits on, so raised and lowered characters do not count as lines."""
    runs = sorted((r for r in page.runs if r.chars and r.size > 0), key=lambda r: r.y)
    groups: list[list[TextRun]] = []
    for r in runs:
        if groups and r.y - groups[-1][-1].y <= tolerance:
            groups[-1].append(r)
        else:
            groups.append([r])
    out = []
    for g in groups:
        big = max(r.size for r in g)
        main = [r for r in g if r.size >= big - 0.01]
        out.append({"y": round(sum(r.y for r in main) / len(main), 2), "x": round(min(r.x for r in g), 2), "size": big})
    return out
