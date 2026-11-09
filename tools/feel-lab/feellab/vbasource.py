"""Reads the modules' source code from an .xlsm (xl/vbaProject.bin) without Office: a small compound-file reader and
MS-OVBA's decompression (2.4.1). Lets a probe read what Excel's macro recorder wrote without Trust Center access to the
VBA project object model (which stays off: D-095)."""
from __future__ import annotations

import struct
import zipfile
from pathlib import Path


def _cfb(data: bytes) -> dict[str, bytes]:
    """Every stream of a compound file by its path ("VBA/dir")."""
    ssz = 1 << struct.unpack_from("<H", data, 30)[0]
    mssz = 1 << struct.unpack_from("<H", data, 32)[0]
    n_fat, first_dir, _, mini_cut, first_mini, n_mini, first_difat, n_difat = struct.unpack_from("<IIIIIIII", data, 44)
    sec = lambda i: data[(i + 1) * ssz:(i + 2) * ssz]  # noqa: E731
    difat = list(struct.unpack_from("<109I", data, 76))
    d = first_difat
    for _ in range(n_difat):
        s = sec(d)
        difat += list(struct.unpack_from(f"<{ssz // 4 - 1}I", s))
        d = struct.unpack_from("<I", s, ssz - 4)[0]
    fat: list[int] = []
    for f in difat[:n_fat]:
        fat += list(struct.unpack_from(f"<{ssz // 4}I", sec(f)))

    def chain(start: int, table: list[int]) -> list[int]:
        out, i = [], start
        while i < 0xFFFFFFFA and len(out) < len(table) + 1:
            out.append(i)
            i = table[i]
        return out

    dirs = b"".join(sec(i) for i in chain(first_dir, fat))
    entries = []
    for k in range(len(dirs) // 128):
        e = dirs[k * 128:(k + 1) * 128]
        nlen = struct.unpack_from("<H", e, 64)[0]
        name = e[:max(0, nlen - 2)].decode("utf-16-le")
        typ = e[66]
        left, right, child = struct.unpack_from("<III", e, 68)
        start, size = struct.unpack_from("<IQ", e, 116)
        entries.append({"name": name, "type": typ, "left": left, "right": right, "child": child, "start": start, "size": size & 0xFFFFFFFF})
    root = entries[0]
    mini_stream = b"".join(sec(i) for i in chain(root["start"], fat))
    minifat: list[int] = []
    for i in chain(first_mini, fat) if n_mini else []:
        minifat += list(struct.unpack_from(f"<{ssz // 4}I", sec(i)))

    def read(e: dict) -> bytes:
        if e["size"] < mini_cut:
            return b"".join(mini_stream[i * mssz:(i + 1) * mssz] for i in chain(e["start"], minifat))[:e["size"]]
        return b"".join(sec(i) for i in chain(e["start"], fat))[:e["size"]]

    out: dict[str, bytes] = {}

    def walk(idx: int, prefix: str) -> None:
        if idx >= len(entries) or idx == 0xFFFFFFFF:
            return
        e = entries[idx]
        walk(e["left"], prefix)
        walk(e["right"], prefix)
        path = prefix + e["name"]
        if e["type"] == 2:
            out[path] = read(e)
        elif e["type"] == 1:
            walk(e["child"], path + "/")

    walk(root["child"], "")
    return out


def decompress(b: bytes) -> bytes:
    """MS-OVBA 2.4.1: a compressed container to bytes."""
    if not b or b[0] != 1:
        raise ValueError("not a compressed container")
    out = bytearray()
    i = 1
    while i + 1 < len(b):
        head = b[i] | (b[i + 1] << 8)
        size = (head & 0x0FFF) + 3
        compressed = head & 0x8000
        end = min(len(b), i + size)
        i += 2
        start = len(out)
        if not compressed:
            out += b[i:i + 4096]
            i += 4096
            continue
        while i < end:
            flags = b[i]
            i += 1
            for bit in range(8):
                if i >= end:
                    break
                if not (flags >> bit) & 1:
                    out.append(b[i])
                    i += 1
                    continue
                token = b[i] | (b[i + 1] << 8)
                i += 2
                pos = len(out) - start
                bits = max(4, (pos - 1).bit_length())
                length = (token & (0xFFFF >> bits)) + 3
                offset = (token >> (16 - bits)) + 1
                for _ in range(length):
                    out.append(out[-offset])
    return bytes(out)


def modules(path: str | Path) -> dict[str, str]:
    """Module name -> source (cp1252), Attribute lines kept, from an .xlsm or a vbaProject.bin."""
    p = Path(path)
    data = p.read_bytes()
    if data[:2] == b"PK":
        with zipfile.ZipFile(p) as z:
            data = z.read("xl/vbaProject.bin")
    files = _cfb(data)
    find = lambda n: next((v for k, v in files.items() if k.lower() == n.lower()), None)  # noqa: E731
    d = decompress(find("VBA/dir"))
    i, mods, cur = 0, [], None
    while i + 6 <= len(d):
        rid, size = struct.unpack_from("<HI", d, i)
        if rid == 0x0009:
            size = 6
        body = d[i + 6:i + 6 + size]
        if rid == 0x0019:
            cur = {"name": body.decode("cp1252"), "stream": body.decode("cp1252"), "offset": 0}
            mods.append(cur)
        elif rid == 0x001A and cur is not None:
            cur["stream"] = body.decode("cp1252")
        elif rid == 0x0031 and cur is not None:
            cur["offset"] = struct.unpack_from("<I", body)[0]
        i += 6 + size
    out = {}
    for m in mods:
        raw = find("VBA/" + m["stream"])
        if raw is None:
            continue
        out[m["name"]] = decompress(raw[m["offset"]:]).decode("cp1252")
    return out
