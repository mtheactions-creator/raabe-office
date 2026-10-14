"""Debug aid for check_shapes.py: Word's paths for one preset (wide variant), relative to the
shape's top left, next to ours.

    .venv/Scripts/python.exe dump_shape.py <prst> [wide|tall|adjusted]
"""
import json
import sys

import check_shapes as cs
from probe_word_shapes import outline

name = sys.argv[1]
variant = sys.argv[2] if len(sys.argv) > 2 else "wide"
first = cs.ours([{"prst": name, "x": 0, "y": 0, "w": 80, "h": 50, "adj": {}, "page": 0}])
spec = cs.layout([name], variant, {name: first[0]["adjust"]})
s = spec[0]
path = cs.TMP / "dump.docx"
cs.docx(spec, path)
outline(path, cs.TMP / "dump.json")
pages = json.loads((cs.TMP / "dump.json").read_text(encoding="utf-8"))
r = lambda v: round(v, 2)  # noqa: E731
print("WORD", s["w"], "x", s["h"], s["adj"])
for p in pages[0]["paths"]:
    segs = [[seg[0], *[r(v - (s["x"] if i % 2 == 0 else s["y"])) for i, v in enumerate(seg[1:])]] for seg in p["d"]]
    print(p["paint"], p.get("fill"), segs)
print("OURS")
for p in cs.ours(spec)[0]["paths"]:
    print(p["fill"], p["stroke"], p["d"])
