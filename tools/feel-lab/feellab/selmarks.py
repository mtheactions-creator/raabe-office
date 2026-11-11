"""Selection marks measured from screen crops (Step 11 part 2, D-170).

The pixels a selection adds to a chart (crop of the selected chart against the same chart with a reference selection)
are split into frame lines (long straight runs) and handles (what is left, grouped), and each is measured: box, size,
ring and fill colour. Office draws two kinds of handles: grey rings with a white fill for parts that can be sized
(chart area, plot area, legend, single data label) and small blue dots for the others. Everything is in device
pixels of the crop."""
from __future__ import annotations

from collections import Counter
from pathlib import Path

import numpy as np

#: A straight run at least this long is a frame line, not part of a handle.
LINE_RUN = 16


def load(path: Path) -> np.ndarray:
    from PIL import Image
    return np.asarray(Image.open(path).convert("RGB")).astype(np.int16)


def diff_mask(sel: np.ndarray, base: np.ndarray, tol: int = 24) -> np.ndarray:
    return np.abs(sel - base).max(axis=2) > tol


def _runs(mask: np.ndarray, length: int) -> np.ndarray:
    """Pixels on a horizontal or vertical run of at least `length` set pixels."""
    out = np.zeros_like(mask)
    for m, o in ((mask, out), (mask.T, out.T)):
        for i in range(m.shape[0]):
            row = m[i]
            if row.sum() < length:
                continue
            d = np.diff(np.concatenate(([0], row.astype(np.int8), [0])))
            starts, ends = np.nonzero(d == 1)[0], np.nonzero(d == -1)[0]
            for s, e in zip(starts, ends):
                if e - s >= length:
                    o[i, s:e] = True
    return out


def _dilate(mask: np.ndarray) -> np.ndarray:
    p = np.pad(mask, 1)
    out = np.zeros_like(mask)
    for dy in (0, 1, 2):
        for dx in (0, 1, 2):
            out |= p[dy:dy + mask.shape[0], dx:dx + mask.shape[1]]
    return out


def _components(mask: np.ndarray) -> list[np.ndarray]:
    """8-connected components as arrays of (y, x)."""
    seen = np.zeros_like(mask)
    h, w = mask.shape
    out = []
    for y0, x0 in zip(*np.nonzero(mask)):
        if seen[y0, x0]:
            continue
        stack = [(y0, x0)]
        seen[y0, x0] = True
        pts = []
        while stack:
            y, x = stack.pop()
            pts.append((y, x))
            for dy in (-1, 0, 1):
                for dx in (-1, 0, 1):
                    yy, xx = y + dy, x + dx
                    if 0 <= yy < h and 0 <= xx < w and mask[yy, xx] and not seen[yy, xx]:
                        seen[yy, xx] = True
                        stack.append((yy, xx))
        out.append(np.array(pts))
    return out


def _hex(c) -> str:
    return "#%02X%02X%02X" % tuple(int(v) for v in c)


def _kind(c) -> str:
    r, g, b = (int(v) for v in c)
    if b - r > 40:
        return "blue"
    if max(r, g, b) - min(r, g, b) < 16:
        return "grey"
    return "other"


def measure_arrays(sel: np.ndarray, base: np.ndarray, origin: tuple[int, int] = (0, 0)) -> dict:
    """The marks in `sel` that `base` does not have; positions relative to `origin` (the chart's top left)."""
    h = min(sel.shape[0], base.shape[0])
    w = min(sel.shape[1], base.shape[1])
    sel, base = sel[:h, :w], base[:h, :w]
    m = diff_mask(sel, base)
    if not m.any():
        return {"summary": "no marks", "handles": [], "lines": None}
    line = _runs(m, LINE_RUN)
    rest = m & ~line
    ox, oy = origin
    handles = []
    groups = []
    for comp in _components(_dilate(rest)):
        ys, xs = comp[:, 0], comp[:, 1]
        own = rest[ys, xs]
        if own.sum():
            groups.append((ys[own], xs[own]))
    # A frame line through a ring cuts it in two: halves whose joint box is still a handle's are one handle.
    merged = True
    while merged:
        merged = False
        for i in range(len(groups)):
            for j in range(i + 1, len(groups)):
                a, b = groups[i], groups[j]
                ys = np.concatenate((a[0], b[0]))
                xs = np.concatenate((a[1], b[1]))
                if xs.max() - xs.min() < 20 and ys.max() - ys.min() < 20:
                    gx = max(a[1].min(), b[1].min()) - min(a[1].max(), b[1].max())
                    gy = max(a[0].min(), b[0].min()) - min(a[0].max(), b[0].max())
                    if gx <= 6 and gy <= 6:
                        groups[i] = (ys, xs)
                        del groups[j]
                        merged = True
                        break
            if merged:
                break
    for ys, xs in groups:
        if len(xs) < 6:
            continue
        x0, x1, y0, y1 = int(xs.min()), int(xs.max()), int(ys.min()), int(ys.max())
        bw, bh = x1 - x0 + 1, y1 - y0 + 1
        if bw < 4 or bh < 4 or bw > 30 or bh > 30:
            continue
        cols = sel[ys, xs]
        # The ring: the darkest sixth of the handle's own pixels (its core, not the smoothed edge); the fill: the middle.
        dark = cols[np.argsort(cols.sum(axis=1))[: max(1, len(cols) // 6)]]
        ring = np.median(dark, axis=0)
        cy, cx = (y0 + y1) // 2, (x0 + x1) // 2
        fill = sel[cy, cx]
        # The ring's thickness: dark pixels on the middle row, from the left edge inwards.
        row = sel[cy, x0:x1 + 1]
        thick = 0
        for px in row:
            if np.abs(px - ring).max() < 40:
                thick += 1
            elif thick:
                break
        handles.append({"x": x0 - ox, "y": y0 - oy, "w": bw, "h": bh, "cx": (x0 + x1) / 2 - ox, "cy": (y0 + y1) / 2 - oy,
                        "ring": _hex(ring), "fill": _hex(fill), "ringWidth": thick, "kind": _kind(ring)})
    handles.sort(key=lambda d: (round(d["cy"]), d["cx"]))
    lines = None
    if line.any():
        ys, xs = np.nonzero(line)
        cols = Counter(_hex(c) for c in sel[ys, xs])
        # Thickness: the widest band of line pixels across one horizontal run's column.
        lines = {"box": [int(xs.min()) - ox, int(ys.min()) - oy, int(xs.max() - xs.min()) + 1, int(ys.max() - ys.min()) + 1],
                 "pixels": int(line.sum()), "colour": cols.most_common(1)[0][0], "colours": dict(cols.most_common(3)),
                 "thickness": _thickness(line), "ink": _ink(line, sel)}
    kinds = Counter(hd["kind"] for hd in handles)
    sizes = Counter(f"{hd['w']}x{hd['h']}" for hd in handles)
    summ = f"{len(handles)} handles {dict(kinds)} sizes {dict(sizes.most_common(2))}"
    if lines:
        summ += f"; lines {lines['colour']} t{lines['thickness']} ink {lines['ink']} box {lines['box']}"
    return {"summary": summ, "handles": handles, "lines": lines}


def _thickness(line: np.ndarray) -> int:
    """The most common thickness of the frame lines (vertical extent of horizontal lines)."""
    t = Counter()
    for x in range(0, line.shape[1], 7):
        col = line[:, x]
        d = np.diff(np.concatenate(([0], col.astype(np.int8), [0])))
        for s, e in zip(np.nonzero(d == 1)[0], np.nonzero(d == -1)[0]):
            if e - s <= 4:
                t[int(e - s)] += 1
    return t.most_common(1)[0][0] if t else 0


def _ink(line: np.ndarray, sel: np.ndarray) -> int:
    """How dark a frame line is across its width (sum of 255 - grey over the line's pixels in one cross-section, the
    median of horizontal and vertical lines' sections): a one pixel line of #7B7B7B and a line smeared over two pixels
    of #BDBDBD both give about 130."""
    grey = sel.mean(axis=2)
    spread_all = sel.max(axis=2) - sel.min(axis=2)
    vals = []
    for m, g, spread in ((line, grey, spread_all), (line.T, grey.T, spread_all.T)):
        for x in range(0, m.shape[1], 5):
            col = m[:, x]
            d = np.diff(np.concatenate(([0], col.astype(np.int8), [0])))
            for s, e in zip(np.nonzero(d == 1)[0], np.nonzero(d == -1)[0]):
                # Only grey sections: where a line runs through a handle or over a coloured part it says nothing.
                if e - s <= 4 and (spread[s:e, x] < 30).all():
                    vals.append(float((255 - g[s:e, x]).sum()))
    return int(np.median(vals)) if vals else 0


def measure(sel_path: Path, base_path: Path, chart_box: list[int]) -> dict:
    return measure_arrays(load(sel_path), load(base_path), (chart_box[0], chart_box[1]))


def compare(ours: dict, office: dict, size_tol: int = 2, colour_tol: int = 40, per_series: int = 1, ink_tol: int = 45,
            over_content: bool = False, count_tol: int = 1) -> dict:
    """Ours against Office's marks of the same part: handle count (within `count_tol`: a handle over busy content can
    hide from the measurement), size, kind and colours, the frame's ink (darkness across its width) and thickness.
    Positions are reported (the charts' own layouts differ) but not judged. `per_series`: ours selects that many
    series' worth of the part where Office selects one (data labels). `over_content`: the frame lies on the element's
    own edges (columns), where its colour mixes with the element's: only its presence is judged."""
    def col(h: str) -> np.ndarray:
        return np.array([int(h[i:i + 2], 16) for i in (1, 3, 5)])
    oh, xh = ours.get("handles") or [], office.get("handles") or []
    fails = []
    if abs(len(oh) - len(xh) * per_series) > count_tol:
        fails.append(f"handles {len(oh)} vs {len(xh)}")
    def mode(hs, key):
        c = Counter(hd[key] for hd in hs)
        return c.most_common(1)[0][0] if c else None
    for key in ("w", "h"):
        a, b = mode(oh, key), mode(xh, key)
        if a is not None and b is not None and abs(a - b) > size_tol:
            fails.append(f"handle {key} {a} vs {b}")
    for key in ("kind",):
        a, b = mode(oh, key), mode(xh, key)
        if a != b:
            fails.append(f"handle {key} {a} vs {b}")
    def median(hs, key):
        return _hex(np.median([col(hd[key]) for hd in hs], axis=0)) if hs else None
    for key in ("ring", "fill"):
        a, b = median(oh, key), median(xh, key)
        if a and b and np.abs(col(a) - col(b)).max() > colour_tol:
            fails.append(f"handle {key} {a} vs {b}")
    ol, xl = ours.get("lines"), office.get("lines")
    if bool(ol) != bool(xl):
        fails.append(f"frame {'present' if ol else 'missing'} vs {'present' if xl else 'missing'}")
    elif ol and xl:
        if not over_content and abs(ol.get("ink", 0) - xl.get("ink", 0)) > ink_tol:
            fails.append(f"frame ink {ol.get('ink')} vs {xl.get('ink')}")
        if not over_content and abs(ol["thickness"] - xl["thickness"]) > 1:
            fails.append(f"frame thickness {ol['thickness']} vs {xl['thickness']}")
    off = None
    if oh and xh:
        n = min(len(oh), len(xh))
        off = max(max(abs(a["cx"] - b["cx"]), abs(a["cy"] - b["cy"])) for a, b in zip(oh[:n], xh[:n]))
    return {"passed": not fails, "fails": fails, "handles": [len(oh), len(xh)],
            "size": [f"{mode(oh, 'w')}x{mode(oh, 'h')}", f"{mode(xh, 'w')}x{mode(xh, 'h')}"],
            "ring": [median(oh, "ring"), median(xh, "ring")], "fill": [median(oh, "fill"), median(xh, "fill")],
            "frame": [ol and ol["colour"], xl and xl["colour"]], "ink": [ol and ol.get("ink"), xl and xl.get("ink")], "maxCentreOffset": off}
