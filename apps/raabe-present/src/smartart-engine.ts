// Our SmartArt layout engine (D-077). PowerPoint's layouts are written in its layout definition language, whose
// definitions are Microsoft's; we lay diagrams out with a few arrangement families of our own instead: a grid of
// boxes, a row, a column, a process (boxes with connectors between them), a cycle, a hierarchy, overlapping
// circles, a pyramid and concentric rings. Each of PowerPoint's layouts is given a family and the family's
// measures (shape kinds and proportions) as measured from PowerPoint's own drawing of that layout with three items
// (probe_powerpoint_smartart.py, smartart-catalog.gen.ts); the family carries them over to any number of items and
// sub-items. Sizes are EMU in the diagram frame (w × h).

import type { DrawShape } from "./smartart-parts";

export interface LayoutSpec {
  family: "grid" | "row" | "column" | "process" | "cycle" | "hierarchy" | "venn" | "pyramid" | "target" | "radial" | "matrix";
  /** The item shape, its height to width, the gap between items as a share of the item size. */
  node: string;
  aspect: number;
  gap: number;
  /** Connector between items (process, cycle, radial): its shape and size relative to the item. */
  conn?: string;
  connSize?: number;
  /** Sub-items: as a text list under the item (box) or as their own shapes (shapes). */
  sub?: "box" | "shapes";
  /** Shapes drawn with each item, relative to its box (fractions of its width and height), before or after it. */
  attach?: { prst: string; x: number; y: number; w: number; h: number; before: boolean; look: "bg" | "conn" | "line" }[];
  /** PowerPoint's drawing with three items, in shares of the frame: the items' boxes [x, y, w, h]. */
  items3?: number[][];
  /** …and its other shapes (connectors, rings, backgrounds). */
  extras3?: { prst: string; box: number[]; rot: number; look: "bg" | "conn" | "line" }[];
}

export interface Node { text: string; level: number; index: number; children: Node[] }

/** The items as a tree (level 0 is the top). */
export function tree(items: { text: string; level: number }[]): Node[] {
  const roots: Node[] = [];
  const stack: Node[] = [];
  items.forEach((it, index) => {
    const n: Node = { text: it.text, level: it.level, index, children: [] };
    while (stack.length && stack[stack.length - 1].level >= it.level) stack.pop();
    if (stack.length) stack[stack.length - 1].children.push(n); else roots.push(n);
    stack.push(n);
  });
  return roots;
}

/** A font size (points) that fits every item's text into its shape, as PowerPoint gives all items one size. */
export function fitSize(shapes: DrawShape[], texts: string[], max = 65): number {
  let size = max;
  for (const s of shapes) {
    if (s.item === undefined) continue;
    const t = texts[s.item] || "M";
    const wPt = (s.w / 12700) * 0.86, hPt = (s.h / 12700) * 0.86;
    // Average letter width 0.55 em; lines of 1.0 em at 90 %; words wrap.
    const words = t.split(/\s+/);
    for (let f = Math.min(size, max); f >= 5; f -= 0.5) {
      const longest = Math.max(...words.map((w) => w.length * 0.55 * f));
      const lines = Math.ceil((t.length * 0.55 * f) / Math.max(1, wPt));
      if (longest <= wPt && lines * f * 0.9 * 1.2 <= hPt) { size = Math.min(size, f); break; }
      if (f <= 5.5) size = 5;
    }
  }
  return Math.max(5, Math.round(size * 2) / 2);
}

const box = (prst: string, x: number, y: number, w: number, h: number, item: number | undefined, look: DrawShape["look"] = "node", extra: Partial<DrawShape> = {}): DrawShape =>
  ({ prst, x, y, w, h, item, look, ...extra });

/** Items with their sub-items: the sub-items' texts go into the item's shape as further lines (box) or as shapes. */
function withSubs(n: Node, shapes: DrawShape[], s: DrawShape, spec: LayoutSpec): void {
  const extra = (before: boolean) => { for (const a of spec.attach ?? []) if (a.before === before) shapes.push(box(a.prst, s.x + a.x * s.w, s.y + a.y * s.h, a.w * s.w, a.h * s.h, undefined, a.look)); };
  extra(true);
  try { place(n, shapes, s, spec); } finally { extra(false); }
}

function place(n: Node, shapes: DrawShape[], s: DrawShape, spec: LayoutSpec): void {
  if (!n.children.length) { shapes.push(s); return; }
  if (spec.sub === "shapes") {
    // The item on top, its sub-items in a pale box below it.
    const head = { ...s, h: s.h * 0.4 };
    shapes.push({ ...s, y: s.y + s.h * 0.35, h: s.h * 0.65, item: undefined, look: "bg" });
    shapes.push(head);
    const each = (s.h * 0.6) / n.children.length;
    n.children.forEach((c, i) => shapes.push(box("rect", s.x + s.w * 0.06, s.y + s.h * 0.4 + each * i, s.w * 0.88, each, c.index, "text", { algn: "l" })));
    return;
  }
  shapes.push(s);
  // The item's text moves up; its sub-items follow below it inside its shape, as bullets in the item's colour.
  s.anchor = "t";
  n.children.forEach((c, i) => shapes.push(box("rect", s.x + s.w * 0.08, s.y + s.h * (0.5 + (0.45 * i) / n.children.length), s.w * 0.84, (s.h * 0.45) / n.children.length, c.index, "inner", { algn: "l" })));
}

/** Lays the items out in a w × h frame. */
export function layout(spec: LayoutSpec, items: { text: string; level: number }[], w: number, h: number): DrawShape[] {
  const roots = tree(items);
  const n = Math.max(1, roots.length);
  const out: DrawShape[] = [];
  const T = spec.items3;
  // Three items: PowerPoint's own arrangement (measured), with its connectors and other shapes.
  if (T && T.length === 3 && n === 3) {
    for (const e of spec.extras3 ?? []) if (e.look === "bg") out.push(box(e.prst, e.box[0] * w, e.box[1] * h, e.box[2] * w, e.box[3] * h, undefined, e.look, { rot: e.rot || undefined }));
    roots.forEach((it, i) => withSubs(it, out, box(spec.node, T[i][0] * w, T[i][1] * h, T[i][2] * w, T[i][3] * h, it.index), spec));
    for (const e of spec.extras3 ?? []) if (e.look !== "bg") out.push(box(e.prst, e.box[0] * w, e.box[1] * h, e.box[2] * w, e.box[3] * h, undefined, e.look, { rot: e.rot || undefined }));
    return out;
  }
  // Other numbers: the family, calibrated by the measured three.
  if (T && T.length === 3) {
    const B = T.map(([x, y, bw, bh]) => ({ x: x * w, y: y * h, w: bw * w, h: bh * h }));
    const cxs = B.map((b) => b.x + b.w / 2), cys = B.map((b) => b.y + b.h / 2);
    const conn = (spec.extras3 ?? []).find((e) => e.look === "conn");
    if (spec.family === "row" || spec.family === "process" || (spec.family === "hierarchy" && !roots.some((r) => r.children.length))) {
      const left = B[0].x, right = B[2].x + B[2].w, gap = B[0].w ? (B[1].x - (B[0].x + B[0].w)) / B[0].w : 0.1;
      const bw = (right - left) / (n + (n - 1) * gap), k = B[0].w ? bw / B[0].w : 1, bh = Math.min(h, B[0].h * k);
      // The item with its attached shapes scales about their common middle (PowerPoint keeps the group centred).
      const att = spec.attach ?? [];
      const gTop = Math.min(B[0].y, ...att.map((a) => B[0].y + a.y * B[0].h)), gBot = Math.max(B[0].y + B[0].h, ...att.map((a) => B[0].y + (a.y + a.h) * B[0].h));
      const gMid = (gTop + gBot) / 2;
      const y = gMid - (gMid - B[0].y) * k;
      const cy = y + bh / 2;
      roots.forEach((it, i) => {
        const x = left + i * bw * (1 + gap);
        withSubs(it, out, box(spec.node, x, y, bw, bh, it.index), spec);
        if (conn && i < n - 1) {
          const k = bw / B[0].w;
          const cw = conn.box[2] * w * k, ch = conn.box[3] * h * k;
          out.push(box(conn.prst, x + bw + (bw * gap - cw) / 2, cy - ch / 2, cw, ch, undefined, "conn", { rot: conn.rot || undefined }));
        }
      });
      return out;
    }
    if (spec.family === "column") {
      const top = B[0].y, bottom = B[2].y + B[2].h, gap = B[0].h ? (B[1].y - (B[0].y + B[0].h)) / B[0].h : 0.1;
      const bh = (bottom - top) / (n + (n - 1) * gap);
      roots.forEach((it, i) => withSubs(it, out, box(spec.node, B[0].x, top + i * bh * (1 + gap), B[0].w, bh, it.index), spec));
      return out;
    }
    if (spec.family === "cycle" || spec.family === "venn") {
      const mx = (cxs[0] + cxs[1] + cxs[2]) / 3, my = (cys[0] + cys[1] + cys[2]) / 3;
      const R = Math.hypot(cxs[0] - mx, cys[0] - my);
      const a0 = Math.atan2(cys[0] - my, cxs[0] - mx);
      const dir = ((Math.atan2(cys[1] - my, cxs[1] - mx) - a0 + 2 * Math.PI) % (2 * Math.PI)) < Math.PI ? 1 : -1;
      const k = Math.min(1.25, Math.sin(Math.PI / Math.max(2, n)) / Math.sin(Math.PI / 3));
      const sw = B[0].w * k, sh = B[0].h * k;
      roots.forEach((it, i) => {
        const a = a0 + (dir * 2 * Math.PI * i) / n;
        withSubs(it, out, box(spec.node, mx + R * Math.cos(a) - sw / 2, my + R * Math.sin(a) - sh / 2, sw, sh, it.index, spec.family === "venn" ? "bg" : "node"), spec);
        if (conn && spec.family === "cycle") {
          const b = a + (dir * Math.PI) / n;
          const cw = conn.box[2] * w * k, ch = conn.box[3] * h * k;
          out.push(box(conn.prst, mx + R * Math.cos(b) - cw / 2, my + R * Math.sin(b) - ch / 2, cw, ch, undefined, "conn", { rot: ((b * 180) / Math.PI + 90 * dir + 360) % 360 }));
        }
      });
      return out;
    }
    // A grid fills the frame (PowerPoint: as many columns as give the largest items).
    if (spec.family === "grid") return family({ ...spec, items3: undefined }, roots, w, h);
    // Others (pyramid, matrix, hierarchy with sub-items): the family in the box the three items took.
    const x0 = Math.min(...B.map((b) => b.x)), y0 = Math.min(...B.map((b) => b.y));
    const x1 = Math.max(...B.map((b) => b.x + b.w)), y1 = Math.max(...B.map((b) => b.y + b.h));
    const fw = Math.max(x1 - x0, w * 0.3), fh = Math.max(y1 - y0, h * 0.3);
    const ox = Math.max(0, Math.min(x0, w - fw)), oy = Math.max(0, Math.min(y0, h - fh));
    return family({ ...spec, items3: undefined }, roots, fw, fh).map((s) => ({ ...s, x: s.x + ox, y: s.y + oy }));
  }
  return family(spec, roots, w, h);
}

function family(spec: LayoutSpec, roots: Node[], w: number, h: number): DrawShape[] {
  const n = Math.max(1, roots.length);
  const out: DrawShape[] = [];
  switch (spec.family) {
    case "grid": {
      // The columns that give the largest items, rows centred, the last row centred.
      let best = { c: 1, bw: 0 };
      for (let c = 1; c <= n; c++) {
        const r = Math.ceil(n / c);
        // The same gap across and down (PowerPoint: 28 pt both ways in Einfache Blockliste).
        const bw = Math.min(w / (c + (c - 1) * spec.gap), h / (r * spec.aspect + (r - 1) * spec.gap));
        if (bw > best.bw) best = { c, bw };
      }
      const { c, bw } = best;
      const bh = bw * spec.aspect, g = bw * spec.gap;
      const rows = Math.ceil(n / c);
      const top = (h - (rows * bh + (rows - 1) * g)) / 2;
      roots.forEach((it, i) => {
        const r = Math.floor(i / c), inRow = Math.min(c, n - r * c);
        const left = (w - (inRow * bw + (inRow - 1) * g)) / 2;
        withSubs(it, out, box(spec.node, left + (i % c) * (bw + g), top + r * (bh + g), bw, bh, it.index), spec);
      });
      break;
    }
    case "column": {
      const bh = h / (n + (n - 1) * spec.gap), g = bh * spec.gap;
      roots.forEach((it, i) => withSubs(it, out, box(spec.node, 0, i * (bh + g), w, bh, it.index), spec));
      break;
    }
    case "row": {
      const bw = Math.min(w / (n + (n - 1) * spec.gap), h / spec.aspect), bh = Math.min(h, bw * spec.aspect), g = bw * spec.gap;
      const left = (w - (n * bw + (n - 1) * g)) / 2, top = (h - bh) / 2;
      roots.forEach((it, i) => withSubs(it, out, box(spec.node, left + i * (bw + g), top, bw, bh, it.index), spec));
      break;
    }
    case "process": {
      // Items in a row with a connector in each gap (its width the gap's 60 %, its height a share of the item).
      const gap = spec.gap;
      const bw = Math.min(w / (n + (n - 1) * gap), h / spec.aspect), bh = bw * spec.aspect, g = bw * gap;
      const left = (w - (n * bw + (n - 1) * g)) / 2, top = (h - bh) / 2;
      roots.forEach((it, i) => {
        withSubs(it, out, box(spec.node, left + i * (bw + g), top, bw, bh, it.index), spec);
        if (i < n - 1 && spec.conn) {
          const cw = g * 0.6, ch = Math.min(bh, cw * (spec.connSize ?? 1.6));
          out.push(box(spec.conn, left + i * (bw + g) + bw + (g - cw) / 2, top + (bh - ch) / 2, cw, ch, undefined, "conn"));
        }
      });
      break;
    }
    case "cycle": case "radial": {
      const radial = spec.family === "radial";
      const cx = w / 2, cy = h / 2;
      const R = Math.min(w, h) * (radial ? 0.36 : 0.38);
      const k = radial ? Math.max(1, n - 1) : n;
      const size = Math.min(Math.min(w, h) * 0.3, ((2 * Math.PI * R) / k) * 0.62);
      const nodes = radial ? roots.slice(1) : roots;
      if (radial && roots[0]) withSubs(roots[0], out, box(spec.node, cx - size * 0.65, cy - size * 0.65, size * 1.3, size * 1.3, roots[0].index), spec);
      nodes.forEach((it, i) => {
        const a = -Math.PI / 2 + (2 * Math.PI * i) / k;
        const x = cx + R * Math.cos(a), y = cy + R * Math.sin(a);
        if (spec.conn) {
          if (radial) {
            const mx = cx + R * 0.55 * Math.cos(a), my = cy + R * 0.55 * Math.sin(a);
            out.push(box("line", Math.min(cx, mx), Math.min(cy, my), Math.abs(mx - cx) || 1, Math.abs(my - cy) || 1, undefined, "line", { flipH: mx < cx, flipV: my < cy }));
          } else {
            const b = a + Math.PI / k;
            const ax = cx + R * Math.cos(b), ay = cy + R * Math.sin(b), cw = size * 0.35;
            out.push(box(spec.conn, ax - cw / 2, ay - cw * 0.4, cw, cw * 0.8, undefined, "conn", { rot: ((b * 180) / Math.PI + 90 + 360) % 360 }));
          }
        }
        withSubs(it, out, box(spec.node, x - size / 2, y - (size * spec.aspect) / 2, size, size * spec.aspect, it.index), spec);
      });
      break;
    }
    case "hierarchy": {
      // A tree: each level a row, each node centred over its children, connectors from parent to child.
      const levels: Node[][] = [];
      const leaves = (nd: Node): number => (nd.children.length ? nd.children.reduce((s, c) => s + leaves(c), 0) : 1);
      const walk = (nd: Node, d: number) => { (levels[d] ??= []).push(nd); nd.children.forEach((c) => walk(c, d + 1)); };
      roots.forEach((r) => walk(r, 0));
      const total = roots.reduce((s, r) => s + leaves(r), 0);
      const depth = levels.length;
      const slotW = w / total;
      const bw = slotW * (1 / (1 + spec.gap)), bh = Math.min(bw * spec.aspect, (h / depth) * 0.62);
      const rowH = depth > 1 ? (h - bh) / (depth - 1) : 0;
      const place = (nd: Node, d: number, x0: number): number => {
        const span = leaves(nd) * slotW;
        const x = x0 + span / 2 - bw / 2, y = d * rowH + (depth === 1 ? (h - bh) / 2 : 0);
        let cx0 = x0;
        for (const c of nd.children) {
          const cs = leaves(c) * slotW;
          const ccx = cx0 + cs / 2, py = y + bh, cy = (d + 1) * rowH;
          out.push(box("line", Math.min(x + bw / 2, ccx), py, Math.abs(ccx - (x + bw / 2)) || 1, cy - py, undefined, "line", { flipH: ccx < x + bw / 2 }));
          place(c, d + 1, cx0);
          cx0 += cs;
        }
        out.push(box(spec.node, x, y, bw, bh, nd.index));
        return span;
      };
      let x0 = 0;
      for (const r of roots) { place(r, 0, x0); x0 += leaves(r) * slotW; }
      break;
    }
    case "venn": {
      const R = Math.min(w, h) * (n === 1 ? 0.45 : n === 2 ? 0.34 : 0.3);
      const d = n === 1 ? 0 : R * (n === 2 ? 0.75 : 0.62);
      roots.forEach((it, i) => {
        const a = -Math.PI / 2 + (2 * Math.PI * i) / n;
        const x = w / 2 + d * Math.cos(a), y = h / 2 + d * Math.sin(a);
        out.push(box("ellipse", x - R, y - R, 2 * R, 2 * R, it.index, "node", { anchor: Math.sin(a) < -0.3 ? "t" : Math.sin(a) > 0.3 ? "b" : "ctr" }));
      });
      for (const s of out) s.look = "bg";
      break;
    }
    case "pyramid": {
      // Slices of a triangle, top to bottom; the top one a triangle.
      const side = Math.min(w, h / 0.866);
      const left = (w - side) / 2, ph = side * 0.866, top = (h - ph) / 2, sh = ph / n;
      roots.forEach((it, i) => {
        const y = top + i * sh;
        const wTop = (side * i) / n, wBot = (side * (i + 1)) / n;
        if (i === 0) out.push(box("triangle", w / 2 - wBot / 2, y, wBot, sh * 0.96, it.index));
        else out.push(box("trapezoid", w / 2 - wBot / 2, y, wBot, sh * 0.96, it.index, "node", { flipV: true, adj: { adj: Math.round(((wBot - wTop) / 2 / sh) * 100000) } }));
      });
      void left;
      break;
    }
    case "target": {
      const R = Math.min(w * 0.5, h) * 0.98;
      roots.forEach((it, i) => {
        const r = (R * (n - i)) / n;
        out.push(box("ellipse", R * 0.02 + (R - r), h / 2 - r / 2 * 1, r, r, undefined, i % 2 ? "conn" : "node"));
        out.push(box("rect", R + R * 0.1, (h * i) / n, w - R * 1.1, h / n * 0.9, it.index, "text", { algn: "l" }));
      });
      break;
    }
    case "matrix": {
      const s = Math.min(w, h);
      const left = (w - s) / 2, top = (h - s) / 2, half = (s - s * 0.03) / 2;
      roots.slice(0, 4).forEach((it, i) => out.push(box(spec.node, left + (i % 2) * (half + s * 0.03), top + Math.floor(i / 2) * (half + s * 0.03), half, half, it.index)));
      break;
    }
  }
  return out;
}
