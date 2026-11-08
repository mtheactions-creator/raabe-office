// Freihand zu Form (D-072, shared with RaabeCalc D-093): while it is on, a pen stroke that looks like a line, a
// triangle, a rectangle (or a rhombus) or an ellipse becomes that shape, outlined in the pen's colour and width (our
// own recognition: the stroke simplified to its corners; Office's is not documented). Points in EMU. Other strokes stay ink.

export interface Recognized { prst: string; x: number; y: number; w: number; h: number; flipV?: boolean; flipH?: boolean }
type P = { x: number; y: number };

/** Ramer-Douglas-Peucker: the stroke's corners within `eps`. */
function simplify(pts: P[], eps: number): P[] {
  if (pts.length < 3) return pts.slice();
  const a = pts[0], b = pts[pts.length - 1];
  const len = Math.hypot(b.x - a.x, b.y - a.y) || 1;
  let worst = 0, at = 0;
  for (let i = 1; i < pts.length - 1; i++) {
    const p = pts[i];
    const d = Math.abs((b.x - a.x) * (a.y - p.y) - (a.x - p.x) * (b.y - a.y)) / len;
    if (d > worst) { worst = d; at = i; }
  }
  if (worst <= eps) return [a, b];
  return [...simplify(pts.slice(0, at + 1), eps).slice(0, -1), ...simplify(pts.slice(at), eps)];
}

export function recognize(pts: P[]): Recognized | null {
  if (pts.length < 4) return null;
  const xs = pts.map((p) => p.x), ys = pts.map((p) => p.y);
  const x0 = Math.min(...xs), y0 = Math.min(...ys), w = Math.max(...xs) - x0, h = Math.max(...ys) - y0;
  const diag = Math.hypot(w, h);
  if (diag < 50000) return null;
  const first = pts[0], last = pts[pts.length - 1];
  const closed = Math.hypot(first.x - last.x, first.y - last.y) < diag * 0.2;
  if (!closed) {
    // A line: no point far from the straight way between the ends.
    const s = simplify(pts, diag * 0.06);
    if (s.length !== 2) return null;
    return { prst: "line", x: x0, y: y0, w: Math.max(1, w), h: Math.max(1, h), flipV: (last.y - first.y) * (last.x - first.x) < 0 };
  }
  // Closed: corners of the loop, simplified in two halves split at the point farthest from the start.
  let far = 0;
  pts.forEach((p, i) => { if (Math.hypot(p.x - first.x, p.y - first.y) > Math.hypot(pts[far].x - first.x, pts[far].y - first.y)) far = i; });
  const eps = diag * 0.08;
  const loop = [...simplify(pts.slice(0, far + 1), eps).slice(0, -1), ...simplify([...pts.slice(far), first], eps)];
  // The start sits on a side, not on a corner, when the loop goes straight through it.
  const corner = (a: P, b: P, c: P) => { const v1 = { x: a.x - b.x, y: a.y - b.y }, v2 = { x: c.x - b.x, y: c.y - b.y }; const cos = (v1.x * v2.x + v1.y * v2.y) / ((Math.hypot(v1.x, v1.y) * Math.hypot(v2.x, v2.y)) || 1); return cos > -0.94; };
  const ring = loop.slice(0, -1);
  if (ring.length > 3 && !corner(ring[ring.length - 1], ring[0], ring[1])) ring.shift();
  const corners = ring.length;
  if (corners === 3) return { prst: "triangle", x: x0, y: y0, w, h, flipV: pts.reduce((s, p) => s + p.y, 0) / pts.length < y0 + h / 2 };
  if (corners === 4) {
    // A rhombus has its corners near the middles of the box's sides.
    const nearMid = ring.filter((c) => Math.abs(c.x - (x0 + w / 2)) < w * 0.2 || Math.abs(c.y - (y0 + h / 2)) < h * 0.2).length;
    return { prst: nearMid >= 3 ? "diamond" : "rect", x: x0, y: y0, w, h };
  }
  // Round: the points keep about the same distance from the middle.
  const cx = x0 + w / 2, cy = y0 + h / 2;
  const r = pts.map((p) => Math.hypot((p.x - cx) / (w / 2 || 1), (p.y - cy) / (h / 2 || 1)));
  const mean = r.reduce((s, v) => s + v, 0) / r.length;
  const spread = Math.sqrt(r.reduce((s, v) => s + (v - mean) ** 2, 0) / r.length);
  return spread < 0.18 ? { prst: "ellipse", x: x0, y: y0, w, h } : null;
}
