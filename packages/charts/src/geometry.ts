// DrawingML preset geometries for shapes (Einfügen > Formen, D-066): the outline of a preset in a
// box of w x h as an SVG path, following the shapes' definitions in ECMA-376 (presetShapeDefinitions)
// with their default adjustments. Lines and arrows of connectors are drawn by the caller.

export interface ShapePreset { prst: string; label: string; group: string }

/** The presets offered in the Formen gallery, grouped as Excel groups them. */
export const SHAPE_PRESETS: ShapePreset[] = [
  { prst: "line", label: "Linie", group: "Linien" },
  { prst: "straightConnector1", label: "Pfeil: gerade", group: "Linien" },
  { prst: "rect", label: "Rechteck", group: "Rechtecke" },
  { prst: "roundRect", label: "Rechteck: abgerundete Ecken", group: "Rechtecke" },
  { prst: "snip1Rect", label: "Rechteck: eine Ecke abgeschnitten", group: "Rechtecke" },
  { prst: "ellipse", label: "Ellipse", group: "Standardformen" },
  { prst: "triangle", label: "Gleichschenkliges Dreieck", group: "Standardformen" },
  { prst: "rtTriangle", label: "Rechtwinkliges Dreieck", group: "Standardformen" },
  { prst: "parallelogram", label: "Parallelogramm", group: "Standardformen" },
  { prst: "trapezoid", label: "Trapez", group: "Standardformen" },
  { prst: "diamond", label: "Raute", group: "Standardformen" },
  { prst: "pentagon", label: "Pfeil: Fünfeck", group: "Blockpfeile" },
  { prst: "hexagon", label: "Sechseck", group: "Standardformen" },
  { prst: "octagon", label: "Achteck", group: "Standardformen" },
  { prst: "plus", label: "Kreuz", group: "Standardformen" },
  { prst: "can", label: "Zylinder", group: "Standardformen" },
  { prst: "cube", label: "Würfel", group: "Standardformen" },
  { prst: "heart", label: "Herz", group: "Standardformen" },
  { prst: "smileyFace", label: "Smiley", group: "Standardformen" },
  { prst: "rightArrow", label: "Pfeil: nach rechts", group: "Blockpfeile" },
  { prst: "leftArrow", label: "Pfeil: nach links", group: "Blockpfeile" },
  { prst: "upArrow", label: "Pfeil: nach oben", group: "Blockpfeile" },
  { prst: "downArrow", label: "Pfeil: nach unten", group: "Blockpfeile" },
  { prst: "leftRightArrow", label: "Pfeil: nach links und rechts", group: "Blockpfeile" },
  { prst: "chevron", label: "Pfeil: Chevron", group: "Blockpfeile" },
  { prst: "mathPlus", label: "Plus", group: "Formelformen" },
  { prst: "mathMinus", label: "Minus", group: "Formelformen" },
  { prst: "mathMultiply", label: "Multiplizieren", group: "Formelformen" },
  { prst: "mathEqual", label: "Gleich", group: "Formelformen" },
  { prst: "flowChartProcess", label: "Flussdiagramm: Prozess", group: "Flussdiagramm" },
  { prst: "flowChartDecision", label: "Flussdiagramm: Verzweigung", group: "Flussdiagramm" },
  { prst: "flowChartTerminator", label: "Flussdiagramm: Grenzstelle", group: "Flussdiagramm" },
  { prst: "star5", label: "Stern: 5 Zacken", group: "Sterne und Banner" },
  { prst: "star6", label: "Stern: 6 Zacken", group: "Sterne und Banner" },
  { prst: "wedgeRectCallout", label: "Sprechblase: rechteckig", group: "Legenden" },
  { prst: "wedgeEllipseCallout", label: "Sprechblase: oval", group: "Legenden" },
];

const f = (v: number) => String(Math.round(v * 100) / 100);
const poly = (pts: [number, number][]) => pts.map(([x, y], i) => `${i ? "L" : "M"}${f(x)} ${f(y)}`).join("") + "Z";

function star(n: number, w: number, h: number, inner: number): string {
  const pts: [number, number][] = [];
  for (let i = 0; i < 2 * n; i++) {
    const a = -Math.PI / 2 + (i * Math.PI) / n;
    const r = i % 2 ? inner : 1;
    pts.push([w / 2 + (w / 2) * r * Math.cos(a), h / 2 + (h / 2) * r * Math.sin(a)]);
  }
  return poly(pts);
}

/** The preset's outline in a w x h box; lines and connectors: the line from corner to corner. */
export function presetPath(prst: string, w: number, h: number): string {
  const m = Math.min(w, h);
  switch (prst) {
    case "line": case "straightConnector1": case "bentConnector3": case "curvedConnector3":
      return `M0 0L${f(w)} ${f(h)}`;
    case "ellipse": case "flowChartConnector":
      return `M0 ${f(h / 2)}A${f(w / 2)} ${f(h / 2)} 0 1 1 ${f(w)} ${f(h / 2)}A${f(w / 2)} ${f(h / 2)} 0 1 1 0 ${f(h / 2)}Z`;
    case "roundRect": {
      const r = m * 0.16667;
      return `M${f(r)} 0H${f(w - r)}A${f(r)} ${f(r)} 0 0 1 ${f(w)} ${f(r)}V${f(h - r)}A${f(r)} ${f(r)} 0 0 1 ${f(w - r)} ${f(h)}H${f(r)}A${f(r)} ${f(r)} 0 0 1 0 ${f(h - r)}V${f(r)}A${f(r)} ${f(r)} 0 0 1 ${f(r)} 0Z`;
    }
    case "flowChartTerminator": {
      const r = Math.min(h / 2, w / 2);
      return `M${f(r)} 0H${f(w - r)}A${f(r)} ${f(h / 2)} 0 0 1 ${f(w - r)} ${f(h)}H${f(r)}A${f(r)} ${f(h / 2)} 0 0 1 ${f(r)} 0Z`;
    }
    case "snip1Rect": { const s = m * 0.16667; return poly([[0, 0], [w - s, 0], [w, s], [w, h], [0, h]]); }
    case "triangle": return poly([[w / 2, 0], [w, h], [0, h]]);
    case "rtTriangle": return poly([[0, 0], [w, h], [0, h]]);
    case "parallelogram": { const d = m * 0.25; return poly([[d, 0], [w, 0], [w - d, h], [0, h]]); }
    case "trapezoid": { const d = m * 0.25; return poly([[d, 0], [w - d, 0], [w, h], [0, h]]); }
    case "diamond": case "flowChartDecision": return poly([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]);
    case "pentagon": { const d = m * 0.5; return poly([[0, 0], [w - d, 0], [w, h / 2], [w - d, h], [0, h]]); }
    case "chevron": { const d = m * 0.5; return poly([[0, 0], [w - d, 0], [w, h / 2], [w - d, h], [0, h], [d, h / 2]]); }
    case "hexagon": { const d = m * 0.25; return poly([[d, 0], [w - d, 0], [w, h / 2], [w - d, h], [d, h], [0, h / 2]]); }
    case "octagon": { const d = m * 0.29289; return poly([[d, 0], [w - d, 0], [w, d], [w, h - d], [w - d, h], [d, h], [0, h - d], [0, d]]); }
    case "plus": { const d = m * 0.25; return poly([[d, 0], [w - d, 0], [w - d, d], [w, d], [w, h - d], [w - d, h - d], [w - d, h], [d, h], [d, h - d], [0, h - d], [0, d], [d, d]]); }
    case "mathPlus": {
      const t = m * 0.2352, cx = w / 2, cy = h / 2, a = m * 0.38;
      return poly([[cx - t / 2, cy - a], [cx + t / 2, cy - a], [cx + t / 2, cy - t / 2], [cx + a, cy - t / 2], [cx + a, cy + t / 2], [cx + t / 2, cy + t / 2],
        [cx + t / 2, cy + a], [cx - t / 2, cy + a], [cx - t / 2, cy + t / 2], [cx - a, cy + t / 2], [cx - a, cy - t / 2], [cx - t / 2, cy - t / 2]]);
    }
    case "mathMinus": { const t = h * 0.2352; return poly([[w * 0.12, h / 2 - t / 2], [w * 0.88, h / 2 - t / 2], [w * 0.88, h / 2 + t / 2], [w * 0.12, h / 2 + t / 2]]); }
    case "mathEqual": {
      const t = h * 0.2352, g = h * 0.1176;
      return poly([[w * 0.12, h / 2 - g / 2 - t], [w * 0.88, h / 2 - g / 2 - t], [w * 0.88, h / 2 - g / 2], [w * 0.12, h / 2 - g / 2]])
        + poly([[w * 0.12, h / 2 + g / 2], [w * 0.88, h / 2 + g / 2], [w * 0.88, h / 2 + g / 2 + t], [w * 0.12, h / 2 + g / 2 + t]]);
    }
    case "mathMultiply": {
      const t = m * 0.1661, cx = w / 2, cy = h / 2, a = m * 0.34;
      const arm = (ang: number) => {
        const c = Math.cos(ang), s = Math.sin(ang);
        const p = (x: number, y: number): [number, number] => [cx + x * c - y * s, cy + x * s + y * c];
        return poly([p(-a, -t / 2), p(a, -t / 2), p(a, t / 2), p(-a, t / 2)]);
      };
      return arm(Math.PI / 4) + arm(-Math.PI / 4);
    }
    case "rightArrow": { const hd = m * 0.5, t = h * 0.25; return poly([[0, t], [w - hd, t], [w - hd, 0], [w, h / 2], [w - hd, h], [w - hd, h - t], [0, h - t]]); }
    case "leftArrow": { const hd = m * 0.5, t = h * 0.25; return poly([[w, t], [hd, t], [hd, 0], [0, h / 2], [hd, h], [hd, h - t], [w, h - t]]); }
    case "upArrow": { const hd = m * 0.5, t = w * 0.25; return poly([[t, h], [t, hd], [0, hd], [w / 2, 0], [w, hd], [w - t, hd], [w - t, h]]); }
    case "downArrow": { const hd = m * 0.5, t = w * 0.25; return poly([[t, 0], [t, h - hd], [0, h - hd], [w / 2, h], [w, h - hd], [w - t, h - hd], [w - t, 0]]); }
    case "leftRightArrow": { const hd = m * 0.5, t = h * 0.25; return poly([[0, h / 2], [hd, 0], [hd, t], [w - hd, t], [w - hd, 0], [w, h / 2], [w - hd, h], [w - hd, h - t], [hd, h - t], [hd, h]]); }
    case "star5": return star(5, w, h, 0.382);
    case "star6": return star(6, w, h, 0.577);
    case "can": {
      const ry = Math.min(h * 0.25, m * 0.25) / 2;
      return `M0 ${f(ry)}A${f(w / 2)} ${f(ry)} 0 0 1 ${f(w)} ${f(ry)}V${f(h - ry)}A${f(w / 2)} ${f(ry)} 0 0 1 0 ${f(h - ry)}Z M0 ${f(ry)}A${f(w / 2)} ${f(ry)} 0 0 0 ${f(w)} ${f(ry)}`;
    }
    case "cube": {
      const d = m * 0.25;
      return poly([[0, d], [d, 0], [w, 0], [w, h - d], [w - d, h], [0, h]]) + `M0 ${f(d)}H${f(w - d)}L${f(w)} 0M${f(w - d)} ${f(d)}V${f(h)}`;
    }
    case "heart":
      return `M${f(w / 2)} ${f(h * 0.25)}C${f(w * 0.75)} ${f(-h * 0.1)} ${f(w * 1.25)} ${f(h * 0.33)} ${f(w / 2)} ${f(h)}C${f(-w * 0.25)} ${f(h * 0.33)} ${f(w * 0.25)} ${f(-h * 0.1)} ${f(w / 2)} ${f(h * 0.25)}Z`;
    case "smileyFace": {
      const e = (x: number, y: number, rx: number, ry: number) => `M${f(x - rx)} ${f(y)}A${f(rx)} ${f(ry)} 0 1 1 ${f(x + rx)} ${f(y)}A${f(rx)} ${f(ry)} 0 1 1 ${f(x - rx)} ${f(y)}Z`;
      return e(w / 2, h / 2, w / 2, h / 2) + e(w * 0.34, h * 0.38, w * 0.05, h * 0.05) + e(w * 0.66, h * 0.38, w * 0.05, h * 0.05)
        + `M${f(w * 0.25)} ${f(h * 0.65)}Q${f(w / 2)} ${f(h * 0.85)} ${f(w * 0.75)} ${f(h * 0.65)}`;
    }
    case "wedgeRectCallout":
      return poly([[0, 0], [w, 0], [w, h], [w * 0.4, h], [w * 0.2, h * 1.3], [w * 0.25, h], [0, h]]);
    case "wedgeEllipseCallout":
      return `M${f(w * 0.3)} ${f(h * 0.95)}L${f(w * 0.2)} ${f(h * 1.3)}L${f(w * 0.45)} ${f(h * 0.99)}A${f(w / 2)} ${f(h / 2)} 0 1 0 ${f(w * 0.3)} ${f(h * 0.95)}Z`;
    default:
      return poly([[0, 0], [w, 0], [w, h], [0, h]]);
  }
}

/** Presets drawn as an open line (no fill). */
export const isLinePreset = (prst: string) => prst === "line" || prst.endsWith("Connector1") || prst.endsWith("Connector3");
