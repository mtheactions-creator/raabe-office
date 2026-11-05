// Custom geometry (a:custGeom: Freihandform, Kurve, Skizze and edited points; D-066) of a drawing
// shape, read into the engine's definition syntax (text-engine geometry.ts); written with
// custGeomXml (../shapes.ts), as RaabeWrite writes it.

const unesc = (s: string) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&amp;/g, "&");
const attr = (tag: string, name: string) => new RegExp(`\\b${name}="([^"]*)"`).exec(tag)?.[1];

/** The custom geometry in a shape's spPr as a definition, or undefined (a preset shape). */
export function custGeomFromXml(spPr: string): string | undefined {
  const cg = /<a:custGeom>([\s\S]*?)<\/a:custGeom>/.exec(spPr)?.[1];
  if (cg === undefined) return undefined;
  const section = (tag: string) => new RegExp(`<a:${tag}>([\\s\\S]*?)</a:${tag}>`).exec(cg)?.[1] ?? "";
  const lines: string[] = [];
  for (const m of section("avLst").matchAll(/<a:gd\b([^>]*)\/>/g)) {
    const v = /^val\s+(-?\S+)/.exec(unesc(attr(m[1], "fmla") ?? ""));
    if (v) lines.push(`av ${attr(m[1], "name")} ${v[1]}`);
  }
  for (const m of section("gdLst").matchAll(/<a:gd\b([^>]*)\/>/g)) lines.push(`gd ${attr(m[1], "name")} ${unesc(attr(m[1], "fmla") ?? "")}`);
  const rect = /<a:rect\b([^>]*)\/>/.exec(cg);
  if (rect) lines.push(`rect ${["l", "t", "r", "b"].map((a) => attr(rect[1], a) ?? "0").join(" ")}`);
  for (const p of section("pathLst").matchAll(/<a:path\b([^>]*?)(?:\/>|>([\s\S]*?)<\/a:path>)/g)) {
    const attrs: string[] = [];
    for (const k of ["w", "h", "fill"]) { const v = attr(p[1], k); if (v) attrs.push(`${k}=${v}`); }
    const stroke = attr(p[1], "stroke");
    if (stroke === "0" || stroke === "false") attrs.push("stroke=false");
    const cmds: string[] = [];
    for (const c of (p[2] ?? "").matchAll(/<a:(moveTo|lnTo|cubicBezTo|quadBezTo|arcTo|close)\b([^>]*?)(?:\/>|>([\s\S]*?)<\/a:\1>)/g)) {
      const pts = [...(c[3] ?? "").matchAll(/<a:pt\b([^>]*)\/>/g)].map((q) => `${attr(q[1], "x") ?? 0} ${attr(q[1], "y") ?? 0}`);
      if (c[1] === "moveTo") cmds.push(`M ${pts[0]}`);
      else if (c[1] === "lnTo") cmds.push(`L ${pts[0]}`);
      else if (c[1] === "cubicBezTo") cmds.push(`C ${pts.join(" ")}`);
      else if (c[1] === "quadBezTo") cmds.push(`Q ${pts.join(" ")}`);
      else if (c[1] === "arcTo") cmds.push(`A ${["wR", "hR", "stAng", "swAng"].map((k) => attr(c[2], k) ?? "0").join(" ")}`);
      else cmds.push("Z");
    }
    lines.push(`path${attrs.length ? " " + attrs.join(" ") : ""}: ${cmds.join(" ")}`);
  }
  return lines.join("\n");
}
