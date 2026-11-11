// Compares a Format pane spec with Excel's recorded pane (results/2026-11-09-excel-chart-pane.json): option sets,
// tabs, sections and controls, by label and order (D-160). Used by paneorder.test.ts.

import type { PaneControl, PaneSpec } from "@raabe/shell";

export interface ExcelControl { t: string; label: string; group?: string; picker?: string; value?: unknown; menu?: boolean; entries?: string[] }
export interface ExcelSection { title: string; controls?: ExcelControl[]; same_as?: string; choices?: Record<string, string | { adds?: ExcelControl[]; removes?: ExcelControl[] }> }
export interface ExcelPane { title: string; sets: { label: string; tabs: { label: string; sections: ExcelSection[] }[] }[] }

export interface PaneDiff {
  pane: string;
  where: string;
  kind: "order" | "excel-only" | "ours-only";
  label?: string;
  excel: string[];
  ours: string[];
}

const norm = (s: string) => s.replace(/\s+/g, " ").replace(/ \.\.\.$/, "...").trim();

/** Excel's controls as the units a student sees: a colour label and its button are one, a slider and its box one. */
export function excelUnits(cs: ExcelControl[]): string[] {
  const out: string[] = [];
  const isColour = (c: ExcelControl | undefined) => !!c && c.t === "MenuItem" && (!!c.picker || String(c.value ?? "").startsWith("RGB("));
  for (let i = 0; i < cs.length; i++) {
    const c = cs[i], next = cs[i + 1];
    if (c.t === "Text") {
      // A label before its colour button; other texts are captions ("%", "Perioden", the trendline's automatic name).
      if (isColour(next)) { out.push(norm(c.label)); i++; }
      continue;
    }
    if (c.t === "Slider" && next?.t === "Spinner" && next.label === c.label) { out.push(norm(c.label)); i++; continue; }
    // The stop buttons belong to the gradient bar; "Auto" belongs to the bound before it.
    if (c.t === "Button" && c.group === "Farbverlaufstopps" && (c.label === "Hinzufügen" || c.label === "Entfernen")) continue;
    if (c.t === "Button" && c.label === "Auto") continue;
    // The pattern tiles are one grid.
    if (c.t === "ListItem") { if (cs[i - 1]?.t !== "ListItem") out.push("Muster"); continue; }
    out.push(norm(c.label));
  }
  return out;
}

/** Our controls as units (rows flattened; headings and free content are not controls). */
export function oursUnits(cs: PaneControl[]): string[] {
  const out: string[] = [];
  for (const c of cs) {
    if (c.type === "row") out.push(...oursUnits(c.items));
    else if (c.type === "heading" || c.type === "custom") continue;
    else out.push(norm(c.label));
  }
  return out;
}

/** The n-th occurrence of a label as "label#n" (n > 1), so repeated labels (Größe, Breite) keep their places. */
function tagged(list: string[]): string[] {
  const seen = new Map<string, number>();
  return list.map((l) => { const n = (seen.get(l) ?? 0) + 1; seen.set(l, n); return n > 1 ? `${l}#${n}` : l; });
}

/** Order of the labels both sides have, and the labels only one side has. */
export function compareLists(pane: string, where: string, excel: string[], ours: string[]): PaneDiff[] {
  const a = tagged(excel), b = tagged(ours);
  const inA = new Set(a), inB = new Set(b);
  const ca = a.filter((x) => inB.has(x)), cb = b.filter((x) => inA.has(x));
  const out: PaneDiff[] = [];
  if (ca.join("\n") !== cb.join("\n")) out.push({ pane, where, kind: "order", excel: ca, ours: cb });
  for (const x of a) if (!inB.has(x)) out.push({ pane, where, kind: "excel-only", label: x, excel: a, ours: b });
  for (const x of b) if (!inA.has(x)) out.push({ pane, where, kind: "ours-only", label: x, excel: a, ours: b });
  return out;
}

function resolve(sec: ExcelSection, set: string, tab: string, panes: Record<string, ExcelPane>, depth = 0): ExcelControl[] {
  if (!sec.same_as) return sec.controls ?? [];
  const p = panes[sec.same_as];
  if (!p || depth > 4) return [];
  const cands: { s: ExcelSection; score: number }[] = [];
  for (const st of p.sets) for (const t of st.tabs) for (const s of t.sections) {
    if (s.title === sec.title) cands.push({ s, score: (st.label === set ? 2 : 0) + (t.label === tab ? 1 : 0) });
  }
  cands.sort((x, y) => y.score - x.score);
  return cands[0] ? resolve(cands[0].s, set, tab, panes, depth + 1) : [];
}

/** Our marker sections (one tab, no page switch) and Excel's Markierung page: the same sections under our titles. */
const MARKER_TITLES: Record<string, string> = { Markierungsoptionen: "Markierungsoptionen", "Füllung": "Markierungsfüllung", Rahmen: "Markierungsrahmen" };

export function comparePane(key: string, excel: ExcelPane, ours: PaneSpec, panes: Record<string, ExcelPane>, markers: Record<string, { sections: ExcelSection[] }>): PaneDiff[] {
  const out: PaneDiff[] = [];
  if (norm(excel.title) !== norm(ours.title)) out.push({ pane: key, where: "Titel", kind: "order", excel: [excel.title], ours: [ours.title] });
  out.push(...compareLists(key, "Optionen", excel.sets.map((s) => norm(s.label)), ours.sets.map((s) => norm(s.label))));
  const marker = markers[`${key}: Füllung und Linie > Markierung`];
  for (const es of excel.sets) {
    const os = ours.sets.find((s) => norm(s.label) === norm(es.label));
    if (!os) continue;
    out.push(...compareLists(key, es.label, es.tabs.map((t) => norm(t.label)), os.tabs.map((t) => norm(t.label))));
    for (const et of es.tabs) {
      const ot = os.tabs.find((t) => norm(t.label) === norm(et.label));
      if (!ot) continue;
      const where = `${es.label} > ${et.label}`;
      const eSecs = et.sections.map((s) => ({ title: norm(s.title), controls: resolve(s, es.label, et.label, panes) }));
      if (marker && et.label === "Füllung und Linie") eSecs.push(...marker.sections.map((s) => ({ title: MARKER_TITLES[s.title] ?? s.title, controls: s.controls ?? [] })));
      // Our marker sections stand where Excel's Markierung page was not recorded: not compared there.
      const oSecs = ot.sections().filter((s) => marker || !Object.values(MARKER_TITLES).includes(s.title));
      out.push(...compareLists(key, where, eSecs.map((s) => s.title), oSecs.map((s) => norm(s.title))));
      for (const s of eSecs) {
        const o = oSecs.find((x) => norm(x.title) === s.title);
        if (!o) continue;
        out.push(...compareLists(key, `${where} > ${s.title}`, excelUnits(s.controls), oursUnits(o.controls())));
      }
    }
  }
  return out;
}

/**
 * The controls a radio choice adds (Excel's "choices": Farbverlauf adds Voreingestellte Farbverläufe, Typ, Richtung …):
 * ours, after picking the same radio in a fresh pane, must show them in Excel's order.
 */
export function compareChoices(key: string, excel: ExcelPane, make: () => PaneSpec, panes: Record<string, ExcelPane>): PaneDiff[] {
  const out: PaneDiff[] = [];
  for (const es of excel.sets) for (const et of es.tabs) for (const sec of et.sections) {
    let src: ExcelSection | undefined = sec;
    // A section recorded as the same as another pane's has that one's choices.
    for (let d = 0; src?.same_as && d < 5; d++) {
      const p: ExcelPane | undefined = panes[src.same_as];
      src = p?.sets.flatMap((s) => s.tabs.flatMap((t) => t.sections)).find((x) => x.title === sec.title);
    }
    for (const [choice, diff] of Object.entries(src?.choices ?? {})) {
      if (typeof diff !== "object" || !diff.adds?.length) continue;
      const radio = norm(choice.split(" > ").pop()!);
      const spec = make();
      const find = () => {
        const os = spec.sets.find((s) => norm(s.label) === norm(es.label));
        const ot = os?.tabs.find((t) => norm(t.label) === norm(et.label));
        return ot?.sections().find((s) => norm(s.title) === norm(sec.title));
      };
      const before = find();
      if (!before) continue;
      const flat = (cs: PaneControl[]): PaneControl[] => cs.flatMap((c) => (c.type === "row" ? flat(c.items) : [c]));
      const r = flat(before.controls()).find((c) => c.type === "radio" && norm(c.label) === radio);
      if (!r || r.type !== "radio") { out.push({ pane: key, where: `${es.label} > ${et.label} > ${sec.title} [${choice}]`, kind: "excel-only", label: radio, excel: [], ours: [] }); continue; }
      r.pick();
      const after = oursUnits(find()?.controls() ?? []);
      const adds = excelUnits(diff.adds);
      const where = `${es.label} > ${et.label} > ${sec.title} [${radio}]`;
      const inOurs = new Set(after);
      const common = adds.filter((x) => inOurs.has(x));
      const set = new Set(common);
      const seen = new Set<string>();
      const oursOrder = after.filter((x) => set.has(x) && !seen.has(x) && seen.add(x));
      if (common.join("|") !== oursOrder.join("|")) out.push({ pane: key, where, kind: "order", excel: common, ours: oursOrder });
      for (const x of adds) if (!inOurs.has(x)) out.push({ pane: key, where, kind: "excel-only", label: x, excel: adds, ours: after });
    }
  }
  return out;
}
