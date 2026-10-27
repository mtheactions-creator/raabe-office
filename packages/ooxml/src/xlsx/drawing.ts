// Charts of a worksheet: read from its drawing part (xl/drawings/drawingN.xml) into the
// sheet's chart models and written back. The drawing and every chart part stay as Excel wrote
// them while the charts are unchanged; pictures and shapes in the drawing are always kept.

import { strFromU8 } from "fflate";
import type { ChartModel } from "@raabe/charts";
import type { Sheet } from "@raabe/sheet-engine";
import { PackageWriter, Rels, relsPath, resolveTarget } from "../package";
import {
  CHART_CONTENT_TYPE, CHART_REL_TYPE, CHARTEX_CONTENT_TYPE, CHARTEX_REL_TYPE, DRAWING_CONTENT_TYPE, DRAWING_REL_TYPE,
  anchorFromXml, anchorXml, chartFromXml, chartToXml, drawingXml, isChartEx,
} from "./chart";
import { isRel } from "./read";
import { canon } from "./styles";
import { children, outer, root } from "./xml";

/** The sheet's charts (the model's optional `charts` field). */
export function chartsOf(s: Sheet): ChartModel[] | undefined {
  return (s as unknown as { charts?: ChartModel[] }).charts;
}

function setCharts(s: Sheet, charts: ChartModel[]): void {
  (s as unknown as { charts?: ChartModel[] }).charts = charts;
}

/** A chart as read: its part, its anchor's XML and the model's corners. */
interface ChartRead { xml: string; part: string; rId: string; anchor: string; place: string }

export interface DrawingRaw {
  part: string;
  xml: string;
  charts: ChartRead[];
  snap: string;
}

const rawXml = (m: ChartModel) => (m.raw as { xml?: unknown } | undefined)?.xml;
const placeOf = (m: ChartModel) => canon([m.from, m.to, m.name ?? null]);

/** Reads the charts of a sheet's drawing; undefined when it has none. */
export function readCharts(files: Record<string, Uint8Array>, sheet: Sheet, sheetPart: string, sheetRels: Rels): DrawingRaw | undefined {
  const rel = sheetRels.list.find((r) => isRel(r.type, DRAWING_REL_TYPE));
  if (!rel) return undefined;
  const part = resolveTarget(sheetPart, rel.target);
  if (!files[part]) return undefined;
  const xml = strFromU8(files[part]);
  const rels = new Rels(files[relsPath(part)] ? strFromU8(files[relsPath(part)]) : null);
  const r = root(xml);
  const charts: ChartModel[] = [];
  const read: ChartRead[] = [];
  for (const a of children(xml, r.open, r.close)) {
    const ax = outer(xml, a);
    if (!/<(?:c|cx):chart\b/.test(ax)) continue;
    const info = anchorFromXml(ax);
    const cr = info && rels.get(info.rId);
    const cpart = cr ? resolveTarget(part, cr.target) : "";
    if (!info || !cr || !files[cpart]) continue;
    try {
      const m = chartFromXml(strFromU8(files[cpart]), { from: info.from, to: info.to, name: info.name });
      charts.push(m);
      read.push({ xml: String(rawXml(m) ?? ""), part: cpart, rId: info.rId, anchor: ax, place: placeOf(m) });
    } catch {
      // A chart the model cannot read stays in the drawing as it is (it is not a model chart).
    }
  }
  if (!charts.length) return undefined;
  setCharts(sheet, charts);
  return { part, xml, charts: read, snap: canon(charts) };
}

/**
 * Writes the sheet's charts when they changed: kept charts reuse their parts and anchors,
 * new ones get parts of their own, removed ones lose their anchor (their parts go unused).
 * Returns the <drawing> element for the sheet, null to remove it, undefined to keep it.
 */
export function writeCharts(pw: PackageWriter, sheet: Sheet, sheetPart: string, raw: DrawingRaw | undefined, current: string | undefined): string | null | undefined {
  const charts = chartsOf(sheet) ?? [];
  if (raw ? canon(charts) === raw.snap : !charts.length) return undefined;
  const sheetRels = pw.relsOf(sheetPart);
  let drel = sheetRels.list.find((r) => isRel(r.type, DRAWING_REL_TYPE));
  let part = drel ? resolveTarget(sheetPart, drel.target) : undefined;
  if (part && !pw.files[part]) part = undefined;
  const old = part ? strFromU8(pw.files[part]) : undefined;
  // The drawing's anchors in their order: pictures and shapes stay, a chart's slot is filled by
  // its model (or dropped); new charts come last.
  const slots: { xml: string; chart?: ChartRead }[] = [];
  let head = "", tail = "";
  if (old) {
    const r = root(old);
    const kids = children(old, r.open, r.close);
    head = old.slice(0, kids.length ? kids[0].start : r.close);
    tail = old.slice(kids.length ? kids[kids.length - 1].end : r.close);
    for (const a of kids) {
      const ax = outer(old, a);
      slots.push({ xml: ax, chart: /<(?:c|cx):chart\b/.test(ax) ? raw?.charts.find((c) => c.anchor === ax) : undefined });
    }
  }
  const others = slots.filter((x) => !x.chart);
  if (!charts.length && !others.length) {
    if (drel) sheetRels.list.splice(sheetRels.list.indexOf(drel), 1), (sheetRels.changed = true);
    return null;
  }
  if (!part) part = pw.freeName("xl/drawings/drawing");
  const drels = pw.relsOf(part);
  let id = Math.max(1, ...slots.map((k) => Number(/<xdr:cNvPr\b[^>]*\bid="(\d+)"/.exec(k.xml)?.[1] ?? 1))) + 1;
  const placed = new Map<ChartRead, string>();
  const added: string[] = [];
  for (const m of charts) {
    const orig = raw?.charts.find((c) => !placed.has(c) && c.xml !== "" && c.xml === rawXml(m) && pw.files[c.part]);
    if (orig) {
      const xml = chartToXml(m);
      if (xml !== strFromU8(pw.files[orig.part])) pw.setPart(orig.part, xml);
      placed.set(orig, placeOf(m) === orig.place ? orig.anchor
        : anchorXml({ from: m.from, to: m.to, id: id++, name: m.name ?? `Diagramm ${id - 1}`, rId: orig.rId, chartEx: isChartEx(m) }));
      continue;
    }
    // A new chart (or a copy): a part of its own. Kept XML that refers to relationships of
    // another part is not reused.
    const ex = isChartEx(m);
    let xml = chartToXml(m);
    if (/\br:id="/.test(xml)) xml = chartToXml({ ...m, raw: undefined });
    const cpart = pw.freeName(ex ? "xl/charts/chartEx" : "xl/charts/chart");
    pw.setPart(cpart, xml, ex ? CHARTEX_CONTENT_TYPE : CHART_CONTENT_TYPE);
    const rId = pw.link(part, ex ? CHARTEX_REL_TYPE : CHART_REL_TYPE, cpart);
    added.push(anchorXml({ from: m.from, to: m.to, id: id++, name: m.name ?? `Diagramm ${id - 1}`, rId, chartEx: ex }));
  }
  const all = [...slots.map((x) => (x.chart ? placed.get(x.chart) ?? "" : x.xml)), ...added].join("");
  // Relationships of removed charts.
  for (const r of [...drels.list]) {
    if ((isRel(r.type, CHART_REL_TYPE) || isRel(r.type, CHARTEX_REL_TYPE)) && !all.includes(`r:id="${r.id}"`)) {
      drels.list.splice(drels.list.indexOf(r), 1);
      drels.changed = true;
    }
  }
  pw.setPart(part, old ? head + all + tail : drawingXml([all]), old ? undefined : DRAWING_CONTENT_TYPE);
  if (!drel) {
    pw.link(sheetPart, DRAWING_REL_TYPE, part);
    drel = sheetRels.list.find((r) => isRel(r.type, DRAWING_REL_TYPE))!;
  }
  return current && current.includes(`r:id="${drel.id}"`) ? current : `<drawing r:id="${drel.id}"/>`;
}
