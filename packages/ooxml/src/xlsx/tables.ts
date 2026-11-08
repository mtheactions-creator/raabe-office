// Excel tables (xl/tables/tableN.xml, the sheet's <tableParts>; D-066): read into the model and
// written as Excel writes them; an unchanged table keeps its part as it was.

import { strFromU8 } from "fflate";
import { fromXlsxFormula, mapStructs, parseRange, rangeName, structToFile, structToShown, tableKey, toXlsxFormula, type Sheet, type SheetTable } from "@raabe/sheet-engine";
import { Rels, resolveTarget, type PackageWriter } from "../package";

export const RT_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
export const CT_TABLE = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const attr = (tag: string, name: string) => new RegExp(`\\b${name}="([^"]*)"`).exec(tag)?.[1];
const unesc = (s: string) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/** A table part as a model. */
export function tableFromXml(xml: string, part?: string): SheetTable | null {
  const top = /<table\b[^>]*>/.exec(xml)?.[0];
  const ref = top ? parseRange(attr(top, "ref") ?? "") : null;
  if (!top || !ref) return null;
  const style = /<tableStyleInfo\b[^>]*>/.exec(xml)?.[0] ?? "";
  const columns = [...xml.matchAll(/<tableColumn\b([^>]*?)(?:\/>|>([\s\S]*?)<\/tableColumn>)/g)].map((m) => {
    const c: SheetTable["columns"][number] = { name: unesc(attr(m[1], "name") ?? "") };
    const fn = attr(m[1], "totalsRowFunction"), label = attr(m[1], "totalsRowLabel");
    if (fn) c.totalsFunction = fn;
    if (label !== undefined) c.totalsLabel = unesc(label);
    const f = /<calculatedColumnFormula\b[^>]*>([\s\S]*?)<\/calculatedColumnFormula>/.exec(m[2] ?? "")?.[1];
    if (f !== undefined) c.formula = unesc(f);
    return c;
  });
  const flag = (n: string, d: boolean) => (attr(style, n) === undefined ? d : attr(style, n) === "1");
  const t: SheetTable = {
    name: unesc(attr(top, "displayName") ?? attr(top, "name") ?? "Tabelle1"), ref,
    header: attr(top, "headerRowCount") !== "0",
    totals: Number(attr(top, "totalsRowCount") ?? 0) > 0,
    columns, style: attr(style, "name") !== undefined ? unesc(attr(style, "name")!) : "",
    firstColumn: flag("showFirstColumn", false), lastColumn: flag("showLastColumn", false),
    rowStripes: flag("showRowStripes", true), columnStripes: flag("showColumnStripes", false),
    // Schaltfläche "Filter": an autoFilter without hidden buttons.
    filterButton: /<autoFilter\b/.test(xml) && !/<filterColumn\b[^>]*hiddenButton="1"/.test(xml),
    id: Number(attr(top, "id") ?? 0) || undefined,
  };
  // The column formula as Excel shows it ([@Note]*2; the file has Noten[[#This Row],[Note]]*2).
  for (const c of t.columns) if (c.formula !== undefined) c.formula = mapStructs(fromXlsxFormula(c.formula), (x) => structToShown(x, t.name));
  if (part) { t.raw = part; t.rawKey = tableKey(t); }
  return t;
}

/** A table as Excel writes a new one. */
export function tableToXml(t: SheetTable, id: number): string {
  const ref = rangeName(t.ref);
  const body = { r0: t.ref.r0, c0: t.ref.c0, r1: t.ref.r1 - (t.totals ? 1 : 0), c1: t.ref.c1 };
  const filter = t.header && t.filterButton ? `<autoFilter ref="${rangeName(body)}"/>` : t.header && !t.filterButton
    ? `<autoFilter ref="${rangeName(body)}">${t.columns.map((_, i) => `<filterColumn colId="${i}" hiddenButton="1"/>`).join("")}</autoFilter>` : "";
  const cols = t.columns.map((c, i) => {
    const a = `id="${i + 1}" name="${esc(c.name)}"${c.totalsLabel !== undefined ? ` totalsRowLabel="${esc(c.totalsLabel)}"` : ""}${c.totalsFunction ? ` totalsRowFunction="${c.totalsFunction}"` : ""}`;
    return c.formula !== undefined ? `<tableColumn ${a}><calculatedColumnFormula>${esc(toXlsxFormula(mapStructs(c.formula, (x) => structToFile(x, t.name))))}</calculatedColumnFormula></tableColumn>` : `<tableColumn ${a}/>`;
  }).join("");
  const style = t.style ? `<tableStyleInfo name="${esc(t.style)}" showFirstColumn="${t.firstColumn ? 1 : 0}" showLastColumn="${t.lastColumn ? 1 : 0}" showRowStripes="${t.rowStripes ? 1 : 0}" showColumnStripes="${t.columnStripes ? 1 : 0}"/>`
    : `<tableStyleInfo showFirstColumn="${t.firstColumn ? 1 : 0}" showLastColumn="${t.lastColumn ? 1 : 0}" showRowStripes="${t.rowStripes ? 1 : 0}" showColumnStripes="${t.columnStripes ? 1 : 0}"/>`;
  return HEAD + `<table ${NS} id="${id}" name="${esc(t.name)}" displayName="${esc(t.name)}" ref="${ref}"${t.header ? "" : ' headerRowCount="0"'}${t.totals ? ' totalsRowCount="1"' : ' totalsRowShown="0"'}>`
    + `${filter}<tableColumns count="${t.columns.length}">${cols}</tableColumns>${style}</table>`;
}

/** The sheet's tables, from the parts its <tableParts> name. */
export function readTables(files: Record<string, Uint8Array>, sheet: Sheet, sheetPart: string, sheetRels: Rels): void {
  const out: SheetTable[] = [];
  for (const r of sheetRels.list.filter((x) => x.type === RT_TABLE)) {
    const part = resolveTarget(sheetPart, r.target);
    if (!files[part]) continue;
    const t = tableFromXml(strFromU8(files[part]), part);
    if (t) out.push(t);
  }
  if (out.length) sheet.tables = out;
}

/**
 * Writes the sheet's tables and returns its <tableParts> (null: none, undefined: as it was).
 * `nextId` hands out workbook-wide table ids for new tables.
 */
export function writeTables(pw: PackageWriter, sheet: Sheet, sheetPart: string, nextId: () => number, current: string | undefined): string | null | undefined {
  const tables = sheet.tables ?? [];
  const rels = pw.relsOf(sheetPart);
  const old = rels.list.filter((r) => r.type === RT_TABLE);
  if (!tables.length && !old.length) return current === undefined ? undefined : null;
  const ids: string[] = [];
  const keep = new Set<string>();
  for (const t of tables) {
    let part = t.raw;
    const same = !!t.raw && t.rawKey === tableKey(t) && !!pw.files[t.raw];
    if (!part || !pw.files[part]) part = pw.freeName("xl/tables/table");
    if (!same) pw.setPart(part, tableToXml(t, t.raw && t.id ? t.id : nextId()), CT_TABLE);
    keep.add(part);
    const rel = rels.list.find((r) => r.type === RT_TABLE && resolveTarget(sheetPart, r.target) === part);
    const rid = rel?.id ?? pw.link(sheetPart, RT_TABLE, part);
    ids.push(rid);
  }
  // Tables removed in RaabeCalc: their relationships (and parts) go.
  for (const r of old) {
    const part = resolveTarget(sheetPart, r.target);
    if (keep.has(part)) continue;
    rels.list.splice(rels.list.indexOf(r), 1);
    rels.changed = true;
    pw.removePart(part);
  }
  return ids.length ? `<tableParts count="${ids.length}">${ids.map((id) => `<tablePart r:id="${id}"/>`).join("")}</tableParts>` : null;
}
