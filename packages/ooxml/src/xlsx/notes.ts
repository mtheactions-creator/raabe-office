// Notizen (legacy comments): the text lives in xl/commentsN.xml, the yellow box in a VML
// drawing (xl/drawings/vmlDrawingN.vml). Notes the user did not touch keep Excel's XML and
// shapes; new ones get Excel's default note box.

import { cellName, parseCell } from "@raabe/sheet-engine";
import { esc, tXml, XML_HEAD } from "./xml";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/** Excel's run format for note text (Segoe UI 9, comment colour). */
const NOTE_RPR = '<rPr><sz val="9"/><color indexed="81"/><rFont val="Segoe UI"/><family val="2"/></rPr>';

export interface NoteIn { row: number; col: number; text: string }

/** comments.xml for the notes, reusing the original <comment> of unchanged ones. */
export function commentsXml(notes: NoteIn[], author: string, orig?: { xml: string; byRef: Record<string, { text: string; xml: string }> }): string {
  let head = `${XML_HEAD}<comments xmlns="${MAIN}">`;
  let authors: string[] = [];
  if (orig?.xml) {
    const m = /^([\s\S]*?<comments\b[^>]*>)/.exec(orig.xml);
    if (m) head = m[1];
    const a = /<authors>([\s\S]*?)<\/authors>/.exec(orig.xml)?.[1] ?? "";
    authors = [...a.matchAll(/<author>([\s\S]*?)<\/author>/g)].map((x) => x[1]);
  }
  const name = esc(author || "RaabeCalc");
  let id = authors.indexOf(name);
  const list = notes.map((n) => {
    const ref = cellName(n.row, n.col);
    const old = orig?.byRef[ref];
    if (old && old.text === n.text) return old.xml;
    if (id < 0) { authors.push(name); id = authors.length - 1; }
    return `<comment ref="${ref}" authorId="${id}" shapeId="0"><text><r>${NOTE_RPR}${tXml(n.text)}</r></text></comment>`;
  });
  if (!authors.length) authors.push(name);
  return `${head}<authors>${authors.map((a) => `<author>${a}</author>`).join("")}</authors><commentList>${list.join("")}</commentList></comments>`;
}

const SHAPETYPE = '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"\n  path="m,l,21600r21600,l21600,xe">\n  <v:stroke joinstyle="miter"/>\n  <v:path gradientshapeok="t" o:connecttype="rect"/>\n </v:shapetype>';

/** Excel's hidden note box for a cell, beside it to the right. */
function noteShape(id: number, row: number, col: number, z: number): string {
  const anchor = `${col + 1}, 15, ${Math.max(0, row - 1)}, 10, ${col + 3}, 15, ${row + 3}, 4`;
  return `<v:shape id="_x0000_s${id}" type="#_x0000_t202" style='position:absolute;\n  margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:${z};\n  visibility:hidden' fillcolor="infoBackground [80]" strokecolor="none [81]"\n  o:insetmode="auto">\n` +
    '  <v:fill color2="infoBackground [80]"/>\n  <v:shadow color="none [81]" obscured="t"/>\n  <v:path o:connecttype="none"/>\n' +
    "  <v:textbox style='mso-direction-alt:auto'>\n   <div style='text-align:left'></div>\n  </v:textbox>\n" +
    `  <x:ClientData ObjectType="Note">\n   <x:MoveWithCells/>\n   <x:SizeWithCells/>\n   <x:Anchor>\n    ${anchor}</x:Anchor>\n   <x:AutoFill>False</x:AutoFill>\n   <x:Row>${row}</x:Row>\n   <x:Column>${col}</x:Column>\n  </x:ClientData>\n </v:shape>`;
}

/** The shapes of a VML drawing with the cell of each note shape. */
function shapes(vml: string): { xml: string; note?: string; id: number }[] {
  return [...vml.matchAll(/<v:shape\b[\s\S]*?<\/v:shape>/g)].map((m) => {
    const xml = m[0];
    const id = Number(/id="_x0000_s(\d+)"/.exec(xml)?.[1] ?? 0);
    if (!/ObjectType="Note"/.test(xml)) return { xml, id };
    const r = Number(/<x:Row>(\d+)<\/x:Row>/.exec(xml)?.[1] ?? -1);
    const c = Number(/<x:Column>(\d+)<\/x:Column>/.exec(xml)?.[1] ?? -1);
    return { xml, id, note: r >= 0 && c >= 0 ? cellName(r, c) : undefined };
  });
}

/**
 * The VML drawing for the notes: other shapes (form controls) and the boxes of notes that
 * stay are kept; null when nothing is left.
 */
export function vmlXml(notes: NoteIn[], orig: string | null, idBlock: number): string | null {
  const old = orig ? shapes(orig) : [];
  const want = new Set(notes.map((n) => cellName(n.row, n.col)));
  const keep = old.filter((s) => !s.note || want.has(s.note));
  const have = new Set(keep.map((s) => s.note).filter(Boolean));
  let next = Math.max(idBlock * 1024, ...old.map((s) => s.id)) + 1;
  const added = notes.filter((n) => !have.has(cellName(n.row, n.col))).map((n, i) => noteShape(next++, n.row, n.col, keep.length + i + 1));
  if (!keep.length && !added.length) return null;
  if (orig && !added.length && keep.length === old.length) return orig;
  if (orig) {
    // Keep the drawing's head (namespaces, shape layout, shape types) and its tail.
    const first = orig.search(/<v:shape\b/);
    const lastEnd = orig.lastIndexOf("</v:shape>");
    const head = first >= 0 ? orig.slice(0, first) : orig.slice(0, orig.lastIndexOf("</xml>"));
    const tail = lastEnd >= 0 ? orig.slice(lastEnd + "</v:shape>".length) : orig.slice(orig.lastIndexOf("</xml>"));
    const withType = /_x0000_t202/.test(head) || !added.length ? head : head + SHAPETYPE;
    return withType + keep.map((s) => s.xml).join("") + added.join("") + tail;
  }
  return `<xml xmlns:v="urn:schemas-microsoft-com:vml"\n xmlns:o="urn:schemas-microsoft-com:office:office"\n xmlns:x="urn:schemas-microsoft-com:office:excel">\n` +
    ` <o:shapelayout v:ext="edit">\n  <o:idmap v:ext="edit" data="${idBlock}"/>\n </o:shapelayout>${SHAPETYPE}${added.join("")}</xml>`;
}

/** The o:idmap block numbers used by VML drawings of the package. */
export function vmlBlocks(vmls: string[]): number[] {
  const out: number[] = [];
  for (const v of vmls) for (const m of v.matchAll(/<o:idmap\b[^>]*data="([^"]*)"/g)) out.push(...m[1].split(/[ ,]+/).map(Number).filter(Number.isFinite));
  return out;
}

export function noteCell(ref: string): { row: number; col: number } | null {
  return parseCell(ref);
}
