// Szenario-Manager and Konsolidieren in the sheet part: <scenarios> (CT_Scenarios) and
// <dataConsolidate> (CT_DataConsolidate), read into the model and written back as Excel writes
// them (recorded: probe_excel_scenarios.py, probe_excel_consolidate.py). The writer leaves both
// elements as they were read while the model's part is unchanged.

import { cellName, parseCell, parseRange, rangeName, type Consolidation, type ConsolidateFunction, type Scenario, type Scenarios } from "@raabe/sheet-engine";
import { attrs, children, esc, type El } from "./xml";

const on = (v: string | undefined) => v === "1" || v === "true";

export function readScenarios(xml: string, el: El): Scenarios {
  const a = attrs(xml, el);
  const out: Scenarios = { list: [] };
  if (a.current !== undefined) out.current = Number(a.current);
  if (a.show !== undefined) out.show = Number(a.show);
  if (a.sqref) out.sqref = a.sqref.split(/\s+/).map((t) => parseRange(t)).filter((r): r is NonNullable<typeof r> => !!r);
  for (const s of children(xml, el.open, el.close)) {
    if (s.name !== "scenario") continue;
    const sa = attrs(xml, s);
    const sc: Scenario = { name: sa.name ?? "", cells: [] };
    if (on(sa.locked)) sc.locked = true;
    if (on(sa.hidden)) sc.hidden = true;
    if (sa.user !== undefined) sc.user = sa.user;
    if (sa.comment !== undefined) sc.comment = sa.comment;
    for (const ic of children(xml, s.open, s.close)) {
      if (ic.name !== "inputCells") continue;
      const ia = attrs(xml, ic);
      const at = parseCell((ia.r ?? "").replace(/\$/g, ""));
      if (!at) continue;
      const cell: Scenario["cells"][number] = { row: at.row, col: at.col, val: ia.val ?? "" };
      if (ia.numFmtId !== undefined) cell.numFmtId = Number(ia.numFmtId);
      if (on(ia.deleted)) cell.deleted = true;
      if (on(ia.undone)) cell.undone = true;
      sc.cells.push(cell);
    }
    out.list.push(sc);
  }
  return out;
}

/** <scenarios> in Excel's attribute order; null when the sheet has none. */
export function scenariosXml(s: Scenarios | undefined): string | null {
  if (!s || !s.list.length) return null;
  const head = `<scenarios current="${s.current ?? 0}"${s.show !== undefined ? ` show="${s.show}"` : ""}${s.sqref?.length ? ` sqref="${s.sqref.map(rangeName).join(" ")}"` : ""}>`;
  const body = s.list.map((sc) => {
    const tag = `<scenario name="${esc(sc.name)}"${sc.locked ? ' locked="1"' : ""}${sc.hidden ? ' hidden="1"' : ""} count="${sc.cells.length}"` +
      `${sc.user !== undefined ? ` user="${esc(sc.user)}"` : ""}${sc.comment ? ` comment="${esc(sc.comment)}"` : ""}>`;
    const cells = sc.cells.map((c) => `<inputCells r="${cellName(c.row, c.col)}"${c.deleted ? ' deleted="1"' : ""}${c.undone ? ' undone="1"' : ""} val="${esc(c.val)}"${c.numFmtId !== undefined ? ` numFmtId="${c.numFmtId}"` : ""}/>`).join("");
    return `${tag}${cells}</scenario>`;
  }).join("");
  return `${head}${body}</scenarios>`;
}

export function readConsolidate(xml: string, el: El): Consolidation {
  const a = attrs(xml, el);
  const out: Consolidation = { refs: [] };
  if (a.function && a.function !== "sum") out.fn = a.function as ConsolidateFunction;
  if (on(a.startLabels)) out.startLabels = true;
  if (on(a.leftLabels)) out.left = true;
  if (on(a.topLabels)) out.top = true;
  if (on(a.link)) out.link = true;
  for (const list of children(xml, el.open, el.close)) {
    if (list.name !== "dataRefs") continue;
    for (const r of children(xml, list.open, list.close)) {
      if (r.name !== "dataRef") continue;
      const ra = attrs(xml, r);
      const ref: Consolidation["refs"][number] = {};
      if (ra.ref !== undefined) ref.ref = ra.ref;
      if (ra.name !== undefined) ref.name = ra.name;
      if (ra.sheet !== undefined) ref.sheet = ra.sheet;
      if (ra["r:id"] !== undefined) ref.rid = ra["r:id"];
      out.refs.push(ref);
    }
  }
  return out;
}

/** <dataConsolidate> as Excel writes it (function only when not sum); null without settings. */
export function consolidateXml(c: Consolidation | undefined): string | null {
  if (!c) return null;
  const attr = `${c.fn && c.fn !== "sum" ? ` function="${c.fn}"` : ""}${c.startLabels ? ' startLabels="1"' : ""}${c.left ? ' leftLabels="1"' : ""}${c.top ? ' topLabels="1"' : ""}${c.link ? ' link="1"' : ""}`;
  if (!c.refs.length) return `<dataConsolidate${attr}/>`;
  const refs = c.refs.map((r) => `<dataRef${r.ref !== undefined ? ` ref="${esc(r.ref)}"` : ""}${r.name !== undefined ? ` name="${esc(r.name)}"` : ""}${r.sheet !== undefined ? ` sheet="${esc(r.sheet)}"` : ""}${r.rid !== undefined ? ` r:id="${esc(r.rid)}"` : ""}/>`).join("");
  return `<dataConsolidate${attr}><dataRefs count="${c.refs.length}">${refs}</dataRefs></dataConsolidate>`;
}
