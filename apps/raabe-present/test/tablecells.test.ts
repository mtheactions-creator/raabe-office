// Zellen verbinden / Zellen teilen on a:tbl (D-072).

import { describe, expect, it } from "vitest";
import { PptxDom } from "@raabe/ooxml";
import * as TC from "../src/tablecells";
import * as T from "../src/textops";

const A = `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`;
const cell = (t: string) => `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p>${t ? `<a:r><a:rPr lang="de-DE"/><a:t>${t}</a:t></a:r>` : ""}<a:endParaRPr lang="de-DE"/></a:p></a:txBody><a:tcPr/></a:tc>`;
function table(rows: string[][], w = 1000, h = 400): Element {
  const grid = rows[0].map(() => `<a:gridCol w="${w}"/>`).join("");
  return PptxDom.parseXml(`<a:tbl ${A}><a:tblPr/><a:tblGrid>${grid}</a:tblGrid>${rows.map((r) => `<a:tr h="${h}">${r.map(cell).join("")}</a:tr>`).join("")}</a:tbl>`).documentElement;
}
const attrs = (tbl: Element) => TC.grid(tbl).map((row) => row.map((tc) => ["gridSpan", "rowSpan", "hMerge", "vMerge"].map((a) => tc.getAttribute(a) ?? "").join(",")));
const text = (tc: Element) => T.paras(PptxDom.kid(tc, "txBody")!).map(T.paraText);

describe("table cells", () => {
  it("merges a 2 x 2 range as PowerPoint writes it, texts as paragraphs of the first cell", () => {
    const tbl = table([["a", "b", "c"], ["d", "", "f"]]);
    TC.merge(tbl, { r0: 0, c0: 0, r1: 1, c1: 1 });
    expect(attrs(tbl)).toEqual([["2,2,,", ",,1,", ",,,"], ["2,,,1", ",,1,1", ",,,"]]);
    expect(text(TC.grid(tbl)[0][0])).toEqual(["a", "b", "d"]);
  });

  it("grows a range to whole merged cells", () => {
    const tbl = table([["a", "b", "c"], ["d", "e", "f"]]);
    TC.merge(tbl, { r0: 0, c0: 0, r1: 1, c1: 1 });
    expect(TC.normalize(TC.grid(tbl), { r: 1, c: 1 }, { r: 1, c: 2 })).toEqual({ r0: 0, c0: 0, r1: 1, c1: 2 });
    expect(TC.isMulti(TC.grid(tbl), { r0: 0, c0: 0, r1: 1, c1: 1 })).toBe(false);
  });

  it("splits a merged cell back into its cells", () => {
    const tbl = table([["a", "b"], ["c", "d"]]);
    TC.merge(tbl, { r0: 0, c0: 0, r1: 1, c1: 1 });
    TC.split(tbl, 0, 0, 2, 2);
    expect(attrs(tbl)).toEqual([[",,,", ",,,"], [",,,", ",,,"]]);
  });

  it("splits a plain cell into two columns: the column halves, the other rows span it", () => {
    const tbl = table([["a", "b"], ["c", "d"]]);
    TC.split(tbl, 0, 0, 2, 1);
    const cols = PptxDom.kids(PptxDom.kid(tbl, "tblGrid")!, "gridCol").map((c) => c.getAttribute("w"));
    expect(cols).toEqual(["500", "500", "1000"]);
    expect(attrs(tbl)).toEqual([[",,,", ",,,", ",,,"], ["2,,,", ",,1,", ",,,"]]);
    expect(text(TC.grid(tbl)[1][0])).toEqual(["c"]);
  });

  it("splits a plain cell into two rows: the row halves, the other columns span it", () => {
    const tbl = table([["a", "b"]]);
    TC.split(tbl, 0, 1, 1, 2);
    expect(PptxDom.kids(tbl, "tr").map((r) => r.getAttribute("h"))).toEqual(["200", "200"]);
    expect(attrs(tbl)).toEqual([[",2,,", ",,,"], [",,,1", ",,,"]]);
  });
});
