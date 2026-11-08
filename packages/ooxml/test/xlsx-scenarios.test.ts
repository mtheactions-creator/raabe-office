// Szenario-Manager and Konsolidieren in .xlsx: Excel's own files (probe_excel_scenarios.py,
// probe_excel_consolidate.py) read into the model, written back unchanged byte for byte, and
// edited: scenarios added, changed and removed, settings of Konsolidieren replaced, in Excel's form
// (<scenarios>, <dataConsolidate>, recorded in results/2026-11-08-excel-*.json).

import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { Workbook } from "@raabe/sheet-engine";
import { readXlsx, writeXlsx } from "../src";

const SCEN = "packages/ooxml/test/fixtures/xlsx-scenarios.xlsx";
const CONS = "packages/ooxml/test/fixtures/xlsx-consolidate.xlsx";
const load = (f: string) => readXlsx(new Uint8Array(readFileSync(f)));
const partOf = (bytes: Uint8Array, n = 1) => strFromU8(unzipSync(bytes)[`xl/worksheets/sheet${n}.xml`]);
const scenXml = (x: string) => /<scenarios\b.*?<\/scenarios>/s.exec(x)?.[0] ?? null;
const consXml = (x: string) => /<dataConsolidate\b[^>]*?(\/>|>.*?<\/dataConsolidate>)/s.exec(x)?.[0] ?? null;
const rec = JSON.parse(readFileSync("packages/sheet-engine/test/fixtures/excel-scenarios.json", "utf-8"));
const cons = JSON.parse(readFileSync("packages/sheet-engine/test/fixtures/excel-consolidate.json", "utf-8"));
const write = (wb: Workbook) => writeXlsx(wb, { author: "RaabeCalc-Test", now: new Date("2026-11-08T12:00:00Z") });

describe("Szenarien read from Excel's file", () => {
  it("has Excel's five scenarios with cells, values, comments and Schutz", () => {
    const s = load(SCEN).sheet(0);
    const list = s.scenarios!.list;
    expect(list.map((x) => x.name)).toEqual(rec.final.map((x: { name: string }) => x.name));
    rec.final.forEach((want: { values: unknown[]; comment: string; locked: boolean; hidden: boolean }, i: number) => {
      const sc = list[i];
      // Excel's COM gives the values as German text, the file holds the same text.
      expect(sc.cells.map((c) => c.val)).toEqual(want.values.map(String));
      expect(sc.comment ?? "").toEqual(want.comment.replace("<user>", "Raabe Office"));
      expect(!!sc.locked).toBe(want.locked);
      expect(!!sc.hidden).toBe(want.hidden);
    });
    expect(list[2].cells.map((c) => [c.row, c.col])).toEqual([[0, 1], [2, 1]]);
    expect(s.scenarios!.current).toBe(0);
    expect(s.scenarios!.show).toBe(0);
    expect(s.scenarios!.sqref).toEqual([{ r0: 3, c0: 1, r1: 4, c1: 1 }]);
  });

  it("writes an unchanged sheet back as it was", () => {
    const bytes = new Uint8Array(readFileSync(SCEN));
    const out = write(readXlsx(bytes));
    expect(scenXml(partOf(out))).toBe(scenXml(partOf(bytes)));
  });

  it("writes changed scenarios as Excel does", () => {
    const bytes = new Uint8Array(readFileSync(SCEN));
    const wb = readXlsx(bytes);
    const s = wb.sheet(0).scenarios!;
    // The recorded XML of step 1 (file_added) is what the writer gives for the same list and current="0".
    s.show = undefined;
    s.sqref = undefined;
    s.list.forEach((sc) => (sc.user = "<user>"));
    for (const sc of s.list) if (sc.comment) sc.comment = sc.comment.replace("Raabe Office", "<user>");
    const x = scenXml(partOf(write(wb)));
    const want = (rec.file_added["xl/worksheets/sheet1.xml"].scenarios as string).replace(/<user>/g, "&lt;user&gt;");
    expect(x).toBe(want);
  });

  it("reads back what it wrote after adding and removing scenarios", () => {
    const wb = readXlsx(new Uint8Array(readFileSync(SCEN)));
    const s = wb.sheet(0);
    s.scenarios!.list.splice(1, 1);
    s.scenarios!.list.push({ name: "Neu & <anders>", cells: [{ row: 9, col: 3, val: "1,5" }, { row: 9, col: 4, val: "Text \"zitiert\"" }], comment: "Zeile 1\nZeile 2", user: "Test", hidden: true });
    s.scenarios!.current = 4;
    const back = readXlsx(write(wb)).sheet(0).scenarios!;
    expect(back.list).toEqual(s.scenarios!.list);
    expect(back.current).toBe(4);
    s.scenarios = undefined;
    expect(scenXml(partOf(write(wb)))).toBeNull();
  });

  it("a new workbook gets the element in the schema's place", () => {
    const wb = Workbook.blank();
    wb.sheet().scenarios = { list: [{ name: "A", cells: [{ row: 0, col: 0, val: "1" }], locked: true }] };
    wb.sheet().protection = { sheet: "1", scenarios: "1" };
    const x = partOf(write(wb));
    expect(scenXml(x)).toBe('<scenarios current="0"><scenario name="A" locked="1" count="1"><inputCells r="A1" val="1"/></scenario></scenarios>');
    expect(x.indexOf("<sheetProtection")).toBeLessThan(x.indexOf("<scenarios"));
    expect(x.indexOf("<scenarios")).toBeLessThan(x.indexOf("<pageMargins"));
  });
});

describe("Konsolidieren's settings in the sheet", () => {
  it("reads Excel's <dataConsolidate> and writes it back unchanged", () => {
    const bytes = new Uint8Array(readFileSync(CONS));
    const wb = readXlsx(bytes);
    expect(wb.sheet(0).consolidate).toEqual({ top: true, left: true, refs: [{ ref: "A1:C4", sheet: "Nord" }, { ref: "A1:C5", sheet: "Süd" }] });
    expect(consXml(partOf(write(wb)))).toBe(consXml(partOf(bytes)));
  });

  it("writes every recorded setting as Excel wrote it", () => {
    for (const c of cons.cases as { dataConsolidate: string | null; name: string }[]) {
      if (!c.dataConsolidate) continue;
      const wb = Workbook.blank();
      const x = c.dataConsolidate;
      const m = (n: string) => new RegExp(`\\b${n}="([^"]*)"`).exec(x.split("<dataRefs")[0])?.[1];
      wb.sheet().consolidate = {
        ...(m("function") ? { fn: m("function") as never } : {}),
        ...(m("leftLabels") ? { left: true } : {}), ...(m("topLabels") ? { top: true } : {}), ...(m("link") ? { link: true } : {}),
        refs: [...x.matchAll(/<dataRef ref="([^"]*)" sheet="([^"]*)"\/>/g)].map((r) => ({ ref: r[1], sheet: r[2] })),
      };
      expect(consXml(partOf(write(wb))), c.name).toBe(x);
    }
  });
});
