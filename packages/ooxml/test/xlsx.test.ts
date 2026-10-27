import { readFileSync, readdirSync, mkdirSync, writeFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { CellError, cellKey, cellName, keyCell, rangeName, type Workbook } from "@raabe/sheet-engine";
import { newWorkbook, readXlsx, widthToFile, writeXlsx } from "../src";

const DIR = "packages/ooxml/test/fixtures";
const FIXTURES = readdirSync(DIR).filter((f) => /^xlsx-.*\.xlsx$/.test(f)).sort();
const fixture = (name: string) => new Uint8Array(readFileSync(`${DIR}/${name}`));

/** A comparable view of a workbook's model. */
export function modelOf(wb: Workbook) {
  return {
    active: wb.active, date1904: wb.date1904, names: wb.names,
    sheets: wb.sheets.map((s) => ({
      name: s.name, hidden: s.hidden, tab: s.tabColor, view: s.view, merges: s.merges, dw: s.defaultColWidth, dh: s.defaultRowHeight,
      cols: [...s.cols].sort((a, b) => a[0] - b[0]), rows: [...s.rows].sort((a, b) => a[0] - b[0]).map(([r, info]) => [r, { ...info, s: info.s ? styleOf(wb, info.s) : undefined }]),
      cells: [...s.cells].sort((a, b) => a[0] - b[0]).map(([k, c]) => [k, { ...c, s: styleOf(wb, c.s) }]),
      autoFilter: s.autoFilter, conditional: s.conditional, page: s.page,
    })),
  };
}

function styleOf(wb: Workbook, id: number | undefined) {
  const { xf: _xf, ...rest } = wb.styles.get(id);
  return rest;
}

/** What Excel's COM interface shows of a workbook, as the model reads it (for xlsx_excel_check.py). */
function dumpForExcel(wb: Workbook) {
  return {
    date1904: wb.date1904, active: wb.active,
    names: wb.names.map((n) => ({ ...n, sheet: n.sheet === undefined ? undefined : wb.sheets[n.sheet].name })),
    sheets: wb.sheets.map((s) => ({
      name: s.name, hidden: s.hidden ?? false, tab: s.tabColor ?? null, view: s.view,
      cols: Object.fromEntries([...s.cols].map(([c, info]) => [c + 1, { width: info.width === undefined ? null : widthToFile(info.width), hidden: !!info.hidden, outline: info.outline ?? 0 }])),
      rows: Object.fromEntries([...s.rows].map(([r, info]) => [r + 1, { height: info.height ?? null, custom: !!info.custom, hidden: !!info.hidden, outline: info.outline ?? 0 }])),
      merges: s.merges.map(rangeName),
      cells: Object.fromEntries([...s.cells].map(([k, c]) => {
        const { row, col } = keyCell(k);
        const st = wb.styles.get(c.s);
        const v = c.v instanceof CellError ? { error: c.v.code } : c.v;
        return [cellName(row, col), { v, f: c.f ?? null, array: c.array ? rangeName(c.array) : null, note: c.note ?? null, link: c.link ?? null, rich: !!c.richXml, style: { ...st, xf: undefined } }];
      })),
    })),
  };
}

describe("xlsx round trip of Excel's own files", () => {
  for (const name of FIXTURES) {
    it(name, () => {
      const bytes = fixture(name);
      const wb = readXlsx(bytes);
      const out = writeXlsx(wb, { now: new Date("2026-10-27T12:00:00Z") });
      if (process.env.ROUNDTRIP_OUT) {
        mkdirSync(process.env.ROUNDTRIP_OUT, { recursive: true });
        writeFileSync(`${process.env.ROUNDTRIP_OUT}/${name}`, out);
      }
      if (process.env.MODEL_OUT) {
        mkdirSync(process.env.MODEL_OUT, { recursive: true });
        writeFileSync(`${process.env.MODEL_OUT}/${name}.json`, JSON.stringify(dumpForExcel(wb)));
      }
      const back = readXlsx(out);
      expect(modelOf(back)).toEqual(modelOf(wb));
      // Every part the model does not rewrite is kept byte for byte.
      const a = unzipSync(bytes), b = unzipSync(out);
      const rewritten = /^(xl\/worksheets\/sheet\d+\.xml|xl\/sharedStrings\.xml|xl\/workbook\.xml|xl\/_rels\/workbook\.xml\.rels|docProps\/core\.xml|\[Content_Types\]\.xml|xl\/calcChain\.xml)$/;
      for (const p of Object.keys(a)) {
        if (rewritten.test(p)) continue;
        expect(b[p], p).toBeDefined();
        expect(strFromU8(b[p]) === strFromU8(a[p]), p).toBe(true);
      }
      expect(Object.keys(b).filter((p) => !a[p])).toEqual([]);
      // Unchanged elements keep Excel's text: only sheetData (and dimension) are rewritten.
      for (const p of Object.keys(a).filter((x) => /^xl\/worksheets\/sheet\d+\.xml$/.test(x))) {
        const strip = (x: string) => x.replace(/<sheetData[\s\S]*?(<\/sheetData>|<sheetData\/>)|<sheetData\/>/, "").replace(/<dimension[^>]*\/>/, "");
        expect(strip(strFromU8(b[p])), p).toBe(strip(strFromU8(a[p])));
      }
    });
  }
});

describe("new workbook", () => {
  it("is Excel's default: Tabelle1, Calibri 11, 15 pt rows, 10.71 characters", () => {
    const wb = newWorkbook();
    expect(wb.sheets.map((s) => s.name)).toEqual(["Tabelle1"]);
    expect(wb.styles.get(0).font).toMatchObject({ name: "Calibri", size: 11, scheme: "minor" });
    expect(wb.sheets[0].defaultRowHeight).toBe(15);
    expect(wb.sheets[0].defaultColWidth).toBeCloseTo(10.71, 2);
    expect(wb.theme.accent1).toBe("#4472C4");
  });

  it("round-trips values, formulas and a new style", () => {
    const wb = newWorkbook();
    const s = wb.sheets[0];
    s.set(0, 0, { v: 1.5 });
    s.set(0, 1, { v: "Grüße", s: wb.styles.derive(0, (st) => { st.font.bold = true; st.fill = "#FFFF00"; }) });
    s.set(1, 0, { v: 3, f: "A1*2" });
    s.set(1, 1, { v: true });
    const back = readXlsx(writeXlsx(wb));
    const t = back.sheets[0];
    expect(t.get(0, 0)?.v).toBe(1.5);
    expect(t.get(0, 1)?.v).toBe("Grüße");
    expect(back.styles.get(t.get(0, 1)?.s)).toMatchObject({ font: { bold: true }, fill: "#FFFF00" });
    expect(t.get(1, 0)).toMatchObject({ v: 3, f: "A1*2" });
    expect(t.cells.get(cellKey(1, 1))?.v).toBe(true);
  });
});
