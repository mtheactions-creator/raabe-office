// CSV against German Excel's own behaviour, recorded by tools/feel-lab/probe_excel_csv.py.

import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { CellError, Sheet, Workbook, decodeCsv, parseRange, readCsv, splitCsv, writeCsv, CSV_BOM } from "../src";

interface Probe {
  save: { cells: Record<string, { set: unknown; format: string | null }>; csv_utf8: { bom: boolean; text: string }; csv: { bom: boolean; text: string } };
  open: { fields: string[]; files: Record<string, { rows: { value: unknown; formula: string; format: string; text: string }[][] }> };
}

const probe: Probe = JSON.parse(readFileSync("packages/sheet-engine/test/fixtures/excel-csv.json", "utf-8"));
const ERR_DE: Record<string, string> = { "#NV": "#N/A", "#DIV/0!": "#DIV/0!", "#WERT!": "#VALUE!" };

/** Excel's English NumberFormat drops escapes and quotes, and shows German currency formats US-style. */
function sameFormat(ours: string, excel: string): boolean {
  const norm = (s: string) => s.replace(/\\/g, "").replace(/"/g, "");
  if (norm(ours) === norm(excel)) return true;
  if (excel === "$#,##0_);[Red]($#,##0)") return ours === '#,##0\\ "€";[Red]\\-#,##0\\ "€"';
  if (excel === "$#,##0.00_);[Red]($#,##0.00)") return ours === '#,##0.00\\ "€";[Red]\\-#,##0.00\\ "€"';
  return false;
}

describe("reading CSV as German Excel", () => {
  for (const label of ["utf8_bom", "cp1252"]) {
    it(`reads every field of ${label} as Excel does`, () => {
      const text = probe.open.fields.join(";");
      const wb = readCsv(label === "utf8_bom" ? CSV_BOM + text : text, { today: new Date(2026, 9, 27) });
      const s = wb.sheets[0];
      const want = probe.open.files[label].rows[0];
      const diffs: string[] = [];
      want.forEach((w, c) => {
        const cell = s.get(0, c);
        const fmt = wb.styles.get(cell?.s).numFmt;
        const v = cell?.v ?? null;
        const field = splitCsv(text)[0][c];
        // cp1252: Excel reads the bytes as Windows-1252; our text is the decoded string already.
        if (label === "cp1252" && /[^\x20-\x7e]/.test(field)) return;
        if (w.formula.startsWith("=") && w.formula !== w.text) {
          if (!cell?.f) diffs.push(`${field}: expected a formula like ${w.formula}`);
          return;
        }
        const wv = w.value as unknown;
        let ok: boolean;
        if (wv && typeof wv === "object" && "error" in wv) ok = v instanceof CellError && v.code === ERR_DE[(wv as { error: string }).error];
        else if (typeof wv === "number" && typeof v === "number") ok = Math.abs(wv - v) < 1e-9 * Math.max(1, Math.abs(wv));
        else ok = (wv ?? null) === v;
        if (!ok) diffs.push(`${JSON.stringify(field)}: value ${JSON.stringify(v)} vs Excel ${JSON.stringify(wv)}`);
        if (!sameFormat(fmt, w.format)) diffs.push(`${JSON.stringify(field)}: format ${fmt} vs Excel ${w.format}`);
      });
      expect(diffs).toEqual([]);
    });
  }

  it("keeps line breaks inside quotes and uses only ';'", () => {
    const wb = readCsv('Kopf;"zwei\r\nZeilen";Ende\r\n1;2;3\r\n');
    expect(wb.sheets[0].get(0, 1)?.v).toBe("zwei\r\nZeilen");
    expect(wb.sheets[0].get(1, 2)?.v).toBe(3);
    const comma = readCsv("1.5,1.000,-3");
    expect(comma.sheets[0].get(0, 0)?.v).toBe("1.5,1.000,-3");
    expect(readCsv("sep=,\n1,2").sheets[0].get(0, 1)?.v).toBe(2);
  });

  it("decodes UTF-8 with or without BOM and falls back to Windows-1252", () => {
    expect(decodeCsv(new Uint8Array([0xef, 0xbb, 0xbf, 0x47, 0xc3, 0xbc]))).toBe("Gü");
    expect(decodeCsv(new Uint8Array([0x47, 0xc3, 0xbc]))).toBe("Gü");
    expect(decodeCsv(new Uint8Array([0x47, 0xfc]))).toBe("Gü");
  });
});

describe("writing CSV as German Excel", () => {
  it("writes the used range as Excel's CSV UTF-8 does", () => {
    const wb = Workbook.blank("Tabelle1");
    const s = wb.sheets[0];
    const values: Record<string, unknown> = {
      B5: new CellError("#DIV/0!"), D5: false, E5: "", F5: 42, H4: "0123",
    };
    for (const [addr, c] of Object.entries(probe.save.cells)) {
      const r = parseRange(addr)!;
      let v = (addr in values ? values[addr] : c.set) as Sheet["cells"] extends Map<number, infer C> ? C extends { v: infer V } ? V : never : never;
      if (typeof v === "string" && v.startsWith("'")) v = v.slice(1);
      s.set(r.r0, r.c0, { v, s: c.format ? wb.styles.derive(0, (st) => { st.numFmt = c.format!; }) : undefined });
    }
    expect(probe.save.csv_utf8.bom).toBe(true);
    expect(probe.save.csv.bom).toBe(false);
    expect(writeCsv(s, { styles: wb.styles })).toBe(probe.save.csv_utf8.text);
  });
});
