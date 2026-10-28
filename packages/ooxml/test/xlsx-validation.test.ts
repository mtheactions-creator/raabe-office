// Datenüberprüfung in .xlsx: Excel's own file (make_xlsx_fixtures.py --validation) read into the
// model, written back unchanged byte for byte, and edited: changed, new and removed validations
// in both of Excel's forms (<dataValidations> and the x14 entry of the extLst). With
// ROUNDTRIP_OUT set the edited file goes to xlsx_excel_check.py, which opens it in Excel.

import { readFileSync, mkdirSync, writeFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { Sheet, literalList, listItems, parseRange, type Validation, type Workbook } from "@raabe/sheet-engine";
import { newWorkbook, readXlsx, writeXlsx } from "../src";

const FILE = "packages/ooxml/test/fixtures/xlsx-validation.xlsx";
const load = () => readXlsx(new Uint8Array(readFileSync(FILE)));
const R = (a: string) => parseRange(a)!;
const at = (s: Sheet, a: string) => s.validations?.find((v) => v.ref.some((r) => r.r0 === R(a).r0 && r.c0 === R(a).c0));
/** Validations compared without regard to their order or the order of their fields. */
const plain = (vs: Validation[] | undefined) => (vs ?? []).map((v) => JSON.stringify(v, Object.keys(v).sort().concat(["r0", "c0", "r1", "c1"]))).sort();
const sheetXml = (bytes: Uint8Array, n = 1) => strFromU8(unzipSync(bytes)[`xl/worksheets/sheet${n}.xml`]);

function save(wb: Workbook, name: string, exp: { allow: string[]; require: [string, unknown][] }): Uint8Array {
  const out = writeXlsx(wb, { author: "RaabeCalc-Test", now: new Date("2026-10-28T12:00:00Z") });
  if (process.env.ROUNDTRIP_OUT) {
    mkdirSync(process.env.ROUNDTRIP_OUT, { recursive: true });
    writeFileSync(`${process.env.ROUNDTRIP_OUT}/${name}.edit.xlsx`, out);
    writeFileSync(`${process.env.ROUNDTRIP_OUT}/${name}.edit.expect.json`, JSON.stringify(exp));
  }
  return out;
}

describe("Datenüberprüfung read from Excel's file", () => {
  it("models every kind Excel wrote", () => {
    const s = load().sheetByName("Prüfung")!;
    expect(s.validations?.length).toBe(14);
    expect(at(s, "B2")).toMatchObject({ type: "list", formula1: '"Ja,Nein"', allowBlank: true, showInputMessage: true, showErrorMessage: true, promptTitle: "Teilnahme", prompt: "Ja oder Nein wählen" });
    expect(listItems(at(s, "B2")!)).toEqual(["Ja", "Nein"]);
    expect(at(s, "C2")).toMatchObject({ type: "list", formula1: "$H$2:$H$5" });
    expect(listItems(at(s, "C2")!)).toBeNull();
    // A list on another sheet: Excel's x14 form in the extLst.
    expect(at(s, "D2")).toMatchObject({ type: "list", formula1: "Listen!$A$1:$A$3", ref: [R("D2:D6")] });
    expect(at(s, "E2")).toMatchObject({ type: "whole", formula1: "1", formula2: "6", errorTitle: "Ungültige Note", error: "Nur ganze Noten von 1 bis 6." });
    expect(at(s, "E2")!.operator).toBeUndefined();
    expect(at(s, "F2")).toMatchObject({ type: "decimal", operator: "greaterThan", errorStyle: "warning", formula1: "0" });
    expect(at(s, "G2")).toMatchObject({ type: "date", errorStyle: "information", formula1: "46235", formula2: "46599" });
    expect(at(s, "I2")).toMatchObject({ type: "time", operator: "lessThan", formula1: "0.5" });
    expect(at(s, "I2")!.showErrorMessage).toBeUndefined();
    expect(at(s, "J2")).toMatchObject({ type: "textLength", operator: "lessThanOrEqual", formula1: "20" });
    expect(at(s, "K2")).toMatchObject({ type: "custom", formula1: "ISNUMBER(K2)" });
    // showDropDown="1": Zellendropdown is off.
    expect(at(s, "L2")).toMatchObject({ type: "list", formula1: "Faecher", showDropDown: true });
    expect(at(s, "M2")).toMatchObject({ type: "list", items: ["1", "2", "2,5"], formula1: '"1,2,2,5"' });
    expect(at(s, "N2")).toMatchObject({ type: "any", promptTitle: "Hinweis", prompt: "Hier steht nur ein Hinweis." });
    expect(at(s, "O2")!.ref).toEqual([R("O2:O4"), R("O8:O9")]);
    expect(listItems(at(s, "O2")!)).toEqual(["a,b", '"c"', "d"]);
    expect(at(s, "P2")).toMatchObject({ type: "custom", errorStyle: "warning", formula1: "P2<=Listen!$B$1" });
    expect(s.validations!.every((v) => !v.rawXml)).toBe(true);
  });

  it("writes unchanged validations back exactly as Excel wrote them", () => {
    const bytes = new Uint8Array(readFileSync(FILE));
    const wb = readXlsx(bytes);
    wb.sheetByName("Prüfung")!.set(0, 20, { v: "neu" });
    const out = writeXlsx(wb);
    const cut = (x: string) => /<dataValidations[\s\S]*<\/dataValidations>/.exec(x)![0] + /<extLst>[\s\S]*<\/extLst>/.exec(x)![0];
    expect(cut(sheetXml(out))).toBe(cut(sheetXml(bytes)));
  });
});

describe("Datenüberprüfung edited", () => {
  it("keeps Excel's text for the rest, writes new ones in Excel's forms", () => {
    const wb = load();
    const s = wb.sheetByName("Prüfung")!;
    const list = s.validations!;
    // Changed: a range moved, a message edited; removed: the hint; new: four kinds.
    at(s, "C2")!.ref = [R("C2:C10")];
    at(s, "E2")!.error = "Bitte eine Note von 1 bis 6 eingeben.";
    s.validations = list.filter((v) => v !== at(s, "N2"));
    const fresh: Validation[] = [
      { ref: [R("Q2:Q6")], type: "list", ...literalList(["rot", "grün, blau"]), allowBlank: true, showInputMessage: true, showErrorMessage: true },
      { ref: [R("R2:R6")], type: "decimal", operator: "notBetween", formula1: "1.5", formula2: "2.5", allowBlank: true, showErrorMessage: true, errorStyle: "information", errorTitle: "Hinweis <1>", error: "Wert & Grenze" },
      { ref: [R("S2:S6")], type: "list", formula1: "Listen!$A$1:$A$2", allowBlank: true, showErrorMessage: true, showDropDown: true },
      { ref: [R("T2")], type: "any", showInputMessage: true, promptTitle: "Titel", prompt: "Zeile 1\nZeile 2" },
    ];
    s.validations.push(...fresh);
    const before = sheetXml(new Uint8Array(readFileSync(FILE)));
    // Excel may show differences only where the edit was; it must read the new ones as written
    // (indices: Type, Operator, Formula1, Formula2, AlertStyle, IgnoreBlank, InCellDropdown, ShowInput, ShowError, …, ErrorMessage).
    const dv = "/sheets/Prüfung/validation/";
    const out = save(wb, "xlsx-validation", {
      allow: ["cells", "C7", "C8", "C9", "C10", "E", "N2", "Q", "R", "S", "T"].map((k) => dv + k),
      require: [
        [`${dv}C10/2`, "=$H$2:$H$5"], [`${dv}E4/12`, "Bitte eine Note von 1 bis 6 eingeben."], [`${dv}N2`, null],
        [`${dv}Q2/2`, "rot;grün, blau"], [`${dv}R3/0`, 2], [`${dv}R3/1`, 2], [`${dv}R3/2`, "1,5"], [`${dv}R3/4`, 3], [`${dv}R3/11`, "Hinweis <1>"],
        [`${dv}S6/2`, "=Listen!$A$1:$A$2"], [`${dv}S6/6`, false], [`${dv}S6/7`, false], [`${dv}T2/10`, "Zeile 1\nZeile 2"], [`${dv}D2/2`, "=Listen!$A$1:$A$3"],
      ],
    });
    const xml = sheetXml(out);
    // Unchanged elements keep Excel's text, xr:uid included; changed ones keep it with the new sqref.
    const el = (x: string, ref: string) => new RegExp(`<dataValidation\\b[^>]*sqref="${ref}"[\\s\\S]*?(?:</dataValidation>|/>)`).exec(x)?.[0];
    expect(el(xml, "B2:B6")).toBe(el(before, "B2:B6"));
    expect(el(xml, "C2:C10")).toBe(el(before, "C2:C6")!.replace('sqref="C2:C6"', 'sqref="C2:C10"'));
    expect(xml).not.toContain('promptTitle="Hinweis"');
    expect(xml).toContain('<x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" xr:uid="{E8154E14-9B45-4615-8799-A1245C5AD04E}">');
    // New ones as Excel writes them.
    expect(xml).toContain('<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="Q2:Q6"><mc:AlternateContent');
    expect(xml).toContain('<x12ac:list>rot,"grün, blau"</x12ac:list></mc:Choice><mc:Fallback><formula1>"rot,grün, blau"</formula1>');
    expect(xml).toContain('<dataValidation type="decimal" errorStyle="information" operator="notBetween" allowBlank="1" showErrorMessage="1" errorTitle="Hinweis &lt;1&gt;" error="Wert &amp; Grenze" sqref="R2:R6"><formula1>1.5</formula1><formula2>2.5</formula2></dataValidation>');
    expect(xml).toContain('<x14:dataValidation type="list" allowBlank="1" showDropDown="1" showErrorMessage="1"><x14:formula1><xm:f>Listen!$A$1:$A$2</xm:f></x14:formula1><xm:sqref>S2:S6</xm:sqref></x14:dataValidation>');
    expect(xml).toContain('<x14:dataValidations count="3"');
    expect(xml).toContain('<dataValidations count="14">');
    // And the model reads back the same.
    const back = readXlsx(out).sheetByName("Prüfung")!;
    expect(plain(back.validations)).toEqual(plain(s.validations));
  });

  it("removes the extLst entry and the element when nothing is left", () => {
    const wb = load();
    wb.sheetByName("Prüfung")!.validations = [];
    const xml = sheetXml(writeXlsx(wb));
    expect(xml).not.toContain("dataValidation");
    expect(xml).not.toContain("<extLst>");
    expect(readXlsx(writeXlsx(wb)).sheetByName("Prüfung")!.validations).toBeUndefined();
  });

  it("in a new workbook, with a list on another sheet", () => {
    const wb = newWorkbook();
    const s = wb.sheets[0];
    const t = new Sheet("Werte");
    wb.sheets.push(t);
    t.set(0, 0, { v: "A" });
    t.set(1, 0, { v: "B" });
    s.validations = [
      { ref: [R("A1:A5")], type: "list", formula1: "Werte!$A$1:$A$2", allowBlank: true, showInputMessage: true, showErrorMessage: true },
      { ref: [R("B1:B5")], type: "whole", operator: "greaterThanOrEqual", formula1: "0", allowBlank: true, showInputMessage: true, showErrorMessage: true },
    ];
    const out = save(wb, "xlsx-blank-validation", { allow: [], require: [] });
    const xml = sheetXml(out);
    expect(xml).toMatch(/<dataValidations count="1"><dataValidation type="whole" operator="greaterThanOrEqual"[^>]*sqref="B1:B5"><formula1>0<\/formula1><\/dataValidation><\/dataValidations><pageMargins/);
    expect(xml).toMatch(/<extLst><ext uri="\{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF\}" xmlns:x14="[^"]+"><x14:dataValidations count="1" xmlns:xm="[^"]+">/);
    expect(xml.endsWith("</extLst></worksheet>")).toBe(true);
    // The extLst's validations come after the sheet's own.
    expect(plain(readXlsx(out).sheets[0].validations)).toEqual(plain(s.validations));
  });
});
