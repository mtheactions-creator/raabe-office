// Blatt schützen against Excel: the fixture xlsx-protection.xlsx was protected by Excel itself
// (make_xlsx_fixtures.py --protection, password "Raabe", Zellen formatieren and Sortieren allowed).

import { readFileSync, writeFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { readXlsx, writeXlsx } from "@raabe/ooxml";
import { legacyHash, protectionAttrs, sha512 } from "../src/protection";

const fixture = () => readXlsx(readFileSync("packages/ooxml/test/fixtures/xlsx-protection.xlsx"));

describe("Blatt schützen", () => {
  it("reads Excel's protection and checks its password with Excel's hash", async () => {
    const wb = fixture();
    const p = wb.sheets[0].protection!;
    expect(p).toMatchObject({ sheet: "1", algorithmName: "SHA-512", formatCells: "0", sort: "0" });
    expect(await sha512("Raabe", Uint8Array.from(atob(p.saltValue), (c) => c.charCodeAt(0)), Number(p.spinCount))).toBe(p.hashValue);
    expect(await sha512("raabe", Uint8Array.from(atob(p.saltValue), (c) => c.charCodeAt(0)), Number(p.spinCount))).not.toBe(p.hashValue);
    expect(wb.sheets[1].protection).toEqual({ sheet: "1", objects: "1", scenarios: "1" });
  }, 60000);

  it("writes the attributes in Excel's order with Excel's defaults", () => {
    const defaults = new Map([["selectLockedCells", true], ["selectUnlockedCells", true], ["formatCells", true], ["sort", true], ["objects", false], ["scenarios", false]]);
    const a = protectionAttrs(defaults, { algorithmName: "SHA-512", hashValue: "h", saltValue: "s", spinCount: "100000" });
    expect(Object.keys(a)).toEqual(["algorithmName", "hashValue", "saltValue", "spinCount", "sheet", "objects", "scenarios", "formatCells", "sort"]);
    expect(protectionAttrs(new Map([["objects", false], ["scenarios", false]]))).toEqual({ sheet: "1", objects: "1", scenarios: "1" });
  });

  it("keeps the protection on a round trip and writes a new one (checked in Excel by check_calc_protection.py)", async () => {
    const wb = fixture();
    const back = readXlsx(writeXlsx(wb));
    expect(back.sheets[0].protection).toEqual(wb.sheets[0].protection);
    const salt = new Uint8Array(16).map((_, i) => i * 7 + 3);
    wb.sheets[1].protection = protectionAttrs(new Map([["formatCells", true], ["objects", false], ["scenarios", false]]),
      { algorithmName: "SHA-512", hashValue: await sha512("Schule", salt, 100000), saltValue: btoa(String.fromCharCode(...salt)), spinCount: "100000" });
    wb.sheets[0].protection = undefined;
    wb.protection = { workbookAlgorithmName: "SHA-512", workbookHashValue: await sha512("Mappe", salt, 100000), workbookSaltValue: btoa(String.fromCharCode(...salt)), workbookSpinCount: "100000", lockStructure: "1" };
    const out = writeXlsx(wb);
    expect(readXlsx(out).protection).toEqual(wb.protection);
    expect(readXlsx(out).sheets[0].protection).toBeUndefined();
    if (process.env.PROTECTION_OUT) writeFileSync(process.env.PROTECTION_OUT, out);
  }, 60000);

  it("checks the old 16-bit password hash", () => {
    expect(legacyHash("")).toBe("CE4B");
    expect(legacyHash("test")).toBe("CBEB");
  });
});

describe("What a protected sheet allows", () => {
  it("knows tonight's commands and reads objects/scenarios as Excel writes them", async () => {
    const { needs } = await import("../src/protection");
    expect(needs("data.subtotal")).toBe("never");
    expect(needs("data.group")).toBe("never");
    expect(needs("review.spelling")).toBe("never");
    expect(needs("insert.pictureFile")).toBe("objects");
    expect(needs("edit.flashFill")).toBe("unlocked");
    expect(protectionAttrs(new Map([["objects", true], ["scenarios", false]]))).toEqual({ sheet: "1", scenarios: "1" });
  });
});
