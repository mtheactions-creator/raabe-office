import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { Doc, builtinStyle, defaultStyleSheet, makePara, validPara } from "@raabe/text-engine";
import * as E from "../../text-engine/src/edits";
import { readDocx, writeDocx } from "../src";

const sheet = () => {
  let s = defaultStyleSheet();
  for (const n of ["Table Grid", "List Paragraph", "header", "footer"]) { const b = builtinStyle(n); if (b) s = s.withStyle(b); }
  return s;
};

describe("tables", () => {
  it("inserts a table, edits rows and columns, and survives a round trip", () => {
    const d = new Doc({ paras: [makePara("Vorher"), makePara("")], styles: sheet() });
    const at = E.insertTable(d, { p: 1, o: 0 }, 2, 3, 450, "Tabellenraster", {});
    expect(at).toEqual({ p: 1, o: 0 });
    expect(d.paras.filter((p) => p.cell).length).toBe(6);
    expect(d.paras[d.paras.length - 1].cell).toBeUndefined();
    d.insertText({ p: 1, o: 0 }, { p: 1, o: 0 }, "A1", {});
    E.insertRows(d, 1, 1, "below");
    let def = d.paras[1].cell!.table;
    expect(def.rows.length).toBe(3);
    E.insertColumns(d, 1, 1, "right");
    def = d.paras[1].cell!.table;
    expect(def.grid.length).toBe(4);
    expect(def.grid.reduce((s, x) => s + x, 0)).toBeCloseTo(450, 5);
    expect(d.paras[1].text).toBe("A1");
    E.deleteColumns(d, 1, 1);
    def = d.paras[1].cell!.table;
    expect(def.grid.length).toBe(3);
    expect(d.paras.every(validPara)).toBe(true);
    const back = readDocx(writeDocx(d.state(), null)).state;
    const cells = back.paras.filter((p) => p.cell);
    expect(cells.length).toBe(9);
    expect(cells[0].cell!.table.rows.length).toBe(3);
    expect(cells[0].cell!.table.style).toBe("Tabellenraster");
  });

  it("merges cells horizontally and vertically", () => {
    const d = new Doc({ paras: [makePara("")], styles: sheet() });
    E.insertTable(d, { p: 0, o: 0 }, 3, 3, 300, "Tabellenraster", {});
    d.insertText({ p: 0, o: 0 }, { p: 0, o: 0 }, "a", {});
    d.insertText({ p: 1, o: 0 }, { p: 1, o: 0 }, "b", {});
    // Cells (0,0) and (0,1): one cell with both texts, spanning two grid columns.
    E.mergeCells(d, 0, 1);
    let def = d.paras[0].cell!.table;
    expect(def.rows[0].cells.length).toBe(2);
    expect(def.rows[0].cells[0].span).toBe(2);
    expect(d.paras.filter((p) => p.cell?.row === 0 && p.cell.col === 0).map((p) => p.text)).toEqual(["a", "b"]);
    // Vertical: (1,2) with (2,2).
    const p12 = d.paras.findIndex((p) => p.cell?.row === 1 && p.cell.col === 2);
    const p22 = d.paras.findIndex((p) => p.cell?.row === 2 && p.cell.col === 2);
    E.mergeCells(d, p12, p22);
    def = d.paras[0].cell!.table;
    expect(def.rows[1].cells[2].vMerge).toBe("restart");
    expect(def.rows[2].cells[2].vMerge).toBe("continue");
    const x = strFromU8(unzipSync(writeDocx(d.state(), null))["word/document.xml"]);
    expect(x).toContain('<w:gridSpan w:val="2"/>');
    expect(x).toContain('<w:vMerge w:val="restart"/>');
  });
});

describe("lists", () => {
  it("bullets on and off, levels, and numbering written to numbering.xml", () => {
    const d = new Doc({ paras: [makePara("Eins"), makePara("Zwei"), makePara("Drei")], styles: sheet() });
    E.toggleList(d, 0, 2, "bullet", "Listenabsatz");
    expect(d.paras.every((p) => p.props.numId === d.paras[0].props.numId && p.props.style === "Listenabsatz")).toBe(true);
    E.changeListLevel(d, 1, 1, 1);
    expect(d.paras[1].props.ilvl).toBe(1);
    E.toggleList(d, 0, 2, "bullet", "Listenabsatz");
    expect(d.paras.every((p) => p.props.numId === undefined && p.props.style === undefined)).toBe(true);
    E.toggleList(d, 0, 1, "number", "Listenabsatz");
    // The next paragraph joins the list above it.
    E.toggleList(d, 2, 2, "number", "Listenabsatz");
    expect(d.paras[2].props.numId).toBe(d.paras[0].props.numId);
    const out = unzipSync(writeDocx(d.state(), null));
    expect(Object.keys(out)).toContain("word/numbering.xml");
    const back = readDocx(writeDocx(d.state(), null)).state;
    expect(back.paras.map((p) => p.props.numId)).toEqual([1, 1, 1].map(() => back.paras[0].props.numId));
    expect(back.numbering?.abstracts.size).toBeGreaterThan(0);
  });

  it("recognises AutoFormat prefixes", () => {
    expect(E.autoListLevels("1.")?.levels[0].text).toBe("%1.");
    expect(E.autoListLevels("3)")?.start).toBe(3);
    expect(E.autoListLevels("a)")?.levels[0].format).toBe("lowerLetter");
    expect(E.autoListLevels("*")?.levels[0].format).toBe("bullet");
    expect(E.autoListLevels("Hallo")).toBeNull();
  });
});

describe("sections, headers and footers", () => {
  it("writes section breaks, landscape pages, columns and a new header with a page number", () => {
    const d = new Doc({ paras: [makePara("Hochformat"), makePara("Quer")], styles: sheet() });
    E.insertSectionBreak(d, { p: 0, o: 10 }, "nextPage", {});
    const sects = d.main.filter((p) => p.sect);
    expect(sects.length).toBe(1);
    const s = d.settings.section;
    d.setSettings({ ...d.settings, section: { ...s, page: { ...s.page, width: s.page.height, height: s.page.width }, columns: { count: 2, space: 35.4, sep: true } } });
    d.setStory("hdr:1", [{ ...makePara("Kopf", { style: "Kopfzeile" }), objects: undefined }]);
    d.setStory("ftr:1", [{
      ...makePara("", { style: "Fuzeile", align: "center" }), text: "\ufffc", runs: [{ len: 1, props: {} }],
      objects: [E.fieldObject("PAGE", "1")],
    }]);
    d.setSettings({ ...d.settings, section: { ...d.settings.section, headers: { default: "hdr:1" }, footers: { default: "ftr:1" } } });
    const bytes = writeDocx(d.state(), null);
    const files = unzipSync(bytes);
    expect(Object.keys(files).some((k) => /^word\/header\d+\.xml$/.test(k))).toBe(true);
    expect(strFromU8(files["word/[Content_Types].xml".replace("word/", "")])).toContain("header+xml");
    const x = strFromU8(files["word/document.xml"]);
    expect(x).toContain('w:orient="landscape"');
    expect(x).toContain('w:num="2"');
    const back = readDocx(bytes).state;
    expect(back.paras.filter((p) => p.sect).length).toBe(1);
    expect(back.settings.section.columns.count).toBe(2);
    const hid = back.settings.section.headers.default!;
    expect(back.stories?.get(hid)?.[0].text).toBe("Kopf");
    const fid = back.settings.section.footers.default!;
    expect(back.stories?.get(fid)?.[0].objects?.[0].instr).toBe("PAGE");
  });
});
