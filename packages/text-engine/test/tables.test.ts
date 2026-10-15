// Step 4c D: Entf and Rücktaste on selected cells, Zellen löschen, Word's shared-edge rule and
// conditional formats by region (results/2026-10-15-word-table-delete.json for Word's results).
import { describe, expect, it } from "vitest";
import { Doc } from "../src/doc";
import * as E from "../src/edits";
import { tableCellLook } from "../src/flow";
import { EMPTY } from "../src/model";

/** A 3 × 3 table "Z1S1" … "Z3S3" between two paragraphs. */
function table(): Doc {
  const doc = new Doc({});
  doc.insertText({ p: 0, o: 0 }, { p: 0, o: 0 }, "Vor", EMPTY);
  E.insertTable(doc, { p: 0, o: 3 }, 3, 3, 450, undefined, EMPTY);
  for (let i = doc.paras.length - 1; i >= 0; i--) {
    const c = doc.paras[i].cell;
    if (c) doc.insertText({ p: i, o: 0 }, { p: i, o: 0 }, `Z${c.row + 1}S${c.col + 1}`, EMPTY);
  }
  return doc;
}
const cells = (doc: Doc) => doc.paras.filter((p) => p.cell).map((p) => `${p.cell!.row + 1},${p.cell!.col + 1}:${p.text}`).join(" ");
const first = (doc: Doc) => doc.paras.findIndex((p) => p.cell);

describe("Zellen löschen (as recorded in Word)", () => {
  it("Entf empties the selected cells", () => {
    const doc = table();
    E.clearCells(doc, first(doc), { r1: 0, r2: 1, c1: 0, c2: 1 });
    expect(cells(doc)).toBe("1,1: 1,2: 1,3:Z1S3 2,1: 2,2: 2,3:Z2S3 3,1:Z3S1 3,2:Z3S2 3,3:Z3S3");
  });
  it("Zellen nach links verschieben takes the cells out of their rows", () => {
    const doc = table();
    E.deleteCells(doc, first(doc), { r1: 0, r2: 1, c1: 0, c2: 1 }, "left");
    expect(cells(doc)).toBe("1,1:Z1S3 2,1:Z2S3 3,1:Z3S1 3,2:Z3S2 3,3:Z3S3");
  });
  it("Zellen nach oben verschieben moves the contents below up", () => {
    const doc = table();
    E.deleteCells(doc, first(doc), { r1: 0, r2: 1, c1: 0, c2: 1 }, "up");
    expect(cells(doc)).toBe("1,1:Z3S1 1,2:Z3S2 1,3:Z1S3 2,1: 2,2: 2,3:Z2S3 3,1: 3,2: 3,3:Z3S3");
  });
});

describe("Tabellenlayout (Step 4c F, as recorded in Word: results/2026-10-15-word-table-layout.json)", () => {
  const texts = (doc: Doc) => doc.paras.map((p) => (p.cell ? `[${p.cell.row + 1},${p.cell.col + 1}]` : "") + p.text).join("|");
  it("Zellen einfügen: nach rechts makes the row longer", () => {
    const doc = table();
    E.insertCells(doc, first(doc), { r1: 1, r2: 1, c1: 1, c2: 1 }, "right");
    expect(cells(doc)).toBe("1,1:Z1S1 1,2:Z1S2 1,3:Z1S3 2,1:Z2S1 2,2: 2,3:Z2S2 2,4:Z2S3 3,1:Z3S1 3,2:Z3S2 3,3:Z3S3");
  });
  it("Zellen einfügen: nach unten moves the column's contents down into a new row", () => {
    const doc = table();
    E.insertCells(doc, first(doc), { r1: 1, r2: 1, c1: 1, c2: 1 }, "down");
    expect(cells(doc)).toBe("1,1:Z1S1 1,2:Z1S2 1,3:Z1S3 2,1:Z2S1 2,2: 2,3:Z2S3 3,1:Z3S1 3,2:Z2S2 3,3:Z3S3 4,1: 4,2:Z3S2 4,3:");
  });
  it("Tabelle teilen puts an empty paragraph between the halves, or before the table in its first row", () => {
    let doc = table();
    E.splitTable(doc, first(doc) + 3);
    expect(texts(doc)).toBe("Vor|[1,1]Z1S1|[1,2]Z1S2|[1,3]Z1S3||[1,1]Z2S1|[1,2]Z2S2|[1,3]Z2S3|[2,1]Z3S1|[2,2]Z3S2|[2,3]Z3S3|");
    doc = table();
    E.splitTable(doc, first(doc));
    expect(texts(doc).startsWith("Vor||[1,1]Z1S1")).toBe(true);
  });
  it("In Text konvertieren joins each row's cells with the separator", () => {
    const doc = table();
    E.tableToText(doc, first(doc), "	");
    expect(texts(doc)).toBe("Vor|Z1S1	Z1S2	Z1S3|Z2S1	Z2S2	Z2S3|Z3S1	Z3S2	Z3S3|");
  });
});

describe("table style borders (Word's rules)", () => {
  const line = (color: string, size = 0.5) => ({ style: "single", size, color, space: 0 });
  const look = { firstRow: true, lastRow: false, firstColumn: true, lastColumn: false, noHBand: false, noVBand: true };
  it("a conditional format's outer sides belong to its region (the header row's bottom line)", () => {
    const style = { borders: { insideH: line("AAAAAA") }, cond: { firstRow: { borders: { bottom: line("000000", 1.5) } } } };
    expect(tableCellLook(style, look, 0, 4, 1, 1, 4).edges.bottom?.size).toBe(1.5);
    expect(tableCellLook(style, look, 1, 4, 1, 1, 4).edges.top?.color).toBe("AAAAAA");
  });
  it("corner cells come last (Einfache Tabelle 3: no line right of the top left cell)", () => {
    const style = { cond: { firstCol: { borders: { right: line("7F7F7F") } }, nwCell: { borders: { right: { style: "nil", size: 0, color: "auto", space: 0 } } } } };
    expect(tableCellLook(style, look, 0, 4, 0, 1, 4).edges.right?.style).toBe("nil");
    expect(tableCellLook(style, look, 1, 4, 0, 1, 4).edges.right?.color).toBe("7F7F7F");
  });
});
