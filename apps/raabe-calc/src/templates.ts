// Raabeschule templates for RaabeCalc (shared shell scope: Neu with school templates; like
// RaabeWrite's D-035, our own designs from ordinary Excel features, so files from them work in
// Excel): a Notenliste, a Notenspiegel with a chart and a table of Messwerte with a trendline.
// The school can replace them without code changes beyond this file.

import { newChart } from "@raabe/charts";
import { Workbook, parseRange, type Sheet } from "@raabe/sheet-engine";
import { newWorkbook } from "@raabe/ooxml";
import * as A from "./actions";
import { History } from "./ops";

export type CalcTemplate = "notenliste" | "notenspiegel" | "messwerte";

export const CALC_TEMPLATES: { id: CalcTemplate; name: string }[] = [
  { id: "notenliste", name: "Notenliste" },
  { id: "notenspiegel", name: "Notenspiegel" },
  { id: "messwerte", name: "Messwerte" },
];

const RED = "#B3141B";

function fill(wb: Workbook, sheet: Sheet, cells: Record<string, string>): void {
  const history = new History(wb);
  const ctx = { wb, history };
  history.step("Vorlage", () => {
    for (const [a, t] of Object.entries(cells)) A.enter(ctx, sheet, parseRange(a)!, t);
  });
}

function style(wb: Workbook, sheet: Sheet, range: string, change: Parameters<typeof A.format>[3]): void {
  const history = new History(wb);
  history.step("Vorlage", () => A.format({ wb, history }, sheet, [parseRange(range)!], change));
}

/** A new workbook with the template's sheets. */
export function templateWorkbook(id: CalcTemplate): Workbook {
  const wb = newWorkbook();
  const sheet = wb.sheet();
  if (id === "notenliste") {
    sheet.name = "Notenliste";
    const cells: Record<string, string> = {
      A1: "Raabeschule – Notenliste", A2: "Klasse:", B2: "[Klasse]", D2: "Fach:", E2: "[Fach]",
      A4: "Name", B4: "Klassenarbeit 1", C4: "Klassenarbeit 2", D4: "Klassenarbeit 3", E4: "Mündlich", F4: "Schriftlich", G4: "Gesamt",
    };
    for (let i = 0; i < 25; i++) {
      const r = 5 + i;
      cells[`A${r}`] = `[Name ${i + 1}]`;
      cells[`F${r}`] = `=WENNFEHLER(MITTELWERT(B${r}:D${r});"")`;
      // Schriftlich and mündlich zählen je zur Hälfte; the school can change the weighting here.
      cells[`G${r}`] = `=WENNFEHLER(RUNDEN((F${r}+E${r})/2;1);"")`;
    }
    cells.A31 = "Durchschnitt";
    for (const c of ["B", "C", "D", "E", "F", "G"]) cells[`${c}31`] = `=WENNFEHLER(RUNDEN(MITTELWERT(${c}5:${c}29);2);"")`;
    fill(wb, sheet, cells);
    style(wb, sheet, "A1", (s) => { s.font.size = 16; s.font.bold = true; s.font.color = RED; });
    style(wb, sheet, "A4:G4", (s) => { s.font.bold = true; s.fill = "#F2F2F2"; s.borders = { bottom: { style: "medium", color: RED } }; s.wrap = true; });
    style(wb, sheet, "B5:G31", (s) => { s.numFmt = "0.0"; s.h = "center"; });
    style(wb, sheet, "A31:G31", (s) => { s.font.bold = true; s.borders = { top: { style: "thin" } }; });
    sheet.cols.set(0, { width: 24 });
    for (let c = 1; c <= 6; c++) sheet.cols.set(c, { width: 13 });
    sheet.rows.set(3, { height: 30 });
    sheet.view.freeze = { row: 4, col: 1 };
    sheet.conditional = [{ ref: [parseRange("G5:G29")!], rules: [{ type: "cellIs", operator: "greaterThan", formulas: ["4.4"], priority: 1, format: { fill: "#FFC7CE", color: "#9C0006" } }] }];
  } else if (id === "notenspiegel") {
    sheet.name = "Notenspiegel";
    const cells: Record<string, string> = { A1: "Raabeschule – Notenspiegel", A2: "Klassenarbeit:", B2: "[Thema]", A4: "Note", B4: "Anzahl", D4: "Noten der Schülerinnen und Schüler" };
    for (let n = 1; n <= 6; n++) { cells[`A${4 + n}`] = String(n); cells[`B${4 + n}`] = `=ZÄHLENWENN($D$5:$D$40;A${4 + n})`; }
    Object.assign(cells, { A12: "Teilnehmer", B12: "=SUMME(B5:B10)", A13: "Durchschnitt", B13: '=WENNFEHLER(RUNDEN(MITTELWERT(D5:D40);2);"")', A14: "Unter 4", B14: '=WENNFEHLER(ZÄHLENWENN(D5:D40;">4")/ANZAHL(D5:D40);"")' });
    fill(wb, sheet, cells);
    style(wb, sheet, "A1", (s) => { s.font.size = 16; s.font.bold = true; s.font.color = RED; });
    style(wb, sheet, "A4:D4", (s) => { s.font.bold = true; s.fill = "#F2F2F2"; s.borders = { bottom: { style: "medium", color: RED } }; });
    style(wb, sheet, "A5:B10", (s) => { s.h = "center"; });
    style(wb, sheet, "A12:B14", (s) => { s.font.bold = true; });
    style(wb, sheet, "B14", (s) => { s.numFmt = "0%"; });
    sheet.cols.set(0, { width: 14 });
    sheet.cols.set(3, { width: 32 });
    const chart = newChart("column", [{ name: { ref: "Notenspiegel!$B$4" }, cat: { ref: "Notenspiegel!$A$5:$A$10", cache: [] }, val: { ref: "Notenspiegel!$B$5:$B$10", cache: [] } }],
      { col: 5, row: 3, x: 0, y: 0 }, { col: 11, row: 18, x: 0, y: 0 });
    chart.title = { show: true, text: "Notenspiegel" };
    chart.legend = "none";
    sheet.charts = [chart];
  } else {
    sheet.name = "Messwerte";
    const cells: Record<string, string> = {
      A1: "Raabeschule – Messwerte", A2: "Versuch:", B2: "[Versuch]",
      A4: "Zeit t in s", B4: "Strecke s in m",
      D4: "Steigung", E4: "=WENNFEHLER(STEIGUNG(B5:B24;A5:A24);\"\")", D5: "Achsenabschnitt", E5: "=WENNFEHLER(ACHSENABSCHNITT(B5:B24;A5:A24);\"\")",
      D6: "Bestimmtheitsmaß", E6: "=WENNFEHLER(BESTIMMTHEITSMASS(B5:B24;A5:A24);\"\")",
    };
    const sample = [[0, 0], [1, 1.9], [2, 4.1], [3, 6], [4, 8.2], [5, 9.9]];
    sample.forEach(([t, s], i) => { cells[`A${5 + i}`] = String(t).replace(".", ","); cells[`B${5 + i}`] = String(s).replace(".", ","); });
    fill(wb, sheet, cells);
    style(wb, sheet, "A1", (s) => { s.font.size = 16; s.font.bold = true; s.font.color = RED; });
    style(wb, sheet, "A4:B4", (s) => { s.font.bold = true; s.fill = "#F2F2F2"; s.borders = { bottom: { style: "medium", color: RED } }; });
    style(wb, sheet, "D4:D6", (s) => { s.font.bold = true; });
    style(wb, sheet, "E4:E6", (s) => { s.numFmt = "0.000"; });
    sheet.cols.set(0, { width: 14 });
    sheet.cols.set(1, { width: 16 });
    sheet.cols.set(3, { width: 18 });
    const chart = newChart("scatter", [{ name: { ref: "Messwerte!$B$4" }, cat: { ref: "Messwerte!$A$5:$A$24", cache: [] }, val: { ref: "Messwerte!$B$5:$B$24", cache: [] },
      trendline: { type: "linear", showEq: true, showR2: true } }], { col: 6, row: 3, x: 0, y: 0 }, { col: 12, row: 18, x: 0, y: 0 });
    chart.title = { show: true, text: "s-t-Diagramm" };
    chart.axisTitles = { cat: "t in s", val: "s in m" };
    sheet.charts = [chart];
  }
  return wb;
}
