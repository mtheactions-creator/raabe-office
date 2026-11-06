// Einfügen > Diagramm (D-070): a chart part with PowerPoint's sample data (Kategorie 1-4,
// Datenreihe 1-3), written by the chart engine's writer shared with RaabeCalc (D-062) with its
// values cached, so it draws without a workbook. The embedded workbook PowerPoint adds for
// "Daten bearbeiten" is not written yet (the chart cannot be edited in PowerPoint's Excel).

import { chartToXml, chartStyleXml, chartColorsXml } from "@raabe/ooxml";
import { newChart, type ChartKind } from "@raabe/charts";
import type { PresentApp } from "./app";
import * as S from "./shapeops";

const CT_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const RT_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";

export function insertChart(app: PresentApp, kind: string): void {
  const slide = app.slidePath();
  if (!slide) return;
  const cats = ["Kategorie 1", "Kategorie 2", "Kategorie 3", "Kategorie 4"];
  const data = [[4.3, 2.5, 3.5, 4.5], [2.4, 4.4, 1.8, 2.8], [2, 2, 3, 5]];
  const col = (i: number) => String.fromCharCode(66 + i);
  const series = (kind === "pie" ? data.slice(0, 1) : data).map((vals, i) => ({
    name: { ref: `Tabelle1!$${col(i)}$1`, text: `Datenreihe ${i + 1}` },
    cat: { ref: "Tabelle1!$A$2:$A$5", cache: cats },
    val: { ref: `Tabelle1!$${col(i)}$2:$${col(i)}$5`, cache: vals },
  }));
  const model = newChart(kind as ChartKind, series, { col: 0, row: 0, x: 0, y: 0 }, { col: 8, row: 15, x: 0, y: 0 });
  const pptx = app.deck.pptx;
  const part = pptx.freeName("ppt/charts/chart#.xml");
  const n = part.match(/(\d+)\.xml$/)![1];
  const size = pptx.slideSize;
  let id = 0;
  app.editor.edit("Diagramm einfügen", () => {
    pptx.put(part, chartToXml(model), CT_CHART);
    pptx.put(`ppt/charts/style${n}.xml`, chartStyleXml(), "application/vnd.ms-office.chartstyle+xml");
    pptx.put(`ppt/charts/colors${n}.xml`, chartColorsXml(), "application/vnd.ms-office.chartcolorstyle+xml");
    const rels = pptx.rels(part);
    rels.ensure("http://schemas.microsoft.com/office/2011/relationships/chartStyle", `style${n}.xml`);
    rels.ensure("http://schemas.microsoft.com/office/2011/relationships/chartColorStyle", `colors${n}.xml`);
    const rid = pptx.link(slide, RT_CHART, part);
    const doc = pptx.doc(slide);
    const tree = S.treeOf(doc);
    id = S.maxId(tree) + 1;
    // PowerPoint's place for a new chart: the content area of the Titel-und-Inhalt layout.
    tree.append(S.newChartFrame(doc, id, rid, { x: 838200, y: 1825625, w: size.cx - 2 * 838200, h: 4351338 }));
  });
  app.editor.select([id]);
}
