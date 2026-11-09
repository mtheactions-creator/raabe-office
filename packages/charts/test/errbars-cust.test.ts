// Benutzerdefinierte Fehlerindikatoren (D-125): the amounts from cells, as Excel's Prognoseblatt writes them for its
// column chart (recorded with probe_excel_forecast.py: c:errValType "cust", c:plus and c:minus on the
// Konfidenzintervall column, line 595959 at alpha 40392). Read, written, refreshed from the cells and drawn.

import { describe, expect, it } from "vitest";
import { chartScene, newChart, refreshCaches, type ChartModel } from "../src/index";
import { chartFromXml, chartToXml } from "../../ooxml/src/xlsx/chart";

const anchor = { col: 0, row: 0, x: 0, y: 0 }, to = { col: 8, row: 15, x: 0, y: 0 };

function model(): ChartModel {
  const amount = { ref: "Tabelle3!$D$2:$D$6", cache: [null, null, null, 1.5, 2] as (number | null)[] };
  return newChart("column", [
    { name: { text: "Besucher" }, val: { ref: "Tabelle3!$B$2:$B$6", cache: [10, 12, 11, null, null] } },
    { name: { text: "Schätzer" }, val: { ref: "Tabelle3!$C$2:$C$6", cache: [null, null, 11, 13, 14] },
      errorBars: { dir: "both", type: "cust", plus: amount, minus: { ...amount, cache: [...amount.cache] }, line: { type: "solid", width: 0.75, color: "#595959", transparency: 59.608 } } },
  ], anchor, to);
}

describe("custom error bars (D-125)", () => {
  it("are written as Excel writes them and read back", () => {
    const xml = chartToXml(model());
    const eb = xml.slice(xml.indexOf("<c:errBars>"), xml.indexOf("</c:errBars>") + 12);
    expect(eb).toBe(
      '<c:errBars><c:errBarType val="both"/><c:errValType val="cust"/><c:noEndCap val="0"/>'
      + '<c:plus><c:numRef><c:f>Tabelle3!$D$2:$D$6</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="5"/><c:pt idx="3"><c:v>1.5</c:v></c:pt><c:pt idx="4"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:plus>'
      + '<c:minus><c:numRef><c:f>Tabelle3!$D$2:$D$6</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="5"/><c:pt idx="3"><c:v>1.5</c:v></c:pt><c:pt idx="4"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:minus>'
      + '<c:spPr><a:noFill/><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="595959"><a:alpha val="40392"/></a:srgbClr></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr></c:errBars>');
    const back = chartFromXml(xml);
    const e = back.series[1].errorBars!;
    expect(e.type).toBe("cust");
    expect(e.plus).toEqual({ ref: "Tabelle3!$D$2:$D$6", cache: [null, null, null, 1.5, 2], format: "General" });
    expect(e.minus?.cache).toEqual([null, null, null, 1.5, 2]);
  });

  it("take their amounts from the cells", () => {
    const m = model();
    refreshCaches(m, (ref) => (ref.includes("$D$") ? { vals: [null, null, null, 3, 4], cols: 1 } : null), (v) => String(v));
    expect(m.series[1].errorBars!.plus!.cache).toEqual([null, null, null, 3, 4]);
    expect(m.series[1].errorBars!.minus!.cache).toEqual([null, null, null, 3, 4]);
  });

  it("are drawn only where the cells give an amount, as long as the amount", () => {
    const m = model();
    const scene = chartScene(m, 600, 300);
    const bars = scene.prims.filter((p) => p.part === "errbars:1") as { d: string }[];
    expect(bars.length).toBe(2);
    // Each bar: a vertical line from value - amount to value + amount, with caps.
    const span = (d: string) => { const ys = [...d.matchAll(/M[\d.]+ ([\d.]+)V([\d.]+)/g)][0]; return Number(ys[1]) - Number(ys[2]); };
    expect(span(bars[1].d) / span(bars[0].d)).toBeCloseTo(2 / 1.5, 2);
  });
});
