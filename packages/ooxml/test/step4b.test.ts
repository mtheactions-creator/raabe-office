import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { Doc, defaultStyleSheet, fromHtml, imageInfo, intern, naturalSizePt, toHtml, addList, bulletLevels } from "@raabe/text-engine";
import { readDocx, writeDocx } from "../src";

describe("emoji (Word's AutoCorrect smileys)", () => {
  it("writes them as Word does (w16se:symEx) and reads them back as text", () => {
    const smile = String.fromCodePoint(0x1f60a);
    const d = new Doc({
      paras: [{ text: `Gut ${smile} `, runs: [{ len: 4, props: intern({}) }, { len: 2, props: intern({ font: "Segoe UI Emoji" }) }, { len: 1, props: intern({}) }], props: {}, mark: intern({}) }],
      styles: defaultStyleSheet(),
    });
    const xml = strFromU8(unzipSync(writeDocx(d.state(), null))["word/document.xml"]);
    expect(xml).toContain('<w16se:symEx w16se:font="Segoe UI Emoji" w16se:char="1F60A"/>');
    expect(xml).toMatch(/xmlns:w16se="[^"]+"/);
    expect(xml).toMatch(/mc:Ignorable="[^"]*w16se/);
    const back = readDocx(writeDocx(d.state(), null)).state.paras[0];
    expect(back.text).toBe(`Gut ${smile} `);
    expect(back.runs[1].props.font).toBe("Segoe UI Emoji");
  });
});

describe("picture headers", () => {
  it("reads the size and resolution of PNG and JPEG files", () => {
    // 2 x 1 PNG with pHYs 19685 px/m (500 dpi).
    const png = new Uint8Array([
      0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 2, 0, 0, 0, 1, 8, 6, 0, 0, 0, 0, 0, 0, 0,
      0, 0, 0, 9, 0x70, 0x48, 0x59, 0x73, 0, 0, 0x4c, 0xe5, 0, 0, 0x4c, 0xe5, 1, 0, 0, 0, 0,
      0, 0, 0, 0, 0x49, 0x44, 0x41, 0x54, 0, 0, 0, 0,
    ]);
    const info = imageInfo(png)!;
    expect(info.px).toEqual({ w: 2, h: 1 });
    expect(Math.round(info.dpi.x)).toBe(500);
    expect(naturalSizePt(info).w).toBeCloseTo(0.288, 2);
    // JPEG with JFIF 300 dpi and a SOF0 of 40 x 30.
    const jpg = new Uint8Array([
      0xff, 0xd8, 0xff, 0xe0, 0, 16, 0x4a, 0x46, 0x49, 0x46, 0, 1, 1, 1, 0x01, 0x2c, 0x01, 0x2c, 0, 0,
      0xff, 0xc0, 0, 17, 8, 0, 30, 0, 40, 3, 1, 0x22, 0, 2, 0x11, 1, 3, 0x11, 1,
    ]);
    const j = imageInfo(jpg)!;
    expect(j.type).toBe("image/jpeg");
    expect(j.px).toEqual({ w: 40, h: 30 });
    expect(j.dpi.x).toBe(300);
  });
});

// The shape of Word's clipboard HTML (style sheet with @list levels, class rules, mso-list
// paragraphs, a table with column widths), shortened.
const WORD_HTML = `<html><head><style>
<!--
p.MsoListParagraphCxSpFirst, li.MsoListParagraphCxSpFirst, div.MsoListParagraphCxSpFirst
	{margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:36.0pt;}
p.MsoListParagraphCxSpMiddle, li.MsoListParagraphCxSpMiddle, div.MsoListParagraphCxSpMiddle
	{margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:36.0pt;}
 /* List Definitions */
 @list l0
	{mso-list-id:1;mso-list-type:hybrid;}
@list l0:level1
	{mso-level-number-format:bullet;
	mso-level-text:@uf0b7;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	text-indent:-18.0pt;
	font-family:Symbol;}
@list l0:level2
	{mso-level-number-format:bullet;
	mso-level-text:o;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	text-indent:-18.0pt;
	font-family:"Courier New";}
@list l1:level1
	{mso-level-tab-stop:none;
	mso-level-number-position:left;
	text-indent:-18.0pt;}
-->
</style></head><body><!--StartFragment-->
<p class=MsoListParagraphCxSpFirst style='text-indent:-18.0pt;mso-list:l0 level1 lfo1'><![if !supportLists]><span style='font-family:Symbol'><span style='mso-list:Ignore'>@u00b7<span style='font:7.0pt "Times New Roman"'>&nbsp; </span></span></span><![endif]>Becherglas<o:p></o:p></p>
<p class=MsoListParagraphCxSpMiddle style='margin-left:72.0pt;text-indent:-18.0pt;mso-list:l0 level2 lfo1'><![if !supportLists]><span style='font-family:"Courier New"'><span style='mso-list:Ignore'>o<span style='font:7.0pt "Times New Roman"'>&nbsp; </span></span></span><![endif]>Leitungswasser<o:p></o:p></p>
<p class=MsoListParagraphCxSpMiddle style='text-indent:-18.0pt;mso-list:l1 level1 lfo2'><![if !supportLists]><span><span style='mso-list:Ignore'>1.<span style='font:7.0pt "Times New Roman"'>&nbsp; </span></span></span><![endif]>Wasser einfüllen<o:p></o:p></p>
<p class=MsoNormal style='margin-left:42.5pt;text-indent:14.2pt'>Eingerückt</p>
<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;border:none'>
 <tr><td width=80 valign=top style='width:60.0pt;border:solid windowtext 1.0pt'><p class=MsoNormal>Z1S1</p></td>
  <td width=267 valign=top style='width:200.0pt;border:solid windowtext 1.0pt'><p class=MsoNormal>Z1S2</p></td></tr>
 <tr style='height:30.0pt'><td width=80 style='width:60.0pt'><p class=MsoNormal>Z2S1</p></td><td width=267 style='width:200.0pt'><p class=MsoNormal>Z2S2</p></td></tr>
</table>
<!--EndFragment--></body></html>`.replace(/@u([0-9a-f]{4})/g, (_, h) => String.fromCharCode(parseInt(h, 16)));

describe("clipboard HTML with Word", () => {
  it("reads Word's list definitions, indents and table widths", () => {
    const { paras, lists } = fromHtml(WORD_HTML, (n) => n);
    const byText = (t: string) => paras.find((p) => p.text === t)!;
    expect(byText("Becherglas")).toMatchObject({ list: "bullet", level: 0, listKey: "l0|1" });
    expect(byText("Leitungswasser")).toMatchObject({ list: "bullet", level: 1, listKey: "l0|1" });
    expect(byText("Wasser einfüllen")).toMatchObject({ list: "number", level: 0, listKey: "l1|2" });
    const l0 = lists.get("l0|1")!;
    expect(l0[0]).toMatchObject({ format: "bullet", text: "\uf0b7", font: "Symbol", indLeft: 36, indFirst: -18 });
    expect(l0[1]).toMatchObject({ text: "o", font: "Courier New", indLeft: 72 });
    expect(lists.get("l1|2")![0]).toMatchObject({ format: "decimal", indLeft: 36 });
    expect(byText("Eingerückt").props).toMatchObject({ indLeft: 42.5, indFirst: 14.2 });
    const cell = byText("Z1S2").cell!;
    expect(cell.table.grid).toEqual([60, 200]);
    expect(cell.table.style).toBe("Table Grid");
    expect(cell.table.rows[1]).toMatchObject({ height: 30, heightRule: "atLeast" });
  });

  it("writes lists and tables the way Word writes them", () => {
    const styles = defaultStyleSheet();
    const made = addList({ abstracts: new Map(), nums: new Map() }, bulletLevels());
    const d = new Doc({
      paras: [{ text: "Punkt", runs: [{ len: 5, props: intern({}) }], props: { numId: made.numId, ilvl: 1 }, mark: intern({}) }],
      styles, numbering: made.numbering,
    });
    const html = toHtml(d.paras, styles, d.numbering);
    expect(html).toContain("mso-list:l0 level2 lfo" + made.numId);
    expect(html).toMatch(/@list l0:level2\s*\{[^}]*mso-level-text:o;[^}]*margin-left:72\.0pt;[^}]*text-indent:-18\.0pt/);
    expect(html).toContain("mso-list:Ignore");
    // Read back: the same list.
    const back = fromHtml(html, (n) => n);
    expect(back.paras[0]).toMatchObject({ list: "bullet", level: 1 });
    expect(back.lists.values().next().value![1]).toMatchObject({ text: "o", font: "Courier New", indLeft: 72, indFirst: -18 });
  });
});
