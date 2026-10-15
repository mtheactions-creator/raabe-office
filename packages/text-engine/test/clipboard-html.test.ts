// Step 4c F: pictures and hyperlinks through the clipboard's HTML (what Word reads and writes).
import { describe, expect, it } from "vitest";
import { fromHtml, toHtml } from "../src/html";
import { OBJ, makePara, type Para } from "../src/model";
import { defaultStyleSheet } from "../src/styles";
import { imageObject } from "../src/edits";

// A 1 x 1 PNG.
const PNG = Uint8Array.from(atob("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="), (c) => c.charCodeAt(0));

describe("clipboard HTML for other programs", () => {
  const styles = defaultStyleSheet();
  it("writes pictures as data: images in pixels and links as <a href>", () => {
    const pic: Para = { ...makePara(`Bild ${OBJ} und Link`), objects: [imageObject({ bytes: PNG, type: "image/png" }, 72, 36)] };
    const link: Para = {
      ...makePara("Ein Link hier"),
      anchors: [{ at: 4, data: undefined, link: { id: "link:1", edge: "start", href: "https://www.example.org/" } }, { at: 8, data: undefined, link: { id: "link:1", edge: "end" } }],
    };
    const html = toHtml([pic, link], styles);
    expect(html).toMatch(/<img width=96 height=48 src="data:image\/png;base64,iVBOR/);
    expect(html).toMatch(/<a href="https:\/\/www\.example\.org\/"><span[^>]*>Link<\/span><\/a>/);
  });

  it("reads <a href> as a link and <img> through the picture callback", () => {
    const html = '<p>Ein <a href="https://www.example.org/"><span style="color:#0563C1">Link</span></a> und <img width=96 height=48 src="data:image/png;base64,AAAA"> Bild</p>';
    const seen: [string, number | undefined, number | undefined][] = [];
    const { paras } = fromHtml(html, (n) => (n === "Hyperlink" ? "Hyperlink" : undefined), undefined, (src, w, h) => {
      seen.push([src.slice(0, 15), w, h]);
      return { kind: "image", data: undefined, display: "", width: (w ?? 0) * 0.75, height: (h ?? 0) * 0.75 };
    });
    const p = paras[0];
    expect(p.text).toBe(`Ein Link und ${OBJ} Bild`);
    expect(seen).toEqual([["data:image/png;", 96, 48]]);
    expect(p.objects?.[0].width).toBe(72);
    const start = p.anchors?.find((a) => a.link?.edge === "start");
    const end = p.anchors?.find((a) => a.link?.edge === "end");
    expect(start?.link?.href).toBe("https://www.example.org/");
    expect([start?.at, end?.at]).toEqual([4, 8]);
    // The link's text has the character style Hyperlink.
    let pos = 0;
    const styled = p.runs.find((r) => { const hit = pos <= 4 && 4 < pos + r.len; pos += r.len; return hit; });
    expect(styled?.props.rStyle).toBe("Hyperlink");
  });
});
