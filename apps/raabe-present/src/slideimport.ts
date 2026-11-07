// Slides from another presentation (D-072): Start > Neue Folie > Folien wiederverwenden... and pasting slides
// copied in another window. A slide comes with what it refers to (pictures, media, charts, links); it takes the
// layout of the same name in this presentation (PowerPoint's "Zieldesign verwenden"), else Titel und Inhalt.
// Jumps to slides of the other presentation, its notes and comments stay behind.

import { h, showDialog } from "@raabe/shell";
import { PptxDom, slideView, type Pptx } from "@raabe/ooxml";
import type { PresentApp } from "./app";
import type { Deck } from "./deck";
import { renderSlide } from "./render/slide";
import { EMU_PX } from "./render/text";

const { kid } = PptxDom;
const NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const SKIP = ["/slideLayout", "/notesSlide", "/comments", "/slide", "/tags"];

function layoutName(p: Pptx, layout: string): string | undefined {
  return kid(p.doc(layout).documentElement, "cSld")?.getAttribute("name") ?? undefined;
}

/** Copies slide `srcPath` of `src` into `deck` after slide index `at`; returns the new slide's path. */
export function importSlide(deck: Deck, src: Pptx, srcPath: string, at: number): string {
  const srcLayout = src.layoutOf(srcPath);
  const name = srcLayout ? layoutName(src, srcLayout) : undefined;
  const layouts = deck.layouts();
  const layout = layouts.find((l) => l.name === name)?.path ?? layouts.find((l) => l.name === "Titel und Inhalt")?.path ?? layouts[0].path;
  const path = deck.addSlide(at, layout);
  const pptx = deck.pptx;
  deck.change("Folien wiederverwenden", [path], () => {
    const doc = PptxDom.parseXml(PptxDom.serialize(src.doc(srcPath)));
    const map = new Map<string, string>();
    for (const r of src.rels(srcPath).list) {
      if (SKIP.some((t) => r.type.endsWith(t))) continue;
      if (r.external) { map.set(r.id, pptx.rels(path).ensure(r.type, r.target, true)); continue; }
      const t = src.target(srcPath, r.id);
      if (!t || !src.has(t)) continue;
      const ext = t.slice(t.lastIndexOf(".") + 1);
      const file = t.slice(t.lastIndexOf("/") + 1).replace(/\d*\.[^.]+$/, "");
      const np = pptx.freeName(`${t.slice(0, t.lastIndexOf("/") + 1)}${file}#.${ext}`);
      const ct = src.contentType(t);
      if (/\.xml$/i.test(t)) {
        // A chart comes without its workbook (as RaabePresent's own charts, D-071).
        const part = PptxDom.parseXml(PptxDom.serialize(src.doc(t)));
        for (const e of Array.from(part.getElementsByTagNameNS(NS_C, "externalData"))) e.remove();
        pptx.put(np, part, ct);
      } else {
        pptx.put(np, src.files[t]);
        if (ct) pptx.setType(np, ct, true);
      }
      map.set(r.id, pptx.link(path, r.type, np));
    }
    // The slide's references to its parts get the new relationship ids; a reference to what stayed behind goes.
    for (const el of Array.from(doc.getElementsByTagName("*"))) {
      for (const at2 of Array.from(el.attributes)) {
        if (at2.namespaceURI !== PptxDom.NS_R) continue;
        const to = map.get(at2.value);
        if (to) el.setAttributeNS(PptxDom.NS_R, at2.name, to);
        else if (el.localName === "hlinkClick" || el.localName === "hlinkMouseOver") el.remove();
        else if (at2.value) el.setAttributeNS(PptxDom.NS_R, at2.name, "");
      }
    }
    pptx.put(path, doc);
  });
  return path;
}

/** Start > Neue Folie > Folien wiederverwenden...: a presentation's slides to pick from, inserted after the current slide. */
export async function reuseSlidesDialog(app: PresentApp): Promise<void> {
  const file = app.platform.openAnyDialog
    ? await app.platform.openAnyDialog("Durchsuchen", [{ name: "Alle PowerPoint-Präsentationen", extensions: ["pptx", "pptm", "ppsx", "potx"] }])
    : await app.platform.openFileDialog("pptx");
  if (!file) return;
  let src: Pptx;
  try { src = (await import("@raabe/ooxml")).Pptx.read(await app.platform.readFile(file)); } catch { return; }
  const urls: string[] = [];
  const image = (p: string): string | undefined => {
    const bytes = src.files[p];
    if (!bytes) return undefined;
    const u = URL.createObjectURL(new Blob([bytes as BlobPart]));
    urls.push(u);
    return u;
  };
  const size = src.slideSize, tw = 160, s = tw / (size.cx / EMU_PX);
  const picked = new Set<number>();
  const tiles = src.slides.map((sl, i) => {
    const thumb = h("div", { style: `width:${tw}px;height:${(size.cy / EMU_PX) * s}px;overflow:hidden;position:relative;background:#fff` });
    try {
      const r = renderSlide(slideView(src, sl.path), size, { editing: false, image, pptx: src });
      r.el.style.transform = `scale(${s})`; r.el.style.transformOrigin = "0 0";
      thumb.append(r.el);
    } catch { /* a slide we cannot draw is still listed */ }
    const tile = h("button", { class: "reuse-tile", "aria-label": `Folie ${i + 1}`, "aria-pressed": "false", onclick: () => {
      if (picked.has(i)) picked.delete(i); else picked.add(i);
      tile.setAttribute("aria-pressed", String(picked.has(i)));
    } }, thumb, h("span", null, `${i + 1}`));
    return tile;
  });
  const all = h("input", { type: "checkbox" }) as HTMLInputElement;
  const ok = await showDialog({ title: "Folien wiederverwenden", body: h("div", { class: "form" },
    h("div", null, file.split(/[/\\]/).pop() ?? file), h("div", { class: "reuse-grid" }, ...tiles), h("label", { class: "dlg-check" }, all, " Alle Folien einfügen")),
    buttons: [{ label: "Einfügen", value: true, primary: true }, { label: "Schließen", value: false }], cancel: false });
  for (const u of urls) URL.revokeObjectURL(u);
  if (!ok) return;
  const list = all.checked ? src.slides.map((_, i) => i) : [...picked].sort((a, b) => a - b);
  let at = app.slideNumber() - 1;
  for (const i of list) { importSlide(app.deck, src, src.slides[i].path, at); at++; }
  if (list.length) app.setCurrent(at);
}

/**
 * Start > Neue Folie > Folien aus Gliederung...: a Word document's headings (Überschrift 1: a slide's title,
 * Überschrift 2 and lower: its bullet levels), or a text file whose lines start with tabs for the level
 * (no tab: a new slide), as PowerPoint imports an outline. The slides come after the current one in the
 * layout Titel und Inhalt.
 */
export async function outlineDialog(app: PresentApp): Promise<void> {
  const file = app.platform.openAnyDialog
    ? await app.platform.openAnyDialog("Gliederung einfügen", [{ name: "Alle Gliederungen", extensions: ["docx", "txt"] }])
    : null;
  if (!file) return;
  const bytes = await app.platform.readFile(file);
  const items: { level: number; text: string }[] = [];
  if (/\.docx$/i.test(file)) {
    const { readDocx } = await import("@raabe/ooxml");
    const st = readDocx(bytes).state;
    for (const p of st.paras) {
      const lvl = p.props.outline ?? st.styles.outlineLevel(st.styles.resolvePara(p.props).style);
      const text = p.text.replace(/[\u000b\u000c￼]/g, " ").trim();
      if (lvl !== undefined && lvl <= 8 && text) items.push({ level: lvl, text });
    }
  } else {
    for (const line of new TextDecoder().decode(bytes).split(/\r?\n/)) {
      const m = /^(\t*)(.*)$/.exec(line)!;
      if (m[2].trim()) items.push({ level: m[1].length, text: m[2].trim() });
    }
  }
  if (!items.length) return;
  const deck = app.deck;
  const layout = deck.layouts().find((l) => l.name === "Titel und Inhalt")?.path ?? deck.layouts()[1]?.path;
  let at = app.slideNumber() - 1;
  const slides: { title: string; body: { level: number; text: string }[] }[] = [];
  for (const it of items) {
    if (it.level === 0 || !slides.length) slides.push({ title: it.level === 0 ? it.text : "", body: it.level === 0 ? [] : [{ level: 0, text: it.text }] });
    else slides[slides.length - 1].body.push({ level: Math.min(8, it.level - 1), text: it.text });
  }
  const T = await import("./textops");
  const S = await import("./shapeops");
  for (const s of slides) {
    const path = deck.addSlide(at, layout);
    at++;
    deck.change("Folien aus Gliederung", [path], () => {
      const tree = S.treeOf(deck.pptx.doc(path));
      const bodies = S.shapeEls(tree).map((e) => ({ e, ph: e.getElementsByTagNameNS(PptxDom.NS_P, "ph")[0] })).filter((x) => x.ph);
      const titleEl = bodies.find((x) => ["title", "ctrTitle"].includes(x.ph.getAttribute("type") ?? ""))?.e;
      const bodyEl = bodies.find((x) => !["title", "ctrTitle", "dt", "ftr", "sldNum"].includes(x.ph.getAttribute("type") ?? ""))?.e;
      const tb = titleEl ? S.txBodyOf(titleEl) : undefined;
      if (tb && s.title) T.insertText(tb, { p: 0, o: 0 }, s.title);
      const bb = bodyEl ? S.txBodyOf(bodyEl) : undefined;
      if (bb && s.body.length) {
        // One paragraph per line of the outline, at its level.
        let pos = { p: 0, o: 0 };
        s.body.forEach((b, i) => { if (i) pos = T.splitPara(bb, pos); pos = T.insertText(bb, pos, b.text); if (b.level) T.paraProps(bb, pos.p, pos.p, (pr) => pr.setAttribute("lvl", String(b.level))); });
      }
    });
  }
  app.setCurrent(at);
}
