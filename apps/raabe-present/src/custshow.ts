// Bildschirmpräsentation > Benutzerdefinierte Bildschirmpräsentation (D-072): PowerPoint's custom shows
// (p:custShowLst in presentation.xml, each a name and its slides by relationship), listed in the menu to run,
// and Zielgruppenorientierte Präsentationen... to make, change and remove them (our dialog with PowerPoint's
// labels from knowledge). A custom show plays only its slides, in its order.

import { h, showDialog, type MenuEntry } from "@raabe/shell";
import { PptxDom } from "@raabe/ooxml";
import type { PresentApp } from "./app";

const { kid, kids } = PptxDom;
const RT_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

export interface CustomShow { name: string; slides: number[] }

/** The custom shows with their slides as indices of the presentation's slides. */
export function customShows(app: PresentApp): CustomShow[] {
  const pptx = app.deck.pptx, pres = pptx.presPath;
  const lst = kid(pptx.doc(pres).documentElement, "custShowLst");
  return kids(lst, "custShow").map((cs) => ({
    name: cs.getAttribute("name") ?? "",
    slides: kids(kid(cs, "sldLst"), "sld").map((s) => {
      const t = pptx.target(pres, s.getAttributeNS(PptxDom.NS_R, "id") ?? "");
      return app.deck.slides.findIndex((x) => x.path === t);
    }).filter((i) => i >= 0),
  }));
}

export function writeShows(app: PresentApp, shows: CustomShow[]): void {
  const pptx = app.deck.pptx, pres = pptx.presPath;
  app.deck.change("Zielgruppenorientierte Präsentationen", [pres], () => {
    const root = pptx.doc(pres).documentElement;
    kid(root, "custShowLst")?.remove();
    if (!shows.length) return;
    const doc = root.ownerDocument;
    const lst = PptxDom.make(doc, "p:custShowLst");
    shows.forEach((s, id) => {
      const cs = PptxDom.make(doc, "p:custShow", { name: s.name, id });
      const sl = PptxDom.make(doc, "p:sldLst");
      for (const i of s.slides) {
        const el = PptxDom.make(doc, "p:sld");
        el.setAttributeNS(PptxDom.NS_R, "r:id", pptx.link(pres, RT_SLIDE, app.deck.slides[i].path));
        sl.append(el);
      }
      cs.append(sl);
      lst.append(cs);
    });
    // Schema order: after notesSz (and smartTags), before photoAlbum, custDataLst, kinsoku …
    const after = kids(root).find((c) => ["photoAlbum", "custDataLst", "kinsoku", "defaultTextStyle", "modifyVerifier", "extLst"].includes(c.localName)) ?? null;
    root.insertBefore(lst, after);
  });
}

async function editShow(app: PresentApp, show: CustomShow | null, names: string[]): Promise<CustomShow | null> {
  const name = h("input", { type: "text", value: show?.name ?? `Benutzerdefinierte Präsentation ${names.length + 1}`, style: "width:260px" }) as HTMLInputElement;
  const title = (i: number) => {
    const v = app.deck.view(app.deck.slides[i].path);
    const t = v.shapes.find((s) => s.ph && (s.ph.type === "title" || s.ph.type === "ctrTitle"));
    return `${i + 1}. ${t?.text ? t.text.paras.map((p) => p.runs.map((r) => r.text).join("")).join(" ") : `Folie ${i + 1}`}`;
  };
  const all = h("select", { size: "10", multiple: true, style: "width:220px" }, ...app.deck.slides.map((_, i) => h("option", { value: String(i) }, title(i)))) as HTMLSelectElement;
  const chosen = h("select", { size: "10", style: "width:220px" }) as HTMLSelectElement;
  const list = [...(show?.slides ?? [])];
  const redraw = () => chosen.replaceChildren(...list.map((i, k) => h("option", { value: String(k) }, title(i))));
  redraw();
  const add = h("button", { onclick: () => { for (const o of Array.from(all.selectedOptions)) list.push(Number(o.value)); redraw(); } }, "Hinzufügen >>");
  const remove = h("button", { onclick: () => { const k = chosen.selectedIndex; if (k >= 0) { list.splice(k, 1); redraw(); } } }, "Entfernen");
  const up = h("button", { onclick: () => { const k = chosen.selectedIndex; if (k > 0) { [list[k - 1], list[k]] = [list[k], list[k - 1]]; redraw(); chosen.selectedIndex = k - 1; } } }, "▲");
  const down = h("button", { onclick: () => { const k = chosen.selectedIndex; if (k >= 0 && k < list.length - 1) { [list[k + 1], list[k]] = [list[k], list[k + 1]]; redraw(); chosen.selectedIndex = k + 1; } } }, "▼");
  const ok = await showDialog({ title: "Zielgruppenorientierte Präsentation definieren", body: h("div", { class: "form" },
    h("label", null, "Name der Zielgruppenorientierten Präsentation: ", name),
    h("div", { style: "display:flex;gap:8px;align-items:center" },
      h("div", null, h("div", null, "Folien in der Präsentation:"), all), h("div", { style: "display:flex;flex-direction:column;gap:4px" }, add, remove),
      h("div", null, h("div", null, "Folien in der Zielgruppenorientierten Präsentation:"), chosen), h("div", { style: "display:flex;flex-direction:column;gap:4px" }, up, down))),
    buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
  if (!ok || !list.length) return null;
  return { name: name.value.trim() || "Benutzerdefinierte Präsentation", slides: list };
}

/** Zielgruppenorientierte Präsentationen...: the list with Neu, Bearbeiten, Entfernen, Kopieren, Vorführen. */
export async function customShowsDialog(app: PresentApp): Promise<void> {
  const shows = customShows(app);
  const sel = h("select", { size: "8", style: "width:280px" }) as HTMLSelectElement;
  const redraw = () => sel.replaceChildren(...shows.map((s, i) => h("option", { value: String(i) }, s.name)));
  redraw();
  let play: number | null = null;
  const btn = (label: string, run: () => void | Promise<void>) => h("button", { onclick: () => void run() }, label);
  const body = h("div", { class: "form", style: "display:flex;gap:10px" }, h("div", null, h("div", null, "Zielgruppenorientierte Präsentationen:"), sel),
    h("div", { style: "display:flex;flex-direction:column;gap:4px" },
      btn("Neu...", async () => { const s = await editShow(app, null, shows.map((x) => x.name)); if (s) { shows.push(s); redraw(); } }),
      btn("Bearbeiten...", async () => { const i = sel.selectedIndex; if (i < 0) return; const s = await editShow(app, shows[i], shows.map((x) => x.name)); if (s) { shows[i] = s; redraw(); } }),
      btn("Entfernen", () => { const i = sel.selectedIndex; if (i >= 0) { shows.splice(i, 1); redraw(); } }),
      btn("Kopieren", () => { const i = sel.selectedIndex; if (i >= 0) { shows.push({ name: `Kopie von ${shows[i].name}`, slides: [...shows[i].slides] }); redraw(); } })));
  const r = await showDialog({ title: "Zielgruppenorientierte Präsentationen", body, buttons: [{ label: "Vorführen", value: "show" }, { label: "Schließen", value: "close", primary: true }], cancel: "close" });
  play = r === "show" ? sel.selectedIndex : null;
  writeShows(app, shows);
  if (play !== null && play >= 0 && shows[play]) void app.startCustomShow(shows[play].slides);
}

export function customShowMenu(app: PresentApp): MenuEntry[] {
  return [...customShows(app).map((s) => ({ label: s.name, run: () => void app.startCustomShow(s.slides) })),
    ...(customShows(app).length ? ["-" as const] : []), { label: "Zielgruppenorientierte Präsentationen...", run: () => void customShowsDialog(app) }];
}
