// Link einfügen (Strg+K, D-070): a web address or file, a slide of this presentation ("Aktuelles
// Dokument") or an e-mail address, on the selected text (or the whole shape's text), written as
// PowerPoint writes it: a:hlinkClick with an external relationship, or with
// action="ppaction://hlinksldjump" and a relationship to the slide. Followed in the show.

import { h, showDialog } from "@raabe/shell";
import { PptxDom } from "@raabe/ooxml";
import type { PresentApp } from "./app";
import * as T from "./textops";

const RT_LINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const RT_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

export async function insertLink(app: PresentApp): Promise<void> {
  const ed = app.editor;
  const id = ed.text?.id ?? ed.sel[0];
  if (id === undefined) return;
  const body = ed.textBody(id);
  const t = ed.text;
  const selected = t && body && T.cmpPos(t.caret, t.anchor) ? T.rangeText(body, t.anchor, t.caret) : "";
  const shown = h("input", { type: "text", value: selected, style: "width:320px" }) as HTMLInputElement;
  const addr = h("input", { type: "text", value: "", style: "width:320px" }) as HTMLInputElement;
  const slides = h("select", { size: 8, style: "width:320px" }, ...app.deck.slides.map((s, i) => h("option", { value: s.path }, `${i + 1}. ${slideTitle(app, s.path) || `Folie ${i + 1}`}`))) as HTMLSelectElement;
  let kind: "web" | "slide" | "mail" = "web";
  const pageWeb = h("div", null, h("label", { class: "dlg-row" }, h("span", { class: "dlg-label" }, "Adresse:"), addr));
  const pageSlide = h("div", { hidden: true }, h("div", null, "Wählen Sie eine Stelle im Dokument:"), slides);
  const mail = h("input", { type: "text", style: "width:320px" }) as HTMLInputElement;
  const pageMail = h("div", { hidden: true }, h("label", { class: "dlg-row" }, h("span", { class: "dlg-label" }, "E-Mail-Adresse:"), mail));
  const nav = h("div", { class: "link-nav" });
  const pick = (k: typeof kind, btn: HTMLElement) => { kind = k; pageWeb.hidden = k !== "web"; pageSlide.hidden = k !== "slide"; pageMail.hidden = k !== "mail"; for (const b of Array.from(nav.children)) b.classList.toggle("on", b === btn); };
  for (const [k, label] of [["web", "Datei oder Webseite"], ["slide", "Aktuelles Dokument"], ["mail", "E-Mail-Adresse"]] as const) {
    const b = h("button", { class: k === "web" ? "on" : "", onclick: () => pick(k, b) }, label);
    nav.append(b);
  }
  const ok = await showDialog({ title: "Link einfügen", body: h("div", { class: "link-dialog" }, h("label", { class: "dlg-row" }, h("span", { class: "dlg-label" }, "Anzuzeigender Text:"), shown), h("div", { class: "link-cols" }, nav, h("div", null, pageWeb, pageSlide, pageMail))),
    buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
  if (!ok) return;
  const path = app.slidePath()!;
  const pptx = app.deck.pptx;
  let target = "", action: string | undefined, external = true;
  if (kind === "web") { target = addr.value.trim(); if (target && !/^[a-z]+:/i.test(target) && !/^[A-Za-z]:\\/.test(target)) target = `http://${target}`; }
  else if (kind === "mail") target = `mailto:${mail.value.trim()}`;
  else { target = slides.value; action = "ppaction://hlinksldjump"; external = false; }
  if (!target) return;
  ed.edit("Link einfügen", () => {
    const rels = pptx.rels(path);
    const rid = external ? rels.ensure(RT_LINK, target, true) : pptx.link(path, RT_SLIDE, target);
    const apply = (rPr: Element) => {
      for (const c of Array.from(rPr.children)) if (c.localName === "hlinkClick") c.remove();
      const link = PptxDom.make(rPr.ownerDocument, "a:hlinkClick");
      link.setAttributeNS(PptxDom.NS_R, "r:id", rid);
      if (action) link.setAttribute("action", action);
      const order = ["hlinkMouseOver", "rtl", "extLst"];
      rPr.insertBefore(link, Array.from(rPr.children).find((c) => order.includes(c.localName)) ?? null);
    };
    if (!body) return;
    if (t && T.cmpPos(t.caret, t.anchor) === 0) {
      // No text selected: the shown text (or the address) is typed at the caret, linked.
      const text = shown.value || (kind === "slide" ? slides.selectedOptions[0]?.textContent ?? "" : target.replace(/^mailto:/, ""));
      const start = t.caret;
      const end = T.insertText(body, start, text);
      T.formatRange(body, start, end, apply);
      t.caret = end; t.anchor = end;
    } else if (t) {
      if (shown.value && shown.value !== selected) {
        const from = T.deleteRange(body, t.anchor, t.caret);
        const end = T.insertText(body, from, shown.value);
        T.formatRange(body, from, end, apply);
        t.caret = end; t.anchor = from;
      } else T.formatRange(body, t.anchor, t.caret, apply);
    } else T.formatAll(body, apply);
  });
}

function slideTitle(app: PresentApp, path: string): string {
  const v = app.deck.view(path);
  const t = v.shapes.find((s) => s.ph && (s.ph.type === "title" || s.ph.type === "ctrTitle"));
  return t?.text ? t.text.paras.map((p) => p.runs.map((r) => r.text).join("")).join(" ") : "";
}

/** Where a link leads: an address (external) or a slide's index. */
export function linkTarget(app: PresentApp, slidePath: string, rid: string): { url?: string; slide?: number } {
  const r = app.deck.pptx.rels(slidePath).get(rid);
  if (!r) return {};
  if (r.external) return { url: r.target };
  const to = app.deck.pptx.target(slidePath, rid);
  const i = app.deck.slides.findIndex((s) => s.path === to);
  return i >= 0 ? { slide: i } : {};
}

// ----- Einfügen > Aktion (D-072): a click on the shape in the show does something -------------------------
// Written as PowerPoint writes it: a:hlinkClick in the shape's p:cNvPr, with action="ppaction://hlinkshowjump?
// jump=nextslide" (and the other jumps), or ppaction://hlinksldjump with a relationship to a slide.

const JUMPS: [string, string][] = [["nextslide", "Nächste Folie"], ["previousslide", "Vorherige Folie"], ["firstslide", "Erste Folie"], ["lastslide", "Letzte Folie"], ["lastslideviewed", "Letzte angezeigte Folie"], ["endshow", "Bildschirmpräsentation beenden"]];

export async function actionDialog(app: PresentApp): Promise<void> {
  const ed = app.editor;
  const v = ed.selectedViews[0];
  const el = v ? ed.shapeEl(v.id) : undefined;
  const nv = el ? PptxDom.kids(el).find((c) => c.localName.startsWith("nv")) : undefined;
  const cNvPr = nv?.firstElementChild;
  if (!el || !cNvPr) return;
  const cur = PptxDom.kid(cNvPr, "hlinkClick");
  const curJump = /jump=(\w+)/.exec(cur?.getAttribute("action") ?? "")?.[1];
  const none = h("input", { type: "radio", name: "act" }) as HTMLInputElement, link = h("input", { type: "radio", name: "act" }) as HTMLInputElement;
  none.checked = !cur; link.checked = !!cur;
  const sel = h("select", { style: "min-width:260px" }) as HTMLSelectElement;
  for (const [k, label] of JUMPS) sel.append(h("option", { value: `j:${k}` }, label));
  app.deck.slides.forEach((s, i) => { const title = app.deck.view(s.path).shapes.find((x) => x.ph && (x.ph.type === "title" || x.ph.type === "ctrTitle"))?.text?.paras.map((p) => p.runs.map((r) => r.text).join("")).join(" ") ?? ""; sel.append(h("option", { value: `s:${i}` }, `Folie ${i + 1}${title ? `: ${title}` : ""}`)); });
  if (curJump) sel.value = `j:${curJump}`;
  sel.addEventListener("change", () => { link.checked = true; });
  const ok = await showDialog({ title: "Aktionseinstellungen", body: h("div", { class: "form" }, h("b", null, "Aktion bei Mausklick"),
    h("label", { class: "dlg-check" }, none, " Keine"), h("label", { class: "dlg-check" }, link, " Hyperlink zu: "), sel),
    buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
  if (!ok) return;
  const path = app.slidePath()!;
  ed.edit("Aktion", () => {
    PptxDom.kid(cNvPr, "hlinkClick")?.remove();
    if (none.checked) return;
    const hl = PptxDom.make(cNvPr.ownerDocument, "a:hlinkClick");
    const [kind, val] = sel.value.split(":");
    if (kind === "j") { hl.setAttributeNS(PptxDom.NS_R, "r:id", ""); hl.setAttribute("action", `ppaction://hlinkshowjump?jump=${val}`); }
    else { const rid = app.deck.pptx.link(path, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide", app.deck.slides[Number(val)].path); hl.setAttributeNS(PptxDom.NS_R, "r:id", rid); hl.setAttribute("action", "ppaction://hlinksldjump"); }
    cNvPr.insertBefore(hl, cNvPr.firstChild);
  });
}

/** A shape's click action for the show: a jump ("nextslide" ...) or a relationship id. */
export function shapeAction(el: Element): { jump?: string; rid?: string } | null {
  const nv = PptxDom.kids(el).find((c) => c.localName.startsWith("nv"));
  const hl = nv?.firstElementChild ? PptxDom.kid(nv.firstElementChild, "hlinkClick") : undefined;
  if (!hl) return null;
  const jump = /jump=(\w+)/.exec(hl.getAttribute("action") ?? "")?.[1];
  return jump ? { jump } : { rid: hl.getAttributeNS(PptxDom.NS_R, "id") ?? undefined };
}
