// Bildschirmpräsentation einrichten (D-070): Art der Präsentation, Wiederholen bis "Esc" gedrückt
// wird, ohne Animation, Folien anzeigen (alle / von-bis), Nächste Folie (manuell / Anzeigedauer
// verwenden), kept in ppt/presProps.xml as PowerPoint keeps them (p:showPr). The show reads them.

import { h, showDialog } from "@raabe/shell";
import { PptxDom } from "@raabe/ooxml";
import type { Deck } from "./deck";

const { kid, at, num, make } = PptxDom;
const RT_PRES_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps";

export interface ShowSetup { kiosk: boolean; loop: boolean; noAnimations: boolean; from?: number; to?: number; useTimings: boolean }

function propsPath(deck: Deck): string | undefined {
  return deck.pptx.related(deck.pptx.presPath, RT_PRES_PROPS);
}

export function readShowSetup(deck: Deck): ShowSetup {
  const p = propsPath(deck);
  const show = p && deck.pptx.has(p) ? at(deck.pptx.doc(p).documentElement, "showPr") : undefined;
  const range = kid(show, "sldRg");
  return {
    kiosk: !!kid(show, "kiosk"), loop: show?.getAttribute("loop") === "1", noAnimations: show?.getAttribute("showAnimation") === "0",
    from: num(range, "st"), to: num(range, "end"), useTimings: show?.getAttribute("useTimings") !== "0",
  };
}

export async function showSetupDialog(deck: Deck): Promise<void> {
  const cur = readShowSetup(deck);
  const radio = (name: string, label: string, on: boolean) => { const r = h("input", { type: "radio", name }) as HTMLInputElement; r.checked = on; return { r, el: h("label", { class: "dlg-check" }, r, ` ${label}`) }; };
  const check = (label: string, on: boolean) => { const c = h("input", { type: "checkbox" }) as HTMLInputElement; c.checked = on; return { c, el: h("label", { class: "dlg-check" }, c, ` ${label}`) }; };
  const speaker = radio("art", "Präsentation durch einen Redner (volle Bildschirmgröße)", !cur.kiosk), kiosk = radio("art", "An einem Kiosk (volle Bildschirmgröße)", cur.kiosk);
  const loop = check("Wiederholen, bis „Esc“ gedrückt wird", cur.loop || cur.kiosk), noAnim = check("Präsentation ohne Animation", cur.noAnimations);
  const all = radio("rg", "Alle", cur.from === undefined), range = radio("rg", "Von:", cur.from !== undefined);
  const from = h("input", { type: "number", min: "1", value: String(cur.from ?? 1), style: "width:50px" }) as HTMLInputElement;
  const to = h("input", { type: "number", min: "1", value: String(cur.to ?? deck.slides.length), style: "width:50px" }) as HTMLInputElement;
  const manual = radio("adv", "Manuell", !cur.useTimings), timings = radio("adv", "Anzeigedauer verwenden, wenn vorhanden", cur.useTimings);
  const body = h("div", { class: "form show-setup" },
    h("fieldset", null, h("legend", null, "Art der Präsentation"), speaker.el, kiosk.el),
    h("fieldset", null, h("legend", null, "Anzeigeoptionen"), loop.el, noAnim.el),
    h("fieldset", null, h("legend", null, "Folien anzeigen"), all.el, h("div", null, range.el, from, " Bis: ", to)),
    h("fieldset", null, h("legend", null, "Nächste Folie"), manual.el, timings.el));
  const ok = await showDialog({ title: "Bildschirmpräsentation einrichten", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
  if (!ok) return;
  const p = propsPath(deck);
  if (!p) return;
  deck.change("Bildschirmpräsentation einrichten", [p], () => {
    const root = deck.pptx.doc(p).documentElement;
    kid(root, "showPr")?.remove();
    const doc = root.ownerDocument;
    const sp = make(doc, "p:showPr");
    if (loop.c.checked || kiosk.r.checked) sp.setAttribute("loop", "1");
    if (noAnim.c.checked) sp.setAttribute("showAnimation", "0");
    if (manual.r.checked) sp.setAttribute("useTimings", "0");
    sp.setAttribute("showNarration", "1");
    sp.append(kiosk.r.checked ? make(doc, "p:kiosk") : make(doc, "p:present"));
    if (range.r.checked) sp.append(make(doc, "p:sldRg", { st: Math.max(1, Number(from.value) || 1), end: Math.max(1, Number(to.value) || 1) }));
    else sp.append(make(doc, "p:sldAll"));
    root.insertBefore(sp, root.firstChild);
  });
}

// ----- the checkboxes of Bildschirmpräsentation > Einrichten (D-072) ---------------------------------------
// Kommentare wiedergeben (showNarration), Anzeigedauern verwenden (useTimings), Mediensteuerelemente anzeigen
// (p14:showMediaCtrls in showPr's extLst), on by default as in a new PowerPoint presentation.

const NS_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const MEDIA_EXT = "{2FDB2607-1784-4EEB-B798-7EB5836EED8A}";
export type ShowFlag = "narration" | "timings" | "media";

export function showFlag(deck: Deck, f: ShowFlag): boolean {
  const p = propsPath(deck);
  const show = p && deck.pptx.has(p) ? at(deck.pptx.doc(p).documentElement, "showPr") : undefined;
  if (!show) return true;
  if (f === "narration") return show.getAttribute("showNarration") === "1";
  if (f === "timings") return show.getAttribute("useTimings") !== "0";
  const m = show.getElementsByTagNameNS(NS_P14, "showMediaCtrls")[0];
  return !m || m.getAttribute("val") !== "0";
}

export function setShowFlag(deck: Deck, f: ShowFlag, on: boolean): void {
  const p = propsPath(deck);
  if (!p) return;
  deck.change("Bildschirmpräsentation", [p], () => {
    const root = deck.pptx.doc(p).documentElement;
    const doc = root.ownerDocument;
    let show = kid(root, "showPr");
    if (!show) { show = make(doc, "p:showPr", { showNarration: 1 }); show.append(make(doc, "p:present"), make(doc, "p:sldAll")); root.insertBefore(show, root.firstChild); }
    if (f === "narration") show.setAttribute("showNarration", on ? "1" : "0");
    else if (f === "timings") { if (on) show.removeAttribute("useTimings"); else show.setAttribute("useTimings", "0"); }
    else {
      let m = show.getElementsByTagNameNS(NS_P14, "showMediaCtrls")[0];
      if (!m) {
        const ext = PptxDom.fragment(doc, `<p:extLst><p:ext uri="${MEDIA_EXT}"><p14:showMediaCtrls xmlns:p14="${NS_P14}" val="1"/></p:ext></p:extLst>`)[0];
        (kid(show, "extLst") ?? show).append(...(kid(show, "extLst") ? Array.from(ext.children) : [ext]));
        m = show.getElementsByTagNameNS(NS_P14, "showMediaCtrls")[0];
      }
      m.setAttribute("val", on ? "1" : "0");
    }
  });
}
