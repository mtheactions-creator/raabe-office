// Makros (D-079): a presentation's VBA project (vbaProject.bin) read and written with @raabe/vba, run by our
// interpreter against a PowerPoint-like object model of the presentation. As in PowerPoint, macros of an opened
// file are off until "Inhalt aktivieren" in the message bar; Ansicht > Makros lists them (Ausführen, Bearbeiten,
// Erstellen, Löschen); a shape's action "Makro ausführen" runs one in the show. Macros cannot reach outside the
// app (D-066): no Shell, files, CreateObject or Declare. One macro run is one undo step.

import { h, showDialog } from "@raabe/shell";
import { PptxDom } from "@raabe/ooxml";
import { NOT_A_METHOD, VbaError, VbaRuntime, readProject, visibleCode, vbaToNumber, vbaToString, writeProject, type VbaForm, type VbaHost, type VbaModule, type VbaObject, type VbaValue } from "@raabe/vba";
import type { PresentApp } from "./app";
import * as S from "./shapeops";

const RT_VBA = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
const CT_VBA = "application/vnd.ms-office.vbaProject";
const EMU = 12700;
const { kid, kids } = PptxDom;

/** Presentations whose macros were switched on (Inhalt aktivieren), by their deck. */
const enabled = new WeakSet<object>();
export const macrosEnabled = (app: PresentApp) => enabled.has(app.deck);
export const enableMacros = (app: PresentApp) => enabled.add(app.deck);

function vbaPath(app: PresentApp): string | undefined {
  return app.deck.pptx.related(app.deck.pptx.presPath, RT_VBA);
}

/** The presentation's modules ([] when it has no VBA project). */
export function modulesOf(app: PresentApp): VbaModule[] {
  const p = vbaPath(app);
  const bytes = p ? app.deck.pptx.files[p] : undefined;
  if (!bytes) return [];
  try { return readProject(bytes); } catch { return []; }
}

/** Writes the modules as the presentation's VBA project (it becomes a presentation with macros, .pptm). */
export function setModules(app: PresentApp, modules: VbaModule[]): void {
  const pptx = app.deck.pptx;
  const path = vbaPath(app) ?? "ppt/vbaProject.bin";
  app.deck.change("Makros", [], () => {
    pptx.put(path, writeProject(modules));
    pptx.setType(path, CT_VBA);
    pptx.link(pptx.presPath, RT_VBA, path);
  });
  enabled.add(app.deck);
}

// ----- the object model ------------------------------------------------------------------------------------

type Get = (args: VbaValue[]) => VbaValue;
/** An object from property getters, setters and methods (names compared without case). */
function obj(props: Record<string, Get>, setters: Record<string, (v: VbaValue) => void> = {}, methods: Record<string, (args: VbaValue[]) => VbaValue | Promise<VbaValue>> = {}, extra: Partial<VbaObject> = {}): VbaObject {
  const low = <T>(r: Record<string, T>) => Object.fromEntries(Object.entries(r).map(([k, v]) => [k.toLowerCase(), v]));
  const P = low(props), SET = low(setters), M = low(methods);
  return {
    kind: "object",
    get: (name, args) => { const g = P[name.toLowerCase()]; if (g) return g(args); const m = M[name.toLowerCase()]; if (m) return m(args) as VbaValue; throw new VbaError(`Eigenschaft oder Methode nicht unterstützt: ${name}`, 438); },
    set: (name, v) => { const s = SET[name.toLowerCase()]; if (!s) throw new VbaError(`Eigenschaft kann nicht gesetzt werden: ${name}`, 438); s(v); },
    call: (name, args) => { const m = M[name.toLowerCase()]; return m ? (m(args) as VbaValue) : NOT_A_METHOD; },
    ...extra,
  };
}

const bgr = (hex: string) => { const n = parseInt(hex, 16); return ((n >> 16) & 255) | (n & 0xff00) | ((n & 255) << 16); };
const hexOf = (rgb: VbaValue) => { const n = vbaToNumber(rgb); return [n & 255, (n >> 8) & 255, (n >> 16) & 255].map((x) => x.toString(16).padStart(2, "0")).join("").toUpperCase(); };
/** msoAutoShapeType to preset geometry (the common ones). */
const SHAPES: Record<number, string> = { 1: "rect", 5: "roundRect", 9: "ellipse", 7: "triangle", 4: "diamond", 33: "rightArrow", 34: "leftArrow", 35: "upArrow", 36: "downArrow", 12: "star5", 17: "smileyFace", 21: "heart", 51: "cloud", 92: "star5", 8: "rtTriangle", 10: "hexagon", 52: "flowChartProcess" };

/**
 * Shapes PowerPoint named itself carry German names in a German file ("Titel 1") but VBA sees them in English
 * ("Title 1"), and either name finds them (measured, macro test FormNamen, D-105).
 */
const DEFAULT_NAMES: [string, string][] = [["Titel", "Title"], ["Untertitel", "Subtitle"], ["Inhaltsplatzhalter", "Content Placeholder"], ["Textplatzhalter", "Text Placeholder"],
  ["Bildplatzhalter", "Picture Placeholder"], ["Datumsplatzhalter", "Date Placeholder"], ["Fußzeilenplatzhalter", "Footer Placeholder"], ["Foliennummernplatzhalter", "Slide Number Placeholder"],
  ["Grafik", "Picture"], ["Textfeld", "TextBox"], ["Rechteck", "Rectangle"], ["Ellipse", "Oval"], ["Tabelle", "Table"], ["Diagramm", "Chart"], ["Gruppieren", "Group"]];
export function vbaShapeName(name: string): string {
  for (const [de, en] of DEFAULT_NAMES) if (new RegExp(`^${de} \\d+$`).test(name)) return en + name.slice(de.length);
  return name;
}

/** The object model over the presentation; `show` is the running show's view (SlideShowWindow.View), if any. */
function model(app: PresentApp, show: ShowControl | null): VbaHost["global"] {
  const deck = () => app.deck;
  const slideObj = (i: number): VbaObject => {
    const ref = () => deck().slides[i];
    const doc = () => deck().pptx.doc(ref().path);
    const shapesColl = (): VbaObject => {
      const list = () => deck().view(ref().path).shapes;
      const item = (k: VbaValue) => { const l = list(); const key = vbaToString(k).toLowerCase(); const v = typeof k === "number" ? l[k - 1] : l.find((s) => s.name.toLowerCase() === key || vbaShapeName(s.name).toLowerCase() === key); if (!v) throw new VbaError("Das Element wurde nicht gefunden.", -2147188160); return shapeObj(i, v.id); };
      const add = (prst: string, l: VbaValue, t: VbaValue, w: VbaValue, hh: VbaValue, text: boolean) => {
        const tree = S.treeOf(doc());
        const id = S.maxId(tree) + 1;
        const x = Math.round(vbaToNumber(l) * EMU), y = Math.round(vbaToNumber(t) * EMU), cx = Math.round(vbaToNumber(w) * EMU), cy = Math.round(vbaToNumber(hh) * EMU);
        const name = text ? `Textfeld ${id - 1}` : `Form ${id - 1}`;
        tree.append(PptxDom.fragment(doc(), `<p:sp xmlns:p="${PptxDom.NS_P}" xmlns:a="${PptxDom.NS_A}"><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr${text ? ' txBox="1"' : ""}/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>${text ? "<a:noFill/>" : ""}</p:spPr>` +
          (text ? "" : `<p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>`) +
          `<p:txBody><a:bodyPr wrap="${text ? "none" : "square"}" rtlCol="0"${text ? "" : ' anchor="ctr"'}>${text ? "<a:spAutoFit/>" : ""}</a:bodyPr><a:lstStyle/><a:p>${text ? "" : '<a:pPr algn="ctr"/>'}<a:endParaRPr lang="de-DE"/></a:p></p:txBody></p:sp>`)[0]);
        deck().invalidate(ref().path);
        return shapeObj(i, id);
      };
      return obj({ Count: () => list().length, Item: (a) => item(a[0]), Title: () => { const v = list().find((s) => s.ph && (s.ph.type === "title" || s.ph.type === "ctrTitle")); if (!v) throw new VbaError("Kein Titel", 438); return shapeObj(i, v.id); } }, {},
        { AddTextbox: (a) => add("rect", a[1], a[2], a[3], a[4], true), AddShape: (a) => add(SHAPES[vbaToNumber(a[0])] ?? "rect", a[1], a[2], a[3], a[4], false) },
        { items: () => list().map((v) => shapeObj(i, v.id)), default: (a) => item(a[0]) });
    };
    return obj({
      Shapes: () => shapesColl(), SlideIndex: () => i + 1, SlideNumber: () => i + 1, SlideID: () => ref().id, Name: () => `Folie${i + 1}`,
      Layout: () => 1, Parent: () => presentation(),
      // NotesPage: its body placeholder (Placeholders(2)) holds the slide's notes (D-105).
      NotesPage: () => {
        const notesShape = () => obj({ HasTextFrame: () => true, Name: () => "Notizenplatzhalter 2", Type: () => 14,
          TextFrame: () => obj({ HasText: () => deck().notesText(ref().path).length > 0,
            TextRange: () => obj({ Text: () => deck().notesText(ref().path).replace(/\n/g, "\r") }, { Text: (v) => deck().setNotes(ref().path, vbaToString(v).replace(/\r\n?/g, "\n")) }) }) });
        const phs = obj({ Count: () => 2 }, {}, {}, { default: (a) => { if (vbaToNumber(a[0]) !== 2) throw new VbaError("Index außerhalb des gültigen Bereichs", 9); return notesShape(); } });
        return obj({ Shapes: () => obj({ Placeholders: () => phs, Count: () => 2 }) }, {}, {}, { default: () => obj({ Shapes: () => obj({ Placeholders: () => phs }) }) });
      },
    }, {}, {
      Delete: () => { deck().deleteSlides([i]); return undefined; },
      Duplicate: () => { deck().duplicateSlides([i]); return slideObj(i + 1); },
      Select: () => { app.setCurrent(i); return undefined; },
    });
  };
  const shapeObj = (slide: number, id: number): VbaObject => {
    const path = () => deck().slides[slide].path;
    const el = () => { const e = Array.from(deck().pptx.doc(path()).getElementsByTagNameNS(PptxDom.NS_P, "cNvPr")).find((c) => c.getAttribute("id") === String(id))?.parentElement?.parentElement; if (!e) throw new VbaError("Die Form gibt es nicht mehr.", -2147188160); return e; };
    const view = () => deck().view(path()).shapes.find((s) => s.id === id)!;
    // LockAspectRatio: a picture's is on unless its file says otherwise; ours for other shapes kept here (D-105).
    const locksEl = () => { const nv = kids(el()).find((c) => c.localName.startsWith("nv")); const c = nv ? kids(nv).find((x) => x.localName.startsWith("cNv") && x.localName !== "cNvPr") : undefined; return c; };
    const locked = () => { const c = locksEl(); const l = c ? kids(c).find((x) => x.localName.endsWith("Locks")) : undefined; return l ? l.getAttribute("noChangeAspect") === "1" : view().kind === "pic"; };
    const setLocked = (on: boolean) => {
      const c = locksEl();
      if (!c) return;
      let l = kids(c).find((x) => x.localName.endsWith("Locks"));
      if (!l) { l = PptxDom.make(c.ownerDocument, view().kind === "pic" ? "a:picLocks" : "a:spLocks"); c.prepend(l); }
      if (on) l.setAttribute("noChangeAspect", "1"); else l.removeAttribute("noChangeAspect");
    };
    const xf = (k: "x" | "y" | "w" | "h", v: VbaValue) => {
      const s = view();
      const box = { x: s.x, y: s.y, w: s.w, h: s.h, [k]: Math.round(vbaToNumber(v) * EMU) };
      // With the proportions locked PowerPoint keeps them: the other side follows (measured in the macro tests).
      if (locked() && s.w && s.h) { if (k === "w") box.h = Math.round((box.w * s.h) / s.w); if (k === "h") box.w = Math.round((box.h * s.w) / s.h); }
      S.setXfrm(el(), box);
      deck().invalidate(path());
    };
    /** msoShapeType as PowerPoint reports it. */
    const shapeType = (): number => { const v = view(); if (v.ph) return 14; if (v.kind === "pic") return 13; if (v.kind === "grpSp") return 6; if (v.frame === "table") return 19; if (v.frame === "chart") return 3; if (v.frame === "diagram") return 24; if (v.kind === "cxnSp" || v.geometry === "line") return 9; return v.textBox ? 17 : 1; };
    const PH_TYPES: Record<string, number> = { title: 1, body: 2, ctrTitle: 3, subTitle: 4, dt: 16, sldNum: 13, ftr: 15, obj: 7, chart: 8, tbl: 12, clipArt: 10, dgm: 11, media: 17, pic: 18 };
    const spPr = () => kid(el(), "spPr")!;
    const setFill = (hex: string | null) => { const sp = spPr(); for (const c of Array.from(sp.children)) if (["noFill", "solidFill", "gradFill", "blipFill", "pattFill"].includes(c.localName)) c.remove(); const f = PptxDom.fragment(sp.ownerDocument, hex ? `<a:solidFill xmlns:a="${PptxDom.NS_A}"><a:srgbClr val="${hex}"/></a:solidFill>` : `<a:noFill xmlns:a="${PptxDom.NS_A}"/>`)[0]; const after = Array.from(sp.children).find((c) => c.localName === "prstGeom" || c.localName === "custGeom"); sp.insertBefore(f, after ? after.nextSibling : null); deck().invalidate(path()); };
    const runs = () => Array.from(el().getElementsByTagNameNS(PptxDom.NS_A, "rPr"));
    const textRange = (): VbaObject => obj({
      Text: () => (view().text?.paras.map((p) => p.runs.map((r) => r.text).join("")).join("\r") ?? ""),
      Font: () => obj({ Size: () => view().text?.paras[0]?.runs[0]?.size ?? 18, Bold: () => !!view().text?.paras[0]?.runs[0]?.bold, Italic: () => !!view().text?.paras[0]?.runs[0]?.italic, Name: () => view().text?.paras[0]?.runs[0]?.font ?? "",
        Color: () => obj({ RGB: () => bgr(view().text?.paras[0]?.runs[0]?.color ?? "000000") }, { RGB: (v) => { for (const r of runs()) { for (const c of Array.from(r.children)) if (c.localName.endsWith("Fill")) c.remove(); r.prepend(PptxDom.fragment(r.ownerDocument, `<a:solidFill xmlns:a="${PptxDom.NS_A}"><a:srgbClr val="${hexOf(v)}"/></a:solidFill>`)[0]); } deck().invalidate(path()); } }) },
        { Size: (v) => { for (const r of runs()) r.setAttribute("sz", String(Math.round(vbaToNumber(v) * 100))); deck().invalidate(path()); }, Bold: (v) => { for (const r of runs()) r.setAttribute("b", vbaToNumber(v) ? "1" : "0"); deck().invalidate(path()); },
          Italic: (v) => { for (const r of runs()) r.setAttribute("i", vbaToNumber(v) ? "1" : "0"); deck().invalidate(path()); }, Name: (v) => { for (const r of runs()) { r.querySelector("latin")?.remove(); r.append(PptxDom.make(r.ownerDocument, "a:latin", { typeface: vbaToString(v) })); } deck().invalidate(path()); } }),
    }, { Text: (v) => { setShapeText(el(), vbaToString(v)); deck().invalidate(path()); } });
    return obj({
      Name: () => vbaShapeName(view().name), Id: () => id, Left: () => view().x / EMU, Top: () => view().y / EMU, Width: () => view().w / EMU, Height: () => view().h / EMU, Rotation: () => view().rot,
      Visible: () => !view().hidden, HasTextFrame: () => !!kid(el(), "txBody") || (!!view().ph && view().kind === "sp"),
      Type: () => shapeType(), LockAspectRatio: () => (locked() ? -1 : 0),
      PlaceholderFormat: () => { const p = view().ph; if (!p) throw new VbaError("Diese Form ist kein Platzhalter.", -2147188160); return obj({ Type: () => PH_TYPES[p.type ?? "body"] ?? 2 }); },
      TextFrame: () => obj({ TextRange: () => textRange(), HasText: () => !!view().text && !view().text!.empty }),
      TextFrame2: () => obj({ TextRange: () => textRange() }),
      Fill: () => obj({ ForeColor: () => obj({ RGB: () => { const f = view().format.fill; return bgr(f && f.kind === "solid" ? f.color : "FFFFFF"); } }, { RGB: (v) => setFill(hexOf(v)) }), Visible: () => view().format.fill?.kind !== "none" }, { Visible: (v) => { if (!vbaToNumber(v)) setFill(null); } }),
      Line: () => obj({ ForeColor: () => obj({ RGB: () => bgr(view().format.line?.color ?? "000000") }, { RGB: (v) => { const sp = spPr(); let ln = kid(sp, "ln"); if (!ln) { ln = PptxDom.make(sp.ownerDocument, "a:ln"); sp.append(ln); } for (const c of Array.from(ln.children)) if (c.localName.endsWith("Fill")) c.remove(); ln.prepend(PptxDom.fragment(sp.ownerDocument, `<a:solidFill xmlns:a="${PptxDom.NS_A}"><a:srgbClr val="${hexOf(v)}"/></a:solidFill>`)[0]); deck().invalidate(path()); } }) }),
    }, {
      Name: (v) => { el().getElementsByTagNameNS(PptxDom.NS_P, "cNvPr")[0]?.setAttribute("name", vbaToString(v)); deck().invalidate(path()); },
      Left: (v) => xf("x", v), Top: (v) => xf("y", v), Width: (v) => xf("w", v), Height: (v) => xf("h", v),
      LockAspectRatio: (v) => { setLocked(vbaToNumber(v) !== 0); deck().invalidate(path()); },
      Rotation: (v) => { const x = kid(spPr(), "xfrm"); x?.setAttribute("rot", String(Math.round(vbaToNumber(v) * 60000))); deck().invalidate(path()); },
      Visible: (v) => { const c = el().getElementsByTagNameNS(PptxDom.NS_P, "cNvPr")[0]; if (vbaToNumber(v)) c?.removeAttribute("hidden"); else c?.setAttribute("hidden", "1"); deck().invalidate(path()); show?.refresh(); },
    }, {
      Delete: () => { el().remove(); deck().invalidate(path()); return undefined; },
      Select: () => { app.editor.select([id]); return undefined; },
      IncrementLeft: (a) => { xf("x", view().x / EMU + vbaToNumber(a[0])); return undefined; },
      IncrementTop: (a) => { xf("y", view().y / EMU + vbaToNumber(a[0])); return undefined; },
      IncrementRotation: (a) => { const x = kid(spPr(), "xfrm"); x?.setAttribute("rot", String(Math.round((view().rot + vbaToNumber(a[0])) * 60000))); deck().invalidate(path()); return undefined; },
    });
  };
  const slidesColl = (): VbaObject => obj({ Count: () => deck().slides.length, Item: (a) => slideObj(vbaToNumber(a[0]) - 1) }, {}, {
    Add: (a) => { const at = Math.max(1, Math.min(deck().slides.length + 1, vbaToNumber(a[0]))); app.setCurrent(at - 2 < 0 ? 0 : at - 2); app.newSlide(); if (at === 1) deck().moveSlides([app.slideNumber() - 1], 0); return slideObj(at - 1); },
  }, { items: () => deck().slides.map((_, i) => slideObj(i)), default: (a) => { const k = vbaToNumber(a[0]) - 1; if (k < 0 || k >= deck().slides.length) throw new VbaError("Index außerhalb des gültigen Bereichs", 9); return slideObj(k); } });
  const view = (): VbaObject => obj({ CurrentShowPosition: () => (show ? show.position() : app.slideNumber()), Slide: () => slideObj((show ? show.position() : app.slideNumber()) - 1), State: () => 1 }, {}, {
    Next: () => { show?.next(); return undefined; }, Previous: () => { show?.prev(); return undefined; },
    GotoSlide: (a) => { show ? show.goto(vbaToNumber(a[0]) - 1) : app.setCurrent(vbaToNumber(a[0]) - 1); return undefined; },
    First: () => { show?.goto(0); return undefined; }, Last: () => { show?.goto(deck().slides.length - 1); return undefined; }, Exit: () => { show?.end(); return undefined; },
  });
  const showWindow = (): VbaObject => obj({ View: () => view() });
  const presentation = (): VbaObject => obj({ Slides: () => slidesColl(), Name: () => app.docName(), SlideShowWindow: () => showWindow(), PageSetup: () => obj({ SlideWidth: () => deck().pptx.slideSize.cx / EMU, SlideHeight: () => deck().pptx.slideSize.cy / EMU }) });
  const application: VbaObject = obj({ ActivePresentation: () => presentation(), Name: () => "RaabePresent", Version: () => "16.0", SlideShowWindows: () => obj({ Count: () => (show ? 1 : 0) }, {}, {}, { default: () => showWindow() }) });
  return (name) => {
    switch (name.toLowerCase()) {
      case "application": return application;
      case "activepresentation": return presentation();
      case "slideshowwindows": return application.get("SlideShowWindows", []);
      case "msoshaperectangle": return 1; case "msoshaperoundedrectangle": return 5; case "msoshapeoval": return 9; case "msoshapeisoscelestriangle": return 7;
      case "msotrue": return -1; case "msofalse": return 0; case "msotextorientationhorizontal": return 1;
      case "pplayoutblank": return 12; case "pplayouttext": return 2; case "pplayouttitle": return 1; case "pplayouttitleonly": return 11;
    }
    return undefined;
  };
}

/** Plain text into a shape's first paragraph (lines as paragraphs), the first run's look kept. */
function setShapeText(el: Element, text: string): void {
  let body = kid(el, "txBody");
  const doc = el.ownerDocument;
  if (!body) { body = PptxDom.fragment(doc, `<p:txBody xmlns:p="${PptxDom.NS_P}" xmlns:a="${PptxDom.NS_A}"><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>`)[0]; el.append(body); }
  const ps = Array.from(body.children).filter((c) => c.localName === "p");
  const first = ps[0];
  const rPr = first?.getElementsByTagNameNS(PptxDom.NS_A, "rPr")[0]?.cloneNode(true) as Element | undefined;
  const pPr = first ? kid(first, "pPr")?.cloneNode(true) as Element | undefined : undefined;
  for (const p of ps) p.remove();
  for (const line of text.split(/\r\n|\r|\n/)) {
    const p = PptxDom.make(doc, "a:p");
    if (pPr) p.append(pPr.cloneNode(true));
    if (line) { const r = PptxDom.make(doc, "a:r"); r.append(rPr ? rPr.cloneNode(true) : PptxDom.make(doc, "a:rPr", { lang: "de-DE" }), PptxDom.make(doc, "a:t", {}, line)); p.append(r); }
    else p.append(PptxDom.make(doc, "a:endParaRPr", { lang: "de-DE" }));
    body.append(p);
  }
}

/** The running show as macros see it (SlideShowWindow.View). */
export interface ShowControl { position(): number; next(): void; prev(): void; goto(i: number): void; end(): void; refresh(): void }

/** Runs a macro by name: one undo step; errors reported as VBA reports them. */
export async function runMacro(app: PresentApp, name: string, show: ShowControl | null = null): Promise<void> {
  if (!macrosEnabled(app)) { await securityNotice(app); if (!macrosEnabled(app)) return; }
  // The code as the editor shows it (no Attribute lines), so error lines and breakpoints match the editor's.
  const mods = modulesOf(app).map((m) => ({ ...m, code: visibleCode(m.code) }));
  let rt: VbaRuntime | null = null;
  const forms = new Map<VbaForm, () => void>();
  const { showVbaForm } = await import("./vbaide");
  const hst: VbaHost = { ...host(app, show), showForm: (f, modal) => { forms.set(f, showVbaForm(() => rt, f, modal)); }, hideForm: (f) => { forms.get(f)?.(); forms.delete(f); } };
  try { rt = new VbaRuntime(hst, mods); } catch (e) { await report(app, e, name); return; }
  const run = rt;
  const paths = app.deck.slides.map((s) => s.path);
  let failure: unknown = null;
  // The whole run is one undo step (the macro's changes are made inside it).
  await app.deck.changeAsync(`Makro ${name}`, paths, async () => { try { await run.run(name); } catch (e) { failure = e; } });
  app.renderAll();
  show?.refresh();
  if (failure) await report(app, failure, name);
}

/** Runs a macro without dialogs (the harness, D-105): macros enabled, MsgBox answered OK, the error returned. */
export async function runMacroQuiet(app: PresentApp, name: string, extra: Partial<VbaHost> = {}): Promise<{ error?: { number: number; message: string; line?: number } }> {
  enabled.add(app.deck);
  const h0 = host(app, null);
  const hst: VbaHost = { ...h0, msgBox: async () => 1, inputBox: async (_p, _t, d) => d, ...extra };
  let failure: unknown = null;
  try {
    const rt = new VbaRuntime(hst, modulesOf(app).map((m) => ({ ...m, code: visibleCode(m.code) })));
    await app.deck.changeAsync(`Makro ${name}`, app.deck.slides.map((s) => s.path), async () => { try { await rt.run(name); } catch (e) { failure = e; } });
  } catch (e) { failure = e; }
  app.renderAll();
  if (!failure) return {};
  const err = failure instanceof VbaError ? failure : new VbaError(String((failure as Error)?.message ?? failure));
  return { error: { number: err.number, message: err.message, line: err.line } };
}

function host(app: PresentApp, show: ShowControl | null): VbaHost {
  const global = model(app, show);
  return {
    global,
    msgBox: async (prompt, buttons, title) => {
      const kind = buttons & 7;
      const opts: [string, number][] = kind === 1 ? [["OK", 1], ["Abbrechen", 2]] : kind === 4 ? [["Ja", 6], ["Nein", 7]] : kind === 3 ? [["Ja", 6], ["Nein", 7], ["Abbrechen", 2]] : kind === 5 ? [["Wiederholen", 4], ["Abbrechen", 2]] : [["OK", 1]];
      return showDialog({ title, body: h("div", { style: "white-space:pre-wrap;max-width:420px" }, prompt), buttons: opts.map(([label, value], i) => ({ label, value, primary: i === 0 })), cancel: opts[opts.length - 1][1] });
    },
    // Files only after asking, every time (D-080, D-100): the macro names the file and what it wants to do.
    fileAccess: async (kind, path, target) => {
      const what: Record<string, string> = { open: "lesen", save: "schreiben", append: "ergänzen", export: "schreiben", delete: "löschen", copy: "kopieren", rename: "umbenennen", mkdir: "als Ordner anlegen", rmdir: "als Ordner löschen", list: "auflisten" };
      return showDialog({ title: "Makro: Dateizugriff", body: h("div", { style: "max-width:460px" }, h("div", null, `Ein Makro dieser Präsentation möchte ${kind === "list" ? "den Ordner" : "die Datei"}`),
        h("div", { style: "margin:6px 0;font-family:Consolas,monospace;word-break:break-all" }, path + (target ? ` \u2192 ${target}` : "")), h("div", null, `${what[kind] ?? "verwenden"}. Zulassen?`)),
        buttons: [{ label: "Zulassen", value: true }, { label: "Nicht zulassen", value: false, primary: true }], cancel: false });
    },
    files: {
      read: (p) => app.platform.readFile(p),
      write: (p, d) => app.platform.writeFile(p, d),
    },
    inputBox: async (prompt, title, def) => {
      const input = h("input", { type: "text", value: def, style: "width:320px" }) as HTMLInputElement;
      const ok = await showDialog({ title, body: h("div", { class: "form" }, h("div", null, prompt), input), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
      return ok ? input.value : "";
    },
  };
}

/** A runtime error as Office's VBA shows it: Laufzeitfehler 'n': and the message; Beenden or Debuggen (D-106). */
async function report(app: PresentApp, e: unknown, name: string): Promise<void> {
  const err = e instanceof VbaError ? e : new VbaError(String((e as Error)?.message ?? e));
  const debug = await showDialog({ title: "RaabePresent - Makros", body: h("div", { style: "max-width:440px" }, h("div", null, `Laufzeitfehler '${err.number}':`), h("div", { style: "margin-top:10px" }, err.message),
    err.line ? h("div", { class: "muted", style: "margin-top:10px" }, `${err.proc ?? name}${err.module ? ` (${err.module})` : ""}, Zeile ${err.line}`) : ""),
    buttons: [{ label: "Beenden", value: false, primary: true }, { label: "Debuggen", value: true }], cancel: false });
  if (debug) await editor(app, err.proc ?? name);
}

/** The message bar of a presentation with macros (SICHERHEITSWARNUNG, Inhalt aktivieren). */
export function messageBar(app: PresentApp, anchor: HTMLElement): void {
  document.querySelector(".macro-bar")?.remove();
  if (!vbaPath(app) || macrosEnabled(app)) return;
  const bar = h("div", { class: "macro-bar", role: "alert" }, h("b", null, "SICHERHEITSWARNUNG"), h("span", null, "Makros wurden deaktiviert."),
    h("button", { class: "btn", onclick: () => { enabled.add(app.deck); bar.remove(); } }, "Inhalt aktivieren"),
    h("button", { class: "macro-bar-close", "aria-label": "Schließen", onclick: () => bar.remove() }, "×"));
  anchor.parentElement?.insertBefore(bar, anchor);
}

async function securityNotice(app: PresentApp): Promise<void> {
  const yes = await showDialog({ title: "RaabePresent", body: h("div", { style: "max-width:420px" }, "Die Makros dieser Präsentation sind deaktiviert. Aktivieren Sie sie nur, wenn Sie der Datei vertrauen."),
    buttons: [{ label: "Inhalt aktivieren", value: true }, { label: "Abbrechen", value: false, primary: true }], cancel: false });
  if (yes) enabled.add(app.deck);
}

/** Ansicht > Makros: PowerPoint's dialog (Makroname, the list, Ausführen, Bearbeiten, Erstellen, Löschen). */
export async function macroDialog(app: PresentApp): Promise<void> {
  const mods = modulesOf(app);
  let names: string[] = [];
  try { names = new VbaRuntime(host(app, null), mods.map((m) => m.code)).macros(); } catch { /* a module that does not parse: listed empty */ }
  const nameInput = h("input", { type: "text", value: names[0] ?? "", style: "width:100%", "aria-label": "Makroname" }) as HTMLInputElement;
  const list = h("select", { size: 10, style: "width:100%", "aria-label": "Makros" }, ...names.map((n) => h("option", { value: n }, n))) as HTMLSelectElement;
  list.addEventListener("change", () => { nameInput.value = list.value; });
  const v = await showDialog<"run" | "edit" | "create" | "delete" | "">({ title: "Makro", body: h("div", { class: "form", style: "width:380px" }, h("label", null, "Makroname:"), nameInput, list, h("div", null, "Makro in: ", h("b", null, app.docName()))),
    buttons: [{ label: "Ausführen", value: "run", primary: true }, { label: "Bearbeiten", value: "edit" }, { label: "Erstellen", value: "create" }, { label: "Löschen", value: "delete" }, { label: "Abbrechen", value: "" }], cancel: "" });
  const name = nameInput.value.trim();
  if (v === "run" && name) await runMacro(app, name);
  if (v === "edit") await editor(app, name);
  if (v === "create" && name) {
    if (!/^[A-Za-zÄÖÜäöüß]\w*$/.test(name)) { await report(app, new VbaError("Ungültiger Makroname", 5), name); return; }
    const all = mods.length ? mods : [{ name: "Modul1", kind: "standard" as const, code: "" }];
    const m = all.find((x) => x.kind === "standard") ?? all[0];
    m.code = `${visibleCode(m.code).trimEnd()}\r\n\r\nSub ${name}()\r\n\r\nEnd Sub\r\n`.replace(/^\r\n\r\n/, "");
    setModules(app, all);
    await editor(app, name);
  }
  if (v === "delete" && name) {
    for (const m of mods) m.code = visibleCode(m.code).replace(new RegExp(`(Public |Private )?Sub ${name}\\b[\\s\\S]*?End Sub\\r?\\n?`, "i"), "");
    setModules(app, mods);
  }
}

/** Our VBA editor: the modules with their code (Attribute lines hidden); Speichern writes the project. */
/** Entwickler > Visual Basic (Alt+F11): the macro editor over the app (vbaide.ts, D-106). */
export async function editor(app: PresentApp, focus?: string): Promise<void> {
  const { openIde } = await import("./vbaide");
  await openIde({
    title: app.docName(),
    modules: () => modulesOf(app),
    save: (mods) => setModules(app, mods),
    vbaHost: (extra) => ({ ...host(app, null), ...extra }),
  }, focus);
}

/** Bildschirmpräsentation: a shape's action ppaction://macro?name=… */
export function macroOfAction(action: string | null): string | null {
  const m = /^ppaction:\/\/macro\?name=(.+)$/i.exec(action ?? "");
  return m ? decodeURIComponent(m[1]) : null;
}
