// A presentation being edited (D-070): the package with its parts as DOM, slide operations as
// PowerPoint does them (Neue Folie with the layout's placeholders, Folie duplizieren, löschen,
// verschieben, Layout, Zurücksetzen) and undo. Each change records the XML of the parts it touches
// before and after (a slide is a few kilobytes), plus the slide list; undoing puts that XML back.

import { Pptx, PPTX_RT, CT_PPTX, PptxDom, Rels, newPptx, phOf, slideView, spTreeOf, type SlideRef, type SlideView } from "@raabe/ooxml";

const { kid, kids, at, num, make, fragment, NS_R, NS_P, serialize, parseXml } = PptxDom;

interface Snapshot {
  slides: SlideRef[];
  parts: Map<string, string | null>;
  rels: Map<string, string | null>;
  sections: string | null;
}

interface Step {
  label: string;
  before: Snapshot;
  after: Snapshot;
  merge?: string;
  time: number;
}

export class Deck {
  readonly pptx: Pptx;
  version = 0;
  private undoStack: Step[] = [];
  private redoStack: Step[] = [];
  private viewCache = new Map<string, { v: number; view: SlideView }>();
  /** Version of each part's last change (thumbnails redraw only what changed). */
  private partVersion = new Map<string, number>();
  onChange: ((paths: string[]) => void) | null = null;

  constructor(pptx?: Pptx) {
    this.pptx = pptx ?? newPptx();
  }

  get slides(): SlideRef[] { return this.pptx.slides; }

  slideIndex(path: string): number { return this.slides.findIndex((s) => s.path === path); }

  /** The slide's view, cached until one of its parts changes. */
  view(path: string): SlideView {
    const deps = [path, this.pptx.layoutOf(path) ?? "", this.pptx.presPath];
    const v = Math.max(...deps.map((d) => this.partVersion.get(d) ?? 0));
    const hit = this.viewCache.get(path);
    if (hit && hit.v === v) return hit.view;
    const view = slideView(this.pptx, path);
    this.viewCache.set(path, { v, view });
    return view;
  }

  partStamp(path: string): number {
    return Math.max(this.partVersion.get(path) ?? 0, this.partVersion.get(this.pptx.layoutOf(path) ?? "") ?? 0);
  }

  // ----- undo -------------------------------------------------------------------------------------

  private snap(paths: string[]): Snapshot {
    const parts = new Map<string, string | null>(), rels = new Map<string, string | null>();
    for (const p of new Set([...paths, this.pptx.presPath])) {
      parts.set(p, this.pptx.has(p) ? serialize(this.pptx.doc(p)) : null);
      rels.set(p, this.pptx.has(p) ? this.pptx.rels(p).xml() : null);
    }
    return { slides: this.slides.map((s) => ({ ...s })), parts, rels, sections: null };
  }

  private restore(s: Snapshot): string[] {
    this.pptx.slides = s.slides.map((x) => ({ ...x }));
    for (const [p, xml] of s.parts) {
      if (xml === null) continue;
      this.pptx.put(p, parseXml(xml));
    }
    for (const [p, xml] of s.rels) {
      if (xml === null) continue;
      const r = this.pptx.rels(p);
      const fresh = new Rels(xml);
      r.list.splice(0, r.list.length, ...fresh.list);
      r.changed = true;
    }
    return [...s.parts.keys()];
  }

  /**
   * One undoable change to `paths` (the parts `fn` edits). With `merge`, a change of the same kind
   * right after the last one joins it (typing a word is one step, as in PowerPoint).
   */
  change(label: string, paths: string[], fn: () => void, merge?: string): void {
    const before = this.snap(paths);
    fn();
    for (const p of paths) this.pptx.touch(p);
    this.pptx.touch(this.pptx.presPath);
    const after = this.snap(paths);
    const last = this.undoStack[this.undoStack.length - 1];
    const now = Date.now();
    if (merge && last?.merge === merge && now - last.time < 4000 && [...after.parts.keys()].every((k) => last.after.parts.has(k))) {
      last.after = after;
      last.time = now;
    } else {
      this.undoStack.push({ label, before, after, merge, time: now });
      if (this.undoStack.length > 200) this.undoStack.shift();
    }
    this.redoStack = [];
    this.changed(paths);
  }

  /** Ends typing's merging (a click, a command). */
  breakMerge(): void {
    const last = this.undoStack[this.undoStack.length - 1];
    if (last) last.merge = undefined;
  }

  get canUndo(): boolean { return this.undoStack.length > 0; }
  get canRedo(): boolean { return this.redoStack.length > 0; }
  get undoLabel(): string | undefined { return this.undoStack[this.undoStack.length - 1]?.label; }

  undo(): boolean {
    const s = this.undoStack.pop();
    if (!s) return false;
    this.redoStack.push(s);
    this.changed(this.restore(s.before));
    return true;
  }

  redo(): boolean {
    const s = this.redoStack.pop();
    if (!s) return false;
    this.undoStack.push(s);
    this.changed(this.restore(s.after));
    return true;
  }

  private changed(paths: string[]): void {
    this.version++;
    for (const p of paths) { this.partVersion.set(p, this.version); this.pptx.touch(p); }
    this.partVersion.set(this.pptx.presPath, this.version);
    this.onChange?.(paths);
  }

  /** A part changed outside `change` (e.g. a picture's bytes added). */
  markChanged(paths: string[]): void { this.changed(paths); }

  // ----- slides ----------------------------------------------------------------------------------

  /** The layouts of the first master (Neue Folie's gallery), with their names and types. */
  layouts(master = this.pptx.masters()[0]): { path: string; name: string; type: string }[] {
    if (!master) return [];
    return this.pptx.layouts(master).map((path) => {
      const root = this.pptx.doc(path).documentElement;
      return { path, name: kid(root, "cSld")?.getAttribute("name") ?? "", type: root.getAttribute("type") ?? "cust" };
    });
  }

  layoutName(slidePath: string): string {
    const l = this.pptx.layoutOf(slidePath);
    return l ? kid(this.pptx.doc(l).documentElement, "cSld")?.getAttribute("name") ?? "" : "";
  }

  /** The layout Neue Folie takes after `index`: after a Titelfolie "Titel und Inhalt", else the same (PowerPoint). */
  nextLayout(index: number): string | undefined {
    const ls = this.layouts();
    const cur = index >= 0 && this.slides[index] ? this.pptx.layoutOf(this.slides[index].path) : undefined;
    const curType = cur ? this.pptx.doc(cur).documentElement.getAttribute("type") : null;
    if (!cur || curType === "title") return ls.find((l) => l.type === "obj")?.path ?? ls[1]?.path ?? ls[0]?.path;
    return cur;
  }

  /** A slide's XML with the layout's placeholders (no footers), as PowerPoint makes a new slide. */
  private slideXml(layout: string): string {
    const tree = spTreeOf(this.pptx.doc(layout));
    let id = 2;
    const shapes = kids(tree).filter((s) => phOf(s) && !["dt", "ftr", "sldNum", "hdr"].includes(phOf(s)!.type)).map((s) => {
      const ph = at(s, "nvSpPr/nvPr/ph");
      const name = at(s, "nvSpPr/cNvPr")?.getAttribute("name") ?? "";
      const phAttrs = Array.from(ph?.attributes ?? []).filter((a) => a.name !== "hasCustomPrompt").map((a) => `${a.name}="${a.value}"`).join(" ");
      const n = id++;
      if (s.localName !== "sp") return "";
      return `<p:sp><p:nvSpPr><p:cNvPr id="${n}" name="${PptxDom.esc(name)}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph ${phAttrs}/></p:nvPr></p:nvSpPr><p:spPr/>` +
        `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="de-DE" dirty="0"/></a:p></p:txBody></p:sp>`;
    }).join("");
    return `${PptxDom.XML_DECL}<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${NS_R}" xmlns:p="${NS_P}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
      `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
  }

  /** Neue Folie: after slide `index` (-1: at the start) with `layout`. Returns the new slide's path. */
  addSlide(index: number, layout?: string): string {
    const lay = layout ?? this.nextLayout(index)!;
    const path = this.pptx.freeName("ppt/slides/slide#.xml");
    this.change("Neue Folie", [path], () => {
      this.pptx.put(path, parseXml(this.slideXml(lay)), CT_PPTX.slide);
      this.pptx.link(path, PPTX_RT.slideLayout, lay);
      this.pptx.slides.splice(index + 1, 0, { id: this.pptx.nextSlideId(), path });
      this.pptx.syncSlideList();
    });
    return path;
  }

  /** Folie duplizieren: copies of the slides right after the last of them. Returns the new paths. */
  duplicateSlides(indices: number[]): string[] {
    const sorted = [...indices].sort((a, b) => a - b);
    const made: string[] = [];
    // Names are reserved up front so the snapshot knows the new parts.
    const plan = sorted.map(() => "");
    for (let i = 0; i < plan.length; i++) {
      let p = this.pptx.freeName("ppt/slides/slide#.xml");
      while (plan.includes(p)) p = p.replace(/(\d+)\.xml$/, (_, n) => `${Number(n) + 1}.xml`);
      plan[i] = p;
    }
    this.change("Folie duplizieren", plan, () => {
      let at2 = sorted[sorted.length - 1];
      sorted.forEach((idx, k) => {
        const src = this.slides[idx].path;
        const path = plan[k];
        this.pptx.put(path, parseXml(serialize(this.pptx.doc(src))), CT_PPTX.slide);
        // The same relationships (layout, pictures); charts and notes get copies of their own.
        for (const r of this.pptx.rels(src).list) {
          const target = r.external ? r.target : this.pptx.target(src, r.id)!;
          if (!r.external && /notesSlide$/.test(r.type)) continue;
          let to = target;
          if (!r.external && /\/chart$/.test(r.type)) to = this.copyChart(target);
          const rels = this.pptx.rels(path);
          rels.list.push({ id: r.id, type: r.type, target: r.external ? r.target : relTarget(path, to), external: r.external });
          rels.changed = true;
        }
        const notes = this.pptx.notesOf(src);
        if (notes) this.copyNotes(notes, path);
        this.pptx.slides.splice(++at2, 0, { id: this.pptx.nextSlideId(), path });
        made.push(path);
      });
      this.pptx.syncSlideList();
    });
    return made;
  }

  private copyChart(chart: string): string {
    const to = this.pptx.freeName("ppt/charts/chart#.xml");
    this.pptx.put(to, this.pptx.files[chart] ? new Uint8Array(this.pptx.files[chart]) : serialize(this.pptx.doc(chart)), CT_PPTX.chart);
    for (const r of this.pptx.rels(chart).list) {
      const rels = this.pptx.rels(to);
      rels.list.push({ ...r, target: r.external ? r.target : relTarget(to, this.pptx.target(chart, r.id)!) });
      rels.changed = true;
    }
    return to;
  }

  private copyNotes(notes: string, slide: string): void {
    const to = this.pptx.freeName("ppt/notesSlides/notesSlide#.xml");
    this.pptx.put(to, parseXml(serialize(this.pptx.doc(notes))), CT_PPTX.notesSlide);
    const master = this.pptx.related(notes, PPTX_RT.notesMaster);
    if (master) this.pptx.link(to, PPTX_RT.notesMaster, master);
    this.pptx.link(to, PPTX_RT.slide, slide);
    this.pptx.link(slide, PPTX_RT.notesSlide, to);
  }

  deleteSlides(indices: number[]): void {
    const drop = new Set(indices);
    this.change("Folie löschen", [], () => {
      this.pptx.slides = this.slides.filter((_, i) => !drop.has(i));
      this.pptx.syncSlideList();
    });
  }

  /** Moves the slides (in order) so they start at `to` among the remaining ones. */
  moveSlides(indices: number[], to: number): void {
    const sorted = [...indices].sort((a, b) => a - b);
    this.change("Folie verschieben", [], () => {
      const moving = sorted.map((i) => this.slides[i]);
      const rest = this.slides.filter((_, i) => !sorted.includes(i));
      rest.splice(Math.max(0, Math.min(to, rest.length)), 0, ...moving);
      this.pptx.slides = rest;
      this.pptx.syncSlideList();
    });
  }

  /** Layout: gives the slides another layout; placeholders are matched by type and index, new ones added empty. */
  setLayout(indices: number[], layout: string): void {
    const paths = indices.map((i) => this.slides[i].path);
    this.change("Layout", paths, () => {
      for (const p of paths) {
        const rels = this.pptx.rels(p);
        const r = rels.list.find((x) => /slideLayout$/.test(x.type));
        if (r) { r.target = relTarget(p, layout); rels.changed = true; }
        // Placeholders the new layout lacks become plain shapes of their old place; missing ones are added.
        const tree = spTreeOf(this.pptx.doc(p))!;
        const lTree = spTreeOf(this.pptx.doc(layout))!;
        const lphs = kids(lTree).filter((s) => phOf(s) && !["dt", "ftr", "sldNum"].includes(phOf(s)!.type));
        const used = new Set<Element>();
        for (const s of kids(tree).filter((x) => phOf(x))) {
          const ph = phOf(s)!;
          const match = lphs.find((l) => !used.has(l) && (phOf(l)!.type === ph.type || (phOf(l)!.idx !== undefined && phOf(l)!.idx === ph.idx && !["title", "ctrTitle"].includes(phOf(l)!.type) && !["title", "ctrTitle"].includes(ph.type))))
            ?? lphs.find((l) => !used.has(l) && ["title", "ctrTitle"].includes(phOf(l)!.type) && ["title", "ctrTitle"].includes(ph.type))
            ?? lphs.find((l) => !used.has(l) && !["title", "ctrTitle"].includes(phOf(l)!.type) && !["title", "ctrTitle"].includes(ph.type) && ["obj", "body"].includes(phOf(l)!.type));
          const phEl = at(s, "nvSpPr/nvPr/ph") ?? at(s, "nvPicPr/nvPr/ph") ?? at(s, "nvGraphicFramePr/nvPr/ph");
          if (match) {
            used.add(match);
            const lph = phOf(match)!;
            if (phEl) {
              phEl.removeAttribute("type"); phEl.removeAttribute("idx");
              if (lph.type !== "obj") phEl.setAttribute("type", lph.type);
              if (lph.idx !== undefined) phEl.setAttribute("idx", String(lph.idx));
            }
          } else if (s.localName === "sp" && isEmptyText(s)) s.remove();
        }
        let next = Math.max(1, ...kids(tree).map((s) => num(at(s, kids(s)[0]?.localName + "/cNvPr"), "id") ?? 0)) + 1;
        for (const l of lphs) {
          if (used.has(l) || l.localName !== "sp") continue;
          const ph = at(l, "nvSpPr/nvPr/ph")!;
          const attrs = Array.from(ph.attributes).filter((a) => a.name !== "hasCustomPrompt").map((a) => `${a.name}="${a.value}"`).join(" ");
          const name = at(l, "nvSpPr/cNvPr")?.getAttribute("name") ?? "";
          tree.append(...fragment(tree.ownerDocument, `<p:sp><p:nvSpPr><p:cNvPr id="${next++}" name="${PptxDom.esc(name)}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph ${attrs}/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="de-DE" dirty="0"/></a:p></p:txBody></p:sp>`));
        }
      }
    });
  }

  /** Zurücksetzen: placeholders back to the layout's place, size and formatting (text kept). */
  resetSlides(indices: number[]): void {
    const paths = indices.map((i) => this.slides[i].path);
    this.change("Zurücksetzen", paths, () => {
      for (const p of paths) {
        for (const s of kids(spTreeOf(this.pptx.doc(p))).filter((x) => phOf(x))) {
          const spPr = kid(s, "spPr");
          if (spPr) while (spPr.firstChild) spPr.firstChild.remove();
          const body = kid(s, "txBody");
          const bodyPr = kid(body, "bodyPr");
          if (bodyPr) { for (const a of Array.from(bodyPr.attributes)) bodyPr.removeAttribute(a.name); while (bodyPr.firstChild) bodyPr.firstChild.remove(); }
          for (const pEl of kids(body, "p")) {
            const pPr = kid(pEl, "pPr");
            if (pPr) { const lvl = pPr.getAttribute("lvl"); for (const a of Array.from(pPr.attributes)) pPr.removeAttribute(a.name); while (pPr.firstChild) pPr.firstChild.remove(); if (lvl && lvl !== "0") pPr.setAttribute("lvl", lvl); }
            for (const r of kids(pEl)) {
              const rPr = kid(r, "rPr") ?? (r.localName === "endParaRPr" ? r : undefined);
              if (!rPr) continue;
              for (const a of Array.from(rPr.attributes)) if (!["lang", "dirty"].includes(a.name)) rPr.removeAttribute(a.name);
              while (rPr.firstChild) rPr.firstChild.remove();
            }
          }
        }
      }
    });
  }

  // ----- notes --------------------------------------------------------------------------------------

  notesText(slidePath: string): string {
    const notes = this.pptx.notesOf(slidePath);
    if (!notes) return "";
    const body = notesBody(this.pptx.doc(notes));
    return kids(kid(body, "txBody"), "p").map((p) => kids(p).filter((r) => r.localName === "r" || r.localName === "br" || r.localName === "fld").map((r) => r.localName === "br" ? "\n" : kid(r, "t")?.textContent ?? "").join("")).join("\n");
  }

  /** Writes the notes of a slide, making its notes page (and the notes master) when it has none. */
  setNotes(slidePath: string, text: string): void {
    let notes = this.pptx.notesOf(slidePath);
    const paths = [slidePath];
    const planned = notes ?? this.pptx.freeName("ppt/notesSlides/notesSlide#.xml");
    paths.push(planned);
    const master = this.ensureNotesMasterPath();
    paths.push(master);
    this.change("Notizen", paths, () => {
      this.ensureNotesMaster(master);
      if (!notes) {
        notes = planned;
        this.pptx.put(notes, parseXml(notesXml()), CT_PPTX.notesSlide);
        this.pptx.link(notes, PPTX_RT.notesMaster, master);
        this.pptx.link(notes, PPTX_RT.slide, slidePath);
        this.pptx.link(slidePath, PPTX_RT.notesSlide, notes);
      }
      const doc = this.pptx.doc(notes);
      const body = kid(notesBody(doc), "txBody")!;
      for (const p of kids(body, "p")) p.remove();
      for (const line of text.split("\n")) {
        const p = make(doc, "a:p");
        if (line) p.append(make(doc, "a:r", {}, make(doc, "a:rPr", { lang: "de-DE", dirty: "0" }), make(doc, "a:t", {}, line)));
        p.append(make(doc, "a:endParaRPr", { lang: "de-DE", dirty: "0" }));
        body.append(p);
      }
    }, `notes:${slidePath}`);
  }

  ensureNotesMasterPath(): string {
    const rels = this.pptx.rels(this.pptx.presPath);
    const r = rels.list.find((x) => /notesMaster$/.test(x.type));
    return r ? this.pptx.target(this.pptx.presPath, r.id)! : "ppt/notesMasters/notesMaster1.xml";
  }

  ensureNotesMaster(path: string): void {
    if (this.pptx.has(path)) return;
    const theme = this.pptx.freeName("ppt/theme/theme#.xml");
    const masterTheme = this.pptx.themeOf(this.pptx.masters()[0]);
    this.pptx.put(theme, this.pptx.text(masterTheme!) ?? "", CT_PPTX.theme);
    this.pptx.put(path, parseXml(notesMasterXml()), CT_PPTX.notesMaster);
    this.pptx.link(path, PPTX_RT.theme, theme);
    const rid = this.pptx.link(this.pptx.presPath, PPTX_RT.notesMaster, path);
    const pres = this.pptx.presentation;
    if (!kid(pres, "notesMasterIdLst")) {
      const list = make(pres.ownerDocument, "p:notesMasterIdLst");
      const id = make(pres.ownerDocument, "p:notesMasterId");
      id.setAttributeNS(NS_R, "r:id", rid);
      list.append(id);
      const after = kid(pres, "sldMasterIdLst");
      pres.insertBefore(list, after ? after.nextSibling : pres.firstChild);
    }
  }

  // ----- sections -------------------------------------------------------------------------------------

  /** Abschnitte (p14:sectionLst in presentation.xml's extLst): names and the slide ids they start with. */
  sections(): { name: string; id: string; slides: number[] }[] {
    const lst = this.pptx.presentation.getElementsByTagNameNS(PptxDom.NS_P14, "sectionLst")[0];
    if (!lst) return [];
    return kids(lst, "section").map((s) => ({
      name: s.getAttribute("name") ?? "", id: s.getAttribute("id") ?? "",
      slides: kids(kid(s, "sldIdLst"), "sldId").map((x) => num(x, "id") ?? 0),
    }));
  }

  /** Writes the sections (each a name and its slide ids, in slide order). */
  setSections(list: { name: string; id: string; slides: number[] }[], label = "Abschnitt"): void {
    this.change(label, [], () => {
      const pres = this.pptx.presentation;
      const doc = pres.ownerDocument;
      let lst = pres.getElementsByTagNameNS(PptxDom.NS_P14, "sectionLst")[0];
      if (!list.length) { const ext = lst?.parentElement; if (ext) ext.remove(); return; }
      if (!lst) {
        let extLst = kid(pres, "extLst");
        if (!extLst) { extLst = make(doc, "p:extLst"); pres.append(extLst); }
        const ext = make(doc, "p:ext", { uri: "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}" });
        lst = doc.createElementNS(PptxDom.NS_P14, "p14:sectionLst");
        ext.append(lst);
        extLst.insertBefore(ext, extLst.firstChild);
      }
      while (lst.firstChild) lst.firstChild.remove();
      for (const s of list) {
        const sec = doc.createElementNS(PptxDom.NS_P14, "p14:section");
        sec.setAttribute("name", s.name);
        sec.setAttribute("id", s.id);
        const ids = doc.createElementNS(PptxDom.NS_P14, "p14:sldIdLst");
        for (const id of s.slides) { const x = doc.createElementNS(PptxDom.NS_P14, "p14:sldId"); x.setAttribute("id", String(id)); ids.append(x); }
        sec.append(ids);
        lst.append(sec);
      }
    });
  }

  write(): Uint8Array {
    return this.pptx.write();
  }
}

function isEmptyText(s: Element): boolean {
  return !kids(kid(s, "txBody"), "p").some((p) => kids(p, "r").some((r) => (kid(r, "t")?.textContent ?? "") !== ""));
}

export function relTarget(from: string, to: string): string {
  const a = from.split("/").slice(0, -1), b = to.split("/");
  let i = 0;
  while (i < a.length && i < b.length - 1 && a[i] === b[i]) i++;
  return [...Array(a.length - i).fill(".."), ...b.slice(i)].join("/");
}

function notesBody(doc: Document): Element | undefined {
  return kids(spTreeOf(doc)).find((s) => phOf(s)?.type === "body");
}

const NSD = `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${NS_R}" xmlns:p="${NS_P}"`;
const TREE_HEAD = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

function notesXml(): string {
  return `${PptxDom.XML_DECL}<p:notes ${NSD}><p:cSld><p:spTree>${TREE_HEAD}` +
    `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Folienbildplatzhalter 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>` +
    `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notizenplatzhalter 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="de-DE"/></a:p></p:txBody></p:sp>` +
    `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`;
}

function notesMasterXml(): string {
  const ph = (id: number, name: string, type: string, idx: number | null, box: number[], extra = "") =>
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="${type}"${idx !== null ? ` idx="${idx}"` : ""}${extra}/></p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${box[0]}" y="${box[1]}"/><a:ext cx="${box[2]}" cy="${box[3]}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    (type === "sldImg" ? "" : `<p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr lang="de-DE"/></a:p></p:txBody>`) + "</p:sp>";
  const lv = Array.from({ length: 9 }, (_, i) => `<a:lvl${i + 1}pPr marL="${i * 457200}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${i + 1}pPr>`).join("");
  return `${PptxDom.XML_DECL}<p:notesMaster ${NSD}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>${TREE_HEAD}` +
    ph(2, "Kopfzeilenplatzhalter 1", "hdr", null, [0, 0, 2971800, 458788], ' sz="quarter"') +
    ph(3, "Datumsplatzhalter 2", "dt", 1, [3884613, 0, 2971800, 458788]) +
    ph(4, "Folienbildplatzhalter 3", "sldImg", 2, [685800, 1143000, 5486400, 3086100], ' hasCustomPrompt="0"').replace(' hasCustomPrompt="0"', "") +
    ph(5, "Notizenplatzhalter 4", "body", 3, [685800, 4400550, 5486400, 3600450], ' sz="quarter"') +
    ph(6, "Fußzeilenplatzhalter 5", "ftr", 4, [0, 8685213, 2971800, 458787], ' sz="quarter"') +
    ph(7, "Foliennummernplatzhalter 6", "sldNum", 5, [3884613, 8685213, 2971800, 458787], ' sz="quarter"') +
    `</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
    `<p:notesStyle>${lv}</p:notesStyle></p:notesMaster>`;
}

export { serialize };
