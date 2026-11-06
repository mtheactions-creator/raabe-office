// A .pptx package in memory (D-070): every part as PowerPoint wrote it, the XML parts that are
// read or edited as DOM documents. Writing serializes only the documents that changed; all other
// parts go back byte for byte. Parts no relationship reaches any more (a deleted slide, its notes)
// are left out, as PowerPoint does.

import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { Rels, relsPath, resolveTarget } from "../package";
import { NS_CT, NS_P, NS_R, at, kid, kids, make, num, parseXml, serialize } from "./dom";

export const RT = {
  officeDocument: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
  slide: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
  slideLayout: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
  slideMaster: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
  notesSlide: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide",
  notesMaster: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster",
  handoutMaster: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster",
  theme: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
  image: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
  chart: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
  hyperlink: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
  presProps: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps",
  viewProps: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps",
  tableStyles: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles",
  package: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
};

export const CT_PPTX = {
  presentation: "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
  slide: "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
  slideLayout: "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
  slideMaster: "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
  notesSlide: "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
  notesMaster: "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml",
  theme: "application/vnd.openxmlformats-officedocument.theme+xml",
  chart: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
};

const isRel = (type: string, want: string) => type === want || type.slice(type.lastIndexOf("/")) === want.slice(want.lastIndexOf("/"));

export interface SlideRef {
  /** p:sldId/@id (256 and up). */
  id: number;
  path: string;
}

export class Pptx {
  readonly files: Record<string, Uint8Array>;
  private docs = new Map<string, Document>();
  private dirty = new Set<string>();
  private relsCache = new Map<string, Rels>();
  readonly presPath: string;
  slides: SlideRef[] = [];
  /** Content types: Default by extension, Override by part. */
  private ctDoc: Document;

  constructor(files: Record<string, Uint8Array>) {
    this.files = files;
    this.ctDoc = parseXml(this.text("[Content_Types].xml") ?? `<Types xmlns="${NS_CT}"/>`);
    const root = this.rels("").list.find((r) => isRel(r.type, RT.officeDocument));
    if (!root) throw new Error("Keine Präsentation (ppt/presentation.xml fehlt)");
    this.presPath = resolveTarget("", root.target);
    const pres = this.doc(this.presPath);
    const rels = this.rels(this.presPath);
    for (const s of kids(at(pres.documentElement, "sldIdLst"), "sldId")) {
      const rid = s.getAttributeNS(NS_R, "id") ?? "";
      const rel = rels.get(rid);
      if (rel && this.has(resolveTarget(this.presPath, rel.target))) this.slides.push({ id: num(s, "id") ?? 0, path: resolveTarget(this.presPath, rel.target) });
    }
  }

  static read(bytes: Uint8Array): Pptx {
    return new Pptx(unzipSync(bytes));
  }

  has(path: string): boolean {
    return path in this.files || this.docs.has(path);
  }

  text(path: string): string | null {
    if (this.docs.has(path)) return serialize(this.docs.get(path)!);
    const b = this.files[path];
    return b ? strFromU8(b) : null;
  }

  /** The part as a DOM document (parsed once). */
  doc(path: string): Document {
    let d = this.docs.get(path);
    if (!d) {
      const t = this.text(path);
      if (t === null) throw new Error(`Teil fehlt: ${path}`);
      d = parseXml(t);
      this.docs.set(path, d);
    }
    return d;
  }

  /** Marks a part as changed, so it is serialized on writing. */
  touch(path: string): void {
    this.dirty.add(path);
  }

  isDirty(path: string): boolean {
    return this.dirty.has(path);
  }

  rels(part: string): Rels {
    const p = part ? relsPath(part) : "_rels/.rels";
    let r = this.relsCache.get(p);
    if (!r) {
      r = new Rels(this.text(p));
      this.relsCache.set(p, r);
    }
    return r;
  }

  /** The part a relationship of `part` points to (resolved path), by id. */
  target(part: string, rid: string): string | undefined {
    const r = this.rels(part).get(rid);
    return r && !r.external ? resolveTarget(part, r.target) : undefined;
  }

  /** The first relationship of `part` of the given type, resolved. */
  related(part: string, type: string): string | undefined {
    const r = this.rels(part).list.find((x) => isRel(x.type, type) && !x.external);
    return r ? resolveTarget(part, r.target) : undefined;
  }

  /** A relationship to `to` written as PowerPoint does ("../slideLayouts/slideLayout2.xml"). */
  link(from: string, type: string, to: string): string {
    return this.rels(from).ensure(type, relTarget(from, to));
  }

  contentType(path: string): string | undefined {
    for (const o of kids(this.ctDoc.documentElement, "Override")) if (o.getAttribute("PartName") === `/${path}`) return o.getAttribute("ContentType") ?? undefined;
    const ext = path.slice(path.lastIndexOf(".") + 1).toLowerCase();
    for (const d of kids(this.ctDoc.documentElement, "Default")) if (d.getAttribute("Extension")?.toLowerCase() === ext) return d.getAttribute("ContentType") ?? undefined;
    return undefined;
  }

  /** Adds or replaces a part; `type` is its content type (an Override), or a Default for its extension. */
  put(path: string, data: string | Uint8Array | Document, type?: string): void {
    if (typeof data === "string") { this.files[path] = strToU8(data); this.docs.delete(path); }
    else if (data instanceof Uint8Array) { this.files[path] = data; this.docs.delete(path); }
    else { this.docs.set(path, data); this.dirty.add(path); }
    if (type) this.setType(path, type);
  }

  setType(path: string, type: string, asDefault = false): void {
    const root = this.ctDoc.documentElement;
    if (asDefault) {
      const ext = path.slice(path.lastIndexOf(".") + 1).toLowerCase();
      if (!kids(root, "Default").some((d) => d.getAttribute("Extension")?.toLowerCase() === ext)) {
        const d = make(this.ctDoc, "Default", { Extension: ext, ContentType: type });
        root.insertBefore(d, kids(root, "Override")[0] ?? null);
        this.ctChanged = true;
      }
      return;
    }
    const hit = kids(root, "Override").find((o) => o.getAttribute("PartName") === `/${path}`);
    if (hit) { if (hit.getAttribute("ContentType") !== type) { hit.setAttribute("ContentType", type); this.ctChanged = true; } return; }
    root.append(make(this.ctDoc, "Override", { PartName: `/${path}`, ContentType: type }));
    this.ctChanged = true;
  }

  private ctChanged = false;

  /** A free part name "ppt/slides/slide7.xml" for the pattern "ppt/slides/slide#.xml". */
  freeName(pattern: string): string {
    for (let i = 1; ; i++) {
      const p = pattern.replace("#", String(i));
      if (!this.has(p)) return p;
    }
  }

  // ----- the presentation's structure ---------------------------------------------------------

  get presentation(): Element {
    return this.doc(this.presPath).documentElement;
  }

  /** Slide size in EMU (p:sldSz), 16:9 (12192000 x 6858000) when missing. */
  get slideSize(): { cx: number; cy: number } {
    const s = at(this.presentation, "sldSz");
    return { cx: num(s, "cx") ?? 12192000, cy: num(s, "cy") ?? 6858000 };
  }

  layoutOf(slidePath: string): string | undefined {
    return this.related(slidePath, RT.slideLayout);
  }

  masterOf(layoutPath: string): string | undefined {
    return this.related(layoutPath, RT.slideMaster);
  }

  themeOf(masterPath: string): string | undefined {
    return this.related(masterPath, RT.theme);
  }

  notesOf(slidePath: string): string | undefined {
    return this.related(slidePath, RT.notesSlide);
  }

  /** The masters in presentation order (p:sldMasterIdLst). */
  masters(): string[] {
    const rels = this.rels(this.presPath);
    return kids(at(this.presentation, "sldMasterIdLst"), "sldMasterId")
      .map((m) => rels.get(m.getAttributeNS(NS_R, "id") ?? ""))
      .filter((r) => !!r)
      .map((r) => resolveTarget(this.presPath, r!.target));
  }

  /** A master's layouts in its order (p:sldLayoutIdLst). */
  layouts(master: string): string[] {
    const rels = this.rels(master);
    const root = this.doc(master).documentElement;
    return kids(at(root, "sldLayoutIdLst"), "sldLayoutId")
      .map((l) => rels.get(l.getAttributeNS(NS_R, "id") ?? ""))
      .filter((r) => !!r)
      .map((r) => resolveTarget(master, r!.target))
      .filter((p) => this.has(p));
  }

  /** Writes p:sldIdLst from `slides` (after adding, deleting or moving slides). */
  syncSlideList(): void {
    const pres = this.presentation;
    const rels = this.rels(this.presPath);
    let list = at(pres, "sldIdLst");
    if (!list) {
      list = make(pres.ownerDocument, "p:sldIdLst");
      const after = kid(pres, "notesMasterIdLst") ?? kid(pres, "sldMasterIdLst");
      pres.insertBefore(list, after ? after.nextSibling : pres.firstChild);
    }
    const old = new Map(kids(list, "sldId").map((s) => [num(s, "id") ?? 0, s]));
    while (list.firstChild) list.firstChild.remove();
    for (const s of this.slides) {
      const rid = rels.ensure(RT.slide, relTarget(this.presPath, s.path));
      const el = old.get(s.id) ?? make(pres.ownerDocument, "p:sldId", { id: s.id });
      el.setAttributeNS(NS_R, "r:id", rid);
      list.append(el);
    }
    // Relationships of removed slides go too.
    const keep = new Set(this.slides.map((s) => relTarget(this.presPath, s.path)));
    for (let i = rels.list.length - 1; i >= 0; i--) {
      const r = rels.list[i];
      if (isRel(r.type, RT.slide) && !keep.has(r.target) && !keep.has(relTarget(this.presPath, resolveTarget(this.presPath, r.target)))) { rels.list.splice(i, 1); rels.changed = true; }
    }
    this.touch(this.presPath);
  }

  nextSlideId(): number {
    return Math.max(255, ...this.slides.map((s) => s.id)) + 1;
  }

  // ----- writing --------------------------------------------------------------------------------

  /** Parts reached through relationships from the package root. */
  private reachable(): Set<string> {
    const seen = new Set<string>();
    const walk = (part: string) => {
      for (const r of this.rels(part).list) {
        if (r.external) continue;
        const t = resolveTarget(part, r.target);
        if (seen.has(t) || !this.has(t)) continue;
        seen.add(t);
        walk(t);
      }
    };
    walk("");
    return seen;
  }

  write(): Uint8Array {
    const out: Record<string, Uint8Array> = {};
    const live = this.reachable();
    for (const p of live) {
      out[p] = this.docs.has(p) && this.dirty.has(p) ? strToU8(serialize(this.docs.get(p)!)) : this.files[p] ?? strToU8(serialize(this.docs.get(p)!));
      const rp = relsPath(p);
      const rc = this.relsCache.get(rp);
      if (rc && (rc.changed || !this.files[rp])) { if (rc.list.length) out[rp] = strToU8(rc.xml()); }
      else if (this.files[rp]) out[rp] = this.files[rp];
    }
    const rootRels = this.relsCache.get("_rels/.rels");
    out["_rels/.rels"] = rootRels?.changed ? strToU8(rootRels.xml()) : this.files["_rels/.rels"];
    // Content types: Overrides only for parts that are written.
    const root = this.ctDoc.documentElement;
    for (const o of kids(root, "Override")) {
      const p = (o.getAttribute("PartName") ?? "").replace(/^\//, "");
      if (!out[p]) { o.remove(); this.ctChanged = true; }
    }
    out["[Content_Types].xml"] = this.ctChanged || !this.files["[Content_Types].xml"] ? strToU8(serialize(this.ctDoc)) : this.files["[Content_Types].xml"];
    // [Content_Types].xml first, as Office writes it.
    const ordered: Record<string, Uint8Array> = { "[Content_Types].xml": out["[Content_Types].xml"] };
    for (const [k, v] of Object.entries(out)) if (k !== "[Content_Types].xml") ordered[k] = v;
    return zipSync(ordered, { level: 6 });
  }
}

/** A relationship target from `from` to `to` as PowerPoint writes it (relative, with ".."). */
export function relTarget(from: string, to: string): string {
  const a = from.split("/").slice(0, -1), b = to.split("/");
  let i = 0;
  while (i < a.length && i < b.length - 1 && a[i] === b[i]) i++;
  return [...Array(a.length - i).fill(".."), ...b.slice(i)].join("/");
}

export { NS_P };
