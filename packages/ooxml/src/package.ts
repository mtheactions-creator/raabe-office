// The parts of a .docx package: relationships, content types and media. Reading resolves
// relationship ids to parts; writing adds relationships and parts for new content (headers,
// footnotes, pictures …) and keeps every existing one.

import { strFromU8, strToU8 } from "fflate";
import { esc, parse } from "./xml";

export const REL = {
  header: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
  footer: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
  footnotes: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
  endnotes: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
  comments: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
  commentsExtended: "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
  numbering: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
  image: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
  hyperlink: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
  styles: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
  settings: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
};

export const CT = {
  header: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
  footer: "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
  footnotes: "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
  endnotes: "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
  comments: "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
  commentsExtended: "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml",
  numbering: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
};

export interface Rel {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

/** The relationships of one part. */
export class Rels {
  readonly list: Rel[] = [];
  changed = false;

  constructor(xml: string | null) {
    if (!xml) return;
    const doc = parse(xml);
    for (const r of Array.from(doc.documentElement.children)) {
      if (r.localName !== "Relationship") continue;
      this.list.push({
        id: r.getAttribute("Id") ?? "", type: r.getAttribute("Type") ?? "", target: r.getAttribute("Target") ?? "",
        external: r.getAttribute("TargetMode") === "External",
      });
    }
  }

  get(id: string): Rel | undefined {
    return this.list.find((r) => r.id === id);
  }

  /** The id of a relationship of `type` to `target`, added when missing. */
  ensure(type: string, target: string, external = false): string {
    const hit = this.list.find((r) => r.type === type && r.target === target);
    if (hit) return hit.id;
    let n = this.list.length + 1;
    const ids = new Set(this.list.map((r) => r.id));
    while (ids.has(`rId${n}`)) n++;
    const id = `rId${n}`;
    this.list.push({ id, type, target, external });
    this.changed = true;
    return id;
  }

  xml(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      this.list.map((r) => `<Relationship Id="${esc(r.id)}" Type="${esc(r.type)}" Target="${esc(r.target)}"${r.external ? ' TargetMode="External"' : ""}/>`).join("") +
      "</Relationships>";
  }
}

/** "word/document.xml" → "word/_rels/document.xml.rels". */
export function relsPath(part: string): string {
  const i = part.lastIndexOf("/");
  return `${part.slice(0, i + 1)}_rels/${part.slice(i + 1)}.rels`;
}

/** Resolves a relationship target relative to its source part ("media/a.png" from word/x.xml). */
export function resolveTarget(part: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const base = part.slice(0, part.lastIndexOf("/") + 1).split("/").filter(Boolean);
  for (const seg of target.split("/")) {
    if (seg === "..") base.pop();
    else if (seg !== ".") base.push(seg);
  }
  return base.join("/");
}

/** The target of `part` written relative to the folder of `from`. */
export function relativeTarget(from: string, part: string): string {
  const dir = from.slice(0, from.lastIndexOf("/") + 1);
  return part.startsWith(dir) ? part.slice(dir.length) : `/${part}`;
}

const MIME_EXT: Record<string, string> = {
  "image/png": "png", "image/jpeg": "jpeg", "image/gif": "gif", "image/bmp": "bmp", "image/svg+xml": "svg",
  "image/tiff": "tiff", "image/x-emf": "emf", "image/x-wmf": "wmf", "image/webp": "webp",
};

export function mimeOf(path: string): string {
  const ext = path.slice(path.lastIndexOf(".") + 1).toLowerCase();
  for (const [mime, e] of Object.entries(MIME_EXT)) if (e === ext || (ext === "jpg" && e === "jpeg")) return mime;
  return "application/octet-stream";
}

/** A package being written: files, relationships per part and content types. */
export class PackageWriter {
  readonly files: Record<string, Uint8Array>;
  private rels = new Map<string, Rels>();
  private defaults = new Map<string, string>();
  private overrides = new Map<string, string>();
  private media = new Map<Uint8Array, string>();
  private typesChanged = false;

  constructor(files: Record<string, Uint8Array>) {
    this.files = { ...files };
    const ct = files["[Content_Types].xml"];
    if (ct) {
      const doc = parse(strFromU8(ct));
      for (const e of Array.from(doc.documentElement.children)) {
        if (e.localName === "Default") this.defaults.set((e.getAttribute("Extension") ?? "").toLowerCase(), e.getAttribute("ContentType") ?? "");
        if (e.localName === "Override") this.overrides.set((e.getAttribute("PartName") ?? "").replace(/^\//, ""), e.getAttribute("ContentType") ?? "");
      }
    } else {
      this.defaults.set("rels", "application/vnd.openxmlformats-package.relationships+xml");
      this.defaults.set("xml", "application/xml");
      this.typesChanged = true;
    }
  }

  relsOf(part: string): Rels {
    let r = this.rels.get(part);
    if (!r) {
      const f = this.files[relsPath(part)];
      r = new Rels(f ? strFromU8(f) : null);
      this.rels.set(part, r);
    }
    return r;
  }

  /** Relationship from `part` to `target` (a package path), returning its id. */
  link(part: string, type: string, target: string): string {
    return this.relsOf(part).ensure(type, relativeTarget(part, target));
  }

  setPart(path: string, xml: string, contentType?: string): void {
    this.files[path] = strToU8(xml);
    if (contentType && this.overrides.get(path) !== contentType) {
      this.overrides.set(path, contentType);
      this.typesChanged = true;
    }
  }

  /** A free part name like word/header3.xml. */
  freeName(prefix: string, ext = "xml"): string {
    let n = 1;
    while (this.files[`${prefix}${n}.${ext}`]) n++;
    return `${prefix}${n}.${ext}`;
  }

  /** Stores picture bytes as a media part (once per picture) and returns its path. */
  addMedia(bytes: Uint8Array, mime: string, prefix = "word/media/image"): string {
    const hit = this.media.get(bytes);
    if (hit) return hit;
    const ext = MIME_EXT[mime] ?? "bin";
    const path = this.freeName(prefix, ext);
    this.files[path] = bytes;
    this.media.set(bytes, path);
    if (!this.defaults.has(ext)) { this.defaults.set(ext, mime); this.typesChanged = true; }
    return path;
  }

  finish(): Record<string, Uint8Array> {
    for (const [part, r] of this.rels) if (r.changed || !this.files[relsPath(part)]) if (r.list.length) this.files[relsPath(part)] = strToU8(r.xml());
    if (this.typesChanged) {
      const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        [...this.defaults].map(([e, t]) => `<Default Extension="${esc(e)}" ContentType="${esc(t)}"/>`).join("") +
        [...this.overrides].map(([p, t]) => `<Override PartName="/${esc(p)}" ContentType="${esc(t)}"/>`).join("") + "</Types>";
      this.files["[Content_Types].xml"] = strToU8(xml);
    }
    return this.files;
  }
}

/** Relationship id attributes in preserved XML (r:embed, r:id, r:link …). */
const REL_ATTR = /\b(r:(?:embed|id|link|pict|dm|lo|qs|cs|href))="([^"]+)"/g;

/**
 * Rewrites the relationship ids in preserved XML that came from part `from` for writing it
 * into part `to` (pictures copied between body and header …).
 */
export function relink(xml: string, from: string, to: string, pkg: PackageWriter): string {
  if (from === to) return xml;
  const src = pkg.relsOf(from);
  return xml.replace(REL_ATTR, (m, name: string, id: string) => {
    const rel = src.get(id);
    if (!rel) return m;
    const target = rel.external ? rel.target : resolveTarget(from, rel.target);
    const nid = rel.external ? pkg.relsOf(to).ensure(rel.type, rel.target, true) : pkg.link(to, rel.type, target);
    return `${name}="${nid}"`;
  });
}
