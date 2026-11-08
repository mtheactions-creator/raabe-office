// A VBA project (vbaProject.bin) as MS-OVBA describes it (D-079): the compression of its streams, the dir stream
// with the modules, and each module's source code. Read: the modules of a project from PowerPoint, Excel or
// Word. Written: a project with our modules (no compiled code, so Office compiles the source when it opens it).

import { readCfb, readCfbStorages, writeCfb } from "./cfb";
import type { VbaFormDesign } from "./forms";
import { CLSID_FORM, newGuid, opaqueForm, readFormStorage, writeFormStorage } from "./oforms";

/** A module of a project. A UserForm is a class module with its form (MS-OVBA stores forms as designer modules,
 *  which have the same module type as class and document modules; the PROJECT stream names them BaseClass=). */
export interface VbaModule { name: string; kind: "standard" | "class" | "document"; code: string; form?: VbaFormDesign }

/** MS-OVBA 2.4.1: a compressed container to bytes. */
export function decompress(b: Uint8Array): Uint8Array {
  if (b[0] !== 1) throw new Error("not a compressed container");
  const out: number[] = [];
  let i = 1;
  while (i + 1 < b.length) {
    const head = b[i] | (b[i + 1] << 8);
    const size = (head & 0x0fff) + 3;
    const compressed = (head & 0x8000) !== 0;
    const end = Math.min(b.length, i + size);
    i += 2;
    const chunkStart = out.length;
    if (!compressed) { for (let k = 0; k < 4096 && i < end; k++) out.push(b[i++]); continue; }
    while (i < end) {
      const flags = b[i++];
      for (let bit = 0; bit < 8 && i < end; bit++) {
        if ((flags & (1 << bit)) === 0) { out.push(b[i++]); continue; }
        const token = b[i] | (b[i + 1] << 8);
        i += 2;
        const pos = out.length - chunkStart;
        let bits = 4;
        while (1 << bits < pos) bits++;
        bits = Math.max(4, Math.min(12, bits));
        const lenMask = 0xffff >> bits;
        const offset = (token >> (16 - bits)) + 1;
        const len = (token & lenMask) + 3;
        for (let k = 0; k < len; k++) out.push(out[out.length - offset]);
      }
    }
  }
  return Uint8Array.from(out);
}

/** MS-OVBA 2.4.1: bytes as a compressed container (copy tokens for repeats, as Office writes them). */
export function compress(src: Uint8Array): Uint8Array {
  const out: number[] = [1];
  for (let start = 0; start < src.length || start === 0; start += 4096) {
    const chunk = src.subarray(start, Math.min(src.length, start + 4096));
    const body: number[] = [];
    let p = 0;
    while (p < chunk.length) {
      const flagAt = body.length;
      body.push(0);
      for (let bit = 0; bit < 8 && p < chunk.length; bit++) {
        // The longest match before p within the reach the token's bits allow.
        let bits = 4;
        while (1 << bits < p) bits++;
        bits = Math.max(4, Math.min(12, bits));
        const maxLen = (0xffff >> bits) + 3, maxOff = 1 << (16 - bits);
        let bestLen = 0, bestOff = 0;
        for (let o = 1; o <= Math.min(p, maxOff); o++) {
          let l = 0;
          while (l < maxLen && p + l < chunk.length && chunk[p + l] === chunk[p + l - o]) l++;
          if (l > bestLen) { bestLen = l; bestOff = o; if (l === maxLen) break; }
        }
        if (bestLen >= 3) {
          const token = ((bestOff - 1) << (16 - bits)) | (bestLen - 3);
          body.push(token & 255, token >> 8);
          body[flagAt] |= 1 << bit;
          p += bestLen;
        } else body.push(chunk[p++]);
      }
    }
    const size = body.length + 2;
    const head = (size - 3) | 0xb000;
    out.push(head & 255, head >> 8, ...body);
    if (!src.length) break;
  }
  return Uint8Array.from(out);
}

const latin1 = (b: Uint8Array) => Array.from(b, (c) => String.fromCharCode(c)).join("");
/** Windows-1252 text for the project's MBCS streams (German letters included). */
function cp1252(s: string): Uint8Array {
  const map: Record<string, number> = { "€": 0x80, "‚": 0x82, "„": 0x84, "…": 0x85, "–": 0x96, "—": 0x97, "‘": 0x91, "’": 0x92, "“": 0x93, "”": 0x94 };
  return Uint8Array.from(Array.from(s), (c) => map[c] ?? (c.charCodeAt(0) < 256 ? c.charCodeAt(0) : 0x3f));
}
function fromCp1252(b: Uint8Array): string {
  const map: Record<number, string> = { 0x80: "€", 0x82: "‚", 0x84: "„", 0x85: "…", 0x96: "–", 0x97: "—", 0x91: "‘", 0x92: "’", 0x93: "“", 0x94: "”" };
  return Array.from(b, (c) => map[c] ?? String.fromCharCode(c)).join("");
}

/** The modules of a project (vbaProject.bin). */
export function readProject(bin: Uint8Array): VbaModule[] {
  const files = readCfb(bin);
  const find = (name: string) => [...files.entries()].find(([k]) => k.toLowerCase() === name.toLowerCase())?.[1];
  const dirRaw = find("VBA/dir");
  if (!dirRaw) return [];
  const dir = decompress(dirRaw);
  const v = new DataView(dir.buffer, dir.byteOffset, dir.byteLength);
  const mods: { name: string; stream: string; offset: number; kind: VbaModule["kind"] }[] = [];
  let cur: { name: string; stream: string; offset: number; kind: VbaModule["kind"] } | null = null;
  let i = 0;
  while (i + 6 <= dir.length) {
    const id = v.getUint16(i, true);
    let size = v.getUint32(i + 2, true);
    // PROJECTVERSION's size field is reserved (4) and the record is 4 + 2 bytes long.
    if (id === 0x0009) size = 6;
    const data = dir.subarray(i + 6, i + 6 + size);
    i += 6 + size;
    if (id === 0x0019) { cur = { name: fromCp1252(data), stream: fromCp1252(data), offset: 0, kind: "standard" }; mods.push(cur); }
    else if (id === 0x001a && cur) cur.stream = fromCp1252(data);
    else if (id === 0x0031 && cur) cur.offset = new DataView(data.buffer, data.byteOffset).getUint32(0, true);
    else if (id === 0x0022 && cur) cur.kind = "class";
    else if (id === 0x0010) break;
  }
  // Document modules (Diese Arbeitsmappe, Tabelle1 …) are classes named in the PROJECT stream as Document=;
  // designer modules (UserForms) are named as BaseClass= and have a storage of their own (the form, oforms.ts).
  const project = find("PROJECT");
  const lines = project ? latin1(project).split(/\r?\n/) : [];
  const docs = new Set(lines.filter((l) => l.startsWith("Document=")).map((l) => l.slice(9).split("/")[0]));
  const designers = new Set(lines.filter((l) => l.startsWith("BaseClass=")).map((l) => l.slice(10).trim().toLowerCase()));
  let clsids: Map<string, string> | null = null;
  return mods.map((m) => {
    const s = find(`VBA/${m.stream}`);
    const code = s && s.length > m.offset ? fromCp1252(decompress(s.subarray(m.offset))) : "";
    const mod: VbaModule = { name: m.name, kind: docs.has(m.name) ? "document" : m.kind, code };
    if (designers.has(m.name.toLowerCase())) {
      const store = [...files.keys()].map((k) => k.split("/")[0]).find((k) => k.toLowerCase() === m.stream.toLowerCase()) ?? m.stream;
      const streams = new Map<string, Uint8Array>();
      for (const [k, v] of files) if (k.startsWith(store + "/")) streams.set(k.slice(store.length + 1), v);
      clsids ??= readCfbStorages(bin);
      const ids = new Map<string, string>();
      for (const [k, v] of clsids) if (k === store) ids.set("", v); else if (k.startsWith(store + "/")) ids.set(k.slice(store.length + 1), v);
      // A form we cannot read is kept as its bytes (written back unchanged).
      try { mod.form = readFormStorage(m.name, streams, ids); } catch { mod.form = opaqueForm(m.name, streams, ids); }
    }
    return mod;
  });
}

/** The source without its Attribute lines (what the editor shows). */
export function visibleCode(code: string): string {
  return code.split(/\r?\n/).filter((l) => !/^Attribute VB_/i.test(l)).join("\n").replace(/^\n+/, "");
}

/** The reference to Microsoft Forms 2.0 that a project with UserForms needs ([MS-OVBA] 2.3.4.2.2.3: the
 *  designer's type library as a REFERENCECONTROL, after the REFERENCEORIGINAL naming FM20.DLL). The extended type
 *  library (MSForms.exd) is a cache Office makes for itself; its path here names no user folder. */
function msFormsReference(rec: (id: number, data: Uint8Array) => Uint8Array[], u16: (n: number) => Uint8Array, u32: (n: number) => Uint8Array, utf16: (s: string) => Uint8Array): Uint8Array[] {
  const name = () => [...rec(0x0016, cp1252("MSForms")), u16(0x003e), u32(14), utf16("MSForms")];
  const original = "*\\G{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0#C:\\Windows\\system32\\FM20.DLL#Microsoft Forms 2.0 Object Library";
  const twiddled = "*\\G{00000000-0000-0000-0000-000000000000}#0.0#0##";
  const extended = `*\\G${newGuid()}#2.0#0#C:\\Windows\\Temp\\VBE\\MSForms.exd#Microsoft Forms 2.0 Object Library`;
  const out: Uint8Array[] = [...name(), ...rec(0x0033, cp1252(original))];
  out.push(u16(0x002f), u32(4 + twiddled.length + 6), u32(twiddled.length), cp1252(twiddled), u32(0), u16(0));
  out.push(...name(), u16(0x0030), u32(4 + extended.length + 6 + 16 + 4), u32(extended.length), cp1252(extended), u32(0), u16(0));
  const g = "0D452EE1E08F101A852E02608C4D0BB4";
  const by = (i: number) => parseInt(g.slice(i * 2, i * 2 + 2), 16);
  out.push(Uint8Array.from([by(3), by(2), by(1), by(0), by(5), by(4), by(7), by(6), ...Array.from({ length: 8 }, (_, i) => by(8 + i))]), u32(1));
  return out;
}

/** A project with these modules (standard and class modules; document modules are kept as classes; a module with
 *  a form is written as a UserForm: designer module, designer storage, BaseClass= and the MSForms reference). */
export function writeProject(modules: VbaModule[], projectName = "VBAProject"): Uint8Array {
  const rec = (id: number, data: Uint8Array) => { const h = new Uint8Array(6); const v = new DataView(h.buffer); v.setUint16(0, id, true); v.setUint32(2, data.length, true); return [h, data]; };
  const u32 = (n: number) => { const b = new Uint8Array(4); new DataView(b.buffer).setUint32(0, n, true); return b; };
  const u16 = (n: number) => { const b = new Uint8Array(2); new DataView(b.buffer).setUint16(0, n, true); return b; };
  const utf16 = (s: string) => { const b = new Uint8Array(s.length * 2); const v = new DataView(b.buffer); for (let k = 0; k < s.length; k++) v.setUint16(k * 2, s.charCodeAt(k), true); return b; };
  const parts: Uint8Array[] = [];
  const add = (...p: Uint8Array[]) => parts.push(...p);
  // PROJECTLCID and PROJECTLCIDINVOKE must be 0x0409 (MS-OVBA): VBA calls the host with LCIDINVOKE, and with German
  // Excel would read a macro's Range.Formula = "=SUM(...)" and NumberFormat as local text (#NAME?).
  add(...rec(0x0001, u32(1)), ...rec(0x0002, u32(0x0409)), ...rec(0x0014, u32(0x0409)), ...rec(0x0003, u16(1252)));
  add(...rec(0x0004, cp1252(projectName)), ...rec(0x0005, new Uint8Array()), ...rec(0x0040, new Uint8Array()));
  add(...rec(0x0006, new Uint8Array()), ...rec(0x003d, new Uint8Array()), ...rec(0x0007, u32(0)), ...rec(0x0008, u32(0)));
  { const h = new Uint8Array(6); const v = new DataView(h.buffer); v.setUint16(0, 0x0009, true); v.setUint32(2, 4, true); add(h, u32(1), u16(0)); }
  add(...rec(0x000c, new Uint8Array()), ...rec(0x003c, new Uint8Array()));
  if (modules.some((m) => m.form)) add(...msFormsReference(rec, u16, u32, utf16));
  add(...rec(0x000f, u16(modules.length)), ...rec(0x0013, u16(0xffff)));
  const streams = new Map<string, Uint8Array>();
  const clsids = new Map<string, string>();
  for (const m of modules) {
    add(...rec(0x0019, cp1252(m.name)), ...rec(0x0047, utf16(m.name)), ...rec(0x001a, cp1252(m.name)), ...rec(0x0032, utf16(m.name)));
    add(...rec(0x001c, new Uint8Array()), ...rec(0x0048, new Uint8Array()), ...rec(0x0031, u32(0)), ...rec(0x001e, u32(0)), ...rec(0x002c, u16(0xffff)));
    if (m.form) {
      // A UserForm: a designer module (type 0x0022, private to the project) with its designer storage.
      add(...rec(0x0022, new Uint8Array()), u16(0x0028), u32(0), u16(0x002b), u32(0));
      const base = /^Attribute VB_Base = "(0\{[^"]*\})"/im.exec(m.code)?.[1] ?? `0${newGuid()}${newGuid()}`;
      const head = [`Attribute VB_Name = "${m.name}"`, `Attribute VB_Base = "${base}"`, "Attribute VB_GlobalNameSpace = False", "Attribute VB_Creatable = False", "Attribute VB_PredeclaredId = True", "Attribute VB_Exposed = False", "Attribute VB_TemplateDerived = False", "Attribute VB_Customizable = False", ""].join("\r\n");
      const body = m.code.replace(/^(Attribute VB_[^\r\n]*\r?\n)+/i, "").replace(/\r?\n/g, "\r\n");
      streams.set(`VBA/${m.name}`, compress(cp1252(head + body)));
      const f = writeFormStorage({ ...m.form, name: m.name });
      for (const [k, v] of f.streams) streams.set(`${m.name}/${k}`, v);
      for (const [k, v] of f.clsids) clsids.set(k ? `${m.name}/${k}` : m.name, v);
      if (!f.clsids.has("")) clsids.set(m.name, CLSID_FORM);
      continue;
    }
    add(...rec(m.kind === "standard" ? 0x0021 : 0x0022, new Uint8Array()), ...rec(0x002b, new Uint8Array()));
    const head = m.kind === "standard" ? `Attribute VB_Name = "${m.name}"\r\n` : `Attribute VB_Name = "${m.name}"\r\nAttribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = False\r\nAttribute VB_Exposed = False\r\n`;
    const body = m.code.replace(/^(Attribute VB_[^\r\n]*\r?\n)+/i, "").replace(/\r?\n/g, "\r\n");
    streams.set(`VBA/${m.name}`, compress(cp1252(head + body)));
  }
  add(u16(0x0010), u32(0));
  const dir = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let o = 0;
  for (const p of parts) { dir.set(p, o); o += p.length; }
  streams.set("VBA/dir", compress(dir));
  // _VBA_PROJECT with version 0xFFFF: no compiled code, Office compiles the source.
  streams.set("VBA/_VBA_PROJECT", Uint8Array.of(0xcc, 0x61, 0xff, 0xff, 0x00, 0x00, 0x00));
  const guid = "{" + [8, 4, 4, 4, 12].map((n) => Array.from(crypto.getRandomValues(new Uint8Array(n / 2)), (x) => x.toString(16).padStart(2, "0")).join("").toUpperCase()).join("-") + "}";
  const items: string[] = [];
  for (const m of modules) {
    if (m.form && !items.some((l) => l.startsWith("Package="))) items.push("Package={AC9F2F90-E877-11CE-9F68-00AA00574A4F}");
    items.push(m.form ? `BaseClass=${m.name}` : m.kind === "standard" ? `Module=${m.name}` : `Class=${m.name}`);
  }
  const lines = [`ID="${guid}"`, ...items, `Name="${projectName}"`, `HelpContextID="0"`, `VersionCompatible32="393222000"`,
    "", "[Host Extender Info]", "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000", "", "[Workspace]", ...modules.map((m) => `${m.name}=0, 0, 0, 0, C`), ""];
  streams.set("PROJECT", cp1252(lines.join("\r\n")));
  const wm: Uint8Array[] = [];
  for (const m of modules) wm.push(cp1252(m.name + "\0"), utf16(m.name + "\0"));
  wm.push(u16(0));
  streams.set("PROJECTwm", Uint8Array.from(wm.flatMap((x) => Array.from(x))));
  return writeCfb(streams, clsids);
}
