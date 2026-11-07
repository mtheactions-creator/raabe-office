// A VBA project (vbaProject.bin) as MS-OVBA describes it (D-079): the compression of its streams, the dir stream
// with the modules, and each module's source code. Read: the modules of a project from PowerPoint, Excel or
// Word. Written: a project with our modules (no compiled code, so Office compiles the source when it opens it).

import { readCfb, writeCfb } from "./cfb";

export interface VbaModule { name: string; kind: "standard" | "class" | "document"; code: string }

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
  // Document modules (Diese Arbeitsmappe, Tabelle1 …) are classes named in the PROJECT stream as Document=.
  const project = find("PROJECT");
  const docs = new Set(project ? latin1(project).split(/\r?\n/).filter((l) => l.startsWith("Document=")).map((l) => l.slice(9).split("/")[0]) : []);
  return mods.map((m) => {
    const s = find(`VBA/${m.stream}`);
    const code = s && s.length > m.offset ? fromCp1252(decompress(s.subarray(m.offset))) : "";
    return { name: m.name, kind: docs.has(m.name) ? "document" : m.kind, code };
  });
}

/** The source without its Attribute lines (what the editor shows). */
export function visibleCode(code: string): string {
  return code.split(/\r?\n/).filter((l) => !/^Attribute VB_/i.test(l)).join("\n").replace(/^\n+/, "");
}

/** A project with these modules (standard and class modules; document modules are kept as classes). */
export function writeProject(modules: VbaModule[], projectName = "VBAProject"): Uint8Array {
  const rec = (id: number, data: Uint8Array) => { const h = new Uint8Array(6); const v = new DataView(h.buffer); v.setUint16(0, id, true); v.setUint32(2, data.length, true); return [h, data]; };
  const u32 = (n: number) => { const b = new Uint8Array(4); new DataView(b.buffer).setUint32(0, n, true); return b; };
  const u16 = (n: number) => { const b = new Uint8Array(2); new DataView(b.buffer).setUint16(0, n, true); return b; };
  const utf16 = (s: string) => { const b = new Uint8Array(s.length * 2); const v = new DataView(b.buffer); for (let k = 0; k < s.length; k++) v.setUint16(k * 2, s.charCodeAt(k), true); return b; };
  const parts: Uint8Array[] = [];
  const add = (...p: Uint8Array[]) => parts.push(...p);
  add(...rec(0x0001, u32(1)), ...rec(0x0002, u32(0x0407)), ...rec(0x0014, u32(0x0407)), ...rec(0x0003, u16(1252)));
  add(...rec(0x0004, cp1252(projectName)), ...rec(0x0005, new Uint8Array()), ...rec(0x0040, new Uint8Array()));
  add(...rec(0x0006, new Uint8Array()), ...rec(0x003d, new Uint8Array()), ...rec(0x0007, u32(0)), ...rec(0x0008, u32(0)));
  { const h = new Uint8Array(6); const v = new DataView(h.buffer); v.setUint16(0, 0x0009, true); v.setUint32(2, 4, true); add(h, u32(1), u16(0)); }
  add(...rec(0x000c, new Uint8Array()), ...rec(0x003c, new Uint8Array()));
  add(...rec(0x000f, u16(modules.length)), ...rec(0x0013, u16(0xffff)));
  const streams = new Map<string, Uint8Array>();
  for (const m of modules) {
    add(...rec(0x0019, cp1252(m.name)), ...rec(0x0047, utf16(m.name)), ...rec(0x001a, cp1252(m.name)), ...rec(0x0032, utf16(m.name)));
    add(...rec(0x001c, new Uint8Array()), ...rec(0x0048, new Uint8Array()), ...rec(0x0031, u32(0)), ...rec(0x001e, u32(0)), ...rec(0x002c, u16(0xffff)));
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
  const lines = [`ID="${guid}"`, ...modules.map((m) => (m.kind === "standard" ? `Module=${m.name}` : `Class=${m.name}`)), `Name="${projectName}"`, `HelpContextID="0"`, `VersionCompatible32="393222000"`,
    "", "[Host Extender Info]", "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000", "", "[Workspace]", ...modules.map((m) => `${m.name}=0, 0, 0, 0, C`), ""];
  streams.set("PROJECT", cp1252(lines.join("\r\n")));
  const wm: Uint8Array[] = [];
  for (const m of modules) wm.push(cp1252(m.name + "\0"), utf16(m.name + "\0"));
  wm.push(u16(0));
  streams.set("PROJECTwm", Uint8Array.from(wm.flatMap((x) => Array.from(x))));
  return writeCfb(streams);
}
