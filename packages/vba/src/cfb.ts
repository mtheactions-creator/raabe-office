// Compound File Binary format (MS-CFB, version 3, 512-byte sectors), as Office keeps a VBA project
// (vbaProject.bin): read a file into its storages and streams, and write streams back into a new file (D-079).

export interface CfbEntry { path: string; data?: Uint8Array; storage: boolean }

const SIG = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const FREE = 0xffffffff, END = 0xfffffffe, FATSECT = 0xfffffffd, NOSTREAM = 0xffffffff;

/** The streams of a compound file by path ("VBA/dir", "PROJECT" …). */
export function readCfb(b: Uint8Array): Map<string, Uint8Array> {
  const v = new DataView(b.buffer, b.byteOffset, b.byteLength);
  for (let i = 0; i < 8; i++) if (b[i] !== SIG[i]) throw new Error("not a compound file");
  const ssz = 1 << v.getUint16(30, true), mssz = 1 << v.getUint16(32, true);
  const nFat = v.getUint32(44, true), dirStart = v.getUint32(48, true), cutoff = v.getUint32(56, true);
  const miniFatStart = v.getUint32(60, true), difatStart = v.getUint32(68, true);
  const sector = (n: number) => 512 + n * ssz;
  // The FAT's sectors: 109 in the header, the rest in DIFAT sectors.
  const fatSectors: number[] = [];
  for (let i = 0; i < 109 && fatSectors.length < nFat; i++) fatSectors.push(v.getUint32(76 + i * 4, true));
  for (let d = difatStart; d !== END && d !== FREE && fatSectors.length < nFat;) {
    const o = sector(d);
    for (let i = 0; i < ssz / 4 - 1 && fatSectors.length < nFat; i++) fatSectors.push(v.getUint32(o + i * 4, true));
    d = v.getUint32(o + ssz - 4, true);
  }
  const fat: number[] = [];
  for (const s of fatSectors) for (let i = 0; i < ssz / 4; i++) fat.push(v.getUint32(sector(s) + i * 4, true));
  const chain = (start: number, table: number[]) => { const out: number[] = []; for (let s = start; s !== END && s !== FREE && s < table.length && out.length <= table.length; s = table[s]) out.push(s); return out; };
  const read = (start: number, size: number) => {
    const out = new Uint8Array(size);
    let o = 0;
    for (const s of chain(start, fat)) { const n = Math.min(ssz, size - o); if (n <= 0) break; out.set(b.subarray(sector(s), sector(s) + n), o); o += n; }
    return out;
  };
  // Directory.
  const dirSecs = chain(dirStart, fat);
  const dir = new Uint8Array(dirSecs.length * ssz);
  dirSecs.forEach((s, i) => dir.set(b.subarray(sector(s), sector(s) + ssz), i * ssz));
  const dv = new DataView(dir.buffer);
  const entries = Array.from({ length: dir.length / 128 }, (_, i) => {
    const o = i * 128;
    const nl = dv.getUint16(o + 64, true);
    const name = String.fromCharCode(...Array.from({ length: Math.max(0, nl / 2 - 1) }, (_, k) => dv.getUint16(o + k * 2, true)));
    return { name, type: dir[o + 66], left: dv.getUint32(o + 68, true), right: dv.getUint32(o + 72, true), child: dv.getUint32(o + 76, true), start: dv.getUint32(o + 116, true), size: dv.getUint32(o + 120, true) };
  });
  // The mini stream (the root's data) and the mini FAT.
  const root = entries[0];
  const miniStream = read(root.start, root.size);
  const miniFat: number[] = [];
  for (const s of chain(miniFatStart, fat)) for (let i = 0; i < ssz / 4; i++) miniFat.push(v.getUint32(sector(s) + i * 4, true));
  const readMini = (start: number, size: number) => {
    const out = new Uint8Array(size);
    let o = 0;
    for (const s of chain(start, miniFat)) { const n = Math.min(mssz, size - o); if (n <= 0) break; out.set(miniStream.subarray(s * mssz, s * mssz + n), o); o += n; }
    return out;
  };
  const out = new Map<string, Uint8Array>();
  const walk = (id: number, prefix: string, seen = new Set<number>()) => {
    if (id === NOSTREAM || id >= entries.length || seen.has(id)) return;
    seen.add(id);
    const e = entries[id];
    walk(e.left, prefix, seen);
    walk(e.right, prefix, seen);
    const path = prefix ? `${prefix}/${e.name}` : e.name;
    if (e.type === 2) out.set(path, e.size < cutoff ? readMini(e.start, e.size) : read(e.start, e.size));
    else if (e.type === 1) walk(e.child, path, seen);
  };
  walk(root.child, "");
  return out;
}

/** Names compare as in MS-CFB: shorter first, then by upper case. */
function cmp(a: string, b: string): number {
  if (a.length !== b.length) return a.length - b.length;
  const A = a.toUpperCase(), B = b.toUpperCase();
  return A < B ? -1 : A > B ? 1 : 0;
}

/** A compound file with these streams (paths with "/" make storages). Small streams go into the mini stream. */
export function writeCfb(streams: Map<string, Uint8Array>): Uint8Array {
  const SSZ = 512, MSZ = 64, CUTOFF = 4096;
  // The tree of storages and streams.
  interface Node { name: string; type: 1 | 2 | 5; data?: Uint8Array; kids: Node[]; id?: number; start?: number }
  const root: Node = { name: "Root Entry", type: 5, kids: [] };
  for (const [path, data] of streams) {
    const parts = path.split("/");
    let cur = root;
    parts.forEach((p, i) => {
      let next = cur.kids.find((k) => k.name === p);
      if (!next) { next = { name: p, type: i === parts.length - 1 ? 2 : 1, kids: [], data: i === parts.length - 1 ? data : undefined }; cur.kids.push(next); }
      cur = next;
    });
  }
  const nodes: Node[] = [];
  const number = (n: Node) => { n.id = nodes.length; nodes.push(n); n.kids.sort((a, b) => cmp(a.name, b.name)); for (const k of n.kids) number(k); };
  number(root);
  // Mini stream: small streams in 64-byte mini sectors.
  const mini: number[] = [];
  const miniData: Uint8Array[] = [];
  let miniCount = 0;
  for (const n of nodes) if (n.type === 2 && n.data && n.data.length < CUTOFF) {
    const k = Math.ceil(n.data.length / MSZ);
    n.start = k ? miniCount : END;
    for (let i = 0; i < k; i++) mini.push(i === k - 1 ? END : miniCount + i + 1);
    const padded = new Uint8Array(k * MSZ); padded.set(n.data);
    miniData.push(padded);
    miniCount += k;
  }
  const miniStream = concat(miniData);
  // Regular sectors: big streams, the mini stream, the mini FAT, the directory, then the FAT.
  const sectors: Uint8Array[] = [];
  const fat: number[] = [];
  const alloc = (data: Uint8Array): number => {
    const k = Math.max(1, Math.ceil(data.length / SSZ));
    const start = sectors.length;
    for (let i = 0; i < k; i++) { const s = new Uint8Array(SSZ); s.set(data.subarray(i * SSZ, (i + 1) * SSZ)); sectors.push(s); fat.push(i === k - 1 ? END : start + i + 1); }
    return start;
  };
  for (const n of nodes) if (n.type === 2 && n.data && n.data.length >= CUTOFF) n.start = alloc(n.data);
  root.start = miniStream.length ? alloc(miniStream) : END;
  const miniFatBytes = new Uint8Array(Math.ceil((mini.length * 4) / SSZ) * SSZ).fill(0xff);
  new Uint32Array(miniFatBytes.buffer, 0, mini.length).set(mini);
  const miniFatStart = mini.length ? alloc(miniFatBytes) : END;
  // Directory entries: siblings as a balanced binary tree under their storage.
  const dir = new Uint8Array(Math.ceil((nodes.length * 128) / SSZ) * SSZ);
  const dv = new DataView(dir.buffer);
  for (let i = nodes.length; i < dir.length / 128; i++) { dv.setUint32(i * 128 + 68, NOSTREAM, true); dv.setUint32(i * 128 + 72, NOSTREAM, true); dv.setUint32(i * 128 + 76, NOSTREAM, true); }
  const tree = (kids: Node[]): number => {
    if (!kids.length) return NOSTREAM;
    const m = Math.floor(kids.length / 2);
    const n = kids[m];
    dv.setUint32(n.id! * 128 + 68, tree(kids.slice(0, m)), true);
    dv.setUint32(n.id! * 128 + 72, tree(kids.slice(m + 1)), true);
    return n.id!;
  };
  for (const n of nodes) {
    const o = n.id! * 128;
    for (let k = 0; k < Math.min(31, n.name.length); k++) dv.setUint16(o + k * 2, n.name.charCodeAt(k), true);
    dv.setUint16(o + 64, (Math.min(31, n.name.length) + 1) * 2, true);
    dir[o + 66] = n.type;
    dir[o + 67] = 1; // black
    dv.setUint32(o + 68, NOSTREAM, true); dv.setUint32(o + 72, NOSTREAM, true);
    dv.setUint32(o + 76, NOSTREAM, true);
    dv.setUint32(o + 116, n.type === 1 ? 0 : n.start ?? END, true);
    dv.setUint32(o + 120, n.type === 5 ? miniStream.length : n.data?.length ?? 0, true);
  }
  for (const n of nodes) if (n.type !== 2) dv.setUint32(n.id! * 128 + 76, tree(n.kids), true);
  const dirStart = alloc(dir);
  // FAT sectors (they hold their own entries too).
  let nFat = Math.ceil(fat.length / (SSZ / 4));
  while (Math.ceil((fat.length + nFat) / (SSZ / 4)) > nFat) nFat++;
  if (nFat > 109) throw new Error("VBA project too large");
  const fatStart = sectors.length;
  for (let i = 0; i < nFat; i++) fat.push(FATSECT);
  const fatBytes = new Uint8Array(nFat * SSZ).fill(0xff);
  new Uint32Array(fatBytes.buffer, 0, fat.length).set(fat);
  for (let i = 0; i < nFat; i++) sectors.push(fatBytes.subarray(i * SSZ, (i + 1) * SSZ));
  // Header.
  const head = new Uint8Array(512);
  const hv = new DataView(head.buffer);
  head.set(SIG);
  hv.setUint16(24, 0x3e, true); hv.setUint16(26, 3, true); hv.setUint16(28, 0xfffe, true); hv.setUint16(30, 9, true); hv.setUint16(32, 6, true);
  hv.setUint32(44, nFat, true); hv.setUint32(48, dirStart, true); hv.setUint32(56, CUTOFF, true);
  hv.setUint32(60, miniFatStart, true); hv.setUint32(64, mini.length ? Math.ceil((mini.length * 4) / SSZ) : 0, true);
  hv.setUint32(68, END, true); hv.setUint32(72, 0, true);
  for (let i = 0; i < 109; i++) hv.setUint32(76 + i * 4, i < nFat ? fatStart + i : FREE, true);
  return concat([head, ...sectors]);
}

function concat(parts: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let o = 0;
  for (const p of parts) { out.set(p, o); o += p.length; }
  return out;
}
