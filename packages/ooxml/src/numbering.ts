// numbering.xml: list definitions (w:abstractNum) and their instances (w:num). Definitions
// read from a file keep their XML and are written back unchanged; new ones are generated.

import type { AbstractNum, NumInstance, NumLevel, Numbering, StyleSheet } from "@raabe/text-engine";
import { readPPr, readRPr, writePPr, writeRPr, type WriteOptions } from "./props";
import { attr, child, esc, kids, num, onOff, outer, parse, W } from "./xml";

function readLevel(el: Element): NumLevel {
  const jc = attr(child(el, "lvlJc"), "val");
  const suff = attr(child(el, "suff"), "val");
  const restart = num(attr(child(el, "lvlRestart"), "val"));
  return {
    start: num(attr(child(el, "start"), "val")) ?? 0,
    format: attr(child(el, "numFmt"), "val") ?? "decimal",
    text: attr(child(el, "lvlText"), "val") ?? "",
    align: jc === "center" ? "center" : jc === "right" || jc === "end" ? "right" : "left",
    suffix: suff === "space" ? "space" : suff === "nothing" ? "nothing" : "tab",
    pPr: readPPr(child(el, "pPr")).props,
    rPr: readRPr(child(el, "rPr")),
    restart,
    legal: onOff(child(el, "isLgl")) || undefined,
    data: outer(el),
  };
}

/** Reads numbering.xml. Lists defined through a list style (numStyleLink) get its levels. */
export function readNumbering(xml: string | null, styles: StyleSheet): Numbering {
  const abstracts = new Map<number, AbstractNum>();
  const nums = new Map<number, NumInstance>();
  if (!xml) return { abstracts, nums };
  const root = parse(xml).documentElement;
  const extra: string[] = [];
  const links = new Map<number, string>();
  for (const el of kids(root)) {
    if (el.namespaceURI !== W) { extra.push(`tail\u0000${outer(el)}`); continue; }
    if (el.localName === "abstractNum") {
      const id = num(attr(el, "abstractNumId")) ?? 0;
      const levels: NumLevel[] = [];
      for (const lvl of kids(el).filter((c) => c.localName === "lvl")) levels[num(attr(lvl, "ilvl")) ?? levels.length] = readLevel(lvl);
      const link = attr(child(el, "numStyleLink"), "val");
      if (link) links.set(id, link);
      abstracts.set(id, { id, levels, data: outer(el) });
    } else if (el.localName === "num") {
      const id = num(attr(el, "numId")) ?? 0;
      const overrides = new Map<number, { start?: number; level?: NumLevel }>();
      for (const o of kids(el).filter((c) => c.localName === "lvlOverride")) {
        const lvl = child(o, "lvl");
        overrides.set(num(attr(o, "ilvl")) ?? 0, {
          start: num(attr(child(o, "startOverride"), "val")),
          level: lvl ? readLevel(lvl) : undefined,
        });
      }
      nums.set(id, { id, abstractId: num(attr(child(el, "abstractNumId"), "val")) ?? 0, overrides: overrides.size ? overrides : undefined, data: outer(el) });
    } else if (el.localName === "numPicBullet") {
      extra.push(`head\u0000${outer(el)}`);
    } else {
      extra.push(`tail\u0000${outer(el)}`);
    }
  }
  // A definition that links to a list style uses the levels of the list the style points to.
  for (const [id, styleId] of links) {
    const st = styles.get(styleId);
    const numId = st?.pPr.numId;
    const target = numId !== undefined ? nums.get(numId) : undefined;
    const levels = target ? abstracts.get(target.abstractId)?.levels : undefined;
    const abs = abstracts.get(id)!;
    if (levels && levels.length) abstracts.set(id, { ...abs, levels });
  }
  return { abstracts, nums, extra: extra.length ? extra : undefined };
}

const SUFFIX: Record<string, string> = { space: "space", nothing: "nothing" };

function levelXml(l: NumLevel, ilvl: number, opts: WriteOptions): string {
  return `<w:lvl w:ilvl="${ilvl}"><w:start w:val="${l.start}"/><w:numFmt w:val="${esc(l.format)}"/>` +
    (l.restart !== undefined ? `<w:lvlRestart w:val="${l.restart}"/>` : "") +
    (l.legal ? "<w:isLgl/>" : "") +
    (SUFFIX[l.suffix] ? `<w:suff w:val="${SUFFIX[l.suffix]}"/>` : "") +
    `<w:lvlText w:val="${esc(l.text)}"/><w:lvlJc w:val="${l.align}"/>${writePPr(l.pPr, {}, opts)}${writeRPr(l.rPr, opts)}</w:lvl>`;
}

/** Writes numbering.xml: unchanged parts as read, new definitions generated. */
export function writeNumbering(n: Numbering, rootTag: string | null, opts: WriteOptions): string {
  const extra = (n.extra ?? []).map(String);
  const head = extra.filter((e) => e.startsWith("head\u0000")).map((e) => e.slice(5));
  const tail = extra.filter((e) => e.startsWith("tail\u0000")).map((e) => e.slice(5));
  const abs = [...n.abstracts.values()].sort((a, b) => a.id - b.id).map((a) => {
    if (typeof a.data === "string") return a.data;
    const mlt = a.levels.length > 1 ? (a.levels.some((l, i) => i > 0 && l.text.includes("%1")) ? "multilevel" : "hybridMultilevel") : "singleLevel";
    return `<w:abstractNum w:abstractNumId="${a.id}"><w:multiLevelType w:val="${mlt}"/>${a.levels.map((l, i) => levelXml(l, i, opts)).join("")}</w:abstractNum>`;
  });
  const nums = [...n.nums.values()].sort((a, b) => a.id - b.id).map((x) => {
    if (typeof x.data === "string") return x.data;
    const ov = [...(x.overrides ?? [])].map(([ilvl, o]) =>
      `<w:lvlOverride w:ilvl="${ilvl}">${o.start !== undefined ? `<w:startOverride w:val="${o.start}"/>` : ""}${o.level ? levelXml(o.level, ilvl, opts) : ""}</w:lvlOverride>`).join("");
    return `<w:num w:numId="${x.id}"><w:abstractNumId w:val="${x.abstractId}"/>${ov}</w:num>`;
  });
  const root = rootTag ?? `<w:numbering xmlns:w="${W}">`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${root}${head.join("")}${abs.join("")}${nums.join("")}${tail.join("")}</w:numbering>`;
}
