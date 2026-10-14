// Equations: Office Math (OMML, m:oMath) in .docx and the equation tree the formula editor
// edits (packages/text-engine/src/math). The converter covers what school mathematics uses:
// fractions, scripts, roots, brackets, sums and integrals, functions, limits, accents,
// matrices and case distinctions. Anything else read from a file keeps its XML (the equation
// is shown but not editable then). Also the document's Formeloptionen (m:mathPr).

export const M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";

import {
  DEFAULT_MATH, lines, mergeText, parseLatex, seqToLatex, seqToMathml, type MathSettings, type MNode, type Seq,
} from "@raabe/text-engine";

export { parseLatex };

/** MathML for showing an equation (kept for callers that have LaTeX). */
export function latexToMathml(latex: string, display: boolean): string {
  return seqToMathml(parseLatex(latex), display);
}

// ----- AST → OMML ------------------------------------------------------------------------------

const xesc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const RPR = '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>';
const val = (tag: string, v: string) => `<m:${tag} m:val="${xesc(v)}"/>`;

function ommlRun(n: Extract<MNode, { k: "t" }>): string {
  let pr = "";
  if (n.nor) pr = `<m:nor/>`;
  else if (n.scr || n.sty) pr = (n.scr ? val("scr", n.scr) : "") + (n.sty ? val("sty", n.sty) : "");
  const wpr = n.nor
    ? (n.sty === "b" ? "<w:rPr><w:b/></w:rPr>" : n.sty === "i" ? "<w:rPr><w:i/></w:rPr>" : "")
    : RPR;
  return `<m:r>${pr ? `<m:rPr>${pr}</m:rPr>` : ""}${wpr}<m:t xml:space="preserve">${xesc(n.s)}</m:t></m:r>`;
}

const e = (s: Seq | undefined, tag = "e") => (s && s.length ? `<m:${tag}>${omml(s)}</m:${tag}>` : `<m:${tag}/>`);

function omml(s: Seq): string {
  let out = "";
  for (const n of s) {
    switch (n.k) {
      case "t": out += ommlRun(n); break;
      case "frac": out += `<m:f>${n.type ? `<m:fPr>${val("type", n.type)}</m:fPr>` : ""}${e(n.num, "num")}${e(n.den, "den")}</m:f>`; break;
      case "scr":
        if (n.pre) out += `<m:sPre>${e(n.sub, "sub")}${e(n.sup, "sup")}${e(n.base)}</m:sPre>`;
        else if (n.sub && n.sup) out += `<m:sSubSup>${n.aln ? `<m:sSubSupPr>${val("alnScr", "1")}</m:sSubSupPr>` : ""}${e(n.base)}${e(n.sub, "sub")}${e(n.sup, "sup")}</m:sSubSup>`;
        else if (n.sup) out += `<m:sSup>${e(n.base)}${e(n.sup, "sup")}</m:sSup>`;
        else out += `<m:sSub>${e(n.base)}${e(n.sub, "sub")}</m:sSub>`;
        break;
      // A root without a degree hides the slot (\sqrt{x}); an empty degree is a slot to fill.
      case "rad": out += `<m:rad>${n.deg ? "" : `<m:radPr>${val("degHide", "1")}</m:radPr>`}${e(n.deg, "deg")}${e(n.e)}</m:rad>`; break;
      case "d": {
        const pr = (n.beg !== "(" ? val("begChr", n.beg) : "") + (n.end !== ")" ? val("endChr", n.end) : "") +
          (n.noGrow ? val("grow", "0") : "") + (n.shp ? val("shp", "match") : "");
        out += `<m:d>${pr ? `<m:dPr>${pr}</m:dPr>` : ""}${n.items.map((x) => e(x)).join("")}</m:d>`;
        break;
      }
      case "nary": {
        // Word writes limLoc only where it differs from the document's default (Formeloptionen:
        // limits above and below a sum, beside an integral).
        const integral = /[∫∬∭∮∯∰∱∲∳]/.test(n.chr);
        const und = (integral ? limits.intLim : limits.naryLim) === "undOvr";
        const pr = (n.chr !== "∫" ? val("chr", n.chr) : "") + (n.und === und ? "" : val("limLoc", n.und ? "undOvr" : "subSup")) +
          (n.grow ? val("grow", "1") : "") + (n.sub ? "" : val("subHide", "1")) + (n.sup ? "" : val("supHide", "1"));
        out += `<m:nary><m:naryPr>${pr}</m:naryPr>${e(n.sub, "sub")}${e(n.sup, "sup")}${e(n.e)}</m:nary>`;
        break;
      }
      case "func": out += `<m:func>${e(n.name, "fName")}${e(n.e)}</m:func>`; break;
      // Word leaves out the circumflex, the default accent.
      case "acc": out += `<m:acc>${n.chr !== "̂" ? `<m:accPr>${val("chr", n.chr)}</m:accPr>` : ""}${e(n.e)}</m:acc>`; break;
      case "bar": out += `<m:bar>${n.top ? `<m:barPr>${val("pos", "top")}</m:barPr>` : ""}${e(n.e)}</m:bar>`; break;
      case "lim": out += `<m:${n.upper ? "limUpp" : "limLow"}>${e(n.e)}${e(n.lim, "lim")}</m:${n.upper ? "limUpp" : "limLow"}>`; break;
      case "group": {
        const pr = (n.chr !== "⏟" ? val("chr", n.chr) : "") + (n.top ? val("pos", "top") + val("vertJc", "bot") : "");
        out += `<m:groupChr>${pr ? `<m:groupChrPr>${pr}</m:groupChrPr>` : ""}${e(n.e)}</m:groupChr>`;
        break;
      }
      case "m": {
        const cols = Math.max(1, ...n.rows.map((r) => r.length));
        const mcs = n.mcJc?.length
          ? Array.from({ length: cols }, (_x, ci) => `<m:mc><m:mcPr>${val("count", "1")}${val("mcJc", n.mcJc![ci] ?? "center")}</m:mcPr></m:mc>`).join("")
          : `<m:mc><m:mcPr>${val("count", String(cols))}${val("mcJc", "center")}</m:mcPr></m:mc>`;
        out += `<m:m><m:mPr>${n.baseJc ? val("baseJc", n.baseJc) : ""}${n.plcHide ? val("plcHide", "1") : ""}<m:mcs>${mcs}</m:mcs></m:mPr>` +
          n.rows.map((r) => `<m:mr>${Array.from({ length: cols }, (_x, i) => e(r[i])).join("")}</m:mr>`).join("") + "</m:m>";
        break;
      }
      case "eqArr": out += `<m:eqArr>${n.rows.map((r) => e(r)).join("")}</m:eqArr>`; break;
      case "box": out += `<m:borderBox>${e(n.e)}</m:borderBox>`; break;
    }
  }
  return out;
}

/** OMML for an equation: m:oMathPara around it when it stands alone on its line. */
export function latexToOmml(latex: string, display: boolean): string {
  return seqToOmml(parseLatex(latex), display);
}

/**
 * OMML for an equation tree. A display equation with several lines (Umschalt+Eingabe) is
 * one m:oMath per line in its m:oMathPara, as Word writes it; inline, an equation array.
 */
export function seqToOmml(seq: Seq, display: boolean, opts: MathOptions = {}): string {
  limits = { intLim: opts.intLim ?? "subSup", naryLim: opts.naryLim ?? "undOvr" };
  // Only an equation's own alignment is written (Word: none means the document's default).
  const pr = opts.jc ? `<m:oMathParaPr>${val("jc", opts.jc)}</m:oMathParaPr>` : "";
  const rows = lines(mergeText(seq));
  // Lines with alignment marks (&) are an equation array (aligned at the marks).
  const aligned = rows.length > 1 && rows.some((r) => r.some((n) => n.k === "t" && n.s.includes("&")));
  if (display && !aligned) return `<m:oMathPara>${pr}${rows.map((r) => `<m:oMath>${omml(r)}</m:oMath>`).join("")}</m:oMathPara>`;
  const body = rows.length > 1 ? omml([{ k: "eqArr", rows }]) : omml(rows[0]);
  const math = `<m:oMath>${body}</m:oMath>`;
  return display ? `<m:oMathPara>${pr}${math}</m:oMathPara>` : math;
}

/** The document's Formeloptionen that decide how an equation is written and read. */
export interface MathOptions {
  intLim?: "subSup" | "undOvr";
  naryLim?: "subSup" | "undOvr";
  jc?: string;
}

/** Where limits go when an equation does not say (the document's Formeloptionen). */
let limits: { intLim: string; naryLim: string } = { intLim: "subSup", naryLim: "undOvr" };

// ----- OMML → tree -------------------------------------------------------------------------------

class Unsupported extends Error {}

const kidsOf = (el: Element) => Array.from(el.children);
const mchild = (el: Element, name: string) => kidsOf(el).find((c) => c.localName === name);
const mval = (el: Element | undefined, name: string): string | null => {
  const c = el ? mchild(el, name) : undefined;
  if (!c) return null;
  return c.getAttributeNS(M_NS, "val") ?? c.getAttribute("m:val") ?? "";
};
const on = (v: string | null) => v !== null && v !== "0" && v !== "off" && v !== "false";

/** The children of an argument element (m:e, m:num …) as a sequence. */
function seqOf(el: Element | undefined): Seq {
  if (!el) return [];
  const out: Seq = [];
  for (const c of kidsOf(el)) out.push(...nodeOf(c));
  return mergeText(out);
}

const STYLES = new Set(["p", "b", "i", "bi"]);

function nodeOf(el: Element): Seq {
  const name = el.localName;
  if (el.namespaceURI !== M_NS) {
    if (name === "bookmarkStart" || name === "bookmarkEnd" || name === "proofErr" || name === "rPr") return [];
    if (name === "r") return [{ k: "t", s: Array.from(el.getElementsByTagName("*")).filter((x) => x.localName === "t").map((x) => x.textContent ?? "").join(""), nor: true }];
    throw new Unsupported(name);
  }
  const e = (n = "e") => seqOf(mchild(el, n));
  switch (name) {
    case "r": {
      const t = kidsOf(el).filter((x) => x.localName === "t").map((x) => x.textContent ?? "").join("");
      const pr = mchild(el, "rPr");
      const node: Extract<MNode, { k: "t" }> = { k: "t", s: t };
      if (pr && mchild(pr, "nor")) node.nor = true;
      const sty = mval(pr, "sty"), scr = mval(pr, "scr");
      if (sty && STYLES.has(sty)) node.sty = sty as "p" | "b" | "i" | "bi";
      if (scr) node.scr = scr;
      // Normal text (m:nor) takes bold/italic from the Word run properties.
      if (node.nor) {
        const wpr = mchild(el, "rPr") && kidsOf(el).find((x) => x.localName === "rPr" && x.namespaceURI !== M_NS);
        if (wpr && kidsOf(wpr).some((x) => x.localName === "b")) node.sty = "b";
        else if (wpr && kidsOf(wpr).some((x) => x.localName === "i")) node.sty = "i";
      }
      return t ? [node] : [];
    }
    case "f": {
      const type = mval(mchild(el, "fPr"), "type");
      return [{ k: "frac", num: e("num"), den: e("den"), type: type === "lin" || type === "noBar" || type === "skw" ? type : undefined }];
    }
    case "sSup": return [{ k: "scr", base: e(), sup: e("sup") }];
    case "sSub": return [{ k: "scr", base: e(), sub: e("sub") }];
    case "sSubSup": return [{ k: "scr", base: e(), sub: e("sub"), sup: e("sup"), ...(on(mval(mchild(el, "sSubSupPr"), "alnScr")) ? { aln: true } : {}) }];
    case "sPre": return [{ k: "scr", base: e(), sub: e("sub"), sup: e("sup"), pre: true }];
    case "rad": {
      const hide = on(mval(mchild(el, "radPr"), "degHide"));
      return [{ k: "rad", deg: hide ? undefined : e("deg"), e: e() }];
    }
    case "d": {
      const pr = mchild(el, "dPr");
      const items = kidsOf(el).filter((x) => x.localName === "e").map((x) => seqOf(x));
      const grow = mval(pr, "grow");
      return [{
        k: "d", beg: mval(pr, "begChr") ?? "(", end: mval(pr, "endChr") ?? ")", items: items.length ? items : [[]],
        ...(grow !== null && !on(grow) ? { noGrow: true } : {}), ...(mval(pr, "shp") === "match" ? { shp: "match" as const } : {}),
      }];
    }
    case "nary": {
      const pr = mchild(el, "naryPr");
      const chr = mval(pr, "chr") ?? "∫";
      const loc = mval(pr, "limLoc");
      const und = (loc ?? (chr.match(/[∫∬∭∮∯∰∱∲∳]/) ? limits.intLim : limits.naryLim)) === "undOvr";
      return [{
        k: "nary", chr, und, ...(on(mval(pr, "grow")) ? { grow: true } : {}),
        sub: on(mval(pr, "subHide")) ? undefined : e("sub"), sup: on(mval(pr, "supHide")) ? undefined : e("sup"), e: e(),
      }];
    }
    case "func": return [{ k: "func", name: e("fName"), e: e() }];
    case "acc": return [{ k: "acc", chr: mval(mchild(el, "accPr"), "chr") ?? "̂", e: e() }];
    case "bar": return [{ k: "bar", top: mval(mchild(el, "barPr"), "pos") === "top", e: e() }];
    case "limLow": case "limUpp": return [{ k: "lim", upper: name === "limUpp", e: e(), lim: e("lim") }];
    case "groupChr": {
      const pr = mchild(el, "groupChrPr");
      return [{ k: "group", chr: mval(pr, "chr") ?? "⏟", top: (mval(pr, "pos") ?? "bot") === "top", e: e() }];
    }
    case "m": {
      const pr = mchild(el, "mPr");
      const rows = kidsOf(el).filter((x) => x.localName === "mr").map((r) => kidsOf(r).filter((x) => x.localName === "e").map((x) => seqOf(x)));
      // Spaltenausrichtung: m:mc groups of columns, each with its count.
      const jc: ("left" | "center" | "right")[] = [];
      const mcs = pr ? mchild(pr, "mcs") : undefined;
      for (const mc of mcs ? kidsOf(mcs).filter((x) => x.localName === "mc") : []) {
        const mcPr = mchild(mc, "mcPr");
        const count = Number(mval(mcPr, "count") ?? "1") || 1;
        const v = mval(mcPr, "mcJc");
        for (let k = 0; k < count; k++) jc.push(v === "left" || v === "right" ? v : "center");
      }
      const base = mval(pr, "baseJc");
      return [{
        k: "m", rows, ...(on(mval(pr, "plcHide")) ? { plcHide: true } : {}), ...(base === "top" || base === "bot" ? { baseJc: base } : {}),
        ...(jc.some((x) => x !== "center") ? { mcJc: jc } : {}),
      }];
    }
    case "eqArr": return [{ k: "eqArr", rows: kidsOf(el).filter((x) => x.localName === "e").map((x) => seqOf(x)) }];
    case "borderBox": return [{ k: "box", e: e() }];
    case "box": case "phant": return e();
    case "e": return seqOf(el);
    case "ctrlPr": case "argPr": case "oMathParaPr": return [];
    case "oMath": return seqOf(el);
    default: throw new Unsupported(name);
  }
}

/**
 * The tree of an m:oMath or m:oMathPara element (lines of a display equation joined by line
 * breaks), or null when it uses something the editor does not know.
 */
export function ommlToSeq(el: Element, opts: MathOptions = {}): Seq | null {
  limits = { intLim: opts.intLim ?? "subSup", naryLim: opts.naryLim ?? "undOvr" };
  try {
    if (el.localName === "oMathPara") {
      const out: Seq = [];
      kidsOf(el).filter((x) => x.localName === "oMath").forEach((x, i) => { if (i) out.push({ k: "t", s: "\n" }); out.push(...seqOf(x)); });
      return out;
    }
    return seqOf(el);
  } catch (err) {
    if (err instanceof Unsupported) return null;
    throw err;
  }
}

/** LaTeX of an m:oMath or m:oMathPara element, or null when it uses something unsupported. */
export function ommlToLatex(el: Element): string | null {
  const seq = ommlToSeq(el);
  return seq ? seqToLatex(seq) : null;
}

// ----- Formeloptionen (m:mathPr in settings.xml) ---------------------------------------------

const mkid = (el: Element | null, name: string) => (el ? Array.from(el.children).find((c) => c.namespaceURI === M_NS && c.localName === name) ?? null : null);
const mv = (el: Element | null) => (el ? el.getAttributeNS(M_NS, "val") || el.getAttribute("m:val") : null);
const twipsPt = (v: string | null, d: number) => (v !== null && v !== "" && Number.isFinite(Number(v)) ? Number(v) / 20 : d);
const isOn = (el: Element | null, d: boolean) => (el ? !["0", "off", "false"].includes(mv(el) ?? "1") : d);
const oneOf = <T extends string>(v: string | null, list: readonly T[], d: T): T => (list.includes(v as T) ? (v as T) : d);

/** The document's Formeloptionen, from the m:mathPr in w:settings (undefined when there is none). */
export function readMathPr(settings: Element): MathSettings | undefined {
  const pr = mkid(settings, "mathPr");
  if (!pr) return undefined;
  const d = DEFAULT_MATH;
  const k = (name: string) => mkid(pr, name);
  return {
    font: mv(k("mathFont")) ?? d.font,
    smallFrac: isOn(k("smallFrac"), false),
    brkBin: oneOf(mv(k("brkBin")), ["before", "after", "repeat"] as const, d.brkBin),
    brkBinSub: oneOf(mv(k("brkBinSub")), ["--", "-+", "+-"] as const, d.brkBinSub),
    dispDef: isOn(k("dispDef"), false),
    lMargin: twipsPt(mv(k("lMargin")), 0),
    rMargin: twipsPt(mv(k("rMargin")), 0),
    defJc: oneOf(mv(k("defJc")), ["left", "right", "center", "centerGroup"] as const, d.defJc),
    wrapIndent: twipsPt(mv(k("wrapIndent")), d.wrapIndent),
    wrapRight: isOn(k("wrapRight"), false),
    intLim: oneOf(mv(k("intLim")), ["subSup", "undOvr"] as const, d.intLim),
    naryLim: oneOf(mv(k("naryLim")), ["subSup", "undOvr"] as const, d.naryLim),
  };
}

/** m:mathPr as Word writes it. */
export function mathPrXml(m: MathSettings): string {
  const tw = (pt: number) => String(Math.round(pt * 20));
  return `<m:mathPr xmlns:m="${M_NS}">${val("mathFont", m.font)}${val("brkBin", m.brkBin)}${val("brkBinSub", m.brkBinSub)}` +
    `${val("smallFrac", m.smallFrac ? "1" : "0")}${m.dispDef ? "<m:dispDef/>" : val("dispDef", "0")}${val("lMargin", tw(m.lMargin))}${val("rMargin", tw(m.rMargin))}` +
    `${val("defJc", m.defJc)}${m.wrapRight ? "<m:wrapRight/>" : val("wrapIndent", tw(m.wrapIndent))}${val("intLim", m.intLim)}${val("naryLim", m.naryLim)}</m:mathPr>`;
}

/** A display equation's own alignment (m:oMathParaPr/m:jc). */
export function mathParaJc(el: Element): "left" | "right" | "center" | "centerGroup" | undefined {
  if (el.localName !== "oMathPara") return undefined;
  const jc = mv(mkid(mkid(el, "oMathParaPr"), "jc"));
  return jc === "left" || jc === "right" || jc === "center" || jc === "centerGroup" ? jc : undefined;
}
