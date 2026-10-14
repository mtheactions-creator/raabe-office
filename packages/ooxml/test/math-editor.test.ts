// The formula editor against Word: every case of tools/feel-lab/catalogue_math.py replayed on
// the editor logic, the result written as OMML and normalised exactly like the catalogue does,
// and compared with what Word wrote (results/<date>-catalogue-math.json).

import { existsSync, readFileSync, readdirSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { MathEditor, type Seq } from "@raabe/text-engine";
import { seqToOmml } from "../src/math";
import { M_NS, norm, repr } from "./omml-norm";

const RESULTS = join(__dirname, "..", "..", "..", "tools", "feel-lab", "results");
/** The catalogue's description of a paragraph: text pieces and equations in order. */
function describePara(parts: ({ text: string } | { math: Seq })[]): string {
  const display = parts.length === 1 && "math" in parts[0];
  const out: string[] = [];
  for (const p of parts) {
    if ("text" in p) { if (p.text) out.push(repr(p.text)); continue; }
    const xml = seqToOmml(p.math, display);
    const doc = new DOMParser().parseFromString(`<r xmlns:m="${M_NS}" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${xml}</r>`, "application/xml");
    const root = doc.documentElement.firstElementChild!;
    if (root.localName === "oMathPara") out.push(`display[${Array.from(root.getElementsByTagNameNS(M_NS, "oMath")).map((m) => norm(m)).join(" | ")}]`);
    else out.push(norm(root));
  }
  return out.join(" ");
}

/** Replays a case like the host app does: an equation in an empty paragraph, then the keys. */
function replay(steps: [string, string, number?][]): string[] {
  const ed = new MathEditor([]);
  let where = "math" as "math" | "before" | "after" | "gone";
  let before = "", after = "";
  const paras: string[] = [];
  const typeText = (ch: string) => {
    if (where === "math") { ed.type(ch); return; }
    if (where === "before") before += ch; else after += ch;
  };
  const key = (k: string) => {
    if (where !== "math") return;
    const sig = (() => {
      switch (k) {
        case "left": return ed.left();
        case "right": return ed.right();
        case "shift+left": return ed.left(true);
        case "shift+right": return ed.right(true);
        case "up": return ed.vertical(-1);
        case "down": return ed.vertical(1);
        case "home": ed.home(); return {};
        case "end": ed.end(); return {};
        case "shift+home": ed.home(true); return {};
        case "shift+end": ed.end(true); return {};
        case "ctrl+right": return ed.word(1);
        case "ctrl+left": return ed.word(-1);
        case "backspace": return ed.backspace();
        case "delete": return ed.del();
        case "shift+enter": ed.newLine(); return {};
        case "enter": return { enter: true };
        default: throw new Error(`key ${k}`);
      }
    })() as { leave?: string; remove?: boolean; enter?: boolean };
    if (sig.leave) where = sig.leave === "before" ? "before" : "after";
    if (sig.remove) where = "gone";
    if (sig.enter) { paras.push("enter"); where = "after"; }
  };
  for (const [op, v, n] of steps) {
    if (op === "type") for (const ch of v) typeText(ch);
    else for (let i = 0; i < (n ?? 1); i++) key(v);
  }
  typeText("§");
  if (where === "gone") return after || before ? [repr(before + after)] : [];
  const eq = { math: ed.value() };
  if (paras.length) return [describePara([eq]), describePara([{ text: after }])];
  return [describePara([{ text: before }, eq, { text: after }].filter((p) => !("text" in p) || p.text))];
}

const latest = existsSync(RESULTS) ? readdirSync(RESULTS).filter((f) => f.endsWith("-catalogue-math.json")).sort().pop() : undefined;

describe.runIf(!!latest)("formula editor as Word (catalogue_math.py)", () => {
  const word = JSON.parse(readFileSync(join(RESULTS, latest!), "utf-8")) as Record<string, { word?: string[] }>;
  const cases = JSON.parse(readFileSync(join(RESULTS, "math-cases.json"), "utf-8")) as { name: string; steps: [string, string, number?][] }[];
  it("behaves like Word in every recorded case", () => {
    const diffs: string[] = [];
    let ok = 0;
    for (const c of cases) {
      const w = word[c.name]?.word;
      if (!w) continue;
      const want = w.map((x) => x.replace(/ ''\)/g, ")"));
      let got: string[];
      try { got = replay(c.steps); } catch (e) { got = [`ERROR ${String(e)}`]; }
      if (JSON.stringify(got) === JSON.stringify(want)) ok++;
      else diffs.push(`${c.name}\n   word: ${want.join(" / ")}\n   ours: ${got.join(" / ")}`);
    }
    console.log(`${ok}/${ok + diffs.length} as in Word`);
    expect(diffs.join("\n")).toBe("");
  });
});
