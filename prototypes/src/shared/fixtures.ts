import type { StyleId } from "./geometry";
import data from "./fixtures.json";

export interface FixturePara {
  style: StyleId;
  text: string;
}

// The texts live in fixtures.json so the feel-lab harness builds the identical document
// in Word. "short" is one page; "long" repeats it (numbering the headings) to ~50 pages.
const BASE = data.base as FixturePara[];

function long(): FixturePara[] {
  const out: FixturePara[] = [];
  for (let i = 1; out.length < data.longMinParagraphs; i++) {
    for (const p of BASE) {
      out.push(p.style === "h1" ? { style: "h1", text: `${i}. ${p.text}` } : p);
    }
  }
  return out;
}

export function fixture(name: string): FixturePara[] {
  return name === "long" ? long() : BASE.map((p) => ({ ...p }));
}

export function fixtureName(): string {
  const fromShell = (window as unknown as { __FIXTURE?: string }).__FIXTURE;
  return fromShell ?? new URLSearchParams(location.search).get("doc") ?? "short";
}
