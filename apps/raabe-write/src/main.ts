// RaabeWrite entry point.

import { browserPlatform, isTauri, tauriPlatform } from "@raabe/shell";
import { installProbe, makePara, probeFor, type DocState, type Para } from "@raabe/text-engine";
import fixtures from "../../../prototypes/src/shared/fixtures.json";
import { WriteApp } from "./app";

/** Test documents for feel-lab runs (identical to the ones the harness builds in Word). */
function fixture(name: string): Partial<DocState> {
  const para = (p: { style: string; text: string }): Para => makePara(p.text, p.style === "h1" ? { style: "berschrift1" } : {});
  const base = fixtures.base;
  if (name !== "long") return { paras: base.map(para) };
  const out: Para[] = [];
  for (let i = 1; out.length < fixtures.longMinParagraphs; i++) {
    for (const p of base) out.push(para(p.style === "h1" ? { style: "h1", text: `${i}. ${p.text}` } : p));
  }
  return { paras: out };
}

const feel = (window as unknown as { __FEEL?: string }).__FEEL;
const app = new WriteApp(isTauri() ? tauriPlatform() : browserPlatform(), feel ? fixture(feel) : undefined);
// The dev server always exposes the probe, so the browser pane can inspect the document.
if (feel || import.meta.env.DEV) {
  installProbe(probeFor(() => app.activeEditor(), fixture));
  (window as unknown as { __app: unknown }).__app = app.testApi();
}
