// RaabeWrite entry point.

import { browserPlatform, isTauri, tauriPlatform } from "@raabe/shell";
import { installProbe, probeFor, type Para } from "@raabe/text-engine";
import fixtures from "../../../prototypes/src/shared/fixtures.json";
import { WriteApp } from "./app";

/** Test documents for feel-lab runs (identical to the ones the harness builds in Word). */
function fixture(name: string): Para[] {
  const base = fixtures.base as Para[];
  if (name !== "long") return base.map((p) => ({ ...p }));
  const out: Para[] = [];
  for (let i = 1; out.length < fixtures.longMinParagraphs; i++) {
    for (const p of base) out.push(p.style === "h1" ? { style: "h1", text: `${i}. ${p.text}` } : p);
  }
  return out;
}

const feel = (window as unknown as { __FEEL?: string }).__FEEL;
const app = new WriteApp(isTauri() ? tauriPlatform() : browserPlatform(), feel ? fixture(feel) : undefined);
if (feel) {
  installProbe(probeFor(() => app.activeEditor(), fixture));
  (window as unknown as { __app: unknown }).__app = app.testApi();
}
