// Formel > Freihandgleichung (D-079): handwriting recognised by Windows' own Math Input Panel (mip.exe), which
// RaabePresent starts; when the equation is written there and inserted ("Einfügen"), its MathML arrives on the
// clipboard and becomes an equation on the slide (our MathML reader). Where Windows' recognition is not there
// (iPad, the browser), the command says so.

import { h, isTauri, showDialog } from "@raabe/shell";
import { mathmlToSeq } from "@raabe/text-engine";
import type { PresentApp } from "./app";
import { insertEquation } from "./equations";

async function invoke<T>(cmd: string): Promise<T> {
  const { invoke: inv } = await import("@tauri-apps/api/core");
  return inv<T>(cmd);
}

/** The MathML on the clipboard (its own formats, or MathML as text), or "". */
async function clipboardMathml(app: PresentApp): Promise<string> {
  try { const m = await invoke<string>("clipboard_read_mathml"); if (m) return m; } catch { /* not there */ }
  const t = await app.platform.clipboardRead().catch(() => "");
  return /<(\w+:)?math[\s>]/.test(t) ? t : "";
}

export async function inkEquation(app: PresentApp): Promise<void> {
  if (!isTauri() || !(await invoke<boolean>("math_input_available").catch(() => false))) {
    await showDialog({ title: "Freihandgleichung", body: h("div", null, "Die Handschrifterkennung für Formeln stellt Windows bereit (Mathematik-Eingabebereich). Auf diesem Gerät ist sie nicht verfügbar."),
      buttons: [{ label: "OK", value: true, primary: true }], cancel: true });
    return;
  }
  const before = await clipboardMathml(app);
  await invoke("math_input_start");
  let done = false;
  let inserted = false;
  const dialog = showDialog({ title: "Freihandgleichung", body: h("div", { style: "max-width:380px" },
    "Schreiben Sie die Formel im Mathematik-Eingabebereich von Windows und klicken Sie dort auf „Einfügen“. RaabePresent übernimmt sie in die Folie."),
    buttons: [{ label: "Abbrechen", value: false }], cancel: false });
  void dialog.then(() => { done = true; });
  // The Math Input Panel puts the equation on the clipboard when it inserts.
  while (!done) {
    await new Promise((r) => setTimeout(r, 400));
    const m = await clipboardMathml(app);
    if (m && m !== before) {
      const seq = mathmlToSeq(m);
      if (seq) { await insertEquation(app, seq); inserted = true; }
      break;
    }
  }
  if (!done) (document.querySelector("[role=dialog] .btn, [role=dialog] button") as HTMLButtonElement | null)?.click();
  void inserted;
}
