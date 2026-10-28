import { describe, expect, it } from "vitest";
import { headerSections } from "../src/print";

describe("Kopf- und Fußzeile", () => {
  const f = { page: 2, pages: 5, sheet: "Tabelle1", file: "Noten", now: new Date(2026, 9, 28, 7, 5) };
  it("fills in Excel's codes per section", () => {
    expect(headerSections("&LKlasse 7b&C&A&RSeite &P von &N", f)).toEqual({ L: "Klasse 7b", C: "Tabelle1", R: "Seite 2 von 5" });
    expect(headerSections("&D &T &F", f)).toEqual({ L: "", C: "28.10.2026 07:05 Noten", R: "" });
  });
  it("leaves format codes out and keeps &&", () => {
    expect(headerSections('&C&"Calibri,Fett"&14Titel && mehr&KFF0000!', f)).toEqual({ L: "", C: "Titel & mehr!", R: "" });
  });
});
