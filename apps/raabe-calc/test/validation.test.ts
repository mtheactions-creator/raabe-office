// Datenüberprüfung against Excel (results/2026-10-28-excel-validation.json, recorded with
// probe_excel_validation.py): Excel's verdict on typed entries for every kind of validation,
// and the criteria as Excel's dialog shows them (Formula1/Formula2 read back through COM).

import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { Workbook, parseRange, type Validation, type ValidationOperator, type ValidationType } from "@raabe/sheet-engine";
import * as A from "../src/actions";
import { History } from "../src/ops";
import { applyValidation, criteriaLabels, isValidEntry, toFileCriterion, toLocalCriterion } from "../src/validation";

interface Case {
  name: string; type: ValidationType; operator: ValidationOperator | null; formula1: string | null; formula2: string | null; ignoreBlank: boolean;
  context: Record<string, string>; read?: { Formula1: string; Formula2: string | null }; refused?: string;
  results?: { entry: string; valid: boolean }[];
}

const data = JSON.parse(readFileSync("tools/feel-lab/results/2026-10-28-excel-validation.json", "utf-8")) as {
  cases: Case[]; files: Record<string, { main: string[]; x14: string[]; reread: { Formula1: string; Formula2: string } }>;
};

/** The case's validation on B2, its criteria typed as in Excel's dialog. */
function setup(c: Case) {
  const wb = Workbook.blank();
  const history = new History(wb);
  const ctx = { wb, history };
  const sheet = wb.sheet();
  history.step("setup", () => { for (const [a, t] of Object.entries(c.context)) A.enter(ctx, sheet, parseRange(a)!, t); });
  const v: Validation = { ref: [parseRange("B2")!], type: c.type, showErrorMessage: true };
  if (c.operator && c.operator !== "between" && c.type !== "list" && c.type !== "custom") v.operator = c.operator;
  if (c.ignoreBlank) v.allowBlank = true;
  for (const [i, f] of [c.formula1, c.formula2].entries()) {
    if (f === null) continue;
    const x = toFileCriterion(c.type, f, wb);
    if (x === null) throw new Error(`criterion ${f} of ${c.name}`);
    if (typeof x === "string") v[i ? "formula2" : "formula1"] = x;
    else Object.assign(v, x);
  }
  sheet.validations = [v];
  return { wb, v };
}

describe("entries checked as Excel checks them", () => {
  for (const c of data.cases.filter((x) => !x.refused)) {
    it(c.name, () => {
      const { wb } = setup(c);
      const got = c.results!.map((r) => ({ entry: r.entry, valid: isValidEntry(wb, 0, 1, 1, r.entry) }));
      expect(got).toEqual(c.results!.map((r) => ({ entry: r.entry, valid: r.valid })));
    });
  }
});

describe("criteria as Excel's dialog shows them", () => {
  it("read back like Excel's Formula1 and Formula2", () => {
    for (const c of data.cases.filter((x) => !x.refused)) {
      const { v } = setup(c);
      expect(toLocalCriterion(v, 1), c.name).toBe(c.read!.Formula1);
      if (c.read!.Formula2) expect(toLocalCriterion(v, 2), c.name).toBe(c.read!.Formula2);
    }
  });

  it("literal lists with commas and quotes (Excel's x12ac:list)", () => {
    for (const name of ["list literal", "list literal numbers", "list literal comma", "list literal quote"]) {
      const typed = data.files[name].reread.Formula1;
      const c = toFileCriterion("list", typed);
      expect(c, name).not.toBeNull();
      const v: Validation = { ref: [], type: "list", ...(c as object) };
      expect(toLocalCriterion(v, 1), name).toBe(typed);
      // Excel's file: x12ac:list exactly when an entry holds a comma or a quote.
      expect(!!v.items, name).toBe(data.files[name].main[0].includes("x12ac:list"));
      if (!v.items) expect(data.files[name].main[0]).toContain(`<formula1>${v.formula1!.replace(/&/g, "&amp;")}</formula1>`);
    }
  });

  it("labels follow the kind and the operator", () => {
    expect(criteriaLabels("whole")).toEqual(["Minimum:", "Maximum:"]);
    expect(criteriaLabels("whole", "equal")).toEqual(["Wert:"]);
    expect(criteriaLabels("textLength", "notEqual")).toEqual(["Länge:"]);
    expect(criteriaLabels("date", "greaterThan")).toEqual(["Anfangsdatum:"]);
    expect(criteriaLabels("time", "lessThanOrEqual")).toEqual(["Endzeit:"]);
    expect(criteriaLabels("list")).toEqual(["Quelle:"]);
    expect(criteriaLabels("custom")).toEqual(["Formel:"]);
    expect(criteriaLabels("any")).toEqual([]);
  });

  it("refuses criteria that do not fit the kind", () => {
    expect(toFileCriterion("whole", "1,5")).toBeNull();
    expect(toFileCriterion("date", "abc")).toBeNull();
    expect(toFileCriterion("date", "15.03.2026")).toBe("46096");
    expect(toFileCriterion("time", "12:00")).toBe("0.5");
    expect(toFileCriterion("decimal", "=$D$1")).toBe("$D$1");
    expect(toFileCriterion("custom", "=ISTZAHL(A1)")).toBe("ISNUMBER(A1)");
  });
});

describe("setting validations on a selection", () => {
  it("takes the cells from older ones and leaves the rest", () => {
    const wb = Workbook.blank();
    const s = wb.sheet();
    s.validations = [{ ref: [parseRange("A1:C3")!], type: "whole", formula1: "1", formula2: "6" }];
    applyValidation(s, [parseRange("B2")!], { type: "list", formula1: '"Ja,Nein"' });
    expect(s.validations[0].ref).toEqual([parseRange("A1:C1"), parseRange("A3:C3"), parseRange("A2"), parseRange("C2")]);
    expect(s.validations[1]).toMatchObject({ type: "list", ref: [parseRange("B2")] });
    applyValidation(s, [parseRange("A1:C3")!], null);
    expect(s.validations).toBeUndefined();
  });

  it("is undone and redone", () => {
    const wb = Workbook.blank();
    const s = wb.sheet();
    const history = new History(wb);
    history.step("dv", () => { history.validations(s); applyValidation(s, [parseRange("A1")!], { type: "whole", formula1: "1", formula2: "6" }); });
    expect(s.validations?.length).toBe(1);
    history.undo();
    expect(s.validations).toBeUndefined();
    history.redo();
    expect(s.validations?.[0]).toMatchObject({ type: "whole", ref: [parseRange("A1")] });
  });
});

describe("the feature in the page", () => {
  /** The parts of CalcApp the feature uses, on a plain workbook. */
  function fakeApp() {
    const wb = Workbook.blank();
    const history = new History(wb);
    const sheet = wb.sheet();
    const el = document.createElement("div");
    document.body.append(el);
    const entered: [string, string][] = [];
    const app = {
      wb, sheet, history, ctx: { wb, history },
      change: (label: string, fn: () => void) => history.step(label, fn),
      grid: {
        el, isEditing: false, commitEdit: () => {}, focus: () => {}, selectMany: (r: unknown) => { app.selected = r; },
        get selection() { return sheet.view.selection; },
        rectOnCanvas: (r: { r0: number; c0: number }) => ({ x: 40 + r.c0 * 80, y: 20 + r.r0 * 20, w: 80, h: 20 }),
      },
      selected: null as unknown,
    };
    return { app, wb, sheet, el, entered, history };
  }
  const button = (label: string) => [...document.querySelectorAll<HTMLButtonElement>(".dialog button")].find((b) => b.textContent === label)!;

  it("shows the arrow and the input message, and enters a chosen entry", async () => {
    const { validationFeature } = await import("../src/validation");
    const { app, sheet, el, entered } = fakeApp();
    for (const [a, t] of [["D1", "Rot"], ["D2", "Grün"]]) sheet.set(parseRange(a)!.r0, 3, { v: t });
    sheet.validations = [{ ref: [parseRange("B2")!], type: "list", formula1: "$D$1:$D$2", showInputMessage: true, promptTitle: "Farbe", prompt: "Eine Farbe wählen" }];
    const f = validationFeature(app as never, (target, text) => entered.push([`${target.r0},${target.c0}`, text]));
    sheet.view.active = { row: 1, col: 1 };
    f.afterSelect();
    const arrow = el.querySelector<HTMLElement>(".dv-arrow")!;
    expect(arrow.style.left).toBe("201px");
    expect(el.querySelector(".dv-prompt")?.textContent).toBe("FarbeEine Farbe wählen");
    expect(f.openList()).toBe(true);
    expect([...el.querySelectorAll(".dv-list div")].map((d) => d.textContent)).toEqual(["Rot", "Grün"]);
    dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown" }));
    dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown" }));
    dispatchEvent(new KeyboardEvent("keydown", { key: "Enter" }));
    expect(entered).toEqual([["1,1", "Grün"]]);
    expect(el.querySelector(".dv-list")).toBeNull();
    sheet.view.active = { row: 0, col: 0 };
    f.afterSelect();
    expect(el.querySelector(".dv-arrow")).toBeNull();
    expect(el.querySelector(".dv-prompt")).toBeNull();
  });

  it("asks with Excel's buttons for an invalid entry", async () => {
    const { validationFeature } = await import("../src/validation");
    const { app, sheet } = fakeApp();
    sheet.validations = [{ ref: [parseRange("A1")!], type: "whole", formula1: "1", formula2: "6", showErrorMessage: true, errorTitle: "Note", error: "Nur 1 bis 6" }];
    const f = validationFeature(app as never);
    expect(await f.check(0, 0, "3")).toBe("ok");
    let answer = f.check(0, 0, "7");
    expect(document.querySelector(".dialog .dlg-title")?.textContent).toBe("Note");
    expect([...document.querySelectorAll(".dialog .dlg-buttons button")].map((b) => b.textContent)).toEqual(["Wiederholen", "Abbrechen"]);
    button("Wiederholen").click();
    expect(await answer).toBe("retry");
    sheet.validations[0].errorStyle = "warning";
    answer = f.check(0, 0, "7");
    expect(document.querySelector(".dialog .dv-msg")?.textContent).toContain("Möchten Sie fortfahren?");
    button("Ja").click();
    expect(await answer).toBe("ok");
    sheet.validations[0].errorStyle = "information";
    sheet.validations[0].error = undefined;
    answer = f.check(0, 0, "7");
    expect([...document.querySelectorAll(".dialog .dlg-buttons button")].map((b) => b.textContent)).toEqual(["OK", "Abbrechen"]);
    button("Abbrechen").click();
    expect(await answer).toBe("cancel");
    sheet.validations[0].showErrorMessage = undefined;
    expect(f.quickCheck(0, 0, "7")).toBe(true);
  });

  it("sets a validation through the dialog, as one undo step", async () => {
    const { validationFeature } = await import("../src/validation");
    const { app, sheet, history } = fakeApp();
    sheet.view.selection = [parseRange("A1:A5")!];
    const f = validationFeature(app as never);
    const done = f.openDialog();
    await Promise.resolve();
    const dlg = document.querySelector(".dialog")!;
    const [zulassen, daten] = [...dlg.querySelectorAll("select")];
    zulassen.value = "decimal";
    zulassen.dispatchEvent(new Event("change"));
    daten.value = "greaterThan";
    daten.dispatchEvent(new Event("change"));
    const inputs = [...dlg.querySelectorAll<HTMLInputElement>(".dv-row input[type=text]")].filter((i) => i.isConnected && !i.closest("[hidden]"));
    expect([...dlg.querySelectorAll(".dv-row > span")].filter((s) => !s.closest("[hidden]")).map((s) => s.textContent)).toContain("Minimum:");
    inputs[0].value = "1,5";
    button("OK").click();
    await done;
    expect(sheet.validations).toEqual([{ ref: [parseRange("A1:A5")], type: "decimal", operator: "greaterThan", formula1: "1.5", allowBlank: true, showInputMessage: true, showErrorMessage: true }]);
    history.undo();
    expect(sheet.validations).toBeUndefined();
  });
});

describe("inserted and deleted rows", () => {
  it("move, grow and shrink the ranges, the criteria follow", async () => {
    const { shiftValidations } = await import("@raabe/sheet-engine");
    const wb = Workbook.blank();
    const s = wb.sheet();
    s.validations = [
      { ref: [parseRange("K2:K6")!], type: "custom", formula1: "ISNUMBER(K2)" },
      { ref: [parseRange("A1:A3")!], type: "list", formula1: "$H$2:$H$5" },
      { ref: [parseRange("B4")!], type: "list", formula1: '"Ja,Nein"' },
    ];
    shiftValidations(wb, s, "rows", 2, 2);
    expect(s.validations.map((v) => [v.ref, v.formula1])).toEqual([
      [[parseRange("K2:K8")], "ISNUMBER(K2)"], [[parseRange("A1:A5")], "$H$2:$H$7"], [[parseRange("B6")], '"Ja,Nein"'],
    ]);
    shiftValidations(wb, s, "rows", 5, -1);
    expect(s.validations.map((v) => v.ref)).toEqual([[parseRange("K2:K7")], [parseRange("A1:A5")]]);
  });
});
