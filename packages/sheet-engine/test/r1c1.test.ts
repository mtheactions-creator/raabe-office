import { describe, expect, it } from "vitest";
import { fromR1C1, toR1C1 } from "../src";

// Recorded from Excel's macro recorder (results/2026-11-09-excel-recorder.json, D-119).
describe("R1C1 as Excel's recorder writes it", () => {
  it("relative references from the formula's cell", () => {
    expect(toR1C1("SUM(A2:A3)", 3, 0)).toBe("SUM(R[-2]C:R[-1]C)");
    expect(toR1C1("A2*2", 6, 1)).toBe("R[-5]C[-1]*2");
    expect(toR1C1("$A$1+B$2+$C3", 4, 1)).toBe("R1C1+R2C+R[-2]C3");
    expect(toR1C1("SUM(A:A)", 0, 1)).toBe("SUM(C[-1])");
    expect(toR1C1("SUM(2:4)", 0, 0)).toBe("SUM(R[1]:R[3])");
    expect(toR1C1("Tabelle2!B2&\"R1C1\"", 0, 0)).toBe("Tabelle2!R[1]C[1]&\"R1C1\"");
  });
  it("back to A1", () => {
    expect(fromR1C1("=SUM(R[-2]C:R[-1]C)", 3, 0)).toBe("=SUM(A2:A3)");
    expect(fromR1C1("=R[-5]C[-1]*2", 6, 1)).toBe("=A2*2");
    expect(fromR1C1("=R1C1+R2C+R[-2]C3", 4, 1)).toBe("=$A$1+B$2+$C3");
    expect(fromR1C1("=SUM(C[-1])", 0, 1)).toBe("=SUM(A:A)");
    expect(fromR1C1("=SUM(R[1]:R[3])", 0, 0)).toBe("=SUM(2:4)");
    expect(fromR1C1("=RC[1]&\"RC\"&CODE(R1C1)", 0, 0)).toBe("=B1&\"RC\"&CODE($A$1)");
  });
});
