import { describe, expect, it } from "vitest";
import { nextStep } from "../src";

const next = (f: string) => {
  const s = nextStep(f);
  return s ? f.slice(s.start, s.end) : null;
};

describe("Formelauswertung: the next part", () => {
  it("references first, then operations by precedence", () => {
    expect(next("=A1+B1*2")).toBe("A1");
    expect(next("=5+B1*2")).toBe("B1");
    expect(next("=5+3*2")).toBe("3*2");
    expect(next("=5+6")).toBe("5+6");
    expect(next("=11")).toBeNull();
    expect(next("=-3")).toBeNull();
    expect(next('="a"')).toBeNull();
  });
  it("functions once their arguments are values; ranges stay", () => {
    expect(next("=SUMME(A1:A3)*2")).toBe("SUMME(A1:A3)");
    expect(next("=RUNDEN(A1/3;2)")).toBe("A1");
    expect(next("=RUNDEN(1,5/3;2)")).toBe("1,5/3");
    expect(next("=RUNDEN(0,5;2)")).toBe("RUNDEN(0,5;2)");
  });
  it("WENN: the condition, then the whole function", () => {
    expect(next("=WENN(A1>2;B1;C1)")).toBe("A1");
    expect(next("=WENN(5>2;B1;C1)")).toBe("5>2");
    expect(next("=WENN(WAHR;B1;C1)")).toBe("WENN(WAHR;B1;C1)");
  });
  it("brackets", () => {
    expect(next("=(1+2)*3")).toBe("1+2");
    expect(next("=(3)*3")).toBe("(3)");
    expect(next("=-A1")).toBe("A1");
  });
});
