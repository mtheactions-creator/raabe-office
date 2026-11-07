// Freihand zu Form (D-072): strokes recognised as shapes.
import { describe, expect, it } from "vitest";
import { recognize } from "../src/inkshape";

const along = (pts: [number, number][], n = 12) => pts.slice(0, -1).flatMap(([x1, y1], i) => {
  const [x2, y2] = pts[i + 1];
  return Array.from({ length: n }, (_, k) => ({ x: x1 + ((x2 - x1) * k) / n + ((k * 7919) % 5) * 300, y: y1 + ((y2 - y1) * k) / n + ((k * 104729) % 5) * 300 }));
}).concat([{ x: pts[pts.length - 1][0], y: pts[pts.length - 1][1] }]);

describe("Freihand zu Form", () => {
  it("a straight stroke becomes a line", () => {
    expect(recognize(along([[0, 0], [2000000, 1000000]]))?.prst).toBe("line");
  });
  it("a closed stroke with four corners becomes a rectangle", () => {
    expect(recognize(along([[0, 0], [2000000, 0], [2000000, 1200000], [0, 1200000], [0, 0]]))?.prst).toBe("rect");
  });
  it("three corners make a triangle", () => {
    expect(recognize(along([[1000000, 0], [2000000, 1500000], [0, 1500000], [1000000, 0]]))?.prst).toBe("triangle");
  });
  it("a round loop becomes an ellipse", () => {
    const pts = Array.from({ length: 60 }, (_, i) => ({ x: 1000000 + 1000000 * Math.cos((i / 59) * 2 * Math.PI), y: 600000 + 600000 * Math.sin((i / 59) * 2 * Math.PI) }));
    expect(recognize(pts)?.prst).toBe("ellipse");
  });
  it("a scribble stays ink", () => {
    const pts = Array.from({ length: 40 }, (_, i) => ({ x: i * 50000, y: (i % 2) * 400000 + i * 10000 }));
    expect(recognize(pts)).toBeNull();
  });
});
