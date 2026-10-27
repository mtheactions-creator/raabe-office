"""Speed of RaabeCalc with a large sheet (feel, D-001/D-013): 20 000 rows of numbers and
formulas built through the app's API, then the time for a full recalculation, for one cell
typed in with its dependents recalculated, and for the grid's redraw while scrolling (frame
times measured in the page). No real input. Result: results/<date>-calc-perf.json.
"""
from __future__ import annotations

import datetime
import json
from pathlib import Path

from feellab.calc_catalogue import CalcSide

JS = r"""(async () => {
  const app = __calc.app, wb = app.wb, sheet = wb.sheet();
  const t0 = performance.now();
  app.change('Test', () => {
    for (let r = 0; r < 20000; r++) {
      for (let c = 0; c < 6; c++) sheet.set(r, c, { v: (r * 7 + c * 13) % 1000 });
      sheet.set(r, 6, { v: null, f: `SUM(A${r + 1}:F${r + 1})` });
      sheet.set(r, 7, { v: null, f: r ? `H${r}+G${r + 1}` : 'G1' });
    }
  });
  const build = performance.now() - t0;
  const t1 = performance.now();
  app.active.calc.recalc();
  const recalc = performance.now() - t1;
  // One typed cell: the step, the recalculation and the redraw of the visible cells.
  const t2 = performance.now();
  app.grid.selectCell(5, 2);
  app.grid.startEdit('enter', '123');
  app.grid.commitEdit('down');
  const edit = performance.now() - t2;
  // A cell nothing depends on, and the parts of an edit: the step and recalculation, the redraw.
  const t3 = performance.now();
  app.grid.selectCell(5, 10);
  app.grid.startEdit('enter', '7');
  app.grid.commitEdit('down');
  const editPlain = performance.now() - t3;
  const t4 = performance.now();
  app.active.calc.changed(0, [{ row: 5, col: 2 }]);
  const chainCalc = performance.now() - t4;
  const t5 = performance.now();
  app.updateChrome();
  const chrome = performance.now() - t5;
  // Scrolling: 200 steps of 3 rows, one frame each.
  const frames = [];
  for (let i = 0; i < 200; i++) {
    const v = sheet.view;
    v.top = Math.min(19000, v.top + 3);
    const s = performance.now();
    app.grid['paint']();
    frames.push(performance.now() - s);
    await new Promise((r) => requestAnimationFrame(r));
  }
  frames.sort((a, b) => a - b);
  return { build, recalc, edit, editPlain, chainCalc, chrome, paintMedian: frames[100], paintP95: frames[190], paintMax: frames[199], last: app.wb.sheet().get(19999, 7)?.v };
})()"""


def main() -> None:
    side = CalcSide()
    try:
        side.js("__calc.reset({}, 'A1', 1)")
        res = side.js(JS)
        print(json.dumps(res, indent=1))
    finally:
        side.close()
    out = Path(__file__).parent / "results" / f"{datetime.date.today().isoformat()}-calc-perf.json"
    out.write_text(json.dumps(res, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
