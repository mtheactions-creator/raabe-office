// Zielwertsuche (Daten > Was-wäre-wenn-Analyse): the value of one cell that makes a formula reach a
// goal. Like Excel it stops as soon as the formula is within 0,001 of the goal (Maximale
// Änderung) or after 100 steps; each step is Newton's with a measured slope, at most as large as
// the value itself (1 at 0) so it does not jump past the nearest solution, halved where the
// formula gives an error. Compared with Excel's results in results/2026-10-28-excel-goalseek.json.

export interface SeekResult { found: boolean; x: number; y: number }

/** `f` computes the formula for a value of the changing cell (NaN for an error). */
export function seek(f: (x: number) => number, goal: number, start: number, opts = { maxSteps: 100, tolerance: 0.001 }): SeekResult {
  let x = start;
  let y = f(x);
  for (let i = 0; i < opts.maxSteps; i++) {
    if (Number.isFinite(y) && Math.abs(y - goal) <= opts.tolerance) return { found: true, x, y };
    const h = Math.max(1e-6, Math.abs(x) * 1e-6);
    const slope = (f(x + h) - y) / h;
    if (!Number.isFinite(slope) || slope === 0 || !Number.isFinite(y)) break;
    let step = (goal - y) / slope;
    const limit = Math.abs(x) > 1e-9 ? Math.abs(x) : 1;
    if (Math.abs(step) > limit) step = Math.sign(step) * limit;
    let nx = x + step, ny = f(nx);
    for (let k = 0; k < 30 && !Number.isFinite(ny); k++) { step /= 2; nx = x + step; ny = f(nx); }
    if (!Number.isFinite(ny)) break;
    x = nx;
    y = ny;
  }
  return { found: Number.isFinite(y) && Math.abs(y - goal) <= opts.tolerance, x, y };
}
