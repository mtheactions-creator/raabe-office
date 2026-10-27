// Criteria of ZÄHLENWENN, SUMMEWENN(S), MITTELWERTWENN(S), MAXWENNS, MINWENNS: a value, or a
// text with an operator in front (">3", "<>apfel", "<=1.1.2024") and wildcards (*, ?, ~).

import { wildcard } from "./fn-text";
import { E, compareText, isErr, numCompare, textToNumber, type V } from "./values";
import { CellError } from "../model";

export type Matcher = (v: V) => boolean;

const OPS = ["<=", ">=", "<>", "<", ">", "="];

function cmpOp(op: string, c: number): boolean {
  switch (op) {
    case "<": return c < 0;
    case ">": return c > 0;
    case "<=": return c <= 0;
    case ">=": return c >= 0;
    case "<>": return c !== 0;
    default: return c === 0;
  }
}

/** A test for cell values from a criterion. */
export function matcher(crit: V): Matcher {
  // An empty criterion cell counts as 0 (=ZÄHLENWENN(C1:C10;C4) with C4 empty is 0).
  if (crit === null) crit = 0;
  if (typeof crit === "number") return (v) => (typeof v === "number" && numCompare(v, crit as number) === 0) || (typeof v === "string" && textNum(v) === crit);
  if (typeof crit === "boolean") return (v) => v === crit;
  if (crit instanceof CellError) return (v) => v === crit;
  let op = "=";
  let rest = crit;
  for (const o of OPS) {
    if (crit.startsWith(o)) {
      op = o;
      rest = crit.slice(o.length);
      break;
    }
  }
  const hasOp = rest !== crit;
  // Empty criterion: "" matches empty cells and empty text, "=" only empty cells, "<>" any non-empty one.
  if (rest === "") {
    if (!hasOp) return (v) => v === null || v === "";
    if (op === "=") return (v) => v === null;
    if (op === "<>") return (v) => v !== null;
    return () => false;
  }
  const x = textToNumber(rest);
  if (x !== null) {
    return (v) => {
      if (typeof v === "number") return cmpOp(op, numCompare(v, x));
      if (op === "=" && typeof v === "string") return textNum(v) === x;
      return op === "<>";
    };
  }
  const u = rest.toUpperCase();
  if (u === "WAHR" || u === "FALSCH") {
    const bv = u === "WAHR";
    return (v) => (typeof v === "boolean" ? cmpOp(op, v === bv ? 0 : v ? 1 : -1) : op === "<>");
  }
  const errCode = errorOf(rest);
  if (errCode) return (v) => (isErr(v) ? cmpOp(op, v === errCode ? 0 : 1) : op === "<>");
  if (op === "=" || op === "<>") {
    const re = new RegExp("^" + wildcard(rest, false) + "$", "i");
    if (op === "=") return (v) => typeof v === "string" && re.test(v);
    return (v) => !(typeof v === "string" && re.test(v));
  }
  return (v) => typeof v === "string" && cmpOp(op, compareText(v, rest));
}

/** Numeric text in a cell counts for a numeric criterion ("10" matches 10). */
function textNum(v: string): number | null {
  return textToNumber(v);
}

const ERR_NAMES: Record<string, CellError> = {
  "#NULL!": E.null, "#DIV/0!": E.div0, "#WERT!": E.value, "#BEZUG!": E.ref, "#NAME?": E.name, "#ZAHL!": E.num, "#NV": E.na,
};

function errorOf(t: string): CellError | undefined {
  return ERR_NAMES[t.toUpperCase()];
}
