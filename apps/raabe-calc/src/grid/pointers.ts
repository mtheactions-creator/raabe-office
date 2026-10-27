// Excel's mouse pointers over the grid, drawn as SVG at 1×, 1.5× and 2× so they are as sharp as
// Windows' own at every display scale (D-057): the thick white cross over cells, the black
// arrows over the column and row headers, the thin black cross on the fill handle, the four
// arrows on a selection's border and the header sizing pointers.

function sharp(body: string, size: number, hx: number, hy: number, fallback: string): string {
  const svg = (px: number) => `url("data:image/svg+xml,${encodeURIComponent(
    `<svg xmlns="http://www.w3.org/2000/svg" width="${px}" height="${px}" viewBox="0 0 ${size} ${size}">${body}</svg>`)}")`;
  const set = `image-set(${svg(size)} 1x, ${svg(size * 1.5)} 1.5x, ${svg(size * 2)} 2x) ${hx} ${hy}, ${fallback}`;
  return typeof CSS !== "undefined" && CSS.supports("cursor", set) ? set : `${svg(size)} ${hx} ${hy}, ${fallback}`;
}

export const POINTERS: Record<string, string> = {
  /** Over cells: Excel's thick white plus with a black outline. */
  cells: sharp('<path d="M8.5 2.5h6v6h6v6h-6v6h-6v-6h-6v-6h6z" fill="#fff" stroke="#000" stroke-width="1"/>', 24, 11, 11, "cell"),
  /** Fill handle: a thin black cross. */
  fill: sharp('<path d="M10.5 3v15M3 10.5h15" stroke="#fff" stroke-width="4"/><path d="M10.5 3v15M3 10.5h15" stroke="#000" stroke-width="2"/>', 21, 10, 10, "crosshair"),
  /** Selection border: arrow with the four-way move sign. */
  move: sharp('<path d="M1.5 1.5v15l4-4 3 6.5 2-1-3-6.5h5.5z" fill="#fff" stroke="#000" stroke-width="1"/>' +
    '<path d="M18 9l-3 3h2v4h-4v-2l-3 3 3 3v-2h4v4h-2l3 3 3-3h-2v-4h4v2l3-3-3-3v2h-4v-4h2z" fill="#000" stroke="#fff" stroke-width=".8"/>', 28, 1, 1, "move"),
  /** Column header: a black arrow pointing down. */
  cols: sharp('<path d="M8 2.5h4v8h3.5L10 17.5 4.5 10.5H8z" fill="#000" stroke="#fff" stroke-width="1"/>', 20, 10, 17, "s-resize"),
  /** Row header: a black arrow pointing right. */
  rows: sharp('<path d="M2.5 8v4h8v3.5L17.5 10 10.5 4.5V8z" fill="#000" stroke="#fff" stroke-width="1"/>', 20, 17, 10, "e-resize"),
  resizeCol: sharp('<path d="M15.5 4v24" stroke="#fff" stroke-width="3"/><path d="M15.5 4v24" stroke="#000" stroke-width="1.5"/>' +
    '<path d="M3 16l5-4.5v3h5v3H8v3z M29 16l-5-4.5v3h-5v3h5v3z" fill="#000" stroke="#fff" stroke-width="1" stroke-linejoin="round"/>', 32, 16, 16, "col-resize"),
  resizeRow: sharp('<path d="M4 15.5h24" stroke="#fff" stroke-width="3"/><path d="M4 15.5h24" stroke="#000" stroke-width="1.5"/>' +
    '<path d="M16 3l-4.5 5h3v5h3V8h3z M16 29l-4.5-5h3v-5h3v5h3z" fill="#000" stroke="#fff" stroke-width="1" stroke-linejoin="round"/>', 32, 16, 16, "row-resize"),
  corner: "default",
  none: "default",
};
