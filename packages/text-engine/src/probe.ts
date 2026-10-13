// The measurement interface the editor exposes as `window.__probe` in feel-lab runs. The feel-lab
// harness reads it over the Chrome DevTools Protocol. All offsets are into the plain
// document text with paragraphs joined by "\n", the same normalisation the harness
// applies to Word's text, so results are directly comparable.

export interface EngineProbe {
  engine: "a" | "b";
  text(): string;
  selection(): { anchor: number; head: number };
  setSelection(anchor: number, head: number): void;
  /** Client-area CSS px of the caret position before `offset`: x, line top y, line height. */
  pointAt(offset: number): { x: number; y: number; h: number };
  /** Offsets at which each visual line starts. */
  lines(): number[];
  /** Offsets at which each page starts. */
  pages(): number[];
  /** Every visual line: start offset, zero-based page, and its box in points from the page's top left. */
  lineBoxes(): LineBox[];
  /** x (points from the page's left edge) of the caret before each offset. */
  xs(offsets: number[]): number[];
  /** Resolved formatting of the character at `offset` (at a paragraph end: its paragraph mark). */
  formatAt(offset: number): ProbeFormat;
  /** Formatting the next typed character would get (Word: Selection.Font with no selection). */
  typingFormat(): ProbeFormat;
  load(name: string): void;
  focus(): void;
  /** Scrolls so that `offset` is visible, without changing the selection. */
  reveal(offset: number): void;
  settled(): boolean;
}

export interface LineBox {
  start: number;
  page: number;
  /** Left edge of the first character. */
  x: number;
  /** Top of the line box. */
  y: number;
  h: number;
  /** Baseline below the top. */
  base: number;
}

/** Formatting as the harness compares it with Word (colour: "auto" or RRGGBB, sizes in points). */
export interface ProbeFormat {
  font: string;
  size: number;
  bold: boolean;
  italic: boolean;
  underline: string;
  strike: boolean;
  vertAlign: string;
  color: string;
  highlight: string;
  caps: boolean;
  smallCaps: boolean;
  hidden: boolean;
  dstrike: boolean;
  style: string;
  align: string;
  /** Paragraph indents, spacing (points) and line spacing (a multiple, or points). */
  indLeft: number;
  indFirst: number;
  before: number;
  after: number;
  line: number;
  lineRule: string;
}

let pendingInput: number | null = null;
const latencySamples: number[] = [];
const channel = new MessageChannel();
channel.port1.onmessage = (e) => latencySamples.push(performance.now() - (e.data as number));

/** Call when an input event has been fully applied to the DOM. Records event-to-frame time. */
export function markInputHandled(eventTimeStamp: number): void {
  if (pendingInput !== null) return;
  pendingInput = eventTimeStamp;
  requestAnimationFrame(() => {
    // Posting from inside rAF lands right after this frame has been produced.
    channel.port2.postMessage(pendingInput);
    pendingInput = null;
  });
}

function frameRecorder() {
  let deltas: number[] = [];
  let longFrames: number[] = [];
  let running = false;
  let last = 0;
  let observer: PerformanceObserver | null = null;
  const tick = (t: number) => {
    if (!running) return;
    if (last) deltas.push(t - last);
    last = t;
    requestAnimationFrame(tick);
  };
  return {
    start() {
      deltas = []; longFrames = []; last = 0; running = true;
      requestAnimationFrame(tick);
      try {
        observer = new PerformanceObserver((list) => {
          for (const e of list.getEntries()) longFrames.push(e.duration);
        });
        observer.observe({ type: "long-animation-frame", buffered: false });
      } catch { observer = null; }
    },
    stop() {
      running = false;
      observer?.disconnect();
      return { deltas, longFrames };
    },
  };
}

export function installProbe(engine: EngineProbe): void {
  let origin: { x: number; y: number } | null = null;
  addEventListener("mousemove", (e) => {
    origin = { x: e.screenX - e.clientX, y: e.screenY - e.clientY };
  }, { capture: true, passive: true });

  (window as unknown as { __probe: unknown }).__probe = {
    engine: engine.engine,
    text: () => engine.text(),
    selection: () => engine.selection(),
    setSelection: (a: number, h: number) => engine.setSelection(a, h),
    pointAt: (o: number) => engine.pointAt(o),
    lines: () => engine.lines(),
    pages: () => engine.pages(),
    lineBoxes: () => engine.lineBoxes(),
    xs: (offsets: number[]) => engine.xs(offsets),
    formatAt: (o: number) => engine.formatAt(o),
    typingFormat: () => engine.typingFormat(),
    load: (n: string) => engine.load(n),
    focus: () => engine.focus(),
    reveal: (o: number) => engine.reveal(o),
    /** True once no font metrics are outstanding (the layout will not change by itself). */
    settled: () => engine.settled(),
    /** Screen origin of the client area in CSS px (from the last mouse move), plus DPR. */
    mouseOrigin: () => (origin ? { ...origin, dpr: devicePixelRatio } : null),
    frames: frameRecorder(),
    latency: { take: () => latencySamples.splice(0) },
  };
}
