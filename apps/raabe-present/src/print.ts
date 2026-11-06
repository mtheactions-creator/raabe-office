// PDF and printing (D-070): every shown slide as one page of the slide's own size, drawn as in the
// show (no prompts, hidden slides left out as PowerPoint's default "Ganzseitige Folien" does).

import { h } from "@raabe/shell";
import type { PresentApp } from "./app";
import { renderSlide } from "./render/slide";
import { EMU_PX } from "./render/text";

export async function printSlides(app: PresentApp, path: string): Promise<void> {
  const deck = app.deck;
  const size = deck.pptx.slideSize;
  const win = size.cx / 914400, hin = size.cy / 914400;
  const root = document.getElementById("print-root") ?? document.body.appendChild(h("div", { id: "print-root" }));
  const pages = deck.slides.map((ref, i) => ({ ref, i })).filter(({ ref }) => deck.pptx.doc(ref.path).documentElement.getAttribute("show") !== "0").map(({ ref, i }) => {
    const page = h("div", { class: "print-page", style: `width:${win}in;height:${hin}in;position:relative;overflow:hidden` });
    const r = renderSlide(deck.view(ref.path), size, { editing: false, slideNumber: i + 1, image: (p) => app.image(p), pptx: deck.pptx });
    // 96 px per inch on screen; the page is in inches, so the slide is drawn at its natural size.
    r.el.style.transformOrigin = "0 0";
    r.el.style.transform = `scale(${(win * 96) / (size.cx / EMU_PX)})`;
    page.append(r.el);
    return page;
  });
  root.replaceChildren(...pages);
  const style = h("style", null, `@page { size: ${win}in ${hin}in; margin: 0 }`);
  document.head.append(style);
  document.body.classList.add("printing");
  // Pictures must have loaded before the page is printed.
  await Promise.all(Array.from(root.querySelectorAll("img")).map((img) => (img as HTMLImageElement).decode().catch(() => {})));
  try {
    await app.platform.printToPdf(path, { width: win, height: hin });
  } finally {
    document.body.classList.remove("printing");
    root.replaceChildren();
    style.remove();
  }
}
