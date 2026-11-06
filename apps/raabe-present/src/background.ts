// Hintergrund formatieren (D-070): the pane at the right with Füllung: Einfarbige Füllung, Farbverlauf,
// Bild- oder Texturfüllung (a picture file), Transparenz, Hintergrundgrafiken ausblenden, and the
// buttons Auf alle anwenden and Hintergrund zurücksetzen. Written into the slide's p:cSld/p:bg as
// PowerPoint does (p:bgPr with the fill and an empty a:effectLst).

import { TaskPane, type PaneControl } from "@raabe/shell";
import { PptxDom } from "@raabe/ooxml";
import type { PresentApp } from "./app";

const { kid, make, fragment } = PptxDom;

export class BackgroundPane {
  readonly pane: TaskPane;
  constructor(private app: PresentApp, host: HTMLElement) {
    this.pane = new TaskPane(host);
  }

  private paths(all = false): string[] {
    const d = this.app.deck;
    return all ? d.slides.map((s) => s.path) : [...this.app.panel.selected].map((i) => d.slides[i]?.path).filter((p): p is string => !!p);
  }

  private setBg(xml: string | null, all = false): void {
    const paths = this.paths(all);
    this.app.deck.change("Hintergrund formatieren", paths, () => {
      for (const p of paths) {
        const cSld = kid(this.app.deck.pptx.doc(p).documentElement, "cSld")!;
        kid(cSld, "bg")?.remove();
        if (xml) cSld.insertBefore(fragment(cSld.ownerDocument, `<p:bg><p:bgPr>${xml}<a:effectLst/></p:bgPr></p:bg>`)[0], cSld.firstChild);
      }
    });
    this.pane.refresh();
  }

  /** The background fill element of the current slide (its own), if any. */
  private own(): Element | undefined {
    const p = this.app.slidePath();
    if (!p) return undefined;
    const pr = PptxDom.at(this.app.deck.pptx.doc(p).documentElement, "cSld/bg/bgPr");
    return pr ? Array.from(pr.children).find((c) => c.localName !== "effectLst") : undefined;
  }

  private current(): string {
    const p = this.app.slidePath();
    const v = p ? this.app.deck.view(p) : null;
    return v?.background.fill.kind === "solid" ? `#${v.background.fill.color}` : "#FFFFFF";
  }

  private async picture(): Promise<void> {
    const path = await this.app.platform.openImageDialog();
    if (!path) return;
    const bytes = await this.app.platform.readFile(path);
    const ext = /\.jpe?g$/i.test(path) ? "jpeg" : "png";
    const media = this.app.deck.pptx.freeName(`ppt/media/image#.${ext}`);
    const paths = this.paths();
    this.app.deck.change("Hintergrund formatieren", paths, () => {
      this.app.deck.pptx.put(media, bytes);
      this.app.deck.pptx.setType(media, ext === "jpeg" ? "image/jpeg" : "image/png", true);
      for (const p of paths) {
        const rid = this.app.deck.pptx.link(p, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", media);
        const cSld = kid(this.app.deck.pptx.doc(p).documentElement, "cSld")!;
        kid(cSld, "bg")?.remove();
        const bg = fragment(cSld.ownerDocument, `<p:bg><p:bgPr><a:blipFill dpi="0" rotWithShape="1"><a:blip/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>`)[0];
        bg.getElementsByTagNameNS(PptxDom.NS_A, "blip")[0].setAttributeNS(PptxDom.NS_R, "r:embed", rid);
        cSld.insertBefore(bg, cSld.firstChild);
      }
    });
    this.pane.refresh();
  }

  private hideGraphics(on: boolean): void {
    const paths = this.paths();
    this.app.deck.change("Hintergrundgrafiken ausblenden", paths, () => {
      for (const p of paths) { const r = this.app.deck.pptx.doc(p).documentElement; if (on) r.setAttribute("showMasterSp", "0"); else r.removeAttribute("showMasterSp"); }
    });
    this.pane.refresh();
  }

  open(): void {
    const kind = () => this.own()?.localName ?? "";
    const controls = (): PaneControl[] => {
      const c: PaneControl[] = [
        { type: "radio", label: "Einfarbige Füllung", checked: () => kind() === "solidFill" || !kind(), pick: () => this.setBg(`<a:solidFill><a:srgbClr val="${this.current().slice(1)}"/></a:solidFill>`) },
        { type: "radio", label: "Farbverlauf", checked: () => kind() === "gradFill", pick: () => this.setBg('<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="bg1"/></a:gs><a:gs pos="100000"><a:schemeClr val="accent1"><a:lumMod val="40000"/><a:lumOff val="60000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill>') },
        { type: "radio", label: "Bild- oder Texturfüllung", checked: () => kind() === "blipFill", pick: () => void this.picture() },
        { type: "check", label: "Hintergrundgrafiken ausblenden", value: () => this.app.deck.pptx.doc(this.app.slidePath()!).documentElement.getAttribute("showMasterSp") === "0", set: (on) => this.hideGraphics(on) },
      ];
      if (kind() === "solidFill" || !kind()) c.push({ type: "color", label: "Farbe", value: () => this.current(), set: (hex) => this.setBg(`<a:solidFill><a:srgbClr val="${(hex ?? "#FFFFFF").slice(1).toUpperCase()}"/></a:solidFill>`) });
      c.push({ type: "button", label: "Auf alle anwenden", run: () => { const f = this.own(); this.setBg(f ? new XMLSerializer().serializeToString(f) : null, true); } },
        { type: "button", label: "Hintergrund zurücksetzen", run: () => this.setBg(null) });
      return c;
    };
    this.pane.show({ title: "Hintergrund formatieren", sets: [{ label: "", tabs: [{ id: "fill", label: "Füllung", icon: "pane-fill-line", sections: () => [{ title: "Füllung", open: true, controls }] }] }] });
  }
}

export { make };
