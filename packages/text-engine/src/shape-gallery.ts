// Office's Formen gallery (D-038, recorded from Word 365 with probe_word_shapes_ui.py; Excel 365
// shows the same gallery, D-066) and its icons, shared by RaabeWrite and RaabeCalc.

import type { DrawTool } from "./editor";
import { flattenPath, presetGeometry } from "./geometry";

export type GalleryEntry = { tool: DrawTool; label: string };
type Entry = GalleryEntry;
const T = (geometry: string, label: string, kind: DrawTool["kind"] = "shape", extra: Partial<DrawTool> = {}): Entry => ({ tool: { geometry, kind, ...extra }, label });
const arrow = (head: boolean, tail: boolean) => ({ pic: { line: { color: "", width: -1, head: head ? { type: "triangle" as const } : undefined, tail: tail ? { type: "triangle" as const } : undefined } } });

/** Office's Formen gallery (the same in Word and Excel), category by category. */
export const SHAPE_GALLERY: { title: string; items: Entry[] }[] = [
  { title: "Linien", items: [
    T("line", "Linie", "connector"), T("straightConnector1", "Linienpfeil", "connector", arrow(false, true)),
    T("straightConnector1", "Linienpfeil: doppelt", "connector", arrow(true, true)), T("bentConnector3", "Verbinder: gewinkelt", "connector"),
    T("bentConnector3", "Verbinder: gewinkelt mit Pfeil", "connector", arrow(false, true)), T("bentConnector3", "Verbinder: gewinkelt mit Doppelpfeil", "connector", arrow(true, true)),
    T("curvedConnector3", "Verbinder: gekrümmt", "connector"), T("curvedConnector3", "Verbinder: gekrümmter Pfeil", "connector", arrow(false, true)),
    T("curvedConnector3", "Verbinder: gekrümmter Doppelpfeil", "connector", arrow(true, true)),
    T("rect", "Kurve", "shape", { freeform: "curve" }), T("rect", "Freihandform: Form", "shape", { freeform: "polygon" }), T("rect", "Freihand: Skizze", "shape", { freeform: "scribble" }),
  ] },
  { title: "Rechtecke", items: [
    T("rect", "Rechteck"), T("roundRect", "Rechteck: abgerundete Ecken"), T("snip1Rect", "Rechteck: eine Ecke abgeschnitten"),
    T("snip2SameRect", "Rechteck: obere Ecken abgeschnitten"), T("snip2DiagRect", "Rechteck: diagonal liegende Ecken abgeschnitten"),
    T("snipRoundRect", "Rechteck: obere Ecken, eine abgerundet, eine abgeschnitten"), T("round1Rect", "Rechteck: eine Ecke abgerundet"),
    T("round2SameRect", "Rechteck: obere Ecken abgerundet"), T("round2DiagRect", "Rechteck: diagonal liegende Ecken abgerundet"),
  ] },
  { title: "Standardformen", items: [
    T("rect", "Textfeld", "textbox"), T("ellipse", "Ellipse"), T("triangle", "Gleichschenkliges Dreieck"), T("rtTriangle", "Rechtwinkliges Dreieck"),
    T("parallelogram", "Parallelogramm"), T("trapezoid", "Trapez"), T("diamond", "Raute"), T("pentagon", "Fünfeck"), T("hexagon", "Sechseck"),
    T("heptagon", "Siebeneck"), T("octagon", "Achteck"), T("decagon", "Zehneck"), T("dodecagon", "Zwölfeck"), T("pie", "Teilkreis"), T("chord", "Sehne"),
    T("teardrop", "Träne"), T("frame", "Rahmen"), T("halfFrame", "Halber Rahmen"), T("corner", "L-Form"), T("diagStripe", "Diagonaler Streifen"),
    T("plus", "Kreuz"), T("plaque", "Zierrahmen"), T("can", "Zylinder"), T("cube", "Würfel"), T("bevel", "Rechteck: abgeschrägt"),
    T("donut", "Kreis: nicht ausgefüllt"), T("noSmoking", "Verbotsymbol"), T("blockArc", "Halbbogen"), T("foldedCorner", "Rechteck: gefaltete Ecke"),
    T("smileyFace", "Smiley"), T("heart", "Herz"), T("lightningBolt", "Gewitterblitz"), T("sun", "Sonne"), T("moon", "Mond"), T("cloud", "Wolke"),
    T("arc", "Bogen"), T("bracketPair", "Runde Klammer links/rechts"), T("bracePair", "Geschweifte Klammer links/rechts"),
    T("leftBracket", "Eckige Klammer links"), T("rightBracket", "Eckige Klammer rechts"), T("leftBrace", "Geschweifte Klammer links"),
    T("rightBrace", "Geschweifte Klammer rechts"),
  ] },
  { title: "Blockpfeile", items: [
    T("rightArrow", "Pfeil: nach rechts"), T("leftArrow", "Pfeil: nach links"), T("upArrow", "Pfeil: nach oben"), T("downArrow", "Pfeil: nach unten"),
    T("leftRightArrow", "Pfeil: nach links und rechts"), T("upDownArrow", "Pfeil: nach oben und unten"), T("quadArrow", "Pfeil: in vier Richtungen"),
    T("leftRightUpArrow", "Pfeil: nach links, rechts und oben"), T("bentArrow", "Pfeil: gebogen"), T("uturnArrow", "Pfeil: 180-Grad"),
    T("leftUpArrow", "Pfeil: nach links und oben"), T("bentUpArrow", "Pfeil: nach oben gebogen"), T("curvedRightArrow", "Pfeil: nach rechts gekrümmt"),
    T("curvedLeftArrow", "Pfeil: nach links gekrümmt"), T("curvedUpArrow", "Pfeil: nach oben gekrümmt"), T("curvedDownArrow", "Pfeil: nach unten gekrümmt"),
    T("stripedRightArrow", "Pfeil: gestreift nach rechts"), T("notchedRightArrow", "Pfeil: eingekerbt nach rechts"), T("homePlate", "Pfeil: Fünfeck"),
    T("chevron", "Pfeil: Chevron"), T("rightArrowCallout", "Legende: mit Pfeil nach rechts"), T("downArrowCallout", "Legende: mit Pfeil nach unten"),
    T("leftArrowCallout", "Legende: mit Pfeil nach links"), T("upArrowCallout", "Legende: mit Pfeil nach oben"),
    T("leftRightArrowCallout", "Legende: mit Pfeil nach links und rechts"), T("quadArrowCallout", "Legende: mit Pfeil in vier Richtungen"),
    T("circularArrow", "Pfeil: gebogen"),
  ] },
  { title: "Formelformen", items: [
    T("mathPlus", "Additionszeichen"), T("mathMinus", "Minuszeichen"), T("mathMultiply", "Multiplikationszeichen"), T("mathDivide", "Divisionszeichen"),
    T("mathEqual", "Gleich"), T("mathNotEqual", "Ungleich"),
  ] },
  { title: "Flussdiagramm", items: [
    T("flowChartProcess", "Flussdiagramm: Prozess"), T("flowChartAlternateProcess", "Flussdiagramm: Alternativer Prozess"), T("flowChartDecision", "Flussdiagramm: Verzweigung"),
    T("flowChartInputOutput", "Flussdiagramm: Daten"), T("flowChartPredefinedProcess", "Flussdiagramm: Vordefinierter Prozess"), T("flowChartInternalStorage", "Flussdiagramm: Zentralspeicher"),
    T("flowChartDocument", "Flussdiagramm: Dokument"), T("flowChartMultidocument", "Flussdiagramm: Mehrere Dokumente"), T("flowChartTerminator", "Flussdiagramm: Grenzstelle"),
    T("flowChartPreparation", "Flussdiagramm: Vorbereitung"), T("flowChartManualInput", "Flussdiagramm: Manuelle Eingabe"), T("flowChartManualOperation", "Flussdiagramm: Manuelle Verarbeitung"),
    T("flowChartConnector", "Flussdiagramm: Verbinder"), T("flowChartOffpageConnector", "Flussdiagramm: Verbinder zu einer anderen Seite"), T("flowChartPunchedCard", "Flussdiagramm: Karte"),
    T("flowChartPunchedTape", "Flussdiagramm: Lochstreifen"), T("flowChartSummingJunction", "Flussdiagramm: Zusammenführung"), T("flowChartOr", "Flussdiagramm: Oder"),
    T("flowChartCollate", "Flussdiagramm: Zusammenstellen"), T("flowChartSort", "Flussdiagramm: Sortieren"), T("flowChartExtract", "Flussdiagramm: Auszug"),
    T("flowChartMerge", "Flussdiagramm: Zusammenführen"), T("flowChartOnlineStorage", "Flussdiagramm: Gespeicherte Daten"), T("flowChartDelay", "Flussdiagramm: Verzögerung"),
    T("flowChartMagneticTape", "Flussdiagramm: Datenträger mit sequenziellem Zugriff"), T("flowChartMagneticDisk", "Flussdiagramm: Magnetplattenspeicher"),
    T("flowChartMagneticDrum", "Flussdiagramm: Datenträger mit direktem Zugriff"), T("flowChartDisplay", "Flussdiagramm: Anzeige"),
  ] },
  { title: "Sterne und Banner", items: [
    T("irregularSeal1", "Explosion: 8 Zacken"), T("irregularSeal2", "Explosion: 14 Zacken"), T("star4", "Stern: 4 Zacken"), T("star5", "Stern: 5 Zacken"),
    T("star6", "Stern: 6 Zacken"), T("star7", "Stern: 7 Zacken"), T("star8", "Stern: 8 Zacken"), T("star10", "Stern: 10 Zacken"), T("star12", "Stern: 12 Zacken"),
    T("star16", "Stern: 16 Zacken"), T("star24", "Stern: 24 Zacken"), T("star32", "Stern: 32 Zacken"), T("ribbon2", "Band: nach oben gekippt"),
    T("ribbon", "Band: nach unten gekippt"), T("ellipseRibbon2", "Band: nach oben gekrümmt und gekippt"), T("ellipseRibbon", "Band: nach unten gekrümmt und gekippt"),
    T("verticalScroll", "Scrollen: vertikal"), T("horizontalScroll", "Scrollen: horizontal"), T("wave", "Welle"), T("doubleWave", "Doppelte Welle"),
  ] },
  { title: "Legenden", items: [
    T("wedgeRectCallout", "Sprechblase: rechteckig"), T("wedgeRoundRectCallout", "Sprechblase: rechteckig mit abgerundeten Ecken"), T("wedgeEllipseCallout", "Sprechblase: oval"),
    T("cloudCallout", "Denkblase: wolkenförmig"), T("borderCallout1", "Legende: Linie"), T("borderCallout2", "Legende: mit gebogener Linie"),
    T("borderCallout3", "Legende: mit gebogener Doppellinie"), T("accentCallout1", "Legende: mit Linie mit Akzentuierungsbalken"),
    T("accentCallout2", "Legende: mit gebogener Linie mit Akzentuierungsbalken"), T("accentCallout3", "Legende: mit gebogener Doppellinie mit Akzentuierungsbalken"),
    T("callout1", "Legende: mit Linie ohne Rahmen"), T("callout2", "Legende: mit gebogener Linie ohne Rahmen"), T("callout3", "Legende: mit gebogener Doppellinie ohne Rahmen"),
    T("accentBorderCallout1", "Legende: mit Linie mit Rahmen und Akzentuierungsbalken"), T("accentBorderCallout2", "Legende: mit gebogener Linie mit Rahmen und Akzentuierungsbalken"),
    T("accentBorderCallout3", "Legende: mit gebogener Doppellinie mit Rahmen und Akzentuierungsbalken"),
  ] },
];

const SVG = "http://www.w3.org/2000/svg";

/** A gallery icon: the preset's outline, as Word draws its gallery (thin dark lines). */
export function shapeIcon(e: Entry, size = 20): SVGSVGElement {
  const svg = document.createElementNS(SVG, "svg");
  svg.setAttribute("width", String(size));
  svg.setAttribute("height", String(size));
  svg.setAttribute("viewBox", `-1.5 -1.5 ${size + 3} ${size + 3}`);
  const t = e.tool;
  let d = "";
  if (t.freeform === "curve") d = `M2 ${size - 4} C${size / 3} 0 ${(2 * size) / 3} ${size} ${size - 2} 4`;
  else if (t.freeform === "polygon") d = `M3 ${size - 3} L5 6 L${size / 2} ${size / 2} L${size - 4} 3 L${size - 3} ${size - 6} Z`;
  else if (t.freeform === "scribble") d = `M2 ${size / 2} C5 2 9 ${size} 12 ${size / 2} S ${size - 2} 4 ${size - 2} ${size - 4}`;
  else if (t.kind === "textbox") d = `M1 1 L${size - 1} 1 L${size - 1} ${size - 1} L1 ${size - 1} Z M6 ${size - 5} L${size / 2} 5 L${size - 6} ${size - 5} M8 ${size / 2 + 2} L${size - 8} ${size / 2 + 2}`;
  else {
    // Callouts reach outside their box: the icon fits the whole outline, as Word's do.
    const box = size * 0.86;
    const g = presetGeometry(t.geometry, box, box * (t.geometry.includes("Callout") || t.geometry.includes("callout") ? 0.62 : 1), t.adjust ?? {});
    d = (g?.paths ?? []).filter((p) => p.stroke || p.fill !== "none").map((p) => p.d).join(" ");
    const pts = flattenPath(d, 4);
    if (pts.length) {
      const xs = pts.map((q) => q[0]), ys = pts.map((q) => q[1]);
      const x0 = Math.min(...xs), y0 = Math.min(...ys), w = Math.max(...xs) - x0 || 1, hh = Math.max(...ys) - y0 || 1;
      const side = Math.max(w, hh);
      svg.setAttribute("viewBox", `${x0 - (side - w) / 2 - 1.5} ${y0 - (side - hh) / 2 - 1.5} ${side + 3} ${side + 3}`);
    }
  }
  const path = document.createElementNS(SVG, "path");
  path.setAttribute("d", d);
  path.setAttribute("fill", "none");
  path.setAttribute("stroke", "currentColor");
  path.setAttribute("stroke-width", "1");
  path.setAttribute("vector-effect", "non-scaling-stroke");
  svg.append(path);
  // Arrowheads at the ends of the line (pointing away from it).
  const head = t.pic?.line?.head, tail = t.pic?.line?.tail;
  const ends = flattenPath(d, 4);
  if (ends.length > 1) {
    const at = (p: [number, number], q: [number, number]) => {
      const ang = (Math.atan2(p[1] - q[1], p[0] - q[0]) * 180) / Math.PI;
      const a = document.createElementNS(SVG, "path");
      a.setAttribute("d", `M${p[0]} ${p[1]} l-4.5 -2.2 l0 4.4 Z`);
      a.setAttribute("transform", `rotate(${ang} ${p[0]} ${p[1]})`);
      a.setAttribute("fill", "currentColor");
      svg.append(a);
    };
    if (tail) at(ends[ends.length - 1], ends[ends.length - 2]);
    if (head) at(ends[0], ends[1]);
  }
  return svg;
}
