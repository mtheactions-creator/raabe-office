// Raabeschule templates for RaabePresent (as RaabeWrite's D-035 and RaabeCalc's D-064): our own
// presentations from ordinary PowerPoint features in the Raabeschule design, so files made from them
// open in PowerPoint unchanged. Offered in Datei > Neu beside the blank presentation.

import { newPptx } from "@raabe/ooxml";
import { Deck } from "./deck";
import { RAABE_DESIGN } from "./designs";
import * as T from "./textops";
import * as S from "./shapeops";

export interface PresentTemplate { id: string; name: string; slides: { layout: string; texts: string[][] }[] }

export const PRESENT_TEMPLATES: PresentTemplate[] = [
  { id: "referat", name: "Referat", slides: [
    { layout: "Titelfolie", texts: [["Thema des Referats"], ["Name · Klasse · Datum"]] },
    { layout: "Titel und Inhalt", texts: [["Gliederung"], ["Einleitung", "Hauptteil", "Ergebnis", "Quellen"]] },
    { layout: "Titel und Inhalt", texts: [["Einleitung"], ["Worum geht es?", "Warum ist das Thema wichtig?"]] },
    { layout: "Zwei Inhalte", texts: [["Hauptteil"], ["Erster Gedanke", "Beleg oder Beispiel"], ["Zweiter Gedanke", "Beleg oder Beispiel"]] },
    { layout: "Titel und Inhalt", texts: [["Ergebnis"], ["Das Wichtigste in einem Satz"]] },
    { layout: "Titel und Inhalt", texts: [["Quellen"], ["Autor: Titel. Verlag, Jahr", "Webseite (aufgerufen am TT.MM.JJJJ)"]] },
  ] },
  { id: "unterricht", name: "Unterrichtsstunde", slides: [
    { layout: "Titelfolie", texts: [["Thema der Stunde"], ["Fach · Klasse · Datum"]] },
    { layout: "Titel und Inhalt", texts: [["Ziele der Stunde"], ["Ihr könnt am Ende …"]] },
    { layout: "Titel und Inhalt", texts: [["Einstieg"], ["Frage oder Impuls"]] },
    { layout: "Titel und Inhalt", texts: [["Arbeitsauftrag"], ["Lest …", "Bearbeitet …", "Zeit: 15 Minuten"]] },
    { layout: "Titel und Inhalt", texts: [["Sicherung"], ["Ergebnisse"]] },
    { layout: "Titel und Inhalt", texts: [["Hausaufgabe"], ["Bis zum …"]] },
  ] },
];

/** A new presentation of the template in the Raabeschule design. */
export function templateDeck(t: PresentTemplate): Deck {
  const deck = new Deck(newPptx(RAABE_DESIGN));
  const layouts = deck.layouts();
  t.slides.forEach((s, i) => {
    const lay = layouts.find((l) => l.name === s.layout)?.path ?? layouts[1].path;
    const path = i === 0 ? deck.slides[0].path : deck.addSlide(i - 1, lay);
    if (i === 0) deck.setLayout([0], lay);
    deck.change("Vorlage", [path], () => {
      const tree = S.treeOf(deck.pptx.doc(path));
      const bodies = S.shapeEls(tree).map((e) => S.txBodyOf(e)).filter((b): b is Element => !!b);
      s.texts.forEach((lines, k) => {
        const body = bodies[k];
        if (!body) return;
        let pos = { p: 0, o: 0 };
        lines.forEach((line, j) => { if (j) pos = T.splitPara(body, pos); pos = T.insertText(body, pos, line); });
      });
    });
  });
  return deck;
}
