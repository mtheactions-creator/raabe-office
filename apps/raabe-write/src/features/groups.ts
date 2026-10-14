// Layout / Formformat / Bildformat > Anordnen > Gruppieren: Word's menu with Gruppieren
// (objects selected with Umschalt+Klick), Gruppierung aufheben and Gruppierung wiederherstellen;
// Formformat > Text > Verknüpfung erstellen / aufheben (text flowing through text boxes).

import { messageBox, type Command, type MenuEntry } from "@raabe/shell";
import type { AppApi } from "./api";

export const GROUP_COMMANDS: Command[] = [
  { id: "shape.group", label: "Gruppieren", icon: "group-objects", tip: "Gruppieren\nObjekte verbinden, damit sie zusammen verschoben und formatiert werden." },
  { id: "shape.ungroup", label: "Gruppierung aufheben", icon: "ungroup-objects" },
  { id: "shape.regroup", label: "Gruppierung wiederherstellen", icon: "group-objects" },
];

export const GROUP_MENU: MenuEntry[] = [{ cmd: "shape.group" }, { cmd: "shape.ungroup" }, { cmd: "shape.regroup" }];

export function groupsFeature(api: AppApi) {
  const ed = () => api.editor();
  const c = api.commands;
  const done = (ok: boolean | undefined) => { if (ok) api.refresh(); api.focus(); };
  c.implement("shape.group", { run: () => done(ed()?.groupFloats()), enabled: () => (ed()?.selectedFloats().length ?? 0) >= 2 });
  c.implement("shape.ungroup", { run: () => done(ed()?.ungroupFloat()), enabled: () => !!ed()?.selectedGroup() });
  c.implement("shape.regroup", { run: () => done(ed()?.regroup()), enabled: () => !!ed()?.canRegroup() });
  c.implement("layout.group", { run: () => {}, enabled: () => (ed()?.selectedFloats().length ?? 0) > 0 || !!ed()?.canRegroup() });

  // Verknüpfung erstellen: the next click on an empty text box links it; the label turns into
  // Verknüpfung aufheben while the box passes its text on.
  const linkCmd = c.get("shape.createLink");
  c.implement("shape.createLink", {
    run: () => {
      const e = ed();
      if (!e) return;
      e.onLinkRefused = (reason) => void messageBox("RaabeWrite", reason);
      if (e.linking) e.cancelLink();
      else if (e.linksOn()) e.breakLink();
      else e.startLink();
      api.refresh();
    },
    enabled: () => { const s = ed()?.selectedShape(); return !!s && s.member === undefined && !!s.obj.shape && s.obj.shape.kind !== "connector"; },
    checked: () => !!ed()?.linking,
  });
  c.onRefresh(() => {
    const on = !!ed()?.linksOn();
    linkCmd.label = on ? "Verknüpfung aufheben" : "Verknüpfung erstellen";
    linkCmd.icon = on ? "break-link" : "create-link";
  });
}
