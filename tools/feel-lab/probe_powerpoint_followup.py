"""Probe: what the RaabePresent audit (D-137) left open, recorded in a PowerPoint of our own (D-152 to D-154).

- smartart: the SmartArt's Format tab (SmartArt-Tools), walked as probe_powerpoint_inventory.py walks a tab (its
  menus, galleries and "..." dialogs listed, never confirmed). The full inventory skipped it: a tab named "Format" had
  been walked for the chart already.
- dauer: Übergänge > Dauer on a slide without a transition: what the spinner shows, a value typed into it, its arrow;
  what PowerPoint writes then (the slide's XML of a saved copy).
- aspect: Tabellenlayout > Seitenverhältnis sperren: the check box, then Höhe typed: what happens to the width, and
  what is written.
- master: Folienmaster: Masterlayout... (its check boxes with a placeholder missing, one put back, where it goes),
  Platzhalter einfügen (its entries; each one picked and a click on the layout, one drawn by dragging: the new
  placeholder's type, size and prompt), Folienmaster einfügen (what the new master is).

Real input into our own PowerPoint only (every click checks the window under the point, D-070): takes the real-input
lock, refuses to run while a POWERPNT.EXE runs (the owner's), never touches a sign-in or consent dialog.

    .venv/Scripts/python.exe probe_powerpoint_followup.py [smartart] [dauer] [aspect] [master]

Result: results/<date>-powerpoint-followup.json (labels and numbers only, feellab.sanitize.clean).
"""
from __future__ import annotations

import datetime
import json
import re
import sys
import time
import zipfile
from pathlib import Path

import probe_powerpoint_inventory as inv_mod
from feellab import uia, winapi
from feellab.powerpoint import powerpoint_pids
from probe_powerpoint_inventory import Blocked, Inventory, Scene, all_desc, end, listed

HERE = Path(__file__).parent
DATE = datetime.date.today().isoformat()
RESULT = HERE / "results" / f"{DATE}-powerpoint-followup.json"
inv_mod.LOCK_TAG = "present-followup"


def save(data: dict) -> None:
    from feellab.sanitize import clean
    RESULT.write_text(json.dumps(clean(data), ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))
    inv_mod.refresh_lock()


def ribbon_ctl(inv: Inventory, name: str, types: tuple[str, ...] | None = None):
    for c in inv.controls():
        if c["name"] == name and (types is None or c["type"] in types):
            return c
    return None


def value_of(el) -> str | None:
    try:
        vp = el.GetCurrentPattern(uia.U.UIA_ValuePatternId)
        return vp.QueryInterface(uia.U.IUIAutomationValuePattern).CurrentValue if vp else None
    except Exception:  # noqa: BLE001
        return None


def toggle_of(el) -> int | None:
    try:
        tp = el.GetCurrentPattern(uia.U.UIA_TogglePatternId)
        return tp.QueryInterface(uia.U.IUIAutomationTogglePattern).CurrentToggleState if tp else None
    except Exception:  # noqa: BLE001
        return None


def spinner_value(c: dict) -> str | None:
    v = value_of(c["el"])
    if v is None:
        for e, d in all_desc(c["el"]):
            if d["type"] == "Edit":
                return value_of(e)
    return v


def type_into(inv: Inventory, c: dict, text: str) -> None:
    """A click into a ribbon box (left part, away from its arrows), Strg+A, the text typed, Enter."""
    x, y, w, h = c["rect"]
    inv.click(x + min(w * 0.3, 30), y + h / 2)
    time.sleep(0.3)
    inv.press("ctrl+a")
    inv.front()
    winapi.type_text(text)
    time.sleep(0.1)
    inv.press("enter")
    time.sleep(0.8)


def slide_xml(inv: Inventory, pres, index: int, part: str | None = None) -> str:
    p = inv.t.tmp / f"followup-{int(time.time() * 1000)}.pptx"
    inv.t.save_copy(pres, p)
    with zipfile.ZipFile(p) as z:
        return z.read(part or f"ppt/slides/slide{index}.xml").decode("utf-8")


def transition_part(xml: str) -> str:
    m = re.search(r"<mc:AlternateContent[^>]*>(?:(?!</mc:AlternateContent>).)*<p:transition.*?</mc:AlternateContent>|<p:transition\b.*?(?:/>|</p:transition>)", xml, re.S)
    return m.group(0) if m else ""


# ----- Übergänge > Dauer --------------------------------------------------------------------------------

def dauer(inv: Inventory, scene: Scene, out: dict) -> None:
    pres, w, s1, s2 = scene.reset()
    w.View.GotoSlide(2)
    w.Selection.Unselect()
    time.sleep(0.4)
    inv.select_tab("Übergänge")
    time.sleep(0.6)
    c = ribbon_ctl(inv, "Dauer", ("Spinner", "Edit", "ComboBox"))
    r: dict = {"before": {"shown": spinner_value(c) if c else None, "enabled": c["enabled"] if c else None,
                          "effect": int(s2.SlideShowTransition.EntryEffect), "duration": float(s2.SlideShowTransition.Duration)}}
    type_into(inv, c, "3")
    c = ribbon_ctl(inv, "Dauer", ("Spinner", "Edit", "ComboBox"))
    r["typed_3"] = {"shown": spinner_value(c), "effect": int(s2.SlideShowTransition.EntryEffect), "duration": float(s2.SlideShowTransition.Duration),
                    "xml": transition_part(slide_xml(inv, pres, 2))}
    # The other slide (no transition either): the spinner's up arrow once.
    w.View.GotoSlide(1)
    w.Selection.Unselect()
    time.sleep(0.5)
    inv.select_tab("Übergänge")
    c = ribbon_ctl(inv, "Dauer", ("Spinner", "Edit", "ComboBox"))
    r["slide1_before"] = {"shown": spinner_value(c), "duration": float(s1.SlideShowTransition.Duration)}
    ups = [(e, d) for e, d in all_desc(c["el"]) if d["type"] == "Button" and d["name"] == "Mehr"] if c else []
    if not ups:
        # The arrows sit beside the box, not in it: the first "Mehr" right of the box on its row.
        ups = [(e, d) for e, d in all_desc(inv.lower()) if d["type"] == "Button" and d["name"] == "Mehr" and abs(d["rect"][1] - c["rect"][1]) < 12 and d["rect"][0] >= c["rect"][0]]
        ups.sort(key=lambda p: p[1]["rect"][0])
    if ups:
        x, y, ww, hh = ups[0][1]["rect"]
        inv.click(x + ww / 2, y + hh / 2)
        time.sleep(0.8)
        c = ribbon_ctl(inv, "Dauer", ("Spinner", "Edit", "ComboBox"))
        r["slide1_up"] = {"shown": spinner_value(c), "effect": int(s1.SlideShowTransition.EntryEffect), "duration": float(s1.SlideShowTransition.Duration),
                          "xml": transition_part(slide_xml(inv, pres, 1))}
    # Then an effect picked on slide 2 (from COM): does it keep the duration typed before?
    w.View.GotoSlide(2)
    s2.SlideShowTransition.EntryEffect = 3844  # ppEffectFade (Verblassen)
    time.sleep(0.4)
    r["fade_after"] = {"duration": float(s2.SlideShowTransition.Duration), "xml": transition_part(slide_xml(inv, pres, 2))}
    out["dauer"] = r


# ----- Seitenverhältnis sperren ----------------------------------------------------------------------------

def aspect(inv: Inventory, scene: Scene, out: dict) -> None:
    pres, w, s1, s2 = scene.reset()
    sh = s2.Shapes.AddTable(3, 3, 120, 120, 600, 150)
    w.View.GotoSlide(2)
    sh.Select()
    time.sleep(0.5)
    inv.select_tab("Tabellenlayout")
    time.sleep(0.5)
    cb = ribbon_ctl(inv, "Seitenverhältnis sperren", ("CheckBox",))
    r: dict = {"before": {"checked": toggle_of(cb["el"]), "lock": int(sh.LockAspectRatio), "w": sh.Width, "h": sh.Height}}
    x, y, ww, hh = cb["rect"]
    inv.click(x + 10, y + hh / 2)
    time.sleep(0.6)
    cb = ribbon_ctl(inv, "Seitenverhältnis sperren", ("CheckBox",))
    r["checked"] = {"checked": toggle_of(cb["el"]), "lock": int(sh.LockAspectRatio), "xml_locks": re.findall(r"<a:graphicFrameLocks[^>]*>", slide_xml(inv, pres, 2))}
    hc = ribbon_ctl(inv, "Höhe", ("Spinner", "Edit"))
    r["height_shown"] = spinner_value(hc)
    type_into(inv, hc, "8")
    hc = ribbon_ctl(inv, "Höhe", ("Spinner", "Edit"))
    bc = ribbon_ctl(inv, "Breite", ("Spinner", "Edit"))
    r["height_8cm"] = {"w": sh.Width, "h": sh.Height, "shown_h": spinner_value(hc), "shown_w": spinner_value(bc),
                       "rows": [sh.Table.Rows(i).Height for i in range(1, 4)], "cols": [sh.Table.Columns(i).Width for i in range(1, 4)]}
    type_into(inv, bc, "20")
    r["width_20cm"] = {"w": sh.Width, "h": sh.Height}
    # Off again: the width alone.
    cb = ribbon_ctl(inv, "Seitenverhältnis sperren", ("CheckBox",))
    x, y, ww, hh = cb["rect"]
    inv.click(x + 10, y + hh / 2)
    time.sleep(0.5)
    bc = ribbon_ctl(inv, "Breite", ("Spinner", "Edit"))
    type_into(inv, bc, "25")
    r["unchecked_width_25cm"] = {"w": sh.Width, "h": sh.Height, "lock": int(sh.LockAspectRatio)}
    out["aspect"] = r


# ----- Folienmaster --------------------------------------------------------------------------------------

def phs(shapes) -> list[dict]:
    got = []
    for i in range(1, shapes.Count + 1):
        s = shapes(i)
        try:
            t = int(s.PlaceholderFormat.Type)
        except Exception:  # noqa: BLE001
            t = None
        txt = ""
        try:
            if s.HasTextFrame:
                txt = s.TextFrame.TextRange.Text[:80]
        except Exception:  # noqa: BLE001
            pass
        got.append({"name": s.Name, "ph": t, "l": round(s.Left, 2), "t": round(s.Top, 2), "w": round(s.Width, 2), "h": round(s.Height, 2), "text": txt})
    return got


def menu_items(inv: Inventory, c: dict) -> list[tuple]:
    """Opens a split button's menu; returns the (element, description) pairs of its entries."""
    x, y, w, h = c["rect"]
    arrow = [uia.describe(k) for k in uia.children(c["el"]) if uia.describe(k)["name"] in ("Weitere Optionen", "Mehr Optionen")]
    if arrow:
        x, y, w, h = arrow[0]["rect"]
    else:
        x, w = x + w - 8, 8
    before = set(inv.popups())
    inv.click(x + w / 2, y + h / 2)
    time.sleep(1.0)
    inv.check_blocking()
    pairs = list(all_desc(c["el"]))
    for hwnd in inv.popups():
        if hwnd not in before:
            pairs += all_desc(uia.element_for(hwnd))
    return [(e, d) for e, d in pairs if d["type"] in ("MenuItem", "ListItem", "Button") and d["name"] and d["rect"][2] > 0
            and d["name"] not in ("Weitere Optionen", "Mehr Optionen", c["name"])]


def master(inv: Inventory, scene: Scene, out: dict) -> None:
    pres, w, s1, s2 = scene.reset()
    r: dict = {}
    w.ViewType = 2  # ppViewSlideMaster
    time.sleep(1.0)
    m = pres.SlideMaster
    r["master_before"] = phs(m.Shapes)
    r["layouts"] = [m.CustomLayouts(i).Name for i in range(1, m.CustomLayouts.Count + 1)]
    # The master itself selected (its first shape, then nothing): Masterlayout... is for the master only.
    m.Shapes(1).Select()
    time.sleep(0.3)
    w.Selection.Unselect()
    time.sleep(0.5)
    inv.select_tab("Folienmaster")
    r["master_selected_controls"] = [{"name": c["name"], "enabled": c["enabled"]} for c in inv.controls() if c.get("group") in ("Masterlayout", "Master bearbeiten")]
    # Datum taken off the master, then Masterlayout...: the check boxes, Datum ticked again, OK.
    for i in range(m.Shapes.Count, 0, -1):
        try:
            if int(m.Shapes(i).PlaceholderFormat.Type) == 16:  # ppPlaceholderDate
                m.Shapes(i).Delete()
        except Exception:  # noqa: BLE001
            pass
    w.Selection.Unselect()
    time.sleep(0.4)
    c = ribbon_ctl(inv, "Masterlayout...")
    if c and c["enabled"]:
        x, y, ww, hh = c["rect"]
        inv.click(x + ww / 2, y + hh / 2)
        end_t = time.time() + 3
        while time.time() < end_t and not inv.dialog_windows():
            time.sleep(0.2)
        dl = inv.dialog_windows()
        if dl:
            root = uia.element_for(dl[0])
            boxes = [(e, d) for e, d in all_desc(root) if d["type"] == "CheckBox"]
            r["dialog"] = {"title": winapi.window_title(dl[0]), "boxes": [{"name": d["name"], "enabled": d["enabled"], "checked": toggle_of(e)} for e, d in boxes]}
            for e, d in boxes:
                if d["name"] == "Datum":
                    bx, by, bw, bh = d["rect"]
                    inv.click(bx + 8, by + bh / 2)
                    time.sleep(0.3)
            okb = [(e, d) for e, d in all_desc(root) if d["type"] == "Button" and d["name"] == "OK"]
            if okb:
                uia.invoke(okb[0][0])
            time.sleep(0.8)
            r["master_after_dialog"] = phs(m.Shapes)
            try:
                mx = slide_xml(inv, pres, 0, "ppt/slideMasters/slideMaster1.xml")
                r["master_after_dialog_xml"] = [x for x in re.findall(r"<p:sp>.*?</p:sp>", mx, re.S) if 'type="dt"' in x]
            except Exception as e:  # noqa: BLE001
                r["master_after_dialog_xml"] = f"error: {str(e)[:80]}"
        else:
            r["dialog"] = "no dialog"
    # A layout selected: Platzhalter einfügen's entries; each picked, then a click on the layout.
    lay = m.CustomLayouts(7)  # Leer
    lay.Select()
    time.sleep(0.6)
    inv.select_tab("Folienmaster")
    r["layout_selected_controls"] = [{"name": c["name"], "enabled": c["enabled"]} for c in inv.controls() if c.get("group") in ("Masterlayout", "Master bearbeiten")]
    c = ribbon_ctl(inv, "Platzhalter einfügen")
    items = menu_items(inv, c) if c else []
    names = []
    for _, d in items:
        if d["name"] not in names:
            names.append(d["name"])
    r["placeholder_menu"] = names
    inv.close_all()
    picks = {}
    for k, name in enumerate(names):
        lay.Select()
        time.sleep(0.4)
        inv.select_tab("Folienmaster")
        c = ribbon_ctl(inv, "Platzhalter einfügen")
        hit = [(e, d) for e, d in menu_items(inv, c) if d["name"] == name]
        if not hit:
            picks[name] = "not found"
            inv.close_all()
            continue
        bx, by, bw, bh = hit[0][1]["rect"]
        inv.click(bx + bw / 2, by + bh / 2)
        time.sleep(0.6)
        n0 = lay.Shapes.Count
        if k == 0:
            x1, y1 = inv.px(100, 100)
            x2, y2 = inv.px(400, 300)
            inv.front()
            winapi.drag(x1, y1, x2, y2)
        else:
            x1, y1 = inv.px(300, 200)
            inv.click(x1, y1)
        time.sleep(0.8)
        got = phs(lay.Shapes)[n0:] if lay.Shapes.Count > n0 else []
        picks[name] = {"dragged" if k == 0 else "clicked": got}
        try:
            # What PowerPoint writes for it: the layout's last shape in a saved copy.
            lx = slide_xml(inv, pres, 0, "ppt/slideLayouts/slideLayout7.xml")
            sps = re.findall(r"<p:sp>.*?</p:sp>", lx, re.S)
            picks[name]["xml"] = sps[-1] if sps and got else ""
        except Exception as e:  # noqa: BLE001
            picks[name]["xml"] = f"error: {str(e)[:80]}"
        for i in range(lay.Shapes.Count, n0, -1):
            lay.Shapes(i).Delete()
        inv.press("esc")
        time.sleep(0.3)
    r["placeholder_picks"] = picks
    # The ribbon's main part of the split button (last pick shown?): one click, then a click on the layout.
    lay.Select()
    time.sleep(0.4)
    inv.select_tab("Folienmaster")
    c = ribbon_ctl(inv, "Platzhalter einfügen")
    if c:
        btns = [d for e, d in all_desc(c["el"]) if d["type"] == "Button" and d["name"] not in ("Weitere Optionen", "Mehr Optionen")]
        x, y, ww, hh = (btns[0]["rect"] if btns else c["rect"])
        inv.click(x + ww / 2, y + min(hh / 2, 30))
        time.sleep(0.6)
        n0 = lay.Shapes.Count
        x1, y1 = inv.px(300, 200)
        inv.click(x1, y1)
        time.sleep(0.8)
        r["placeholder_main_button"] = phs(lay.Shapes)[n0:]
        for i in range(lay.Shapes.Count, n0, -1):
            lay.Shapes(i).Delete()
    # Folienmaster einfügen.
    d0 = pres.Designs.Count
    c = ribbon_ctl(inv, "Folienmaster einfügen")
    if c:
        x, y, ww, hh = c["rect"]
        inv.click(x + ww / 2, y + hh / 2)
        time.sleep(1.5)
        r["insert_master"] = {"designs_before": d0, "designs_after": pres.Designs.Count}
        if pres.Designs.Count > d0:
            dn = pres.Designs(pres.Designs.Count)
            mm = dn.SlideMaster
            r["insert_master"].update({"design_name": dn.Name, "preserved": int(dn.Preserved), "master_shapes": phs(mm.Shapes),
                                       "layouts": [mm.CustomLayouts(i).Name for i in range(1, mm.CustomLayouts.Count + 1)],
                                       "theme_colors_accent1": int(mm.Theme.ThemeColorScheme(5).RGB),
                                       "fonts": [mm.Theme.ThemeFontScheme.MajorFont(1).Name, mm.Theme.ThemeFontScheme.MinorFont(1).Name]})
            try:
                p = inv.t.tmp / "followup-master.pptx"
                inv.t.save_copy(pres, p)
                with zipfile.ZipFile(p) as z:
                    names_ = z.namelist()
                    r["insert_master"]["parts"] = sorted(n for n in names_ if re.match(r"ppt/(slideMasters|theme)/[^/]+\.xml$", n))
                    r["insert_master"]["presentation_masters"] = re.findall(r"<p:sldMasterId [^>]*>", z.read("ppt/presentation.xml").decode("utf-8"))
            except Exception as e:  # noqa: BLE001
                r["insert_master"]["save_error"] = str(e)[:120]
    w.ViewType = 9
    out["master"] = r


# ----- SmartArt Format ---------------------------------------------------------------------------------------

def smartart(inv: Inventory, scene: Scene, out: dict) -> None:
    sa = scene.build("SmartArt")
    time.sleep(1.0)
    got = inv.walk_tab("Format")
    try:
        inv.second_pass("Format", got["controls"])
    except Blocked:
        raise
    except Exception as e:  # noqa: BLE001
        got["second_pass_error"] = str(e)[:100]
    for c in got["controls"]:
        c.pop("el", None)
    out["smartart_format"] = got
    out["smartart_format_dialogs"] = {k: v for k, v in inv.dialogs.items() if k.startswith("Format/")}
    # Also the controls with a shape inside the SmartArt selected (PowerPoint enables Form ändern, Größer, Kleiner).
    try:
        sa.SmartArt.AllNodes(1).Shapes(1).Select()
        time.sleep(0.6)
        inv.select_tab("Format")
        out["smartart_format_node_selected"] = [{"group": c["group"], "name": c["name"], "type": c["type"], "enabled": c["enabled"]} for c in inv.controls()]
    except Exception as e:  # noqa: BLE001
        out["smartart_format_node_selected"] = f"error: {str(e)[:100]}"


PARTS = {"dauer": dauer, "aspect": aspect, "master": master, "smartart": smartart}


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    pick = [a for a in sys.argv[1:] if a in PARTS] or list(PARTS)
    if powerpoint_pids():
        print("POWERPNT.EXE is running (not ours): no PowerPoint probe now", flush=True)
        return 76
    if not inv_mod.acquire_lock():
        return 75
    out = json.loads(RESULT.read_text(encoding="utf-8")) if RESULT.exists() else {}
    out.update({"date": DATE, "built_by": "PowerPoint (Microsoft 365, German UI)"})
    inv = None
    try:
        inv = Inventory()
        scene = Scene(inv)
        for p in pick:
            try:
                PARTS[p](inv, scene, out)
            except Blocked:
                raise
            except Exception as e:  # noqa: BLE001
                out[f"{p}_error"] = str(e)[:300]
                print(p, "failed:", str(e)[:300], flush=True)
                try:
                    inv.close_all()
                except Exception:  # noqa: BLE001
                    pass
            save(out)
            print(p, "done", flush=True)
    except Blocked as e:
        out["blocked"] = str(e)[:80]
        print("BLOCKED:", e, flush=True)
    finally:
        out["close"] = end(inv, kill=True) if inv else None
        inv_mod.release_lock()
        save(out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
