"""Round trips with PowerPoint (D-070), both directions, invisible (COM without windows, headless Edge):

1. PowerPoint's fixtures through RaabePresent: each file opened and saved by RaabePresent unchanged,
   then opened by PowerPoint; PowerPoint must read the same as from its own file (read-back JSON).
2. Files made by RaabePresent: scenarios built through the app's commands (text, levels, formatting,
   shapes, table, group, notes, sections, hidden slide, transition, animation, design, picture, chart),
   saved, opened by PowerPoint; PowerPoint's read-back is checked against what the scenario made.

    .venv/Scripts/python.exe check_present_roundtrip.py

Needs the dev server (npm run dev:present) and no running POWERPNT.EXE (the owner's: then nothing is
done). Result: results/<date>-present-roundtrip.json.
"""
from __future__ import annotations

import base64
import datetime
import json
import os
import sys
import tempfile
from pathlib import Path

from feellab.powerpoint import OwnerPowerPointRunning, PowerPointTarget
from feellab.present_page import PresentPage

HERE = Path(__file__).parent
FIX = HERE.parent.parent / "packages" / "ooxml" / "test" / "fixtures"
TMP = Path(tempfile.mkdtemp(prefix="raabe-present-rt-"))
# RT_ONLY="name,name": only these scenarios, no fixtures, nothing written (for checking one change).
ONLY = [x.strip() for x in os.environ.get("RT_ONLY", "").split(",") if x.strip()]
FIXTURES = ["pptx-blank", "pptx-layouts", "pptx-bullets", "pptx-textbox", "pptx-shapes", "pptx-group", "pptx-picture", "pptx-table", "pptx-autofit", "pptx-show", "pptx-equation"]
# PowerPoint's own default design (Office, Aptos) as well (added in the overnight run, D-072).
FIXTURES += [f.replace("pptx-", "pptx-office-") for f in FIXTURES]

# A tiny red PNG (2x2) for the picture scenario.
PNG = "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEklEQVR4nGP4z8DwnwEIGBgYAAA+5gT8uY8z6AAAAABJRU5ErkJggg=="

SCENARIOS: dict[str, str] = {
    "title text": """p.newDoc(); const a = p.app; const ed = a.editor;
        ed.enterText(2); ed.typeText('Mein Titel'); ed.endText(); ed.enterText(3); ed.typeText('Untertitel'); ed.endText();""",
    "bullets levels": """p.newDoc(); const a = p.app; const ed = a.editor; a.newSlide(); const id = a.deck.view(a.slidePath()).shapes[1].id;
        ed.enterText(id); ed.typeText('Eins\\nZwei'); ed.changeLevel(1); ed.typeText('\\nDrei'); ed.changeLevel(-1); ed.endText();""",
    "shapes": """p.setupCase({layout:'Leer', shapes:[{kind:'rect',box:[100,100,200,100],text:'Rechteck'},{kind:'ellipse',box:[400,100,150,150]},{kind:'textbox',box:[100,300,300,40],text:'Textfeld'}]});""",
    "rotate flip": """p.setupCase({layout:'Leer', shapes:[{kind:'rect',box:[100,100,200,100]}]}); const a = p.app; a.editor.select([2]); await p.run('arrange.rotRight'); await p.run('arrange.flipH');""",
    "bold part": """p.setupCase({layout:'Leer', shapes:[{kind:'textbox',box:[100,100,400,40],text:'normal fett'}]}); const ed = p.app.editor; ed.enterText(2, {p:0,o:7}); ed.text.anchor = {p:0,o:11}; await p.run('font.bold'); ed.endText();""",
    "size color": """p.setupCase({layout:'Leer', shapes:[{kind:'textbox',box:[100,100,400,40],text:'Gross und rot'}]}); const a = p.app; a.editor.select([2]); a.ribbonContext().setSize('32'); a.ribbonContext().setFontColor('#FF0000');""",
    "align center": """p.setupCase({layout:'Leer', shapes:[{kind:'textbox',box:[100,100,400,40],text:'Mitte'}]}); p.app.editor.select([2]); await p.run('para.center');""",
    "table": """p.newDoc(); const a = p.app; a.deck.deleteSlides([0]); a.deck.addSlide(-1, a.deck.layouts()[6].path); a.setCurrent(0); a.insertTable(2, 3);""",
    "table merge": """p.newDoc(); const a = p.app; a.deck.setLayout([0], a.deck.layouts()[6].path); a.setCurrent(0); a.insertTable(2, 3); const id = a.deck.view(a.slidePath()).shapes.find(s => s.frame === 'table').id;
        a.editor.enterText(id, undefined, {r:0,c:0}); a.editor.typeText('eins'); a.editor.setRange({r:1,c:1}); await p.run('table.merge'); a.editor.endText();""",
    "table split": """p.newDoc(); const a = p.app; a.deck.setLayout([0], a.deck.layouts()[6].path); a.setCurrent(0); a.insertTable(2, 2); const id = a.deck.view(a.slidePath()).shapes.find(s => s.frame === 'table').id;
        a.editor.enterText(id, undefined, {r:0,c:0}); const m = await import('/src/tablecells.ts'); a.editor.edit('Zellen teilen', () => m.split(a.editor.tableOf(), 0, 0, 2, 1)); a.editor.endText();""",
    "equation": """p.newDoc(); const a = p.app; a.deck.setLayout([0], a.deck.layouts()[6].path); a.setCurrent(0); const m = await import('/src/equations.ts'); const b = await import('/src/equations-builtin.ts'); await m.insertEquation(a, b.BUILT_IN.find(e => e.name === 'Satz des Pythagoras').seq);""",
        "equation typed": """p.newDoc(); const a = p.app; a.deck.setLayout([0], a.deck.layouts()[6].path); a.setCurrent(0); const m = await import('/src/equations.ts'); await m.insertEquation(a); await new Promise(r => setTimeout(r, 300));
        const key = (k) => window.dispatchEvent(new KeyboardEvent('keydown', {key: k, bubbles: true, cancelable: true})); for (const ch of 'a^2 +b^2 =c^2 ') key(ch); key('Enter'); await new Promise(r => setTimeout(r, 300));""",
        "table style": """p.newDoc(); const a = p.app; a.deck.setLayout([0], a.deck.layouts()[6].path); a.setCurrent(0); a.insertTable(3, 3); const m = await import('/src/tablestyles.ts'); m.applyTableStyle(a, m.tableStyles().find(s => s.name === 'Helle Formatvorlage 2 - Akzent 2'));""",
        "motion path": """p.setupCase({layout:'Leer', shapes:[{kind:'rect',box:[100,100,200,100]}]}); const a = p.app; a.editor.select([2]); const m = await import('/src/animations.ts'); const g = m.ANIMATION_GROUPS.find(x => x.cls === 'path'); m.addAnimation(a, 'path', g.items[0], false);""",
        "custom path": """p.setupCase({layout:'Leer', shapes:[{kind:'rect',box:[100,100,200,100]}]}); const a = p.app; a.editor.select([2]); const m = await import('/src/animations.ts'); const g = m.ANIMATION_GROUPS.find(x => x.cls === 'path'); m.addAnimation(a, 'path', {...g.items.find(i => i.preset === 0), path: 'M 0 0 L 0.2 0 L 0.2 0.25 L 0.4 0.25 E'}, false);""",
        "freeform": """p.setupCase({layout:'Leer', shapes:[]}); const a = p.app; a.editor['addFreeform']([{x:1000000,y:1000000},{x:3000000,y:1500000},{x:1500000,y:3000000}], true, false); a.editor['addFreeform']([{x:5000000,y:1000000},{x:6000000,y:2000000},{x:7000000,y:1200000},{x:8000000,y:2500000}], false, true);""",
        "zoom": """p.newDoc(); const a = p.app; a.newSlide(); a.newSlide(); a.setCurrent(0); const m = await import('/src/zoom.ts'); await m.insertSlideZooms(a, [2]); await new Promise(r => setTimeout(r, 300)); const v = a.deck.view(a.slidePath()).shapes.find(s => s.frame === 'zoom'); a.editor.select([v.id]); a.editor.edit('Verschieben', () => { const S = a.editor.shapeEl(v.id); const x = S.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'off')[0]; x.setAttribute('x', '1000000'); });""",
        "section zoom": """p.newDoc(); const a = p.app; a.newSlide(); a.newSlide(); a.deck.setSections([{name:'Anfang', id:'{11111111-1111-1111-1111-111111111111}', slides:[a.deck.slides[0].id]},{name:'Teil 2', id:'{22222222-2222-2222-2222-222222222222}', slides:[a.deck.slides[1].id, a.deck.slides[2].id]}]); a.setCurrent(0); const m = await import('/src/zoom.ts'); await m.insertZooms(a, [{slide: 1, section: '{22222222-2222-2222-2222-222222222222}'}]);""",
        "summary zoom": """p.newDoc(); const a = p.app; a.newSlide(); a.newSlide(); const ed = a.editor; for (const i of [1, 2]) { a.setCurrent(i); const t = a.deck.view(a.slidePath()).shapes[0]; ed.enterText(t.id); ed.typeText('Teil ' + i); ed.endText(); } const m = await import('/src/zoom.ts'); await m.insertSummaryZoom(a, [1, 2]);""",
        "file pptm": """p.newDoc(); const a = p.app; const {PRES_TYPES} = await import('/src/app.ts'); a.deck.pptx.setType(a.deck.pptx.presPath, PRES_TYPES.pptm);""",
        "file ppsx": """p.newDoc(); const a = p.app; a.newSlide(); const {PRES_TYPES} = await import('/src/app.ts'); a.deck.pptx.setType(a.deck.pptx.presPath, PRES_TYPES.ppsx);""",
        "file potx": """p.newDoc(); const a = p.app; const {PRES_TYPES} = await import('/src/app.ts'); a.deck.pptx.setType(a.deck.pptx.presPath, PRES_TYPES.potx);""",
        "shape effects": """p.setupCase({layout:'Leer', shapes:[{kind:'rect',box:[100,100,200,120]},{kind:'ellipse',box:[450,100,160,160]},{kind:'textbox',box:[100,320,400,60],text:'Texteffekt'}]}); const a = p.app; const m = await import('/src/shapefx.ts');
        const pick = (menu, path) => { let entries = menu; for (const lab of path.slice(0, -1)) { const e = entries.find(x => x.label === lab); entries = typeof e.submenu === 'function' ? e.submenu() : e.submenu; } entries.find(x => x.label === path[path.length - 1]).run(); };
        a.editor.select([2]); pick(m.shapeEffectsMenu(a), ['Abschrägung', 'Rund']); pick(m.shapeEffectsMenu(a), ['3D-Drehung', 'Perspektive: von links']); pick(m.shapeEffectsMenu(a), ['Schatten', 'Offset: unten rechts']);
        a.editor.select([3]); pick(m.shapeEffectsMenu(a), ['Leuchteffekt', 'Leuchteffekt: 11 Pt.; Orange, Akzentfarbe 2']); pick(m.shapeEffectsMenu(a), ['Spiegelung', 'Halbe Spiegelung: mit Berührung']);
        a.editor.select([4]); pick(m.transformMenu(a), ['Arch']); pick(m.textThreeD(a, 'bevel'), ['Konvex']);""",
        "action button": """p.setupCase({layout:'Leer', shapes:[]}); const a = p.app; a.newSlide(); a.setCurrent(0); const S = await import('/src/shapeops.ts');
        a.editor.edit('Form einfügen', () => { const doc = a.deck.pptx.doc(a.slidePath()); const tree = S.treeOf(doc); const el = S.newShape(doc, S.maxId(tree) + 1, 'actionButtonForwardNext', {x: 1000000, y: 1000000, w: 900000, h: 900000, rot: 0, flipH: false, flipV: false}); tree.append(el); a.actionButtonDrawn(el, 'actionButtonForwardNext'); });""",
        "fills and background": """p.setupCase({layout:'Leer', shapes:[{kind:'rect',box:[100,100,200,120]},{kind:'ellipse',box:[400,100,160,160]}]}); const a = p.app;
        const pick = (menu, path) => { let entries = menu; for (const lab of path.slice(0, -1)) { const e = entries.find(x => x.label === lab); entries = typeof e.submenu === 'function' ? e.submenu() : e.submenu; } entries.find(x => x.label === path[path.length - 1]).run(); };
        a.editor.select([2]); pick(a.fillExtra(), ['Farbverlauf', 'Linear unten']); a.editor.select([3]); a.ribbonContext().setFill('accent2,60000,40000');
        const bg = await import('/src/background.ts'); bg.bgStylesMenu(a)[0].block().querySelectorAll('button')[6].click();""",
        "table shading": """p.newDoc(); const a = p.app; a.deck.setLayout([0], a.deck.layouts()[6].path); a.setCurrent(0); a.insertTable(2, 2); const id = a.deck.view(a.slidePath()).shapes.find(s => s.frame === 'table').id;
        a.editor.enterText(id, undefined, {r:0,c:0}); const m = await import('/src/contextcmds.ts'); m.setShading(a, 'C00000'); a.editor.endText();""",
        "table effects": """p.newDoc(); const a = p.app; a.deck.setLayout([0], a.deck.layouts()[6].path); a.setCurrent(0); a.insertTable(2, 2); const id = a.deck.view(a.slidePath()).shapes.find(s => s.frame === 'table').id;
        a.editor.select([id]); const m = a.host2().tableEffects(); m[1].submenu()[1].run(); a.editor.enterText(id, undefined, {r:0,c:0}); m[0].submenu()[1].run(); a.editor.endText();""",
        "table pen": """p.newDoc(); const a = p.app; a.deck.setLayout([0], a.deck.layouts()[6].path); a.setCurrent(0); a.insertTable(2, 2); const id = a.deck.view(a.slidePath()).shapes.find(s => s.frame === 'table').id;
        a.tablePen = { dash: 'sysDash', w: 38100, color: 'C00000' }; a.penCellLine(id, 0, 0, 'r'); await new Promise(r => setTimeout(r, 300));""",
        "action": """p.setupCase({layout:'Leer', shapes:[{kind:'rect',box:[100,100,200,100]}]}); const a = p.app; a.newSlide(); a.setCurrent(0); a.editor.select([2]); const doc = a.deck.pptx.doc(a.slidePath()); const nv = doc.getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'cNvPr')[1]; a.editor.edit('Aktion', () => { const hl = doc.createElementNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'a:hlinkClick'); hl.setAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'r:id', ''); hl.setAttribute('action', 'ppaction://hlinkshowjump?jump=nextslide'); nv.insertBefore(hl, nv.firstChild); });""",
        "animation copy": """p.setupCase({layout:'Leer', shapes:[{kind:'rect',box:[100,100,200,100]},{kind:'ellipse',box:[400,100,150,150]}]}); const a = p.app; a.editor.select([2]); const m = await import('/src/animations.ts'); m.addAnimation(a, 'entr', m.ANIMATION_GROUPS[0].items[2], false); m.copyAnimation(a, 2, 3);""",
        "trigger": """p.setupCase({layout:'Leer', shapes:[{kind:'rect',box:[100,100,200,100]},{kind:'ellipse',box:[400,100,150,150]}]}); const a = p.app; a.editor.select([3]); const m = await import('/src/animations.ts'); m.addAnimation(a, 'entr', m.ANIMATION_GROUPS[0].items[2], false); m.setTrigger(a, 2);""",
        "group": """p.setupCase({layout:'Leer', shapes:[{kind:'rect',box:[100,100,100,100]},{kind:'ellipse',box:[300,150,100,100]}]}); p.app.editor.select([2,3]); await p.run('arrange.group');""",
    "notes": """p.newDoc(); p.app.deck.setNotes(p.app.deck.slides[0].path, 'Meine Notiz\\nZweite Zeile');""",
    "hidden slide": """p.newDoc(); const a = p.app; a.newSlide(); a.panel.selected = new Set([1]); await p.run('slide.hide');""",
    "sections": """p.newDoc(); const a = p.app; a.newSlide(); a.newSlide(); a.deck.setSections([{name:'Anfang', id:'{11111111-1111-1111-1111-111111111111}', slides:[a.deck.slides[0].id]},{name:'Teil 2', id:'{22222222-2222-2222-2222-222222222222}', slides:[a.deck.slides[1].id, a.deck.slides[2].id]}]);""",
    "move duplicate delete": """p.newDoc(); const a = p.app; a.newSlide(); a.newSlide(); a.deck.moveSlides([2], 0); a.deck.duplicateSlides([0]); a.deck.deleteSlides([3]);""",
    "layout change": """p.newDoc(); const a = p.app; a.deck.setLayout([0], a.deck.layouts()[1].path);""",
    "transition": """p.newDoc(); const a = p.app; const m = await import('/src/transitions.ts'); a.panel.selected = new Set([0]); m.setTransition(a, 'fade');""",
    "animation": """p.setupCase({layout:'Leer', shapes:[{kind:'rect',box:[100,100,200,100]}]}); const a = p.app; a.editor.select([2]); const m = await import('/src/animations.ts'); m.addAnimation(a, 'entr', m.ANIMATION_GROUPS[0].items[2], false);""",
    "design raabe": """p.newDoc(); const a = p.app; const d = await import('/src/designs.ts'); a.applyDesign(d.RAABE_DESIGN); await new Promise(r => setTimeout(r, 300));""",
    "picture": f"""p.newDoc(); const a = p.app; a.deck.setLayout([0], a.deck.layouts()[6].path); const bin = atob('{PNG}'); const b = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) b[i] = bin.charCodeAt(i); a.insertPictureBytes(b, 'image/png'); await new Promise(r => setTimeout(r, 500));""",
    "picture effects": f"""p.newDoc(); const a = p.app; a.deck.setLayout([0], a.deck.layouts()[6].path); const bin = atob('{PNG}'); const b = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) b[i] = bin.charCodeAt(i); a.insertPictureBytes(b, 'image/png'); await new Promise(r => setTimeout(r, 500));
        const v = a.deck.view(a.slidePath()).shapes.find(s => s.kind === 'pic'); a.editor.select([v.id]); const m = await import('/src/picfx.ts');
        await m.setEffect(a, 'saturation', {{name: 'saturation', attrs: {{sat: '0'}}}}, 'Farbsättigung'); await m.setEffect(a, 'artistic', {{name: 'artisticBlur', attrs: {{}}}}, 'Künstlerische Effekte'); await m.setEffect(a, 'sharpenSoften', {{name: 'sharpenSoften', attrs: {{amount: '25000'}}}}, 'Schärfen');""",
    "background removal": f"""p.newDoc(); const a = p.app; a.deck.setLayout([0], a.deck.layouts()[6].path); const bin = atob('{PNG}'); const b = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) b[i] = bin.charCodeAt(i); a.insertPictureBytes(b, 'image/png'); await new Promise(r => setTimeout(r, 500));
        const v = a.deck.view(a.slidePath()).shapes.find(s => s.kind === 'pic'); a.editor.select([v.id]); await p.run('pic.removeBg'); await new Promise(r => setTimeout(r, 500));
        await p.run('bg.markKeep'); a.editor.tool.done([{{x: v.x + 1, y: v.y + 1}}, {{x: v.x + v.w / 2, y: v.y + v.h / 2}}], false); await new Promise(r => setTimeout(r, 500)); await p.run('bg.keep');""",
    "reuse slides": f"""p.newDoc(); const a = p.app; const r = await fetch('/@fs/{(FIX / "pptx-picture.pptx").as_posix()}'); const bytes = new Uint8Array(await r.arrayBuffer());
        const src = a.deck.pptx.constructor.read(bytes); const m = await import('/src/slideimport.ts'); m.importSlide(a.deck, src, src.slides[0].path, 0); a.setCurrent(1);""",
    # SmartArt made by PowerPoint (probe_powerpoint_smartart.py, kept in captures/, not committed): its texts edited in the Textbereich.
    "smartart text": f"""const r = await fetch('/@fs/{(HERE / "captures" / "smartart" / "smartart.pptx").as_posix()}'); p.openBytes('smartart.pptx', new Uint8Array(await r.arrayBuffer())); await new Promise(r => setTimeout(r, 800));
        const a = p.app; a.setCurrent(1); await new Promise(r => setTimeout(r, 300)); const v = a.deck.view(a.slidePath()).shapes.find(s => s.frame === 'diagram'); a.editor.select([v.id]); await a.commands.run('smartart.textPane'); await new Promise(r => setTimeout(r, 300));
        const inputs = [...document.querySelectorAll('[role=dialog] input[type=text]')]; inputs[0].value = 'Planen'; inputs[1].value = 'Bauen'; [...document.querySelectorAll('[role=dialog] button')].find(x => x.textContent === 'OK').click(); await new Promise(r => setTimeout(r, 400));""",
    "smartart convert": f"""const r = await fetch('/@fs/{(HERE / "captures" / "smartart" / "smartart.pptx").as_posix()}'); p.openBytes('smartart.pptx', new Uint8Array(await r.arrayBuffer())); await new Promise(r => setTimeout(r, 800));
        const a = p.app; a.setCurrent(1); await new Promise(r => setTimeout(r, 300)); let v = a.deck.view(a.slidePath()).shapes.find(s => s.frame === 'diagram'); a.editor.select([v.id]); await a.commands.run('smartart.toShapes');
        a.setCurrent(3); await new Promise(r => setTimeout(r, 300)); v = a.deck.view(a.slidePath()).shapes.find(s => s.frame === 'diagram'); a.editor.select([v.id]); await a.commands.run('smartart.toText'); await new Promise(r => setTimeout(r, 300));""",
    "picture crop": f"""p.newDoc(); const a = p.app; a.deck.setLayout([0], a.deck.layouts()[6].path); const bin = atob('{PNG}'); const b = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) b[i] = bin.charCodeAt(i); a.insertPictureBytes(b, 'image/png'); await new Promise(r => setTimeout(r, 500));
        const v = a.deck.view(a.slidePath()).shapes.find(s => s.kind === 'pic'); const S = await import('/src/shapeops.ts'); a.editor.edit('Größe', () => S.setXfrm(a.editor.shapeEl(v.id), {{x: 1000000, y: 1000000, w: 3810000, h: 3810000}})); a.editor.select([v.id]); await a['cropToRatio'](16 / 9); a.changeShape('ellipse', true);""",
    "picture effects removed": f"""p.newDoc(); const a = p.app; a.deck.setLayout([0], a.deck.layouts()[6].path); const bin = atob('{PNG}'); const b = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) b[i] = bin.charCodeAt(i); a.insertPictureBytes(b, 'image/png'); await new Promise(r => setTimeout(r, 500));
        const v = a.deck.view(a.slidePath()).shapes.find(s => s.kind === 'pic'); a.editor.select([v.id]); const m = await import('/src/picfx.ts');
        await m.setEffect(a, 'colorTemperature', {{name: 'colorTemperature', attrs: {{colorTemp: '4700'}}}}, 'Farbton'); await m.setEffect(a, 'colorTemperature', null, 'Farbton');""",
    "custom show": """p.newDoc(); const a = p.app; a.newSlide(); a.newSlide(); const m = await import('/src/custshow.ts'); m.writeShows(a, [{name: 'Kurz', slides: [2, 0]}]);""",
    "chart": """p.newDoc(); const a = p.app; a.deck.setLayout([0], a.deck.layouts()[6].path); const m = await import('/src/charts.ts'); m.insertChart(a, 'column');""",
    "chart area": """p.newDoc(); const a = p.app; a.deck.setLayout([0], a.deck.layouts()[6].path); const m = await import('/src/charts.ts'); m.insertChart(a, 'column');
        const ch = a.deck.view(a.slidePath()).shapes.find(s => s.frame === 'chart'); a.editor.select([ch.id]); const o = await import('/src/chartops.ts'); o.setChartArea(a, 'fill', 'ED7D31'); o.setChartArea(a, 'line', 'C00000');""",
    "comment": """p.newDoc(); p.app.comments.add('Bitte prüfen');""",
    "undo redo": """p.setupCase({layout:'Leer', shapes:[{kind:'rect',box:[100,100,200,100]}]}); const a = p.app; a.editor.select([2]); a.editor.deleteSelected(); await p.run('edit.undo'); await p.run('edit.redo'); await p.run('edit.undo');""",
}


def ppt_shapes(slide) -> list[dict]:
    out = []
    for i in range(1, slide.Shapes.Count + 1):
        s = slide.Shapes(i)
        d = {"name": str(s.Name), "type": int(s.Type), "box": [round(float(s.Left), 1), round(float(s.Top), 1), round(float(s.Width), 1), round(float(s.Height), 1)],
             "rot": round(float(s.Rotation), 1), "flipH": bool(s.HorizontalFlip)}
        try:
            if s.HasTextFrame and s.TextFrame.HasText:
                tr = s.TextFrame.TextRange
                d["text"] = str(tr.Text)
                d["paras"] = [{"text": str(tr.Paragraphs(k).Text).rstrip("\r"), "level": int(tr.Paragraphs(k).IndentLevel), "bold": int(tr.Paragraphs(k).Font.Bold),
                               "size": float(tr.Paragraphs(k).Font.Size), "align": int(tr.Paragraphs(k).ParagraphFormat.Alignment)} for k in range(1, tr.Paragraphs().Count + 1)]
                d["runs"] = [{"text": str(tr.Runs(k).Text), "bold": int(tr.Runs(k).Font.Bold), "color": int(tr.Runs(k).Font.Color.RGB)} for k in range(1, tr.Runs().Count + 1)]
        except Exception:  # noqa: BLE001
            pass
        try:
            if int(s.ThreeD.BevelTopType) > 1 or int(s.Glow.Radius) or int(s.Reflection.Type) or int(s.Shadow.Visible):
                d["fx"] = {"bevel": int(s.ThreeD.BevelTopType), "camera": int(s.ThreeD.PresetCamera), "shadow": int(s.Shadow.Visible), "glow": round(float(s.Glow.Radius)), "reflection": int(s.Reflection.Type)}
            if s.HasTextFrame and int(s.TextFrame2.WarpFormat) not in (-2, 36):
                d["warp"] = int(s.TextFrame2.WarpFormat)
                d["textBevel"] = int(s.TextFrame2.ThreeD.BevelTopType)
        except Exception:  # noqa: BLE001
            pass
        if int(s.Type) == 13:
            try:
                fx = s.Fill.PictureEffects
                d["fx"] = [int(fx(k).Type) for k in range(1, fx.Count + 1)]
            except Exception:  # noqa: BLE001
                d["fx"] = None
        if int(s.Type) == 6:
            d["members"] = int(s.GroupItems.Count)
        try:
            d["auto"] = int(s.AutoShapeType)
        except Exception:  # noqa: BLE001
            pass
        try:
            d["action"] = int(s.ActionSettings(1).Action)
        except Exception:  # noqa: BLE001
            pass
        if int(s.Type) == 1:
            try:
                d["filltype"] = int(s.Fill.Type)
                d["fillcolor"] = int(s.Fill.ForeColor.ObjectThemeColor)
            except Exception:  # noqa: BLE001
                pass
        try:
            if int(s.HasSmartArt):
                d["sa"] = [str(s.SmartArt.AllNodes(i).TextFrame2.TextRange.Text) for i in range(1, s.SmartArt.AllNodes.Count + 1)]
        except Exception:  # noqa: BLE001
            pass
        if bool(s.HasTable):
            d["table"] = [int(s.Table.Rows.Count), int(s.Table.Columns.Count)]
            t = s.Table
            # Merged cells: PowerPoint reports the merged area's width for each of its cells.
            d["tstyle"] = str(t.Style.Name)
            d["cellw"] = [[round(float(t.Cell(r, c).Shape.Width)) for c in range(1, t.Columns.Count + 1)] for r in range(1, t.Rows.Count + 1)]
            d["cellt"] = [[str(t.Cell(r, c).Shape.TextFrame.TextRange.Text) for c in range(1, t.Columns.Count + 1)] for r in range(1, t.Rows.Count + 1)]
            try:
                b = t.Cell(1, 1).Borders(4)  # ppBorderRight
                d["border11r"] = [round(float(b.Weight), 1), int(b.DashStyle), int(b.ForeColor.RGB)]
                d["fill11"] = int(t.Cell(1, 1).Shape.Fill.ForeColor.RGB)
            except Exception:  # noqa: BLE001
                pass
            try:
                d["bevel11"] = int(t.Cell(1, 1).Shape.ThreeD.BevelTopType)
            except Exception as e:  # noqa: BLE001
                d["bevel11"] = repr(e)[:80]
            try:
                d["tshadow"] = int(s.Shadow.Visible)
            except Exception as e:  # noqa: BLE001
                d["tshadow"] = repr(e)[:80]
            try:
                pass
            except Exception:  # noqa: BLE001
                pass
        if bool(s.HasChart):
            d["chart"] = int(s.Chart.ChartType)
            try:
                ca = s.Chart.ChartArea.Format
                d["chartarea"] = [int(ca.Fill.ForeColor.RGB), int(ca.Line.ForeColor.RGB)]
            except Exception:  # noqa: BLE001
                pass
        out.append(d)
    return out


def ppt_read(pres) -> dict:
    slides = []
    for i in range(1, pres.Slides.Count + 1):
        sl = pres.Slides(i)
        notes = ""
        try:
            for k in range(1, sl.NotesPage.Shapes.Placeholders.Count + 1):
                ph = sl.NotesPage.Shapes.Placeholders(k)
                if int(ph.PlaceholderFormat.Type) == 2:
                    notes = str(ph.TextFrame.TextRange.Text)
        except Exception:  # noqa: BLE001
            pass
        slides.append({"comments": int(sl.Comments.Count), "layout": str(sl.CustomLayout.Name), "hidden": bool(sl.SlideShowTransition.Hidden), "transition": int(sl.SlideShowTransition.EntryEffect),
                       "animations": int(sl.TimeLine.MainSequence.Count),
                       "interactive": [int(sl.TimeLine.InteractiveSequences(k).Count) for k in range(1, sl.TimeLine.InteractiveSequences.Count + 1)],
                       "motion": [str(sl.TimeLine.MainSequence(k).Behaviors(1).MotionEffect.Path) for k in range(1, sl.TimeLine.MainSequence.Count + 1) if int(sl.TimeLine.MainSequence(k).Behaviors(1).Type) == 1], "notes": notes, "shapes": ppt_shapes(sl)})
    custom = []
    try:
        ns = pres.SlideShowSettings.NamedSlideShows
        # SlideIDs: the show's slide ids (PowerPoint's array may carry an empty first element).
        custom = [[str(ns(k).Name), [int(x) for x in ns(k).SlideIDs if x]] for k in range(1, ns.Count + 1)]
    except Exception:  # noqa: BLE001
        pass
    sp = pres.SectionProperties
    return {"custom": custom, "slides": slides, "sections": [str(sp.Name(k)) for k in range(1, sp.Count + 1)], "design": str(pres.SlideMaster.Theme.ThemeColorScheme.Colors(5).RGB) if False else str(pres.SlideMaster.Design.Name),
            "accent1": int(pres.SlideMaster.Theme.ThemeColorScheme(5).RGB)}


def check(name: str, r: dict, ours: dict) -> list[str]:
    """What the scenario must show in PowerPoint (beyond opening without an error)."""
    s0 = r["slides"][0] if r["slides"] else {"shapes": []}
    sh = s0["shapes"]
    err: list[str] = []
    def want(cond, msg):
        if not cond:
            err.append(msg)
    if name == "title text":
        want([x.get("text") for x in sh] == ["Mein Titel", "Untertitel"], f"texts {[x.get('text') for x in sh]}")
    elif name == "bullets levels":
        paras = r["slides"][1]["shapes"][1].get("paras", [])
        want([(p["text"], p["level"]) for p in paras] == [("Eins", 1), ("Zwei", 2), ("Drei", 1)], f"paras {paras}")
    elif name == "shapes":
        want([x["box"] for x in sh][:2] == [[100.0, 100.0, 200.0, 100.0], [400.0, 100.0, 150.0, 150.0]], f"boxes {[x['box'] for x in sh]}")
        want(sh[0].get("text") == "Rechteck" and sh[2].get("text") == "Textfeld", "texts")
    elif name == "rotate flip":
        want(sh[0]["rot"] == 90.0 and sh[0]["flipH"], f"rot {sh[0]['rot']} flip {sh[0]['flipH']}")
    elif name == "bold part":
        want([(x["text"], x["bold"]) for x in sh[0].get("runs", [])] == [("normal ", 0), ("fett", -1)], f"runs {sh[0].get('runs')}")
    elif name == "size color":
        want(sh[0]["paras"][0]["size"] == 32.0 and all(x["color"] == 255 for x in sh[0]["runs"]), f"size/color {sh[0].get('paras')} {sh[0].get('runs')}")
    elif name == "align center":
        want(sh[0]["paras"][0]["align"] == 2, f"align {sh[0]['paras'][0]['align']}")
    elif name == "table":
        want(any(x.get("table") == [2, 3] for x in sh), f"table {[x.get('table') for x in sh]}")
    elif name == "table merge":
        t = next((x for x in sh if x.get("table")), {})
        w = t.get("cellw", [[0, 0, 0]])
        want(t.get("table") == [2, 3] and w[0][0] == w[0][1] == w[1][0] == w[1][1] and w[0][0] > w[0][2] and t["cellt"][0][0] == "eins", f"merge {t.get('cellw')} {t.get('cellt')}")
    elif name == "table split":
        t = next((x for x in sh if x.get("table")), {})
        w = t.get("cellw", [[0, 0, 0]])
        want(t.get("table") == [2, 3] and w[0][0] == w[0][1] and w[0][0] * 2 - 1 <= w[0][2] <= w[0][0] * 2 + 1 and w[1][0] == w[1][1] == w[0][2], f"split {t.get('cellw')}")
    elif name == "equation typed":
        texts = [x.get("text", "") for x in sh]
        want(any("=" in t and "+" in t for t in texts), f"equation {texts}")
    elif name == "equation":
        texts = [x.get("text", "") for x in sh]
        want(any("=" in t and "2" in t for t in texts), f"equation {texts}")
    elif name == "table style":
        t = next((x for x in sh if x.get("table")), {})
        want(t.get("tstyle") == "Helle Formatvorlage 2 - Akzent 2", f"style {t.get('tstyle')}")
    elif name == "motion path":
        want(s0["animations"] == 1 and len(s0.get("motion", [])) == 1, f"motion {s0.get('animations')} {s0.get('motion')}")
    elif name == "custom path":
        want(s0["animations"] == 1 and s0.get("motion", [""])[0].startswith("M 0 0 L 0.2 0"), f"motion {s0.get('animations')} {s0.get('motion')}")
    elif name == "freeform":
        want([x["type"] for x in sh] == [5, 5], f"types {[x['type'] for x in sh]}")
    elif name in ("file pptm", "file potx"):
        want(len(r["slides"]) == 1, f"slides {len(r['slides'])}")
    elif name == "file ppsx":
        want(len(r["slides"]) == 2, f"slides {len(r['slides'])}")
    elif name == "shape effects":
        f = [x.get("fx") for x in sh]
        # msoBevelCircle 3, msoCameraPerspectiveLeft 41 (any perspective camera is fine), glow 11 pt, a reflection; text: msoWarpFormat1 (Arch) 0, msoBevelConvex 8.
        ok = f[0] and f[0]["bevel"] == 3 and f[0]["camera"] != 1 and f[0]["shadow"] == -1 and f[1] and f[1]["glow"] == 11 and f[1]["reflection"] != 0 and sh[2].get("warp") is not None and sh[2].get("textBevel") == 8
        want(bool(ok), f"effects {f} warp {sh[2].get('warp')} {sh[2].get('textBevel')}")
    elif name == "summary zoom":
        s1 = r["slides"][1]["shapes"] if len(r["slides"]) > 1 else []
        want(len(r["slides"]) == 4 and [x["type"] for x in s1].count(-2) == 1 and r["sections"] == ["Standardabschnitt", "Zusammenfassungsabschnitt", "Teil 1", "Teil 2"],
             f"slides {len(r['slides'])} types {[x['type'] for x in s1]} sections {r['sections']}")
    elif name in ("zoom", "section zoom"):
        # PowerPoint reports its zooms as type -2 (a picture, 13, would mean it fell back to the Fallback).
        z = [x for x in sh if x["type"] == -2]
        want(len(z) == 1, f"types {[x['type'] for x in sh]}")
    elif name == "fills and background":
        # msoFillGradient 3; the second shape msoThemeColorAccent2 6 (with lumMod/lumOff); the master's background a style (bgRef).
        want(sh[0].get("filltype") == 3 and sh[1].get("filltype") == 1 and sh[1].get("fillcolor") == 6, f"{[(x.get('filltype'), x.get('fillcolor')) for x in sh]}")
    elif name == "table shading":
        t = next((x for x in sh if x.get("table")), {})
        want(t.get("fill11") == 192, f"cell fill {t.get('fill11')}")
    elif name == "table effects":
        t = next((x for x in sh if x.get("table")), {})
        # The table's shadow visible (msoTrue -1); a cell's ThreeD is not readable through COM (it raises),
        # so the cell bevel is checked by PowerPoint opening the file with it.
        want(t.get("tshadow") == -1, f"shadow {t.get('tshadow')}")
    elif name == "table pen":
        t = next((x for x in sh if x.get("table")), {})
        # 3 pt, msoLineDash 4 (sysDash) or another dash, red C00000 (BGR 0x0000C0 = 192).
        want(t.get("border11r", [0])[0] == 3.0 and t.get("border11r", [0, 0, 0])[2] == 192, f"border {t.get('border11r')}")
    elif name == "action button":
        # msoShapeActionButtonForwardorNext 130, ppActionNextSlide 1.
        want(any(x.get("action") == 1 and x.get("auto") == 130 for x in sh), f"{[(x.get('auto'), x.get('action')) for x in sh]}")
    elif name == "action":
        want(any(x.get("action") == 1 for x in sh), f"action {[x.get('action') for x in sh]}")
    elif name == "animation copy":
        want(s0["animations"] == 2, f"animations {s0['animations']}")
    elif name == "trigger":
        want(s0["animations"] == 0 and s0.get("interactive") == [1], f"trigger main {s0['animations']} interactive {s0.get('interactive')}")
    elif name == "group":
        want(len(sh) == 1 and sh[0].get("members") == 2, f"group {sh}")
    elif name == "notes":
        want(s0["notes"].replace("\r", "\n") == "Meine Notiz\nZweite Zeile", f"notes {s0['notes']!r}")
    elif name == "hidden slide":
        want([s["hidden"] for s in r["slides"]] == [False, True], f"hidden {[s['hidden'] for s in r['slides']]}")
    elif name == "sections":
        want(r["sections"] == ["Anfang", "Teil 2"], f"sections {r['sections']}")
    elif name == "move duplicate delete":
        want(len(r["slides"]) == 3, f"slides {len(r['slides'])}")
    elif name == "layout change":
        want(s0["layout"] == "Titel und Inhalt", f"layout {s0['layout']}")
    elif name == "transition":
        want(s0["transition"] != 0, f"transition {s0['transition']}")
    elif name == "animation":
        want(s0["animations"] == 1, f"animations {s0['animations']}")
    elif name == "design raabe":
        want(r["accent1"] == 0x1B14B3, f"accent1 {r['accent1']:06X}")
    elif name == "picture":
        want(any(x["type"] == 13 for x in sh), f"types {[x['type'] for x in sh]}")
    elif name == "picture effects":
        # MsoPictureEffectType: 2 Blur, 24 Saturation, 25 SharpenSoften (PowerPoint lists them in its own order).
        fx = next((x.get("fx") for x in sh if x["type"] == 13), None)
        want(fx is not None and sorted(fx) == [2, 24, 25], f"effects {fx}")
    elif name == "background removal":
        fx = next((x.get("fx") for x in sh if x["type"] == 13), None)
        want(fx == [1], f"effects {fx}")
    elif name == "reuse slides":
        s1 = r["slides"][1]["shapes"] if len(r["slides"]) > 1 else []
        want(len(r["slides"]) == 2 and any(x["type"] == 13 for x in s1), f"slides {len(r['slides'])} types {[x['type'] for x in s1]}")
    elif name == "smartart text":
        s1 = r["slides"][1]["shapes"] if len(r["slides"]) > 1 else []
        sa = next((x.get("sa") for x in s1 if x.get("sa")), None)
        want(sa is not None and sa[:3] == ["Planen", "Bauen", "Punkt 3"], f"smartart {sa}")
    elif name == "smartart convert":
        s1 = r["slides"][1]["shapes"] if len(r["slides"]) > 3 else []
        s3 = r["slides"][3]["shapes"] if len(r["slides"]) > 3 else []
        want([x.get("members") for x in s1] == [5] and len(s3) == 1 and s3[0].get("text", "").count("Punkt") >= 4, f"slide 2 {[(x['type'], x.get('members')) for x in s1]} slide 4 {[x.get('text') for x in s3]}")
    elif name == "picture crop":
        p = next((x for x in sh if x["type"] == 13), None)
        want(p is not None and abs(p["box"][2] / p["box"][3] - 16 / 9) < 0.02 and p.get("auto") == 9, f"picture {p}")
    elif name == "picture effects removed":
        fx = next((x.get("fx") for x in sh if x["type"] == 13), None)
        want(fx == [], f"effects {fx}")
    elif name == "custom show":
        c = r.get("custom") or []
        want(len(c) == 1 and c[0][0] == "Kurz" and len(c[0][1]) == 2, f"custom shows {c}")
    elif name == "chart area":
        # ED7D31 and C00000 as BGR.
        want(any(x.get("chartarea") == [0x317DED, 0x0000C0] for x in sh), f"chart area {[x.get('chartarea') for x in sh]}")
    elif name == "chart":
        want(any("chart" in x for x in sh), f"chart {[x.get('chart') for x in sh]}")
    elif name == "comment":
        want(s0["comments"] == 1, f"comments {s0['comments']}")
    elif name == "undo redo":
        want(len(sh) == 1, f"shapes {len(sh)}")
    # The shapes' places as RaabePresent has them.
    if ours and ours.get("shapes") is not None and name not in ("group",):
        for o, t in zip(ours["shapes"], sh):
            b = [round(o["x"] / 12700, 1), round(o["y"] / 12700, 1), round(o["w"] / 12700, 1), round(o["h"] / 12700, 1)]
            if t["rot"] in (90.0, 270.0):
                continue
            if any(abs(p - q) > 0.6 for p, q in zip(b, t["box"])) and t["type"] not in (14,):
                err.append(f"box ours {b} PowerPoint {t['box']}")
    return err


def main() -> None:
    try:
        ppt = PowerPointTarget(visible=False)
    except OwnerPowerPointRunning as e:
        raise SystemExit(f"skipped: {e}")
    ppt.app.DisplayAlerts = 1  # ppAlertsNone: a damaged file raises instead of asking
    page = PresentPage(1600, 1000, 1.0)
    out: dict = {"date": datetime.date.today().isoformat(), "fixtures": {}, "scenarios": {}}
    try:
        # 1. PowerPoint's files, written back by RaabePresent unchanged.
        for name in ([] if ONLY else FIXTURES):
            page.open_file(FIX / f"{name}.pptx")
            b = base64.b64decode(page.js("window.__present.bytes()"))
            f = TMP / f"{name}-rt.pptx"
            f.write_bytes(b)
            try:
                a = ppt_read(ppt.open_copy(FIX / f"{name}.pptx", with_window=False))
                c = ppt_read(ppt.open_copy(f, with_window=False))
                same = a == c
                out["fixtures"][name] = {"result": "same" if same else "different"}
                if not same:
                    out["fixtures"][name]["diff"] = [k for k in ("slides", "sections", "design") if a.get(k) != c.get(k)]
            except Exception as e:  # noqa: BLE001
                out["fixtures"][name] = {"result": "error", "error": repr(e)[:300]}
            print(name, out["fixtures"][name]["result"], flush=True)
        # 2. Files RaabePresent made.
        for name, js in SCENARIOS.items():
            if ONLY and name not in ONLY:
                continue
            if name in ("smartart text", "smartart convert") and not (HERE / "captures" / "smartart" / "smartart.pptx").exists():
                continue
            try:
                page.js(f"(async () => {{ const p = window.__present; {js} }})()")
                ours = page.js("window.__present.state()")
                b = base64.b64decode(page.js("window.__present.bytes()"))
                # File types (D-072): the scenario's name ends with the extension it is saved as.
                ext = name.split()[-1] if name.startswith("file ") else "pptx"
                f = TMP / f"scenario-{len(out['scenarios'])}.{ext}"
                f.write_bytes(b)
                r = ppt_read(ppt.open_copy(f, with_window=False))
                errs = check(name, r, ours)
                out["scenarios"][name] = {"result": "same" if not errs else "different", "diff": errs}
            except Exception as e:  # noqa: BLE001
                out["scenarios"][name] = {"result": "error", "error": repr(e)[:400]}
            print(name, out["scenarios"][name], flush=True)
    finally:
        page.close()
        print("PowerPoint:", ppt.close(kill_if_unreachable=True))
    fx = sum(1 for v in out["fixtures"].values() if v["result"] == "same")
    sc = sum(1 for v in out["scenarios"].values() if v["result"] == "same")
    out["summary"] = {"fixtures": f"{fx}/{len(out['fixtures'])}", "scenarios": f"{sc}/{len(out['scenarios'])}"}
    print(out["summary"])
    if ONLY:
        return  # a partial run (RT_ONLY) is not written as the results
    (HERE / "results" / f"{out['date']}-present-roundtrip.json").write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))


if __name__ == "__main__":
    main()
