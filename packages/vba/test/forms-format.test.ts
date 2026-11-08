// @raabe/vba: UserForms in the VBA project (MS-OFORMS designer storages, MS-OVBA designer modules).
import { describe, expect, it } from "vitest";
import { VbaRuntime, decompress, newControl, newForm, readCfb, readProject, writeProject, type VbaControlDesign, type VbaFormDesign } from "../src";
import { readCfbStorages } from "../src/cfb";
import { CLSID_FORM, CLSID_FRAME, designSnapshot, readFormStorage, writeFormStorage } from "../src/oforms";

const latin1 = (b: Uint8Array) => Array.from(b, (c) => String.fromCharCode(c)).join("");
const hex = (s: string) => Uint8Array.from(s.replace(/\s+/g, "").match(/../g)!.map((x) => parseInt(x, 16)));
/** A design without what the reader keeps (for comparing). */
const plain = (d: VbaFormDesign | VbaControlDesign) => JSON.parse(designSnapshot(d));

function sampleForm(): VbaFormDesign {
  const d = newForm("frmTest");
  d.caption = "Prüfung (Übung)";
  d.width = 300; d.height = 200.25;
  d.backColor = 0xffffe0;
  d.font = { name: "Arial", size: 10, bold: true, italic: false, underline: false, strikethrough: false };
  const btn = newControl("CommandButton", "cmdOk", 210, 170); btn.caption = "OK"; btn.accelerator = "O"; btn.controlTipText = "Übernehmen";
  const lbl = newControl("Label", "lblFrage", 12, 12); lbl.caption = "Wie heißt die Hauptstadt? €"; lbl.width = 200; lbl.foreColor = 0xff; lbl.textAlign = 2; lbl.font = { name: "Tahoma", size: 12, bold: false, italic: true, underline: false, strikethrough: false };
  const txt = newControl("TextBox", "txtAntwort", 12, 36); txt.value = "Paris"; txt.multiLine = true; txt.maxLength = 40; txt.tag = "a1";
  const pwd = newControl("TextBox", "txtCode", 12, 60); pwd.passwordChar = "*"; pwd.locked = true;
  const chk = newControl("CheckBox", "chkHilfe", 12, 84); chk.caption = "Hilfe"; chk.value = true;
  const lst = newControl("ListBox", "lstStadt", 160, 36); lst.multiSelect = 1; lst.columnCount = 2;
  const cbo = newControl("ComboBox", "cboLand", 160, 120); cbo.style = 2; cbo.value = "";
  const tgl = newControl("ToggleButton", "tglModus", 12, 108); tgl.caption = "Modus"; tgl.enabled = false;
  const spn = newControl("SpinButton", "spnZahl", 100, 84); spn.min = 1; spn.max = 10; spn.value = 3;
  const scr = newControl("ScrollBar", "scrWert", 260, 12); scr.max = 50; scr.largeChange = 5; scr.orientation = 0; scr.visible = false;
  const img = newControl("Image", "imgLogo", 240, 100); img.borderStyle = 0;
  const fra = newControl("Frame", "fraWahl", 12, 132); fra.caption = "Wahl"; fra.width = 180; fra.height = 40;
  const o1 = newControl("OptionButton", "optA", 6, 6); o1.caption = "A"; o1.groupName = "g"; o1.value = true;
  const o2 = newControl("OptionButton", "optB", 90, 6); o2.caption = "B"; o2.groupName = "g";
  fra.controls!.push(o1, o2);
  d.controls.push(btn, lbl, txt, pwd, chk, lst, cbo, tgl, spn, scr, img, fra);
  d.controls.forEach((c, i) => { c.tabIndex = i; });
  fra.controls!.forEach((c, i) => { c.tabIndex = i; });
  return d;
}

describe("UserForm storage (MS-OFORMS)", () => {
  it("writes a form and reads the same design back", () => {
    const d = sampleForm();
    const { streams, clsids } = writeFormStorage(d);
    expect([...streams.keys()].sort()).toEqual(["\u0001CompObj", "\u0003VBFrame", "f", "i12/\u0001CompObj", "i12/f", "i12/o", "o"].sort());
    expect(clsids.get("")).toBe(CLSID_FORM);
    expect(clsids.get("i12")).toBe(CLSID_FRAME);
    const back = readFormStorage("frmTest", streams, clsids);
    const want = plain(d), got = plain(back);
    // Runtime-only fields and fields we do not set are compared through the reader's view of both.
    expect(got.caption).toBe(want.caption);
    expect([got.width, got.height, got.backColor]).toEqual([300, 200.25, 0xffffe0]);
    expect(got.font).toEqual(want.font);
    const byName = (l: VbaControlDesign[]) => new Map(l.map((c) => [c.name, c]));
    const g = byName(back.controls), w = byName(d.controls);
    for (const [n, c] of w) {
      const r = g.get(n)!;
      expect(r.type, n).toBe(c.type);
      for (const k of ["left", "top", "width", "height", "caption", "value", "visible", "enabled", "locked", "foreColor", "backColor", "multiLine", "maxLength", "passwordChar", "multiSelect", "columnCount", "style", "min", "max", "largeChange", "orientation", "tag", "controlTipText", "accelerator", "groupName", "tabIndex", "textAlign", "borderStyle"] as const) {
        if (c[k] !== undefined) expect(r[k], `${n}.${k}`).toEqual(c[k]);
      }
      if (c.font && c.type !== "Image" && c.type !== "SpinButton" && c.type !== "ScrollBar") expect(r.font, `${n}.font`).toEqual(c.font);
    }
    const fr = g.get("fraWahl")!;
    expect(fr.controls!.map((c) => [c.name, c.caption, c.groupName, c.value, c.left])).toEqual([["optA", "A", "g", true, 6], ["optB", "B", "g", false, 90]]);
  });

  it("follows the format: versions, masks, defaults left out, strings compressed, CompObj and VBFrame", () => {
    const d = newForm("UserForm1");
    const b = newControl("CommandButton", "CommandButton1", 48, 60); b.width = 126; b.height = 36;
    d.controls.push(b);
    const { streams } = writeFormStorage(d);
    const f = streams.get("f")!, o = streams.get("o")!;
    expect([f[0], f[1]]).toEqual([0x00, 0x04]);
    const fv = new DataView(f.buffer, f.byteOffset);
    // NextAvailableID, DisplayedSize, LogicalSize, Caption, Font and DrawBuffer are stored.
    expect(fv.getUint32(4, true)).toBe((1 << 3) | (1 << 10) | (1 << 11) | (1 << 19) | (1 << 20) | (1 << 27));
    // The button: version 0.2, Caption and Size (and nothing at its default), the caption one byte per character.
    expect(Array.from(o.subarray(0, 8))).toEqual([0x00, 0x02, 0x20, 0x00, 0x28, 0x00, 0x00, 0x00]);
    expect(Array.from(o.subarray(8, 12))).toEqual([14, 0, 0, 0x80]);
    expect(latin1(o.subarray(12, 26))).toBe("CommandButton1");
    const ov = new DataView(o.buffer, o.byteOffset);
    expect([ov.getInt32(28, true), ov.getInt32(32, true)]).toEqual([4445, 1270]); // 126 × 36 points in HIMETRIC
    // TextProps follow (version 0.2, the font name).
    expect([o[36], o[37]]).toEqual([0x00, 0x02]);
    const co = streams.get("\u0001CompObj")!;
    expect(Array.from(co.subarray(0, 12))).toEqual([0x01, 0x00, 0xfe, 0xff, 0x03, 0x0a, 0x00, 0x00, 0xff, 0xff, 0xff, 0xff]);
    const lp = (t: string) => String.fromCharCode(t.length + 1, 0, 0, 0) + t + String.fromCharCode(0);
    expect(latin1(co.subarray(28))).toBe(lp("Microsoft Forms 2.0 Form") + lp("Embedded Object") + lp("Forms.Form.1") + String.fromCharCode(0xf4, 0x39, 0xb2, 0x71, ...new Array(12).fill(0)));
    const vbf = latin1(streams.get("\u0003VBFrame")!);
    expect(vbf).toMatch(/^VERSION 5\.00\r\nBegin \{C62A69F0-16DC-11CE-9E98-00AA00574A4F\} UserForm1 \r\n {3}Caption {9}= {3}"UserForm1"\r\n/);
    expect(vbf).toContain("ClientWidth     =   4710\r\n");
    expect(vbf).toContain("ClientHeight    =   3210\r\n");
    expect(vbf.endsWith("End\r\n")).toBe(true);
  });

  it("reads the specification's UserForm example (a RefEdit control kept as its bytes)", () => {
    // The "f" stream of [MS-OFORMS] 3.4, assembled from its tables: Verdana, 166.5 × 132 points, one RefEdit site.
    const f = hex(`00042400 080C100C 01000000 FFFF0000 01000000 007D0000 F2160000 31120000 9C130100 CE890000
      0352E30B918FCE119DE300AA004BB851 01 0000 00 9001 44420100 07 56657264616E61
      0100 00003800 4B000000 00000000 12450200000000 00C000000000000046 121DD28B42ECCE119E0D00AA006002F3 18450200000000 00C000000000000046
      01000000 2C000000 0001 9D9D
      00002400 E5010000 08000080 01000000 38000000 00000080 5265664564697431 45080000 9D060000`);
    expect(f.length).toBe(0xbc);
    const o = new Uint8Array(0x38).map((_, i) => i);
    const streams = new Map([["f", f], ["o", o]]);
    const d = readFormStorage("UserForm1", streams);
    expect([d.width, d.height]).toEqual([166.5, 132]);
    expect(d.font).toEqual({ name: "Verdana", size: 8.25, bold: false, italic: false, underline: false, strikethrough: false });
    expect(d.controls).toEqual([]);
    // Not edited: written back as it was read.
    const same = writeFormStorage(d);
    expect(Array.from(same.streams.get("f")!)).toEqual(Array.from(f));
    // Edited: the RefEdit's site, its class table entry and its data stay; the new button gets the next ID.
    d.caption = "Neu";
    const b = newControl("CommandButton", "cmdNeu", 6, 6);
    d.controls.push(b);
    const edited = writeFormStorage(d);
    const again = readFormStorage("UserForm1", edited.streams);
    expect(again.caption).toBe("Neu");
    expect(again.controls.map((c) => c.name)).toEqual(["cmdNeu"]);
    expect((again.controls[0].raw as { id: number }).id).toBe(2);
    const ef = edited.streams.get("f")!;
    expect(latin1(ef)).toContain("RefEdit1");
    expect(latin1(ef)).toContain(latin1(hex("12450200000000 00C000000000000046")));
    expect(Array.from(edited.streams.get("o")!.subarray(0, 0x38))).toEqual(Array.from(o));
  });
});

describe("UserForms in the VBA project (MS-OVBA)", () => {
  const code = `Private Sub cmdOk_Click()\r\n    Me.Hide\r\nEnd Sub\r\n`;

  it("writes the designer module, its storage, BaseClass=, Package= and the MSForms reference", () => {
    const bin = writeProject([{ name: "Modul1", kind: "standard", code: "Sub Zeigen()\r\n    frmTest.Show\r\nEnd Sub" }, { name: "frmTest", kind: "class", code, form: sampleForm() }]);
    const files = readCfb(bin);
    expect(files.has("frmTest/f")).toBe(true);
    expect(files.has("frmTest/\u0003VBFrame")).toBe(true);
    expect(readCfbStorages(bin).get("frmTest")).toBe(CLSID_FORM);
    const project = latin1(files.get("PROJECT")!);
    expect(project).toMatch(/\r\nModule=Modul1\r\nPackage=\{AC9F2F90-E877-11CE-9F68-00AA00574A4F\}\r\nBaseClass=frmTest\r\n/);
    const src = latin1(decompress(files.get("VBA/frmTest")!));
    expect(src).toMatch(/^Attribute VB_Name = "frmTest"\r\nAttribute VB_Base = "0\{[0-9A-F-]{36}\}\{[0-9A-F-]{36}\}"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = False\r\nAttribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = False\r\nPrivate Sub cmdOk_Click/);
    const dir = latin1(decompress(files.get("VBA/dir")!));
    expect(dir).toContain("*\\G{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0#C:\\Windows\\system32\\FM20.DLL#Microsoft Forms 2.0 Object Library");
    expect(dir).toContain("*\\G{00000000-0000-0000-0000-000000000000}#0.0#0##");
    expect(dir).not.toMatch(/Users/);
  });

  it("reads the project back with its form, and keeps an unedited form byte for byte", () => {
    const bin = writeProject([{ name: "frmTest", kind: "class", code, form: sampleForm() }]);
    const mods = readProject(bin);
    expect(mods[0].kind).toBe("class");
    expect(mods[0].form?.controls.map((c) => c.name)).toEqual(["cmdOk", "lblFrage", "txtAntwort", "txtCode", "chkHilfe", "lstStadt", "cboLand", "tglModus", "spnZahl", "scrWert", "imgLogo", "fraWahl"]);
    const bin2 = writeProject(mods);
    const a = readCfb(bin), b = readCfb(bin2);
    for (const k of [...a.keys()].filter((k) => k.startsWith("frmTest/"))) expect(Array.from(b.get(k)!), k).toEqual(Array.from(a.get(k)!));
    // The module keeps its VB_Base (the form's class and type library ids).
    const base = (x: Map<string, Uint8Array>) => /VB_Base = "([^"]+)"/.exec(latin1(decompress(x.get("VBA/frmTest")!)))![1];
    expect(base(b)).toBe(base(a));
  });

  it("writes an edited form again and the runtime can use what was read", async () => {
    const mods = readProject(writeProject([{ name: "frmTest", kind: "class", code, form: sampleForm() }]));
    const form = mods[0].form!;
    form.controls.find((c) => c.name === "lblFrage")!.caption = "Neue Frage";
    form.controls.push(newControl("CommandButton", "cmdAbbrechen", 120, 170));
    const back = readProject(writeProject(mods))[0].form!;
    expect(back.controls.find((c) => c.name === "lblFrage")!.caption).toBe("Neue Frage");
    expect(back.controls.map((c) => c.name)).toContain("cmdAbbrechen");
    const ids = new Set<number>();
    const walk = (l: VbaControlDesign[]) => { for (const c of l) { ids.add((c.raw as { id: number }).id); if (c.controls) walk(c.controls); } };
    walk(back.controls);
    expect(ids.size).toBe(15);
    // The runtime takes the modules as read: the form's controls are there for VBA.
    const said: string[] = [];
    const rt = new VbaRuntime({ global: () => undefined, msgBox: async (p) => { said.push(p); return 1; }, inputBox: async () => "" }, [...readProject(writeProject(mods)), { name: "Modul1", kind: "standard", code: "Sub Test()\r\n    MsgBox frmTest.lblFrage.Caption & \"|\" & frmTest.fraWahl.Controls(\"optA\").Value & frmTest.optB.Value & \"|\" & frmTest.txtAntwort.Text\r\nEnd Sub" }]);
    await rt.run("Test");
    expect(said).toEqual(["Neue Frage|WahrFalsch|Paris"]);
  });
});
