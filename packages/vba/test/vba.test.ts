// @raabe/vba (D-079): the project format and the interpreter.
import { describe, expect, it } from "vitest";
import { compress, decompress, readCfb, readProject, visibleCode, writeCfb, writeProject, VbaRuntime, type VbaHost, type VbaObject, type VbaValue } from "../src";

const enc = new TextEncoder();

describe("VBA project format", () => {
  it("compression round-trips (repeats, long text, empty)", () => {
    for (const t of ["", "abc", "Sub Test()\r\n  MsgBox \"Hallo\"\r\nEnd Sub\r\n".repeat(200), "x".repeat(9000)]) {
      const b = enc.encode(t);
      expect(Array.from(decompress(compress(b)))).toEqual(Array.from(b));
    }
  });

  it("compresses repeated text with copy tokens", () => {
    const b = enc.encode("abcabcabcabcabcabcabcabc".repeat(50));
    expect(compress(b).length).toBeLessThan(b.length / 3);
  });

  it("compound file round-trips small and large streams in storages", () => {
    const big = new Uint8Array(10000).map((_, i) => i % 251);
    const files = writeCfb(new Map([["PROJECT", enc.encode("ID=\"x\"")], ["VBA/dir", enc.encode("dir")], ["VBA/Modul1", big]]));
    const back = readCfb(files);
    expect(new TextDecoder().decode(back.get("PROJECT"))).toBe("ID=\"x\"");
    expect(Array.from(back.get("VBA/Modul1")!)).toEqual(Array.from(big));
    expect(back.get("VBA/dir")!.length).toBe(3);
  });

  it("writes a project and reads its modules back", () => {
    const bin = writeProject([{ name: "Modul1", kind: "standard", code: "Sub Hallo()\r\n    MsgBox \"Grüße\"\r\nEnd Sub" }, { name: "Quiz", kind: "standard", code: "Public Punkte As Integer" }]);
    const mods = readProject(bin);
    expect(mods.map((m) => m.name)).toEqual(["Modul1", "Quiz"]);
    expect(visibleCode(mods[0].code)).toContain('MsgBox "Grüße"');
    expect(mods[0].code).toMatch(/^Attribute VB_Name = "Modul1"/);
  });
});

/** A small object model for the tests: a slide list with shapes. */
function model(log: string[]) {
  const shape = (name: string): VbaObject & { text: string; left: number } => {
    const o = { kind: "object" as const, text: "", left: 10,
      get: (n: string): VbaValue => (n.toLowerCase() === "name" ? name : n.toLowerCase() === "left" ? o.left : n.toLowerCase() === "textframe" ? tf : undefined),
      set: (n: string, v: VbaValue) => { if (n.toLowerCase() === "left") o.left = Number(v); },
    };
    const tr: VbaObject = { kind: "object", get: (n) => (n.toLowerCase() === "text" ? o.text : undefined), set: (n, v) => { if (n.toLowerCase() === "text") o.text = String(v); } };
    const tf: VbaObject = { kind: "object", get: (n) => (n.toLowerCase() === "textrange" ? tr : undefined) };
    return o;
  };
  const shapes = [shape("Frage"), shape("Antwort")];
  const shapesColl: VbaObject = { kind: "object", get: (n, a) => (n.toLowerCase() === "count" ? shapes.length : n.toLowerCase() === "item" ? (typeof a[0] === "number" ? shapes[a[0] - 1] : shapes.find((s) => s.get("Name", []) === a[0])) : undefined), items: () => shapes, default: (a) => (typeof a[0] === "number" ? shapes[a[0] - 1] : shapes.find((s) => s.get("Name", []) === a[0])) };
  const slide: VbaObject = { kind: "object", get: (n) => (n.toLowerCase() === "shapes" ? shapesColl : undefined) };
  const slides: VbaObject = { kind: "object", get: (n) => (n.toLowerCase() === "count" ? 1 : undefined), default: () => slide, items: () => [slide] };
  const pres: VbaObject = { kind: "object", get: (n) => (n.toLowerCase() === "slides" ? slides : undefined) };
  const host: VbaHost = {
    global: (n) => (n.toLowerCase() === "activepresentation" ? pres : undefined),
    msgBox: async (p) => { log.push(`MsgBox: ${p}`); return 1; },
    inputBox: async () => "42",
  };
  return { host, shapes };
}

const run = async (code: string, name = "Test") => {
  const log: string[] = [];
  const m = model(log);
  const rt = new VbaRuntime(m.host, [code]);
  await rt.run(name);
  return { log, shapes: m.shapes, rt };
};

describe("VBA interpreter", () => {
  it("runs loops, conditions and string functions", async () => {
    const { log } = await run(`
Sub Test()
    Dim i As Integer, s As String
    For i = 1 To 5
        If i Mod 2 = 0 Then s = s & "g" Else s = s & "u"
    Next i
    MsgBox s & " " & UCase(Left("hallo", 2)) & " " & Len(s)
End Sub`);
    expect(log).toEqual(["MsgBox: ugugu HA 5"]);
  });

  it("calls functions with ByRef and ByVal parameters", async () => {
    const { log } = await run(`
Function Doppelt(ByVal x As Integer) As Integer
    Doppelt = x * 2
End Function
Sub Erhoehe(n As Integer)
    n = n + 1
End Sub
Sub Test()
    Dim a As Integer
    a = 3
    Erhoehe a
    Call Erhoehe(a)
    MsgBox Doppelt(a)
End Sub`);
    expect(log).toEqual(["MsgBox: 10"]);
  });

  it("handles Select Case, Do loops, arrays and module variables", async () => {
    const { log } = await run(`
Dim Punkte As Integer
Sub Test()
    Dim noten(1 To 3) As Integer, k As Integer
    noten(1) = 1: noten(2) = 4: noten(3) = 6
    Do While k < 3
        k = k + 1
        Select Case noten(k)
            Case 1 To 2: Punkte = Punkte + 10
            Case Is >= 5: Punkte = Punkte - 1
            Case Else: Punkte = Punkte + 1
        End Select
    Loop
    MsgBox Punkte & "/" & UBound(noten)
End Sub`);
    expect(log).toEqual(["MsgBox: 10/3"]);
  });

  it("works on the app's objects (With, For Each, properties)", async () => {
    const { shapes } = await run(`
Sub Test()
    Dim s
    For Each s In ActivePresentation.Slides(1).Shapes
        s.Left = s.Left + 5
    Next
    With ActivePresentation.Slides(1).Shapes("Antwort").TextFrame.TextRange
        .Text = "Richtig!"
    End With
End Sub`);
    expect(shapes.map((s) => s.left)).toEqual([15, 15]);
    expect(shapes[1].text).toBe("Richtig!");
  });

  it("uses InputBox, IIf, Format and On Error Resume Next", async () => {
    const { log } = await run(`
Sub Test()
    On Error Resume Next
    Dim x
    x = 1 / 0
    MsgBox IIf(Val(InputBox("Zahl?")) = 42, "ok", "nein") & " " & Format(0.5, "0%")
End Sub`);
    expect(log).toEqual(["MsgBox: ok 50%"]);
  });

  it("lists the macros a user can run and stops endless loops", async () => {
    const rt = new VbaRuntime(model([]).host, [`Sub A()\nEnd Sub\nPrivate Sub B()\nEnd Sub\nSub C(x)\nEnd Sub\nSub Endlos()\nDo\nLoop\nEnd Sub`], 10000);
    expect(rt.macros()).toEqual(["A", "Endlos"]);
    await expect(rt.run("Endlos")).rejects.toThrow(/zu lange/);
  });

  it("refuses to reach outside the app", async () => {
    await expect(run(`Sub Test()\nShell "cmd.exe"\nEnd Sub`)).rejects.toThrow(/nicht erlaubt/);
    expect(() => new VbaRuntime(model([]).host, [`Declare Function X Lib "kernel32" ()`])).toThrow(/Declare/);
  });
});
