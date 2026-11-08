// @raabe/vba: class modules, events, error handling, the debugger, file access through the app, UserForms.
import { describe, expect, it } from "vitest";
import { VbaRuntime, newControl, newForm, type VbaBreak, type VbaForm, type VbaHost, type VbaModule, type VbaStepMode, type VbaValue } from "../src";

function host(log: string[], extra: Partial<VbaHost> = {}): VbaHost {
  return {
    global: () => undefined,
    msgBox: async (p) => { log.push(p); return 1; },
    inputBox: async () => "",
    debugPrint: (t) => log.push(`> ${t}`),
    ...extra,
  };
}
const std = (code: string, name = "Modul1"): VbaModule => ({ name, kind: "standard", code });
const cls = (name: string, code: string): VbaModule => ({ name, kind: "class", code });
async function run(mods: VbaModule[], macro = "Test", extra: Partial<VbaHost> = {}) {
  const log: string[] = [];
  const rt = new VbaRuntime(host(log, extra), mods);
  await rt.run(macro);
  return { log, rt };
}

describe("class modules", () => {
  it("creates objects with New, member variables, properties, methods and Me", async () => {
    const konto = cls("Konto", `
Private mStand As Currency
Public Inhaber As String
Public Property Get Stand() As Currency
    Stand = mStand
End Property
Public Property Let Stand(ByVal v As Currency)
    If v < 0 Then Err.Raise vbObjectError + 1, "Konto", "Kein negativer Stand"
    mStand = v
End Property
Public Sub Einzahlen(betrag As Currency)
    Me.Stand = Me.Stand + betrag
End Sub
Public Function Beschreibung() As String
    Beschreibung = Inhaber & ": " & Stand
End Function`);
    const { log } = await run([konto, std(`
Sub Test()
    Dim k As Konto
    Set k = New Konto
    k.Inhaber = "Ada"
    k.Einzahlen 10
    k.Einzahlen 2.5
    MsgBox k.Beschreibung()
    Dim b As New Konto
    b.Stand = 3
    MsgBox TypeName(b) & " " & b.Stand & " " & (TypeOf b Is Konto) & " " & (b Is k) & " " & (Nothing Is Nothing)
    On Error Resume Next
    b.Stand = -1
    MsgBox Err.Number - vbObjectError & " " & Err.Description & " " & b.Stand
End Sub`)]);
    expect(log).toEqual(["Ada: 12,5", "Konto 3 Wahr Falsch Wahr", "1 Kein negativer Stand 3"]);
  });

  it("runs Class_Initialize and Class_Terminate (Set Nothing, end of procedure)", async () => {
    const c = cls("Zaehler", `
Public Name As String
Private Sub Class_Initialize()
    Debug.Print "init"
End Sub
Private Sub Class_Terminate()
    Debug.Print "ende " & Name
End Sub`);
    const { log } = await run([c, std(`
Sub Lokal()
    Dim z As New Zaehler
    z.Name = "lokal"
End Sub
Sub Test()
    Dim a As Zaehler, b As Zaehler
    Set a = New Zaehler
    a.Name = "a"
    Set b = a
    Set a = Nothing
    Debug.Print "noch da"
    Set b = Nothing
    Lokal
    Debug.Print "fertig"
End Sub`)]);
    expect(log).toEqual(["> init", "> noch da", "> ende a", "> init", "> ende lokal", "> fertig"]);
  });

  it("has Collection with keys, Before/After, Remove and For Each, and Scripting.Dictionary", async () => {
    const { log } = await run([std(`
Sub Test()
    Dim c As New Collection, x, s As String
    c.Add "b", "kb"
    c.Add "a", Key:="ka", Before:=1
    c.Add "c", After:="kb"
    For Each x In c: s = s & x: Next
    s = s & c.Count & c("KB") & c.Item(3)
    c.Remove "ka"
    s = s & c.Count & c(1)
    On Error Resume Next
    c.Add "d", "kb"
    s = s & " " & Err.Number
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("x") = 1
    d.Add "y", 2
    d("x") = d("x") + 10
    s = s & " " & d.Count & d("x") & d.Exists("y") & Join(d.Keys, "+")
    MsgBox s
End Sub`)]);
    expect(log).toEqual(["abc3bc2b 457 211Wahrx+y"]);
  });

  it("supports default members, Implements, predeclared classes and user-defined types", async () => {
    const liste = cls("Liste", `
Private c As New Collection
Public Sub Add(v)
    c.Add v
End Sub
Public Property Get Item(i)
Attribute Item.VB_UserMemId = 0
    Item = c(i)
End Property
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = c.[_NewEnum]
End Property`);
    const tier = cls("ITier", "Public Function Laut() As String\nEnd Function");
    const hund = cls("Hund", `Implements ITier
Private Function ITier_Laut() As String
    ITier_Laut = "Wuff"
End Function`);
    const cfg = cls("Konfig", `Attribute VB_PredeclaredId = True
Public Wert As Integer`);
    const { log } = await run([liste, tier, hund, cfg, std(`
Type Punkt
    X As Double
    Y As Double
End Type
Sub Test()
    Dim l As New Liste, v, s As String
    l.Add "x": l.Add "y"
    For Each v In l: s = s & v: Next
    Dim t As ITier
    Set t = New Hund
    Konfig.Wert = 7
    Dim p As Punkt, q As Punkt
    p.X = 1: q = p: q.X = 2
    MsgBox l(2) & s & t.Laut() & (TypeOf t Is ITier) & Konfig.Wert & p.X & q.X
End Sub`)]);
    expect(log).toEqual(["yxyWuffWahr712"]);
  });
});

describe("events", () => {
  it("raises class events to WithEvents variables", async () => {
    const timer = cls("Wecker", `
Public Event Klingeln(ByVal n As Integer, Abbrechen As Boolean)
Public Sub Start()
    Dim i As Integer, stopp As Boolean
    For i = 1 To 5
        RaiseEvent Klingeln(i, stopp)
        If stopp Then Exit For
    Next
    Debug.Print "Ende bei " & i
End Sub`);
    const user = cls("Schlaefer", `
Private WithEvents w As Wecker
Public Sub Los()
    Set w = New Wecker
    w.Start
End Sub
Private Sub w_Klingeln(ByVal n As Integer, Abbrechen As Boolean)
    Debug.Print "ring " & n
    If n = 3 Then Abbrechen = True
End Sub`);
    const { log } = await run([timer, user, std(`
Sub Test()
    Dim s As New Schlaefer
    s.Los
End Sub`)]);
    expect(log).toEqual(["> ring 1", "> ring 2", "> ring 3", "> Ende bei 3"]);
  });

  it("fires event procedures by name and the app's own events, writing ByRef arguments back", async () => {
    const log: string[] = [];
    const app = { kind: "object" as const, get: () => undefined };
    const rt = new VbaRuntime(host(log, { global: (n) => (n.toLowerCase() === "application" ? app : undefined) }), [
      { name: "DieseArbeitsmappe", kind: "document", code: `
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Cancel = True
    Debug.Print "nicht schließen"
End Sub` },
      cls("AppEvents", `
Public WithEvents App As Application
Private Sub App_NewPresentation(ByVal Pres)
    Debug.Print "neu: " & Pres
End Sub`),
      std(`
Dim ev As AppEvents
Sub Init()
    Set ev = New AppEvents
    Set ev.App = Application
End Sub`),
    ]);
    const args: VbaValue[] = [false];
    expect(await rt.fire("DieseArbeitsmappe", "Workbook_BeforeClose", args)).toBe(true);
    expect(args[0]).toBe(true);
    expect(await rt.fire("DieseArbeitsmappe", "Workbook_Open")).toBe(false);
    await rt.run("Init");
    expect(await rt.raise(app, "NewPresentation", ["P1"])).toBe(1);
    expect(log).toEqual(["> nicht schließen", "> neu: P1"]);
  });
});

describe("errors", () => {
  it("handles On Error GoTo, Resume Next, Resume and Err", async () => {
    const { log } = await run([std(`
Function Teile(a, b)
    On Error GoTo Fehler
    Teile = a / b
    Exit Function
Fehler:
    Debug.Print "Fehler " & Err.Number & ": " & Err.Description & " in Zeile " & Erl
    Teile = 0
    Resume Next
End Function
Sub Test()
    Dim versuch As Integer
    Debug.Print Teile(6, 3); Teile(1, 0)
    On Error GoTo Nochmal
    If versuch < 3 Then Err.Raise 1000, , "eigener Fehler"
    Debug.Print "Versuche:"; versuch
    Exit Sub
Nochmal:
    versuch = versuch + 1
    Resume
End Sub`)]);
    expect(log).toEqual(["> Fehler 11: Division durch Null in Zeile 4", ">  2  0 ", "> Versuche: 3 "]);
  });

  it("reports the module, procedure and line of an unhandled error", async () => {
    const rt = new VbaRuntime(host([]), [std("Sub A()\n    B\nEnd Sub", "Haupt"), std("Sub B()\n    Dim x\n    x = 1 / 0\nEnd Sub", "Hilfe")]);
    const e = await rt.run("A").catch((x) => x);
    expect([e.number, e.line, e.module, e.proc]).toEqual([11, 3, "Hilfe", "B"]);
  });

  it("End stops everything and resets module variables", async () => {
    const { log, rt } = await run([std(`
Dim n As Integer
Sub Test()
    n = n + 1
    Debug.Print "n=" & n
    Weiter
    Debug.Print "nie"
End Sub
Sub Weiter()
    End
End Sub`)]);
    await rt.run("Test");
    expect(log).toEqual(["> n=1", "> n=1"]);
  });
});

describe("debugger", () => {
  const code = `Sub Test()
    Dim i As Integer, s As Integer
    For i = 1 To 3
        s = s + Zwei(i)
    Next
    Debug.Print s
End Sub
Function Zwei(x As Integer) As Integer
    Zwei = x * 2
End Function`;

  it("pauses at breakpoints with the call stack and local variables, and continues", async () => {
    const seen: string[] = [];
    let rt!: VbaRuntime;
    const log: string[] = [];
    rt = new VbaRuntime(host(log, {
      debugBreak: (b: VbaBreak) => {
        const loc = rt.locals();
        seen.push(`${b.reason} ${b.proc}:${b.line} ${rt.callStack().map((f) => f.proc).join("<")} ${loc.map((v) => `${v.name}=${v.value}`).join(",")}`);
        return Promise.resolve("continue" as VbaStepMode);
      },
    }), [std(code)]);
    rt.setBreakpoint("Modul1", 9);
    expect(rt.statementLines("Modul1")).toEqual([3, 4, 6, 9]);
    await rt.run("Test");
    expect(seen).toEqual(["breakpoint Zwei:9 Zwei<Test x=1,Zwei=0", "breakpoint Zwei:9 Zwei<Test x=2,Zwei=0", "breakpoint Zwei:9 Zwei<Test x=3,Zwei=0"]);
    expect(log).toEqual([">  12 "]);
  });

  it("steps into, over and out, and evaluates in the paused procedure", async () => {
    const lines: number[] = [];
    const modes: VbaStepMode[] = ["into", "into", "into", "into", "out", "over", "over", "stop"];
    let rt!: VbaRuntime;
    let probe: VbaValue;
    rt = new VbaRuntime(host([], {
      debugBreak: async (b) => {
        lines.push(b.line);
        if (b.line === 9 && probe === undefined) probe = await rt.evaluate("x * 10 + 1");
        return modes.shift() ?? "continue";
      },
    }), [std(code)]);
    rt.setBreakpoint("Modul1", 3);
    await rt.run("Test");
    // The For line, into the body, into Zwei, into the next pass of the body, into Zwei, out of it (the third
    // pass), over the call to the Debug.Print line.
    expect(lines).toEqual([3, 4, 9, 4, 9, 4, 6]);
    expect(probe).toBe(11);
  });

  it("pauses at Stop and Debug.Assert False; without a debugger they do nothing", async () => {
    const reasons: string[] = [];
    const src = std("Sub Test()\n    Stop\n    Debug.Assert 1 = 2\n    Debug.Assert True\n    Debug.Print \"ok\"; 1; \"-\";\n    Debug.Print \"x\", \"y\"\nEnd Sub");
    const log: string[] = [];
    await new VbaRuntime(host(log, { debugBreak: (b) => { reasons.push(`${b.reason}:${b.line}`); return Promise.resolve("continue" as VbaStepMode); } }), [src]).run("Test");
    expect(reasons).toEqual(["stop:2", "assert:3"]);
    expect(log).toEqual(["> ok 1 -x       y"]);
    const quiet: string[] = [];
    await new VbaRuntime(host(quiet), [src]).run("Test");
    expect(quiet).toEqual(["> ok 1 -x       y"]);
  });

  it("pauses on request and resumes through the runtime", async () => {
    let rt!: VbaRuntime;
    const at: number[] = [];
    rt = new VbaRuntime(host([], { debugBreak: (b) => { at.push(b.line); setTimeout(() => rt.resume("continue"), 0); } }), [std("Sub Test()\n    Dim i As Long\n    Do While i < 100\n        i = i + 1\n    Loop\nEnd Sub")]);
    rt.pause();
    await rt.run("Test");
    expect(at).toEqual([3]);
  });
});

describe("file access through the app", () => {
  function memFs() {
    const files = new Map<string, Uint8Array>([["C:\\Daten\\noten.csv", new TextEncoder().encode("Anna,2\r\nBen,1\r\n")]]);
    const asked: string[] = [];
    return {
      files, asked,
      extra: (allow: boolean): Partial<VbaHost> => ({
        fileAccess: async (kind, path) => { asked.push(`${kind} ${path}`); return allow; },
        files: {
          read: async (p) => { const b = files.get(p); if (!b) throw new Error("missing"); return b; },
          write: async (p, d) => { files.set(p, d); },
          list: async (folder) => [...files.keys()].filter((k) => k.startsWith(folder + "\\")).map((k) => ({ name: k.slice(folder.length + 1), dir: false })),
          stat: async (p) => (files.has(p) ? { size: files.get(p)!.length, dir: false } : null),
          remove: async (p) => { if (!files.delete(p)) throw new Error("missing"); },
        },
      }),
    };
  }
  const code = std(`
Sub Test()
    Dim f As Integer, zeile As String, name As String, note As Integer, n As Integer
    f = FreeFile
    Open "C:\\Daten\\noten.csv" For Input As #f
    Do Until EOF(f)
        Input #f, name, note
        n = n + note
    Loop
    Close #f
    Open "C:\\Daten\\schnitt.txt" For Output As #1
    Print #1, "Schnitt:"; n / 2
    Write #1, "Anna", 2.5, True
    Close #1
    Open "C:\\Daten\\schnitt.txt" For Input As #2
    Line Input #2, zeile
    Close
    Debug.Print zeile; "|"; Dir("C:\\Daten\\*.txt"); "|"; FileLen("C:\\Daten\\schnitt.txt")
    Kill "C:\\Daten\\schnitt.txt"
End Sub`);

  it("reads and writes files only after the app asked", async () => {
    const m = memFs();
    const { log } = await run([code], "Test", m.extra(true));
    expect(log).toEqual(["> Schnitt: 1,5 |schnitt.txt| 34 "]);
    expect(m.asked).toEqual(["open C:\\Daten\\noten.csv", "save C:\\Daten\\schnitt.txt", "open C:\\Daten\\schnitt.txt", "list C:\\Daten\\*.txt", "list C:\\Daten\\schnitt.txt", "delete C:\\Daten\\schnitt.txt"]);
    expect(m.files.has("C:\\Daten\\schnitt.txt")).toBe(false);
  });

  it("refuses with Err 70 when the user says no or the app offers no files", async () => {
    const m = memFs();
    const e = await run([code], "Test", m.extra(false)).catch((x) => x);
    expect(e.number).toBe(70);
    const e2 = await run([code]).catch((x) => x);
    expect(e2.number).toBe(70);
    const { log } = await run([std(`Sub Test()\n    On Error Resume Next\n    Open "x.txt" For Output As #1\n    MsgBox Err.Number & " " & (Err.Number = 70)\nEnd Sub`)]);
    expect(log).toEqual(["70 Wahr"]);
  });

  it("keeps Shell, CreateObject, Environ and Declare blocked; settings go to the app's store", async () => {
    for (const s of ['Shell "cmd"', 'Set o = CreateObject("WScript.Shell")', 'x = Environ("USERNAME")', 'Set o = GetObject(, "Excel.Application")']) {
      const e = await run([std(`Sub Test()\n    ${s}\nEnd Sub`)]).catch((x) => x);
      expect(e.message).toMatch(/nicht erlaubt/);
    }
    expect(() => new VbaRuntime(host([]), [std('Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long')])).toThrow(/Declare/);
    const store = new Map<string, string>();
    const { log } = await run([std(`Sub Test()\n    SaveSetting "Quiz", "Stand", "Punkte", 12\n    MsgBox GetSetting("Quiz", "Stand", "Punkte", "0") + 1 & GetSetting("Quiz", "Stand", "Fehlt", "leer")\nEnd Sub`)], "Test", {
      settings: { get: (a, s, k) => store.get(`${a}/${s}/${k}`), set: (a, s, k, v) => { store.set(`${a}/${s}/${k}`, v); }, delete: () => undefined },
    });
    expect(log).toEqual(["13leer"]);
    expect(store.get("Quiz/Stand/Punkte")).toBe("12");
  });
});

describe("UserForms", () => {
  function quizForm(): VbaModule {
    const design = newForm("frmQuiz");
    design.caption = "Quiz";
    design.controls.push(newControl("Label", "lblFrage", 6, 6), newControl("TextBox", "txtAntwort", 6, 30), newControl("CommandButton", "cmdOk", 6, 60), newControl("ListBox", "lstAuswahl", 90, 6), newControl("CheckBox", "chkHilfe", 90, 90));
    return { name: "frmQuiz", kind: "class", form: design, code: `
Public Ergebnis As String
Private Sub UserForm_Initialize()
    lblFrage.Caption = "Hauptstadt von Frankreich?"
    lstAuswahl.AddItem "Paris"
    lstAuswahl.AddItem "Lyon"
    Me.Caption = "Frage 1"
End Sub
Private Sub cmdOk_Click()
    Ergebnis = txtAntwort.Text & "/" & lstAuswahl.Value & "/" & chkHilfe.Value
    Me.Hide
End Sub
Private Sub txtAntwort_Change()
    Debug.Print "Change: " & txtAntwort.Text
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu And txtAntwort.Text = "" Then Cancel = True
End Sub
Private Sub UserForm_Terminate()
    Debug.Print "Terminate"
End Sub` };
  }

  it("shows a modal form; the app's input runs the events and Hide returns from Show", async () => {
    const log: string[] = [];
    let rt!: VbaRuntime;
    const shown: string[] = [];
    rt = new VbaRuntime(host(log, {
      showForm: (form: VbaForm, modal: boolean) => {
        shown.push(`${form.design.caption} ${modal} ${form.design.controls[0].caption} ${form.design.controls[3].list?.map((r) => r[0]).join(",")}`);
        // The user: types, picks Paris, ticks the box, tries the X (cancelled: empty answer?), clicks OK.
        setTimeout(async () => {
          expect(await rt.closeForm(form)).toBe(false);
          await rt.setFormValue(form, "txtAntwort", "Text", "Paris");
          await rt.setFormValue(form, "lstAuswahl", "ListIndex", 0);
          await rt.setFormValue(form, "chkHilfe", "Value", true);
          await rt.fire(form, "cmdOk_Click");
        }, 0);
      },
      hideForm: (form) => shown.push(`hide ${form.name}`),
    }), [quizForm(), std(`
Sub Test()
    frmQuiz.txtAntwort.Text = ""
    frmQuiz.Show
    MsgBox frmQuiz.Ergebnis & " " & frmQuiz.Visible
    Unload frmQuiz
    MsgBox frmQuiz.lblFrage.Caption & " " & frmQuiz.Ergebnis & "."
End Sub`)]);
    await rt.run("Test");
    expect(shown).toEqual(["Frage 1 true Hauptstadt von Frankreich? Paris,Lyon", "hide frmQuiz"]);
    expect(log).toEqual(["> Change: Paris", "Paris/Paris/Wahr Falsch", "> Terminate", "Hauptstadt von Frankreich? ."]);
  });

  it("gives Controls, control properties, fonts and option groups to VBA", async () => {
    const design = newForm("UserForm1");
    const fr = newControl("Frame", "Frame1", 6, 6);
    fr.controls!.push(newControl("OptionButton", "optA", 6, 6), newControl("OptionButton", "optB", 6, 30));
    design.controls.push(fr, newControl("SpinButton", "spn", 160, 6), newControl("ComboBox", "cbo", 160, 40));
    const changes: string[] = [];
    const { log } = await run([{ name: "UserForm1", kind: "class", form: design, code: `
Private Sub optB_Click()
    Debug.Print "B gewählt"
End Sub` }, std(`
Sub Test()
    Dim c, s As String
    Load UserForm1
    For Each c In UserForm1.Controls
        s = s & TypeName(c) & ":" & c.Name & " "
    Next
    With UserForm1
        .optA.Value = True
        .optB.Value = True
        .spn.Max = 10
        .spn.Value = 50
        .cbo.AddItem "rot"
        .cbo.AddItem "grün"
        .cbo.ListIndex = 1
        .cbo.Font.Bold = True
        .Controls("Frame1").Caption = "Wahl"
        MsgBox s & .optA.Value & .optB & " " & .spn.Value & " " & .cbo.Text & .cbo.ListCount & .cbo.Font.Bold & .Frame1.Caption & .Controls.Count
    End With
End Sub`)], "Test", { showForm: () => undefined });
    void changes;
    expect(log).toEqual(["> B gewählt", "Frame:Frame1 OptionButton:optA OptionButton:optB SpinButton:spn ComboBox:cbo FalschWahr 10 grün2WahrWahl5"]);
  });

  it("refuses Show when the app has no forms", async () => {
    const e = await run([quizForm(), std("Sub Test()\n    frmQuiz.Show\nEnd Sub")]).catch((x) => x);
    expect(e.message).toMatch(/UserForms/);
  });
});

describe("language details", () => {
  it("keeps typed variables, Static, ParamArray, named arguments, Enum and Option Base", async () => {
    const { log } = await run([std(`
Option Base 1
Enum Farbe
    Rot = 1
    Gruen
    Blau = 10
End Enum
Function Summe(ParamArray w())
    Dim x, t
    For Each x In w: t = t + x: Next
    Summe = t
End Function
Function Gruss(Optional ByVal Name As String = "Welt", Optional Zeichen As String = "!") As String
    Gruss = "Hallo " & Name & Zeichen
End Function
Function Zaehle() As Integer
    Static n As Integer
    n = n + 1
    Zaehle = n
End Function
Function Fak(n As Integer) As Long
    If n <= 1 Then Fak = 1 Else Fak = n * Fak(n - 1)
End Function
Sub Test()
    Dim i As Integer, s As String, a(3) As Integer, b
    i = 2.5: s = 12
    a(1) = 7.6
    b = Array(1, 2)
    Call Zaehle: Call Zaehle
    MsgBox i & " " & TypeName(s) & " " & a(1) & LBound(a) & LBound(b) & " " & Summe(1, 2, 3) & " " & Gruss(Zeichen:="?") & " " & Zaehle() & " " & Farbe.Gruen & Blau & " " & Fak(5)
End Sub`)]);
    expect(log).toEqual(["2 String 811 6 Hallo Welt? 3 210 120"]);
  });

  it("resolves conditional compilation as 64-bit Office does", async () => {
    const { log } = await run([std(`
#If VBA7 And Win64 Then
Const Bits = 64
#Else
Const Bits = 32
#End If
Sub Test()
    MsgBox Bits
End Sub`)]);
    expect(log).toEqual(["64"]);
  });
});

describe("Rnd and Randomize as VBA", () => {
  it("draws VBA's numbers after Rnd -1 and Randomize 7 (PowerPoint: 0 3 0 2 1)", async () => {
    const out: string[] = [];
    const rt = new VbaRuntime({ global: () => undefined, msgBox: async (p) => { out.push(p); return 1; }, inputBox: async () => "" },
      ["Sub T()\r\n  Dim i As Integer, r As String\r\n  Rnd -1\r\n  Randomize 7\r\n  For i = 1 To 5\r\n    r = r & Int(Rnd * 4)\r\n  Next i\r\n  MsgBox r\r\nEnd Sub"]);
    await rt.run("T");
    expect(out).toEqual(["03021"]);
  });
});

describe("named arguments to the app's objects", () => {
  it("passes them by name (and positionally when the object names its parameters)", async () => {
    const seen: unknown[] = [];
    const sheet = { kind: "object", get: () => undefined, call: (n: string, a: VbaValue[], named?: Record<string, VbaValue>) => { seen.push([n, a, named]); return undefined; } } as const;
    const sorter = { kind: "object", get: () => undefined, argNames: () => ["Key1", "Order1", "Header"], call: (n: string, a: VbaValue[], named?: Record<string, VbaValue>) => { seen.push([n, a, named]); return undefined; } } as const;
    const host: VbaHost = { global: (g) => (g.toLowerCase() === "blatt" ? sheet : g.toLowerCase() === "sortierer" ? sorter : undefined) as never, msgBox: async () => 1, inputBox: async () => "" };
    const rt = new VbaRuntime(host, ["Sub T()\r\n  Blatt.Add 1, After:=2, Count:=3\r\n  Sortierer.Sort Header:=1, Key1:=\"A1\"\r\nEnd Sub"]);
    await rt.run("T");
    expect(seen[0]).toEqual(["Add", [1], { after: 2, count: 3 }]);
    expect(seen[1]).toEqual(["Sort", ["A1", undefined, 1], { header: 1, key1: "A1" }]);
  });
});
