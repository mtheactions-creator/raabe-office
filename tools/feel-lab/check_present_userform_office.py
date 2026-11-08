"""UserForms written by RaabePresent, opened by PowerPoint (D-100, D-106). RaabePresent (headless Edge on the dev
server) writes a presentation with two forms (UserForm1: a button, a text box, a label, a check box and a list;
UserForm2: a Frame holding an option button and a text box) and a module that loads each form, reads its
controls' properties into the first slide's title and unloads it. PowerPoint (our own, only when no POWERPNT.EXE
runs; its window on the virtual display, D-073) opens it, runs the module through COM (no security setting
changed) and saves it again; our reader then reads PowerPoint's forms back.

    .venv/Scripts/python.exe check_present_userform_office.py      (npm run dev:present must be running)

Result: results/<date>-present-userform-office.json.
"""
from __future__ import annotations

import base64
import datetime
import json
from pathlib import Path

from check_present_macro_suite import Watchdog
from feellab.present_page import PresentPage

HERE = Path(__file__).parent
CAP = HERE / "captures" / "userform"
CODE = ("Sub Pruefe()\r\n"
        "    Dim r As String\r\n"
        "    Load UserForm1\r\n"
        "    r = UserForm1.Controls.Count & \"|\" & UserForm1.CommandButton1.Caption & \"|\" & UserForm1.Label1.Caption & \"|\" & UserForm1.TextBox1.Left & \"|\" & UserForm1.CheckBox1.Caption\r\n"
        "    Unload UserForm1\r\n"
        "    ActivePresentation.Slides(1).Shapes(1).TextFrame.TextRange.Text = r\r\n"
        "End Sub\r\n\r\n"
        "Sub PruefeRahmen()\r\n"
        "    Load UserForm2\r\n"
        "    ActivePresentation.Slides(1).Shapes(1).TextFrame.TextRange.Text = UserForm2.Controls.Count & \"|\" & UserForm2.Frame1.Controls.Count & \"|\" & UserForm2.OptionButton1.Caption\r\n"
        "    Unload UserForm2\r\n"
        "End Sub\r\n")

BUILD = r"""(async () => { const P = window.__present; P.newDoc(); const a = P.app; const m = await import('/src/macros.ts'); const v = await import('/@fs/""" + "{VBA}" + r"""');
  const f1 = v.newForm('UserForm1'); f1.caption = 'Anmeldung';
  const b = v.newControl('CommandButton', 'CommandButton1', 120, 90); b.caption = 'Los';
  const l = v.newControl('Label', 'Label1', 12, 12); l.caption = 'Name:';
  const t = v.newControl('TextBox', 'TextBox1', 60, 12);
  const c = v.newControl('CheckBox', 'CheckBox1', 12, 40); c.caption = 'Gelesen';
  const li = v.newControl('ListBox', 'ListBox1', 12, 66);
  f1.controls.push(b, l, t, c, li);
  const f2 = v.newForm('UserForm2');
  const fr = v.newControl('Frame', 'Frame1', 12, 12); fr.caption = 'Auswahl'; fr.width = 150; fr.height = 80;
  const o = v.newControl('OptionButton', 'OptionButton1', 6, 6); o.caption = 'Ja';
  fr.controls = [o, v.newControl('TextBox', 'TextBox2', 6, 30)];
  f2.controls.push(fr);
  m.setModules(a, [{ name: 'Modul1', kind: 'standard', code: CODE }, { name: 'UserForm1', kind: 'class', code: '', form: f1 }, { name: 'UserForm2', kind: 'class', code: '', form: f2 }]);
  a.deck.pptx.setType(a.deck.pptx.presPath, 'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml');
  return P.bytes(); })()"""


def main() -> None:
    from feellab import winapi
    from feellab.powerpoint import PowerPointTarget
    from check_present_screens import display, monitors
    CAP.mkdir(parents=True, exist_ok=True)
    out: dict = {"date": datetime.date.today().isoformat(), "checks": {}}
    vba = (HERE.parents[1] / "packages/vba/src/index.ts").resolve().as_posix()
    page = PresentPage()
    try:
        b64 = page.js(BUILD.replace("{VBA}", vba).replace("CODE", json.dumps(CODE)))
        ours = CAP / "userforms.pptm"
        ours.write_bytes(base64.b64decode(b64))
        before = monitors()
        mons = display("extend", 2)
        virtual = [m for m in mons if m not in before] or mons[1:]
        t = PowerPointTarget(visible=True, place=False)
        dog = Watchdog(t.pid)
        saved = CAP / "userforms-powerpoint.pptm"
        try:
            pres = t.open_copy(ours, with_window=True)
            if virtual and t.hwnd:
                winapi.place(t.hwnd, *virtual[0])
            out["opened"] = True
            for name in ("Pruefe", "PruefeRahmen"):
                n0 = len(dog.seen)
                try:
                    t.app.Run(f"{pres.Name}!Modul1.{name}")
                except Exception as e:  # noqa: BLE001
                    out[f"{name}_com_error"] = str(e)[:200]
                out[name] = str(pres.Slides(1).Shapes(1).TextFrame.TextRange.Text)
                if len(dog.seen) > n0:
                    out[f"{name}_vba_error"] = dog.seen[n0:]
                pres.Slides(1).Shapes(1).TextFrame.TextRange.Text = ""
            if saved.exists():
                saved.unlink()
            pres.SaveCopyAs(str(saved), 25)
            pres.Saved = -1
        finally:
            dog.stop = True
            out["close"] = t.close(kill_if_unreachable=True)
            display("internal", 1)
        # The same macros in RaabePresent, on the file we wrote.
        page.open_file(ours)
        out["ours"] = page.js("""(async () => { const a = window.__present.app; const m = await import('/src/macros.ts'); const r = [];
            for (const n of ['Pruefe', 'PruefeRahmen']) { const e = await m.runMacroQuiet(a, n); r.push(e.error ? 'Fehler ' + e.error.number + ': ' + e.error.message : a.deck.view(a.deck.slides[0].path).shapes[0].text.paras.map((p) => p.runs.map((x) => x.text).join('')).join('')); }
            return r; })()""")
        if saved.exists():
            import zipfile
            with zipfile.ZipFile(saved) as z:
                data = base64.b64encode(z.read(next(n for n in z.namelist() if n.endswith("vbaProject.bin")))).decode()
            out["read_back"] = page.js(f"""(async () => {{ const v = await import('/@fs/{vba}');
                const b = Uint8Array.from(atob({json.dumps(data)}), (c) => c.charCodeAt(0));
                return v.readProject(b).filter((m) => m.form).map((m) => ({{ name: m.name, caption: m.form.caption,
                  controls: v.allControls(m.form.controls).map((c) => c.type + ':' + c.name + ':' + (c.caption ?? '')) }})); }})()""")
    finally:
        page.close()
    out["checks"] = {
        "powerpoint_opens": bool(out.get("opened")),
        "form_loads_and_reads": out.get("Pruefe") == "5|Los|Name:|60|Gelesen",
        # A form's Controls holds the controls inside its frames too (PowerPoint: 3).
        "frame_form_loads": out.get("PruefeRahmen") == "3|2|Ja",
        "ours_same": out.get("ours") == [out.get("Pruefe"), out.get("PruefeRahmen")],
        "powerpoint_keeps_forms": any(f["name"] == "UserForm1" and "CommandButton:CommandButton1:Los" in f["controls"] for f in out.get("read_back") or []),
        "frame_read_back": any(f["name"] == "UserForm2" and "OptionButton:OptionButton1:Ja" in f["controls"] for f in out.get("read_back") or []),
    }
    out["summary"] = f"{sum(out['checks'].values())}/{len(out['checks'])}"
    p = HERE / "results" / f"{out['date']}-present-userform-office.json"
    p.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))
    print(json.dumps({k: out.get(k) for k in ("Pruefe", "PruefeRahmen", "ours", "Pruefe_vba_error", "PruefeRahmen_vba_error", "checks", "summary")}, ensure_ascii=False))


if __name__ == "__main__":
    main()
