"""Macros (Step 8b, D-079): a presentation with a VBA project written by RaabePresent (our compound file, our
compression, our dir stream) is opened by PowerPoint (invisible, only when no POWERPNT.EXE runs) and saved again as
.pptm (PowerPoint compiles the source and writes its own project); our reader then reads PowerPoint's project. The
module and its code must come back. PowerPoint runs nothing: presentations have no macros that start by themselves,
and no security setting is changed.

    .venv/Scripts/python.exe check_present_macros.py      (npm run dev:present must be running)

Result: results/<date>-present-macros.json.
"""
from __future__ import annotations

import base64
import datetime
import json
import tempfile
import zipfile
from pathlib import Path

from feellab.powerpoint import OwnerPowerPointRunning, PowerPointTarget
from feellab.present_page import PresentPage

HERE = Path(__file__).parent
CODE = 'Sub Quiz()\r\n    MsgBox "Grüße aus RaabePresent"\r\nEnd Sub\r\n\r\nFunction Doppelt(x)\r\n    Doppelt = 2 * x\r\nEnd Function'


def main() -> None:
    out: dict = {"date": datetime.date.today().isoformat()}
    tmp = Path(tempfile.gettempdir())
    ours = tmp / "raabe-macro-check.pptm"
    page = PresentPage()
    try:
        b64 = page.js(f"""(async () => {{ const P = window.__present; P.newDoc(); const a = P.app; const m = await import('/src/macros.ts');
            m.setModules(a, [{{ name: 'Modul1', kind: 'standard', code: {json.dumps(CODE)} }}]);
            a.deck.pptx.setType(a.deck.pptx.presPath, 'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml');
            return P.bytes(); }})()""")
        ours.write_bytes(base64.b64decode(b64))
        try:
            t = PowerPointTarget(visible=False)
        except OwnerPowerPointRunning as e:
            out["powerpoint"] = {"skipped": str(e)}
            t = None
        if t:
            saved = tmp / "raabe-macro-check-saved.pptm"
            try:
                pres = t.open_copy(ours, with_window=False)
                out["opened"] = True
                if saved.exists():
                    saved.unlink()
                pres.SaveCopyAs(str(saved), 25)  # ppSaveAsOpenXMLPresentationMacroEnabled
            except Exception as e:  # noqa: BLE001
                out["error"] = str(e)[:300]
            finally:
                out["close"] = t.close(kill_if_unreachable=True)
            if saved.exists():
                with zipfile.ZipFile(saved) as z:
                    names = z.namelist()
                    bin_name = next((n for n in names if n.endswith("vbaProject.bin")), None)
                    out["powerpoint_kept_project"] = bool(bin_name)
                    if bin_name:
                        data = base64.b64encode(z.read(bin_name)).decode()
                        out["read_back"] = page.js(f"""(async () => {{ const v = await import('/@fs/{(HERE.parents[1] / 'packages/vba/src/index.ts').resolve().as_posix()}');
                            const b = Uint8Array.from(atob({json.dumps(data)}), (c) => c.charCodeAt(0));
                            return v.readProject(b).map((m) => ({{ name: m.name, kind: m.kind, code: v.visibleCode(m.code) }})); }})()""")
    finally:
        page.close()
    mods = out.get("read_back") or []
    m1 = next((m for m in mods if m["name"] == "Modul1"), None)
    out["checks"] = {
        "powerpoint_opens": bool(out.get("opened")),
        "project_kept": bool(out.get("powerpoint_kept_project")),
        "module_back": bool(m1),
        "code_back": bool(m1) and "Grüße aus RaabePresent" in m1["code"] and "Function Doppelt" in m1["code"],
    }
    out["summary"] = f"{sum(out['checks'].values())}/{len(out['checks'])}"
    p = HERE / "results" / f"{out['date']}-present-macros.json"
    p.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(json.dumps({k: out[k] for k in ("checks", "summary")}, ensure_ascii=False))


if __name__ == "__main__":
    main()
