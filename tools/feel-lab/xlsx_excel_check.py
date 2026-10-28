"""Checks RaabeCalc's .xlsx reader and writer with Excel itself (D-024, D-061).

1. Runs the round-trip tests (`packages/ooxml/test/xlsx*.test.ts`), which read every Excel-built
   fixture `packages/ooxml/test/fixtures/xlsx-*.xlsx`, write it back (and, for the edit cases,
   change it first) into a temporary folder, and dump what the model read.
2. Opens the original and our file in an invisible Excel of our own. A file Excel has to repair
   fails to open (Excel refuses it while alerts are off; checked with broken files in --selftest)
   and counts as a failure.
3. Compares what Excel sees in both: every cell's Formula, Value2, NumberFormat, font, fill,
   borders, alignment and protection, column widths, row heights, merged cells, sheets (names,
   visibility, tab colours, standard sizes), defined names, windows (frozen panes, zoom,
   gridlines, selection), charts, notes, hyperlinks, conditional formats, data validation,
   AutoFilter, tables and pictures.
4. Compares the model as read with Excel's view of the original ("read check").

    .venv/Scripts/python.exe xlsx_excel_check.py              # all fixtures
    .venv/Scripts/python.exe xlsx_excel_check.py --selftest   # broken files must fail

Result: results/<date>-xlsx-excel-check.json (file names only, no paths).
"""

from __future__ import annotations

import datetime
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

import pythoncom

from feellab.excel import ExcelTarget

ROOT = Path(__file__).resolve().parents[2]
FIXTURES = ROOT / "packages" / "ooxml" / "test" / "fixtures"
RESULTS = Path(__file__).parent / "results"

UNDERLINE = {None: -4142, "single": 2, "double": -4119, "singleAccounting": 4, "doubleAccounting": 5}
H_ALIGN = {None: 1, "general": 1, "left": -4131, "center": -4108, "right": -4152, "fill": 5, "justify": -4130,
           "centerContinuous": 7, "distributed": -4117}
V_ALIGN = {None: -4107, "bottom": -4107, "top": -4160, "center": -4108, "justify": -4130, "distributed": -4117}
BORDER = {"thin": (1, 2), "medium": (1, -4138), "thick": (1, 4), "dashed": (-4115, 2), "dotted": (-4118, 2),
          "double": (-4119, 4), "hair": (1, 1), "dashDot": (4, 2), "dashDotDot": (5, 2), "mediumDashed": (-4115, -4138),
          "mediumDashDot": (4, -4138), "mediumDashDotDot": (5, -4138), "slantDashDot": (13, -4138)}
SIDES = {"left": 7, "top": 8, "bottom": 9, "right": 10, "diagonal_down": 5, "diagonal_up": 6}


def english(obj, name: str):
    """A property in Excel's English form (pywin32 passes the user's locale otherwise)."""
    o = obj._oleobj_
    return o.Invoke(o.GetIDsOfNames(name), 1033, pythoncom.DISPATCH_PROPERTYGET, 1)


def bgr(hex_: str | None) -> int:
    h = (hex_ or "#000000").lstrip("#")
    return int(h[0:2], 16) + (int(h[2:4], 16) << 8) + (int(h[4:6], 16) << 16)


def safe(fn, default=None):
    try:
        return fn()
    except Exception:  # noqa: BLE001
        return default


def cell_props(r) -> dict:
    f = r.Font
    i = r.Interior
    out = {
        "Formula": r.Formula, "Value2": r.Value2, "NumberFormat": english(r, "NumberFormat"), "Text": r.Text,
        "Font": (f.Name, f.Size, f.Bold, f.Italic, f.Underline, f.Strikethrough, f.Superscript, f.Subscript, f.Color),
        "Interior": (i.Pattern, i.Color if i.Pattern != -4142 else None, i.PatternColor if i.Pattern not in (1, -4142) else None),
        "Align": (r.HorizontalAlignment, r.VerticalAlignment, r.WrapText, r.IndentLevel, r.Orientation, r.ShrinkToFit),
        "Protection": (r.Locked, r.FormulaHidden), "Merge": r.MergeArea.GetAddress(False, False) if r.MergeCells else None,
        "Style": r.Style.Name, "HasArray": r.HasArray, "Array": r.CurrentArray.GetAddress(False, False) if r.HasArray else None,
    }
    for name, idx in SIDES.items():
        b = r.Borders(idx)
        out["Border " + name] = (b.LineStyle, b.Weight, b.Color) if b.LineStyle != -4142 else None
    return out


DV_PROPS = ("Type", "Operator", "Formula1", "Formula2", "AlertStyle", "IgnoreBlank", "InCellDropdown", "ShowInput", "ShowError",
            "InputTitle", "InputMessage", "ErrorTitle", "ErrorMessage")


def validation_props(ws) -> dict | None:
    """Datenüberprüfung: the cells that have one and, per cell (up to 400), its settings."""
    try:
        rng = ws.Cells.SpecialCells(-4174)
    except Exception:  # noqa: BLE001 - no cell has a validation
        return None
    out: dict = {"cells": rng.GetAddress(False, False)}
    n = 0
    for area in rng.Areas:
        for c in area.Cells:
            n += 1
            if n > 400:
                return out
            v = c.Validation
            out[c.GetAddress(False, False)] = [safe(lambda p=p: getattr(v, p)) for p in DV_PROPS]
    return out


def sheet_props(wb, ws, cells: list[str]) -> dict:
    used = ws.UsedRange
    out: dict = {
        "Name": ws.Name, "Visible": ws.Visible, "Tab": ws.Tab.Color, "StandardWidth": ws.StandardWidth,
        "StandardHeight": ws.StandardHeight, "UsedRange": used.GetAddress(False, False),
    }
    rows = used.Row + used.Rows.Count
    cols = used.Column + used.Columns.Count
    for c in cells:
        rr = ws.Range(c)
        rows = max(rows, rr.Row + 1)
        cols = max(cols, rr.Column + 1)
    cols = max(cols, 12)
    out["cols"] = [(ws.Columns(c).ColumnWidth, ws.Columns(c).Hidden, ws.Columns(c).OutlineLevel) for c in range(1, cols + 1)]
    out["rows"] = [(ws.Rows(r).RowHeight, ws.Rows(r).Hidden, ws.Rows(r).OutlineLevel) for r in range(1, rows + 1)]
    out["cells"] = {c: cell_props(ws.Range(c)) for c in cells}
    out["charts"] = [(co.Name, co.Chart.ChartType, co.Left, co.Top, co.Width, co.Height,
                      [s.Formula for s in co.Chart.SeriesCollection()], co.Chart.HasTitle and co.Chart.ChartTitle.Text)
                     for co in ws.ChartObjects()]
    out["shapes"] = [(s.Name, s.Type, s.Left, s.Top, s.Width, s.Height) for s in ws.Shapes]
    out["notes"] = sorted((c.Parent.GetAddress(False, False), c.Text(), c.Visible) for c in ws.Comments)
    out["links"] = sorted((h.Range.GetAddress(False, False), h.Address, h.SubAddress, h.TextToDisplay, h.ScreenTip) for h in ws.Hyperlinks)
    out["conditional"] = safe(lambda: [(fc.Type, safe(lambda fc=fc: fc.AppliesTo.GetAddress(False, False))) for fc in ws.Cells.FormatConditions], [])
    out["validation"] = safe(lambda: validation_props(ws))
    out["protection"] = (ws.ProtectContents, ws.Protection.AllowFormattingCells, ws.Protection.AllowSorting, ws.Protection.AllowInsertingRows)
    out["autofilter"] = (ws.AutoFilterMode, safe(lambda: ws.AutoFilter.Range.GetAddress(False, False)))
    out["tables"] = [(lo.Name, lo.Range.GetAddress(False, False), lo.ShowTotals) for lo in ws.ListObjects]
    # Header codes through the English interface (German COM shows &P as &S, &N as &A).
    out["page"] = (english(ws.PageSetup, "PrintArea"), english(ws.PageSetup, "CenterHeader"), english(ws.PageSetup, "RightFooter"), ws.PageSetup.Orientation,
                   english(ws.PageSetup, "PrintTitleRows"), english(ws.PageSetup, "PrintTitleColumns"))
    ws.Activate()
    w = wb.Windows(1)
    out["window"] = {
        "FreezePanes": w.FreezePanes, "SplitRow": w.SplitRow, "SplitColumn": w.SplitColumn, "ScrollRow": w.ScrollRow,
        "ScrollColumn": w.ScrollColumn, "Zoom": w.Zoom, "Gridlines": w.DisplayGridlines, "Headings": w.DisplayHeadings,
        "ActiveCell": wb.Application.ActiveCell.GetAddress(False, False), "Selection": wb.Application.Selection.GetAddress(False, False),
    }
    return out


def book_props(wb, model: dict | None) -> dict:
    active = wb.ActiveSheet.Name
    sheets = {}
    for ws in wb.Worksheets:
        cells: set[str] = set()
        if model:
            m = next((s for s in model["sheets"] if s["name"] == ws.Name), None)
            if m:
                cells |= set(m["cells"])
        for kind in (2, -4123):  # constants, formulas
            a = safe(lambda kind=kind: ws.Cells.SpecialCells(kind).GetAddress(False, False))
            if a:
                for part in a.split(","):
                    rng = ws.Range(part)
                    if rng.Count <= 5000:
                        cells |= {c.GetAddress(False, False) for c in rng.Cells}
        sheets[ws.Name] = sheet_props(wb, ws, sorted(cells))
    names = sorted((n.Name, english(n, "RefersTo"), n.Visible) for n in wb.Names)
    return {"Date1904": wb.Date1904, "ActiveSheet": active, "Sheets": [s.Name for s in wb.Sheets],
            "names": names, "sheets": sheets}


def diff(a, b, path="") -> list[str]:
    if isinstance(a, dict) and isinstance(b, dict):
        out = []
        for k in sorted(set(a) | set(b), key=str):
            out += diff(a.get(k), b.get(k), f"{path}/{k}")
        return out
    if isinstance(a, (list, tuple)) and isinstance(b, (list, tuple)) and len(a) == len(b) and not path.endswith(("Font", "Interior", "Align")):
        out = []
        for i, (x, y) in enumerate(zip(a, b)):
            out += diff(x, y, f"{path}/{i}")
        return out
    if isinstance(a, float) and isinstance(b, float) and abs(a - b) < 1e-9:
        return []
    return [] if a == b else [f"{path}: {a!r} -> {b!r}"]


def open_book(app, path: Path):
    """Opens read-only without updating links; raises when Excel finds the file damaged."""
    return app.Workbooks.Open(str(path), 0, True, None, None, None, True, None, None, False, False, None, False)


def formula_shown(f: str) -> str:
    """A formula in file form as Excel's Formula property shows it."""
    f = re.sub(r"(?:_xlfn[.])?SINGLE[(]([^()]*)[)]", lambda m: "@" + m.group(1), f)
    return "=" + f.replace("_xlfn._xlws.", "").replace("_xlfn.", "").replace("_xlws.", "")


def same_numfmt(ours: str, excel: str) -> bool:
    """Excel's NumberFormat drops escapes and shows locale currency formats in US form."""
    norm = lambda s: s.replace("\\", "").replace('"', "")
    return norm(ours) == norm(excel) or "€" in ours or "[$" in ours


def neighbour(cells: dict, addr: str, side: str) -> bool:
    """Whether the cell beside `side` has a border on the shared edge (it may win in Excel)."""
    m = re.match(r"([A-Z]+)(\d+)$", addr)
    col = sum((ord(ch) - 64) * 26 ** i for i, ch in enumerate(reversed(m.group(1))))
    row = int(m.group(2))
    dr, dc, other = {"top": (-1, 0, "bottom"), "bottom": (1, 0, "top"), "left": (0, -1, "right"), "right": (0, 1, "left")}[side]
    r, c = row + dr, col + dc
    if r < 1 or c < 1:
        return False
    name = ""
    while c:
        c, rem = divmod(c - 1, 26)
        name = chr(65 + rem) + name
    n = cells.get(f"{name}{r}")
    return bool(n and other in n["style"]["borders"])


def read_check(model: dict, excel: dict) -> list[str]:
    """The model as read against Excel's view of the original file."""
    out: list[str] = []
    if model["date1904"] != excel["Date1904"]:
        out.append("Date1904")
    for ms in model["sheets"]:
        es = excel["sheets"].get(ms["name"])
        if es is None:
            out.append(f"sheet {ms['name']} missing")
            continue
        hidden = {False: -1, True: 0, "very": 2}[ms["hidden"]]
        if hidden != es["Visible"]:
            out.append(f"{ms['name']} Visible {hidden} vs {es['Visible']}")
        if ms["tab"] and bgr(ms["tab"]) != es["Tab"]:
            out.append(f"{ms['name']} tab {ms['tab']} vs {es['Tab']:06X}")
        for addr, mc in ms["cells"].items():
            ec = es["cells"].get(addr)
            if ec is None:
                continue
            p = f"{ms['name']}!{addr}"
            st = mc["style"]
            v = mc["v"]
            ev = ec["Value2"]
            if isinstance(v, dict):
                if not isinstance(ev, int) or ev >= 0:
                    out.append(f"{p} value {v} vs {ev!r}")
            elif isinstance(v, bool):
                if ev != v:
                    out.append(f"{p} value {v} vs {ev!r}")
            elif v is None:
                if ev not in (None, ""):
                    out.append(f"{p} value None vs {ev!r}")
            elif v != ev and not (isinstance(v, (int, float)) and isinstance(ev, float) and abs(v - ev) < 1e-12 * max(1, abs(v))):
                out.append(f"{p} value {v!r} vs {ev!r}")
            if mc["f"] is not None:
                ours = formula_shown(mc["f"])
                theirs = ec["Formula"]
                # Excel shows a table's own columns without the table name ([Note]).
                if ours != theirs and re.sub(r"[A-Za-z_][\w.]*\[", "[", ours) != re.sub(r"[A-Za-z_][\w.]*\[", "[", theirs):
                    out.append(f"{p} formula {ours!r} vs {ec['Formula']!r}")
            # Dynamic arrays (spills) are not arrays to COM.
            if mc["array"] and ec["HasArray"] and mc["array"] != ec["Array"]:
                out.append(f"{p} array {mc['array']} vs {ec['Array']}")
            if not same_numfmt(st["numFmt"], ec["NumberFormat"]):
                out.append(f"{p} numFmt {st['numFmt']!r} vs {ec['NumberFormat']!r}")
            f = st["font"]
            ef = ec["Font"]
            mine = (f["name"], f["size"], bool(f.get("bold")), bool(f.get("italic")), UNDERLINE[f.get("underline")],
                    bool(f.get("strike")), f.get("vertAlign") == "superscript", f.get("vertAlign") == "subscript", bgr(f.get("color")))
            if not mc["rich"] and tuple(mine) != tuple(ef):
                out.append(f"{p} font {mine} vs {tuple(ef)}")
            ei = ec["Interior"]
            if st.get("fill"):
                if ei[0] != 1 or ei[1] != bgr(st["fill"]):
                    out.append(f"{p} fill {st['fill']} vs {ei}")
            elif not st.get("fillXml") and ei[0] != -4142:
                out.append(f"{p} fill none vs {ei}")
            al = (H_ALIGN[st.get("h")], V_ALIGN[st.get("v")], bool(st.get("wrap")), st.get("indent", 0))
            if al != tuple(ec["Align"][:4]):
                out.append(f"{p} align {al} vs {tuple(ec['Align'][:4])}")
            rot = st.get("rotation", 0)
            erot = ec["Align"][4]
            want = {255: [-4166], 90: [90, -4171], 180: [-90, -4170]}.get(rot, [rot if rot <= 90 else -(rot - 90)])
            if (erot if erot != -4128 else 0) not in want:
                out.append(f"{p} rotation {rot} vs {erot}")
            if bool(st.get("shrink")) != ec["Align"][5]:
                out.append(f"{p} shrink")
            if (st.get("locked", True), bool(st.get("hidden"))) != tuple(ec["Protection"]):
                out.append(f"{p} protection")
            bd = st["borders"]
            for side, key in (("left", "Border left"), ("right", "Border right"), ("top", "Border top"), ("bottom", "Border bottom")):
                # Excel shows the border a cell shares with its neighbour from either side.
                shared = neighbour(ms["cells"], addr, side)
                if side in bd and not shared:
                    ls, wt = BORDER.get(bd[side]["style"], (None, None))
                    eb = ec[key]
                    if not eb or (eb[0], eb[1], eb[2]) != (ls, wt, bgr(bd[side].get("color"))):
                        out.append(f"{p} border {side} {bd[side]} vs {eb}")
        for c, info in ms["cols"].items():
            ew = es["cols"][int(c) - 1] if int(c) - 1 < len(es["cols"]) else None
            if ew and bool(info["hidden"]) != bool(ew[1]):
                out.append(f"{ms['name']} col {c} hidden")
        for r, info in ms["rows"].items():
            er = es["rows"][int(r) - 1] if int(r) - 1 < len(es["rows"]) else None
            if er is None:
                continue
            if info["hidden"] != er[1]:
                out.append(f"{ms['name']} row {r} hidden {info['hidden']} vs {er[1]}")
            elif info["height"] is not None and info["custom"] and not info["hidden"] and abs(info["height"] - er[0]) > 1e-9:
                out.append(f"{ms['name']} row {r} height {info['height']} vs {er[0]}")
        notes = sorted((a, c["note"]) for a, c in ms["cells"].items() if c["note"] is not None)
        if notes != sorted((a, t) for a, t, _ in es["notes"]):
            out.append(f"{ms['name']} notes {notes} vs {es['notes']}")
        w = es["window"]
        v = ms["view"]
        if v["zoom"] != w["Zoom"] or v["showGrid"] != w["Gridlines"] or v["showHeaders"] != w["Headings"]:
            out.append(f"{ms['name']} window zoom/grid/headings")
        if bool(v.get("freeze")) != bool(w["FreezePanes"]):
            out.append(f"{ms['name']} freeze {v.get('freeze')} vs {w['FreezePanes']}")
    return out


def lookup(props, path: str):
    """A value of Excel's view by a path like /sheets/Werte/cells/A2/Value2."""
    cur = props
    for key in path.strip("/").split("/"):
        if isinstance(cur, dict):
            cur = cur.get(key)
        elif isinstance(cur, (list, tuple)) and key.isdigit() and int(key) < len(cur):
            cur = cur[int(key)]
        else:
            return "<missing>"
    return cur


def same(a, b) -> bool:
    if isinstance(b, dict) and "approx" in b:
        # Sizes Excel shows depend on the screen's dpi.
        return isinstance(a, (int, float)) and abs(a - b["approx"]) <= b["tol"]
    if isinstance(a, (int, float)) and isinstance(b, (int, float)) and not isinstance(a, bool):
        return abs(a - b) < 1e-9
    return a == b


def check_one(app, orig: Path | None, ours: Path, model: dict | None, exp: dict, edit: bool) -> dict:
    """Opens the original and our file and compares what Excel sees."""
    a = None
    if orig:
        wa = open_book(app, orig)
        try:
            a = book_props(wa, model)
        finally:
            wa.Close(False)
    try:
        wb = open_book(app, ours)
    except Exception as e:  # noqa: BLE001
        return {"opened": False, "error": str(e)[:200]}
    try:
        b = book_props(wb, model)
    finally:
        wb.Close(False)
    d = diff(a, b) if a else []
    allowed = [x for x in d if any(x.split(": ", 1)[0].startswith(p) for p in exp["allow"])]
    d = [x for x in d if x not in allowed]
    rc = read_check(model, a) if model and a and not edit else []
    missing = [f"{p}: want {v!r}, Excel shows {lookup(b, p)!r}" for p, v in exp["require"] if not same(lookup(b, p), v)]
    return {"opened": True, "differences": len(d), "allowed": len(allowed), "allowed_details": allowed[:50], "details": d[:50], "read_check": len(rc), "read_details": rc[:50],
            "missing": missing, "cells": sum(len(s["cells"]) for s in b["sheets"].values())}


def run_tests(rt: Path, models: Path) -> None:
    env = dict(os.environ, ROUNDTRIP_OUT=str(rt), MODEL_OUT=str(models))
    npx = shutil.which("npx") or "npx"
    r = subprocess.run([npx, "vitest", "run", "packages/ooxml/test/xlsx"], cwd=ROOT, env=env, capture_output=True, text=True, encoding="utf-8", errors="replace", shell=os.name == "nt")
    tail = (r.stdout + r.stderr).strip().splitlines()[-6:]
    print("\n".join(tail).encode("ascii", "replace").decode("ascii"))
    if r.returncode:
        raise SystemExit("the round-trip tests failed")


def broken_variants(src: Path, out: Path) -> list[Path]:
    """Files Excel must refuse: a bad style index, cells out of order."""
    cases = {
        "broken-style.xlsx": (b'<c r="A2"><v>42</v></c>', b'<c r="A2" s="999"><v>42</v></c>'),
        "broken-order.xlsx": (b'<c r="A2"><v>42</v></c>', b'<c r="B2"><v>1</v></c><c r="A2"><v>42</v></c>'),
    }
    paths = []
    for name, (old, new) in cases.items():
        p = out / name
        with zipfile.ZipFile(src) as z, zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as o:
            for i in z.infolist():
                d = z.read(i)
                if i.filename == "xl/worksheets/sheet1.xml":
                    assert old in d
                    d = d.replace(old, new)
                o.writestr(i, d)
        paths.append(p)
    return paths


def main() -> int:
    tmp = Path(tempfile.mkdtemp(prefix="raabe-xlsx-"))
    rt, models = tmp / "roundtrip", tmp / "model"
    report: dict = {"date": datetime.date.today().isoformat(), "fixtures": {}}
    failures = 0
    x = ExcelTarget(visible=False)
    try:
        if "--selftest" in sys.argv:
            for p in broken_variants(FIXTURES / "xlsx-values.xlsx", tmp):
                try:
                    open_book(x.app, p).Close(False)
                    print(f"SELFTEST FAIL {p.name}: Excel opened a damaged file without complaint")
                    failures += 1
                except Exception:  # noqa: BLE001
                    print(f"selftest ok {p.name}: Excel refuses it")
            return failures
        run_tests(rt, models)
        for ours in sorted(rt.glob("*.xlsx")):
            name = ours.name
            base = name.replace(".edit.", ".")
            edit = base != name
            orig = FIXTURES / base
            model_file = models / f"{base}.json"
            model = json.loads(model_file.read_text(encoding="utf-8")) if model_file.exists() else None
            exp = json.loads((rt / f"{name[:-5]}.expect.json").read_text(encoding="utf-8")) if edit else {"allow": [], "require": []}
            entry = check_one(x.app, orig if orig.exists() else None, ours, model, exp, edit)
            report["fixtures"][name] = entry
            ok = entry.get("opened") and not entry["differences"] and not entry["read_check"] and not entry["missing"]
            if not ok:
                failures += 1
            if not entry.get("opened"):
                print(f"FAIL {name}: Excel could not open it cleanly: {entry['error']}")
                continue
            print(f"{'PASS' if ok else 'FAIL'} {name}: {entry['cells']} cells, {entry['differences']} differences"
                  f"{' (+%d allowed)' % entry['allowed'] if edit else ''}, read check {entry['read_check']}, requirements missed {len(entry['missing'])}")
            for line in (entry["details"] + entry["read_details"] + entry["missing"])[:25]:
                print("   ", line)
    finally:
        x.close()
        shutil.rmtree(tmp, ignore_errors=True)
    RESULTS.mkdir(exist_ok=True)
    out = RESULTS / f"{report['date']}-xlsx-excel-check.json"
    out.write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print("result:", out.name, "failures:", failures)
    return failures


if __name__ == "__main__":
    sys.exit(main())
