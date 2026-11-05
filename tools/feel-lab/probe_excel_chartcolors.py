"""Probe: the colour schemes of Diagrammentwurf > Farben ändern (ChartColor 10-26) as Excel saves them
in the chart's colour part (cs:colorStyle): the theme colours used and the variations applied to later
series. Own invisible Excel, COM and the saved file only (D-066).
Result: results/<date>-excel-chartcolors.json.
"""
from __future__ import annotations

import datetime
import json
import re
import tempfile
import zipfile
from pathlib import Path

from feellab.excel import ExcelTarget

HERE = Path(__file__).parent


def main() -> None:
    x = ExcelTarget(visible=False)
    out: dict = {}
    tmp = Path(tempfile.gettempdir())
    try:
        ws = x.sheet
        for r, row in enumerate([("M", "A", "B"), ("J", 3, 4), ("F", 5, 2)], start=1):
            for c, v in enumerate(row, start=1):
                ws.Cells(r, c).Value = v
        for k in range(10, 27):
            for s in list(ws.Shapes):
                s.Delete()
            ch = ws.Shapes.AddChart2(-1, 51, 300, 20, 360, 216).Chart
            ch.SetSourceData(ws.Range("A1:C3"))
            ch.ChartColor = k
            path = tmp / f"raabe-colors-{k}.xlsx"
            x.wb.SaveCopyAs(str(path))
            with zipfile.ZipFile(path) as z:
                names = [n for n in z.namelist() if re.search(r"xl/charts/colors\d+\.xml$", n)]
                xml = z.read(names[0]).decode("utf-8") if names else ""
            meth = re.search(r'meth="(\w+)"', xml)
            colors = re.findall(r"<a:(schemeClr|srgbClr) val=\"(\w+)\"", xml.split("<cs:variation")[0])
            variations = [re.findall(r"<a:(\w+) val=\"(\d+)\"", v) for v in xml.split("<cs:variation")[1:]]
            out[k] = {"meth": meth.group(1) if meth else None, "colors": [c[1] for c in colors], "variations": variations}
            path.unlink(missing_ok=True)
            print(k, out[k]["meth"], out[k]["colors"], len(variations), flush=True)
    finally:
        x.close()
    (HERE / "results" / f"{datetime.date.today().isoformat()}-excel-chartcolors.json").write_text(json.dumps(out, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
