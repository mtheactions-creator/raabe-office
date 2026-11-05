"""Generated behaviour cases for RaabeCalc against Excel (D-066): every combination of editing mode,
input and situation that matters, instead of hand-picked cases. `catalogue_calc.py` runs them like
its own. Each group is an area of its own, so it can be recorded and compared alone.
"""

from __future__ import annotations

from feellab.calc_catalogue import Case as C


def _point_cases() -> list[C]:
    """Clicking, dragging and arrow keys while a formula is typed or edited."""
    out: list[C] = []
    # After which characters a click inserts a reference, and after which it ends the entry.
    for label, typed in [
        ("plus", "=1+"), ("minus", "=1-"), ("times", "=1*"), ("divide", "=1/"), ("power", "=1^"), ("ampersand", '="a"&'),
        ("less", "=1<"), ("greater", "=1>"), ("less equal", "=1<="), ("greater equal", "=1>="), ("not equal", "=1<>"), ("equals", "=1="),
        ("bracket", "=SUMME("), ("semicolon", "=SUMME(1;"), ("colon", "=A1:"), ("space", "=A1 "), ("after space", "=1+ "),
        ("percent", "=1%"), ("closing bracket", "=(1)"), ("number", "=12"), ("reference", "=A1"), ("name", "=SUMME"), ("open string", '="a'),
        ("closed string", '="a"'), ("leading minus", "-"), ("leading plus", "+"), ("leading at", "@"), ("minus then number", "-1+"), ("plain text", "abc"),
        ("text with plus", "a+"), ("number comma", "=1;"), ("nested", "=WENN(A1>1;"), ("sheet bang", "=Tabelle1!"),
    ]:
        out.append(C(f"click after {label}", {}, [("type", typed), ("click", "C3"), ("keys", "enter")], "point", start="B5"))
        out.append(C(f"arrow after {label}", {}, [("type", typed), ("keys", "up"), ("keys", "enter")], "point", start="B5"))
    # Bearbeiten: entered by F2, by a double-click or in the formula bar.
    for how, enter in [("F2", [("keys", "f2")]), ("double-click", [("dblclick", "B5")]), ("bar", [("fxclick",)])]:
        for label, setup, typed in [
            ("after an operator", "=1+2", "+"), ("after a number", "=1+2", ""), ("in text", "abc", ""), ("in text with plus", "abc", "+"),
            ("after a bracket", "=SUMME(1", ";"), ("after a reference", "=A1", ""),
        ]:
            steps = [*enter]
            if typed:
                steps.append(("type", typed))
            out.append(C(f"{how} click {label}", {"B5": setup}, [*steps, ("click", "C3"), ("keys", "enter")], "pointedit", start="B5"))
            out.append(C(f"{how} Down {label}", {"B5": setup}, [*steps, ("keys", "down"), ("keys", "enter")], "pointedit", start="B5"))
        out.append(C(f"{how} click with caret inside", {"B5": "=1+2"}, [*enter, ("keys", "left"), ("click", "C3"), ("keys", "enter")], "pointedit", start="B5"))
        out.append(C(f"{how} click with caret at the start", {"B5": "=1+2"}, [*enter, ("keys", "home"), ("click", "C3"), ("keys", "enter")], "pointedit", start="B5"))
        out.append(C(f"{how} click twice", {"B5": "=1+2"}, [*enter, ("type", "+"), ("click", "C3"), ("click", "D4"), ("keys", "enter")], "pointedit", start="B5"))
        out.append(C(f"{how} click then Esc", {"B5": "=1+2"}, [*enter, ("type", "+"), ("click", "C3"), ("keys", "esc")], "pointedit", start="B5"))
        out.append(C(f"{how} column header", {"B5": "=1+2"}, [*enter, ("type", "+"), ("colhead", "C"), ("keys", "enter")], "pointedit", start="B5"))
        out.append(C(f"{how} row header", {"B5": "=1+2"}, [*enter, ("type", "+"), ("rowhead", "3"), ("keys", "enter")], "pointedit", start="B5"))
    out.append(C("F2 twice then Down points", {"B5": "=1+2"}, [("keys", "f2"), ("type", "+"), ("keys", "f2"), ("keys", "down"), ("keys", "enter")], "pointedit", start="B5"))
    out.append(C("F2 in Eingeben then Down moves the caret", {}, [("type", "=1+"), ("keys", "f2"), ("keys", "down"), ("keys", "enter")], "pointedit", start="B5"))
    out.append(C("bar typing from empty", {}, [("fxclick",), ("type", "=1+"), ("click", "C3"), ("keys", "enter")], "pointedit", start="B5"))
    out.append(C("bar click on text then Esc", {"B5": "abc"}, [("fxclick",), ("keys", "esc")], "pointedit", start="B5"))
    out.append(C("bar click then click a cell", {"B5": "abc"}, [("fxclick",), ("click", "C3")], "pointedit", start="B5"))
    # What a pointed reference can become.
    for label, steps in [
        ("Strg+click adds", [("click", "A1"), ("click", "C3", "ctrl")]),
        ("Strg+drag adds", [("drag", "A1", "A3"), ("dragctrl", "C1", "C3")]),
        ("Umschalt+click extends", [("click", "A1"), ("click", "C3", "shift")]),
        ("Umschalt+arrows extend", [("click", "A1"), ("keys", "shift+down"), ("keys", "shift+right")]),
        ("Strg+Umschalt+Unten extends", [("click", "A1"), ("keys", "ctrl+shift+down")]),
        ("Strg+Unten jumps", [("click", "A1"), ("keys", "ctrl+down")]),
        ("column header", [("colhead", "C")]),
        ("column headers Umschalt", [("colhead", "C"), ("colhead", "E", "shift")]),
        ("column headers Strg", [("colhead", "C"), ("colhead", "E", "ctrl")]),
        ("row header", [("rowhead", "3")]),
        ("row headers Umschalt", [("rowhead", "3"), ("rowhead", "5", "shift")]),
        ("Strg+Leertaste", [("keys", "ctrl+space")]),
        ("Umschalt+Leertaste", [("keys", "shift+space")]),
        ("Strg+A", [("keys", "ctrl+a")]),
        ("click twice replaces", [("click", "A1"), ("click", "C3")]),
        ("click then typed operator then click", [("click", "A1"), ("type", ";"), ("click", "C3")]),
        ("click then F4", [("click", "A1"), ("keys", "f4")]),
        ("drag then F4", [("drag", "A1", "C3"), ("keys", "f4")]),
        ("drag then F4 twice", [("drag", "A1", "C3"), ("keys", "f4"), ("keys", "f4")]),
        ("click then Backspace", [("click", "A1"), ("keys", "backspace")]),
        ("arrow then Pos1", [("keys", "up"), ("keys", "home")]),
        ("arrow then Bild ab", [("keys", "up"), ("keys", "pagedown")]),
        ("Pos1 points", [("keys", "home")]),
        ("Strg+Pos1 points", [("keys", "ctrl+home")]),
        ("Strg+Ende points", [("keys", "ctrl+end")]),
        ("Bild auf points", [("keys", "pageup")]),
        ("click then wheel", [("click", "A1"), ("wheel", "C3", -3)]),
        ("click on the own cell", [("click", "D9")]),
    ]:
        out.append(C(f"pointing {label}", {"A1": "1", "A2": "2", "A3": "3", "C1": "4", "C2": "5", "C3": "6"}, [("type", "=SUMME("), *steps, ("keys", "enter")], "pointed", start="D9"))
    # The coloured frames of a formula's references can be moved and sized (Excel).
    for label, setup, steps in [
        ("corner of a cell", "=C3", [("dragfrom", "C3", "br", "E6")]),
        ("border of a cell", "=C3", [("dragfrom", "C3", "right", "E3")]),
        ("corner of a range", "=SUMME(C3:D5)", [("dragfrom", "D5", "br", "F7")]),
        ("corner of a range smaller", "=SUMME(C3:E6)", [("dragfrom", "E6", "br", "D4")]),
        ("border of a range", "=SUMME(C3:D5)", [("dragfrom", "D5", "right", "F5")]),
        ("second reference", "=C3+D5", [("dragfrom", "D5", "right", "F5")]),
        ("absolute reference moved", "=$C$3", [("dragfrom", "C3", "right", "E3")]),
    ]:
        out.append(C(f"frame {label} in Bearbeiten", {"A1": setup}, [("keys", "f2"), *steps, ("keys", "enter")], "frames"))
        out.append(C(f"frame {label} after a double-click", {"A1": setup}, [("dblclick", "A1"), *steps, ("keys", "enter")], "frames"))
    out.append(C("frame while typing", {}, [("type", "=C3+"), ("dragfrom", "C3", "br", "E6"), ("keys", "enter")], "frames"))
    return out


MATRIX: list[C] = _point_cases()
