"""Probe: the symbol characters Word's formatted AutoCorrect entries insert (arrows, smileys)."""
import re

import win32com.client

app = win32com.client.DispatchEx("Word.Application")
app.Visible = False
try:
    d = app.Documents.Add()
    for name in ["-->", "<--", "==>", "<==", "<=>", ":)", ":-)", ":(", ":-(", ":|", ":-|", ";)", ";-)"]:
        d.Content.Text = ""
        r = d.Range(0, 0)
        app.AutoCorrect.Entries(name).Apply(r)
        xml = d.Content.WordOpenXML
        syms = re.findall(r'<w:sym w:font="([^"]+)" w:char="([0-9A-F]+)"/>', xml)
        fonts = re.findall(r'<w:rFonts [^>]*/>', xml)
        print(name, syms or [hex(ord(c)) for c in d.Content.Text], fonts[:1] if not syms else "")
    d.Close(0)
finally:
    app.Quit(0)
