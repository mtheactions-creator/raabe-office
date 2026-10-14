"""Probe: what Word's Strg+Eingabe (InsertBreak wdPageBreak) does to the text and the caret,
at the end of a paragraph and in the middle. Own invisible Word instance."""
import win32com.client

app = win32com.client.DispatchEx("Word.Application")
app.Visible = False
try:
    for cmd in ("InsertPageBreak", "InsertColumnBreak"):
     for case, at in (("end", None), ("middle", 5), ("start", 0)):
         d = app.Documents.Add()
         d.Content.Text = "Erster Absatz\rZweiter Absatz"
         s = app.Selection
         pos = len("Erster Absatz") if at is None else at
         s.SetRange(pos, pos)
         # Strg+Eingabe is bound to InsertPageBreak; WordBasic runs the same command.
         import pythoncom
         o = app.WordBasic._oleobj_
         o.Invoke(o.GetIDsOfNames(cmd), 0, pythoncom.DISPATCH_METHOD, False)
         t = d.Content.Text
         print(cmd, case, repr(t), "sel", s.Start, s.End, "page", s.Information(3), "paras", d.Paragraphs.Count)
         d.Close(0)
finally:
    app.Quit(0)
