"""Probe: Word's AutoCorrect and AutoFormat-as-you-type settings, the symbol entries of the
AutoCorrect list, and the abbreviation exceptions (own invisible Word instance)."""
import json

import win32com.client

app = win32com.client.DispatchEx("Word.Application")
app.Visible = False
try:
    ac = app.AutoCorrect
    o = app.Options
    flags = {k: getattr(ac, k) for k in ("CorrectInitialCaps", "CorrectSentenceCaps", "CorrectDays", "CorrectCapsLock",
                                         "ReplaceText", "ReplaceTextFromSpellingChecker", "CorrectKeyboardSetting", "CorrectTableCells",
                                         "FirstLetterAutoAdd", "TwoInitialCapsAutoAdd")}
    fmt = {k: getattr(o, k) for k in dir(o) if k.startswith("AutoFormatAsYouType")}
    entries = [(e.Name, e.Value, e.RichText) for e in ac.Entries]
    symbols = [x for x in entries if not any(ch.isalpha() for ch in x[0]) or len(x[1]) <= 2 and not x[1].isalpha()]
    first = [e.Name for e in ac.FirstLetterExceptions]
    two = [e.Name for e in ac.TwoInitialCapsExceptions]
    print(json.dumps({"flags": flags, "autoformat": fmt, "entries": len(entries), "symbol_entries": symbols[:80],
                      "first_letter_exceptions": len(first), "first_sample": first[:60], "two_caps": two[:20]}, ensure_ascii=False, indent=1))
finally:
    app.Quit(0)
