# VBA-JSON (test data)

`JsonConverter.bas` is VBA-JSON v2.3.1 by Tim Hall, unchanged, from
https://github.com/VBA-tools/VBA-JSON (file `JsonConverter.bas`, branch `master`, fetched 2026-11-08).
Licence: MIT (`LICENSE` in this folder); the file also carries the BSD notice of the code it grew from.

It is not part of Raabe Office: `check_present_macro_suite.py` uses it as a stress test of our VBA interpreter
(conditional compilation, Scripting.Dictionary, Collection, string and number functions), running the same
JSON round trip in PowerPoint and in RaabePresent (D-105).
