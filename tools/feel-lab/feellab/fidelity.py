"""F4 layout fidelity: do lines and pages break at the same characters as in Word?
Needs no keyboard or mouse input."""

from __future__ import annotations

from .targets import Target


def compare(reference: list[int], candidate: list[int], text: str, limit: int = 8) -> dict:
    ref, cand = set(reference), set(candidate)
    matched = len(ref & cand)
    mismatches = []
    for o in sorted(ref ^ cand)[:limit]:
        side = "only in Word" if o in ref else "only in candidate"
        mismatches.append({"offset": o, "where": side, "context": text[max(0, o - 25):o] + "|" + text[o:o + 25]})
    return {
        "reference": len(reference),
        "candidate": len(candidate),
        "identical": matched,
        "percent": round(100 * matched / max(len(ref), 1), 2),
        "first_differences": mismatches,
    }


def run(word: Target, candidates: list[Target], fixtures: list[str]) -> dict:
    out: dict = {}
    for name in fixtures:
        word.load(name)
        text = word.text()
        ref_lines, ref_pages = word.lines(), word.pages()
        out[name] = {"word": {"lines": len(ref_lines), "pages": len(ref_pages)}}
        for c in candidates:
            c.load(name)
            assert c.text() == text, f"{c.name}: document text differs from Word's"
            out[name][c.name] = {
                "lines": compare(ref_lines, c.lines(), text),
                "pages": compare(ref_pages, c.pages(), text),
            }
    return out
