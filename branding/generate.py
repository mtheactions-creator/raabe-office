"""Generates the Raabe Office logo SVGs.

Geometry was measured from original-logo-reference.png (least-squares circle
fits, max residual < 0.7 px). Letters are drawn from simple strokes so the
logos contain no font outlines and need no font licence.
"""
from pathlib import Path

RED = "#ED1C24"     # exact colour of the original logo
SHADOW = "#231F20"  # exact colour of the original drop shadow
SHADOW_OFFSET = (11, 12)

CRESCENT = "M547.1 81.3A329.4 329.4 0 1 0 165.9 614.3A328.5 328.5 0 0 1 547.1 81.3Z"
DISC = '<circle cx="382.5" cy="367.6" r="259.2"/>'
R_GLYPH = ('<path d="M261 195H342V241C358 214 388 190 425 189C450 189 468 195 477 204'
           'L449 281C436 270 422 266 408 266C390 266 373 276 365 290C355 307 350 330 349 360'
           'L348 395V531H261Z"/>')
DOT = '<circle cx="451" cy="478" r="56.5"/>'

# Small letters in place of the dot (variant A). Cap height 384..531, stroke 41.
STROKE = 'fill="none" stroke="currentColor" stroke-width="41"'
LETTERS = {
    "W": f'<path clip-path="url(#capband)" {STROKE} stroke-miterlimit="10" '
         'd="M389 345L420 552L459 430L498 552L529 345"/>',
    "C": f'<path {STROKE} d="M501.4 422A53 53 0 1 0 501.4 493"/>',
    "P": f'<rect x="400" y="384" width="41" height="147" fill="currentColor"/>'
         f'<path {STROKE} d="M420 404.5H470A30 30 0 0 1 470 464.5H420"/>',
}

# Small-size icon (<= 32 px): disc only, one big letter, no crescent / shadow.
BIG = 'fill="none" stroke="#FFFFFF" stroke-width="64"'
BIG_LETTERS = {
    "W": f'<path clip-path="url(#bigband)" {BIG} stroke-miterlimit="10" '
         'd="M222 140L276 520L342 330L408 520L462 140"/>',
    "C": f'<path {BIG} d="M418 270A118 118 0 1 0 418 465"/>',
    "P": f'<rect x="229" y="185" width="64" height="300" fill="#FFFFFF"/>'
         f'<path {BIG} d="M261 217H354A70 70 0 0 1 354 357H261"/>',
}


def logo(title: str, mark: str) -> str:
    dx, dy = SHADOW_OFFSET
    glyphs = R_GLYPH + mark
    return f'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 642 627">
  <title>{title}</title>
  <defs><clipPath id="capband"><rect x="300" y="384" width="300" height="147"/></clipPath></defs>
  <g fill="{RED}"><path d="{CRESCENT}"/>{DISC}</g>
  <g fill="{SHADOW}" color="{SHADOW}" transform="translate({dx} {dy})">{glyphs}</g>
  <g fill="#FFFFFF" color="#FFFFFF">{glyphs}</g>
</svg>
'''


def small_icon(title: str, letter: str) -> str:
    return f'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 684 684">
  <title>{title}</title>
  <defs><clipPath id="bigband"><rect x="0" y="185" width="684" height="300"/></clipPath></defs>
  <circle cx="342" cy="342" r="342" fill="{RED}"/>
  {BIG_LETTERS[letter]}
</svg>
'''


def main() -> None:
    out = Path(__file__).parent
    (out / "raabe-logo.svg").write_text(logo("Raabeschule", DOT), encoding="utf-8")
    for app, letter in (("RaabeWrite", "W"), ("RaabeCalc", "C"), ("RaabePresent", "P")):
        slug = app.lower().replace("raabe", "raabe-")
        (out / f"{slug}.svg").write_text(logo(app, LETTERS[letter]), encoding="utf-8")
        (out / f"{slug}-small.svg").write_text(small_icon(app, letter), encoding="utf-8")


if __name__ == "__main__":
    main()
