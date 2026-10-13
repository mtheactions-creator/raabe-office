# Branding

**These files are not covered by the GPL-3.0 licence of this repository.** The name
„Raabeschule“, the school logo and the derived app logos belong to the Raabeschule and
are used here with the school leadership's permission. Forks must replace them.

## Files

| File | Use |
|---|---|
| `raabe-logo.svg` | School logo, vector rebuild of `original-logo-reference.png` |
| `raabe-write.svg`, `raabe-calc.svg`, `raabe-present.svg` | App logos (variant A, D-008), from 48 px upwards: splash screen, About dialog, Start menu tiles, website |
| `raabe-*-small.svg` | Simplified app icons for 32 px and below: taskbar, title bar, file icons |
| `generate.py` | Generates all SVGs. Edit this file, not the SVGs. |

## Colours

| Name | Hex | Use |
|---|---|---|
| Raabe red | `#ED1C24` | Logo, title bar accent, ribbon accent, selected states |
| Shadow | `#231F20` | Drop shadow in the logo only |

Both values were sampled from the original logo. The UI palette derived from Raabe red
(hover, pressed, light tints for selection) will be defined in the shell package.
