# Raabe Office: notes for Claude

You are the maintainer of this project (D-016). You do not remember earlier sessions;
this repository is the memory. Read this file first, then `docs/DECISIONS.md`.

## What this is
A lean, offline office suite for the Raabeschule (Gymnasium in Niedersachsen, Germany)
that looks and behaves like Microsoft 365 so students can use their Office muscle memory:
**RaabeWrite** (Word), **RaabeCalc** (Excel), **RaabePresent** (PowerPoint).
Windows first, iPad close behind (new students all use iPads). Open source, GPL-3.0.

## Working agreement with the project owner (a teacher at the school)
- **No new step starts without explicit agreement.** Discuss, then wait for a clear "go".
- Ask before installing software on the development PC.
- Never handle passwords or credentials. The owner authenticates `gh` themselves.
- The owner writes in English; the apps' UI is German, using Microsoft's exact German
  labels and shortcuts.
- Feel is judged by measurement against real Word (`docs/FEEL.md`), not by asking the
  owner to test consciously. Running the harness takes over keyboard and screen: tell
  the owner first and only run it when they are not using the PC.

## Where things are
- `docs/DECISIONS.md`: every decision with its reasoning. Add an entry for each new one.
- `docs/SCOPE.md`: what each app does (MVP / later / out of scope).
- `docs/ARCHITECTURE.md`: technology and structure.
- `docs/ROADMAP.md`: steps and their status.
- `docs/FEEL.md`: feel metrics and targets.
- `branding/`: logos; regenerate with `python branding/generate.py`. Not GPL.
- `tools/feel-lab/`: measurement harness (Python venv in `tools/feel-lab/.venv`).
- `prototypes/`: Step 0 feel prototypes (throwaway code).

## Environment (development PC)
Windows 10 Pro, PowerShell + Git Bash, Node 24, Rust (stable, MSVC), Python 3.14,
Visual Studio 2022 C++ tools, WebView2, Microsoft Word 365 (COM: `Word.Application`).
Display: 4K at 150 % scaling, currently 30 Hz.

## Conventions
- Internal docs and code comments in English; UI strings in German.
- Commit messages in English, imperative mood.
- Keep `docs/` current in the same commit as the change it describes.
