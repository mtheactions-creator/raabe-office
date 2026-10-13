# Decision log

Every decision that shapes the project gets an entry here: what was decided, why, and
what it rules out. Newest at the bottom. Never rewrite an old entry. If a decision is
reversed, add a new entry that supersedes it and mark the old one `Superseded by D-xxx`.

| ID | Date | Decision |
|---|---|---|
| D-001 | 2026-10-13 | Build our own suite instead of adopting ONLYOFFICE |
| D-002 | 2026-10-13 | Three apps: RaabeWrite, RaabeCalc, RaabePresent |
| D-003 | 2026-10-13 | Look and feel of current Microsoft 365 (German UI) |
| D-004 | 2026-10-13 | OOXML (.docx/.xlsx/.pptx) is the native file format |
| D-005 | 2026-10-13 | Fully local, no network access by default |
| D-006 | 2026-10-13 | Tauri 2 + TypeScript web UI (provisional, pending D-014) |
| D-007 | 2026-10-13 | Windows first, iPad close behind; touch-first design |
| D-008 | 2026-10-13 | Branding: logo variant A, all three apps pure Raabe red |
| D-009 | 2026-10-13 | GPL-3.0 for code; school logo and name excluded |
| D-010 | 2026-10-13 | Document tabs in one window (deliberate deviation from Word) |
| D-011 | 2026-10-13 | Default font Calibri 11, Carlito as metric-compatible fallback |
| D-012 | 2026-10-13 | Scope additions: comments, track changes, formula editor, school templates |
| D-013 | 2026-10-13 | "Feel" is measured objectively against real Word, not judged by gut |
| D-014 | 2026-10-13 | Editing-engine choice decided by the feel prototypes (Step 0) |
| D-015 | 2026-10-13 | Updates only on explicit user request; code signing later |
| D-016 | 2026-10-13 | Claude is the maintainer; the repository is the project's memory |

---

### D-001 Build our own suite instead of adopting ONLYOFFICE
The project owner evaluated ONLYOFFICE Desktop Editors. Verdict: a convincing Word
alternative, much easier for Word users than LibreOffice, but it "feels like a web-based
document editor packaged as a desktop application". Selecting text, dragging objects,
adjusting tables and navigating long documents feel less immediate than in Word.
**Consequence:** fluid, native-feeling interaction is our primary quality bar, ahead of
feature count. ONLYOFFICE remains the benchmark we have to beat.

### D-002 Three apps
RaabeWrite (Word), RaabeCalc (Excel), RaabePresent (PowerPoint). They share one shell
(ribbon, backstage, file handling, dialogs). Build order: shell first, then RaabeWrite.

### D-003 Look and feel of current Microsoft 365, German UI
Students used Microsoft 365 until mid-2026, so the reference is the current M365 desktop
look with the classic (full) ribbon. Labels use Microsoft's exact German wording
("Start", "Einfügen", "Layout", "Referenzen", "Überprüfen", "Ansicht") and Word's German
keyboard shortcuts, because that is where the muscle memory lives.
**Not copied:** Microsoft icons, artwork, logos, product names. Icons come from openly
licensed sets (e.g. Fluent UI System Icons, MIT) or are drawn by us.

### D-004 OOXML is the native file format
Files must move freely between students, teachers, home PCs with Word, and LibreOffice.
We support a defined subset well. **Rule:** content we do not understand is preserved
untouched when saving and shown as a placeholder. We never silently destroy parts of a
teacher's document.

### D-005 Fully local
No accounts, no cloud, no telemetry, no automatic network calls. Translation models,
dictionaries and templates are bundled or installed locally. This also keeps the apps
GDPR-safe (DSGVO) without any data-protection paperwork.

### D-006 Tauri 2 + TypeScript (provisional)
Small installers (~10 MB), fast start, uses the WebView2 runtime built into Windows,
and the same code base can target iPadOS via Tauri's iOS support. The risk is the
"web app in a box" feel criticised in D-001. That risk is tested by D-014 before we
commit. Fallback if the prototypes fail: a native Windows stack, at the cost of a much
harder iPad port.

### D-007 Windows first, iPad close behind
All new students must buy iPads; grade 8 and up use Windows laptops/tablets (BYOD, parents
manage the devices). The iPad share grows every year. **Consequence:** every interaction
must work by touch from day one, the ribbon needs a touch mode, and the architecture must
leave room for Apple Pencil ink later. iOS builds run on GitHub's macOS runners (no Mac
needed locally).

### D-008 Branding
Logo variant A: the school logo (red crescent + red disc, white "r" with dark drop
shadow) with the dot replaced by the app letter: W, C, P. All three apps use pure Raabe
red `#ED1C24` (measured from the original); no per-app accent colours. Below 48 px a
simplified icon is used (red disc + large white letter), because the full logo is
illegible at 16-32 px. See `branding/README.md`.

### D-009 Licence
Code: GPL-3.0, so the suite stays open. The Raabeschule name and logo are excluded from
the licence; forks must rebrand. The school leadership approved publishing the name and
logo in the public repository.

### D-010 Document tabs
Several documents open as tabs in one window, like ONLYOFFICE. This is a deliberate
improvement over Word, requested by the project owner. Everything else follows Word.

### D-011 Default font
Microsoft 365 defaults to Aptos, which we cannot ship. Default is Calibri 11 (installed
with Windows, familiar from years of older Word versions). Where Calibri is missing
(iPad), we render with Carlito, which has identical metrics, so documents never reflow.
Same approach for Arial/Times/Courier (Liberation fonts) and Cambria (Caladea).

### D-012 Scope additions
Confirmed as needed: comments and track changes (teachers correct work digitally);
formula editor (high priority, maths and physics); Raabeschule templates (worksheet with
school header, Facharbeit template following the school's formal rules; the owner will
supply the material later).

### D-013 Feel is measured objectively
The project owner pointed out that consciously testing puts you in a different mental
mode than everyday use, so their judgement of subtle feel would be unreliable. Instead
we measure against Microsoft Word, which is installed on the development PC (M365,
16.0.20430, scriptable via COM): keystroke-to-screen latency, frame drops while
scrolling and dragging, a catalogue of micro-behaviours recorded from Word, layout
fidelity (line and page breaks) and start/load times. See `docs/FEEL.md`. The final
check is real use by a pilot class.

### D-014 Editing engine chosen by prototypes
Before writing product code we build two small editor prototypes with different
architectures, measure both with the D-013 harness, and pick the winner. Recorded as a
new decision when done.

### D-015 Updates and signing
Updates are only fetched when the user explicitly clicks "Nach Updates suchen", or come
through the Microsoft Store / App Store. Code signing (Microsoft Store or SignPath's free
open-source signing) is agreed but postponed until the first release.

### D-016 Maintenance
Claude maintains the code. Claude does not remember earlier sessions, so everything
needed to continue must live in this repository: `CLAUDE.md`, this decision log, the
architecture notes, and automated tests. Any session must be able to pick up the work
from the repository alone.
