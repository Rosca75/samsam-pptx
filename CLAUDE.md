# CLAUDE.md — SamSam Tools (PowerPoint productivity add-in)

Read this file completely before touching anything. Every constraint below was
**empirically discovered** in the target environment (corporate locked-down Windows +
Microsoft 365 PowerPoint). Treat them as ground truth — do not "improve" the project in a
way that violates any of them.

## What this project is

A standalone PowerPoint productivity add-in ("SamSam Tools") written in **VBA**, packaged
as a **`.ppam`** add-in, with a ribbon tab defined in **`ribbon/customUI14.xml`** that is
injected into the `.ppam` package by a **Python 3 standard-library-only** script
(`build/inject_ribbon.py`). The approved feature scope is the reduced keep-list in
`docs/research.md` ("Shipped feature list (after the ribbon trim)") — the ribbon was
trimmed from an earlier broader build to only those functions plus one-click EN / FR
deck-language presets.

## Fixed technology stack (non-negotiable)

- **VBA packaged as `.ppam`.** Not Office.js (the JS API cannot do the shape/geometry
  work). Not VSTO or a COM DLL (those need Visual Studio, a signing certificate and
  registry writes — all blocked in the target environment).
- **Ribbon UI in `customUI14.xml`**, namespace
  `http://schemas.microsoft.com/office/2009/07/customui`, injected into the `.ppam`
  (a `.ppam` is an OPC ZIP package).
- **Ribbon injection via Python 3 using ONLY the standard library** (`zipfile`,
  `xml.etree.ElementTree`). Never add a pip dependency: the corporate TLS-inspection
  proxy blocks pip, npm, and often `git push`.

## Hard environment constraints (ground truth — encode these, never assume them away)

1. **Macro policy is "Disable all macros with notification", locked by Group Policy.**
   Unsigned VBA runs only after a per-session "Enable Macros" prompt, or silently if the
   `.ppam` lives in a **Trusted Location**. Design for this. Never assume signed macros
   or "enable all macros".
2. **Mark-of-the-Web hard-blocks downloaded macro files** — there is no enable button at
   all for a downloaded `.ppam`. The `.ppam` must always be **built locally on disk** —
   never downloaded, emailed, or pulled through a browser. The build script produces the
   file in place on the local machine.
3. **Saving as `.ppam` silently redirects**: PowerPoint ignores the folder chosen in the
   Save As dialog and writes the file to `%AppData%\Microsoft\AddIns` with no error.
   All build/inject tooling targets that path (`build/build.bat` does).
4. **VBA authoring gotchas:**
   - Ribbon callbacks must be declared **`(control As Object)`**, NOT
     `(control As IRibbonControl)`. `IRibbonControl` depends on an unreliable Office
     object-library reference; `Object` late-binds and always works. This applies to
     every callback signature (galleries, toggles, getters — all parameters late-bound).
   - **ALL module-level declarations** (`Const`, `Dim`, `Private`/`Public` variables)
     must appear **before the first `Sub`/`Function`** in a module, or VBA raises
     "Only comments may appear after End Sub".
   - **Debug → Compile VBAProject shows no success message** — a clean compile just greys
     the menu item out. That greyed-out state IS the success signal.
   - **"Trust access to the VBA project object model" is disabled and cannot be
     enabled**, so macros CANNOT import/export VBA modules programmatically. Module
     import is a **manual step in the VBA editor** (File → Import File…). Never build a
     pipeline that assumes programmatic VBA import.
5. **Ribbon XML gotchas:**
   - `imageMso` icon IDs must be valid **for PowerPoint specifically** (IDs differ per
     Office app). Example: `ShapeFill` is invalid; `FillColorPicker` is valid. Use ONLY
     IDs from the vetted list in `ribbon/imageMso_verified.md`, and add newly confirmed
     IDs there as you verify them. **One invalid ID raises a Custom UI Runtime Error**
     and can take the whole tab down.
   - Enable **File → Options → Advanced → "Show add-in user interface errors"** so
     ribbon faults show a dialog with line/column instead of failing silently. (Also
     documented in README.)
   - **PowerPoint caches the ribbon.** After re-injecting the XML, ALL PowerPoint
     windows must be fully closed and reopened, or the old ribbon persists.
6. **Source-of-truth workflow:** because programmatic VBA import is blocked, the repo's
   `src/*.bas` (and `.cls`) files are the source of truth, and the developer round-trips
   them through the VBA editor manually:
   - **Load:** VBA editor → File → Import File… for each changed module (Remove the old
     copy of the module first — import does not replace).
   - **Save back:** VBA editor → select module → File → Export File… into `src/`.
   README spells this out step by step for the full build loop.
7. **`git push` over the proxy can fail.** Keep the repo fully self-contained (no
   submodules, no fetch-at-build-time anything); pushing may require Azure DevOps Repos
   or a zip-based workaround.

## Repository layout

```
CLAUDE.md                    this file
README.md                    non-developer install & usage guide
docs/research.md             approved feature list (Phase 0)
docs/testing-checklist.md    per-feature manual test checklist (no automated rig exists)
src/*.bas, src/*.cls         VBA source of truth (manually imported/exported)
ribbon/customUI14.xml        ribbon definition (injected, never edited inside the ppam)
ribbon/imageMso_verified.md  vetted PowerPoint imageMso IDs — use only these
build/inject_ribbon.py       stdlib-only ribbon injector (zipfile + ElementTree)
build/build.bat              delete old ppam → inject → print restart reminder
```

### Module map

| Module | Scope |
|---|---|
| `modCommon.bas` | selection guards, cm/pt conversion, input helpers, theme access — **everything uses this** |
| `modSizing.bas` | all-like-largest/smallest sizing; stretch-edge-to-reference |
| `modArrange.bas` | distribute (fixed gap), stack/abut, grid arrange |
| `modRibbon.bas` | shared ribbon plumbing only: `Ribbon_OnLoad` and the shared `gRibbon` (IRibbonUI) reference. Feature callbacks live in their own feature modules. (Pruned in the ribbon trim: the colour-gallery getters, align-mode toggle and sticky-painter toggle callbacks and their swatch-bitmap helpers were removed with their controls.) |
| `modSelect.bas` | select-same (fill/font/size/type) |
| `modText.bas` | deck proofing language (interactive + one-click EN / FR presets), text cleanup, margins/wrap |
| `modClean.bas` | notes removal, unused masters/layouts, shape inspector |
| `modAudit.bas` | deck statistics |

## VBA coding conventions (enforced)

- `Option Explicit` in every module.
- Every user-facing macro is a `Public Sub Name(control As Object)` (ribbon-callable).
  Tag-dispatched subs read `control.Tag` to select a variant (e.g. align L/R/T/B).
- Module-level `Const`/`Dim`/`Public`/`Private` declarations at the very top of each
  module, before any procedure (see constraint 4).
- **Guard every selection** via `modCommon.GuardShapes(minCount, opName)`; it MsgBoxes a
  clear message and returns `Nothing` on bad selections. Never fail silently, never
  throw an unhandled error — every entry point has `On Error` → friendly MsgBox.
- **Reference-shape convention: the FIRST shape in the selection is the reference.**
  Documented in README and in every relevant screentip.
- **Theme colours via `ObjectThemeColor`** (`msoThemeColorAccent1` …), never hardcoded
  RGB, so colours track theme changes. Two clearly-marked exceptions:
  1. Tint/shade **gallery swatch images** use RGB calculated live from the theme (the
     applied colour still uses `ObjectThemeColor` + `ColorFormat.Brightness`, which
     stays theme-linked).
  2. Brand-fixed status colours (RAG traffic lights in `modVisuals`) — commented as such.
- Distances shown to the user are in **cm**; the object model works in points
  (`modCommon.CmToPt`/`PtToCm`, 28.3465 pt/cm). Number input accepts `,` or `.` decimals.
- Fragile techniques (Application events, WithEvents) are quarantined in
  `modPainter.bas` + `clsAppEvents.cls` so a failure there cannot break anything else.
  Keep it that way.

## Build & install loop (summary — full detail in README)

1. Open `SamSam.pptm` working file (or any pptm), VBA editor, import changed modules
   from `src/` (remove old copies first).
2. Debug → Compile VBAProject (success = menu item greys out).
3. File → Save As → type **PowerPoint Add-in (*.ppam)** → name `SamSamTools_src.ppam`.
   PowerPoint silently writes it to `%AppData%\Microsoft\AddIns` regardless of the
   folder shown (constraint 3).
4. **Close ALL PowerPoint windows** (ribbon cache + file lock).
5. Run `build\build.bat` — deletes the old injected `SamSamTools.ppam`, injects
   `ribbon/customUI14.xml` into a copy of `SamSamTools_src.ppam`, writes
   `%AppData%\Microsoft\AddIns\SamSamTools.ppam`, prints the restart reminder.
6. Reopen PowerPoint. First time only: File → Options → Add-ins → Manage
   "PowerPoint Add-ins" → Go… → Add New → pick `SamSamTools.ppam`.
7. Test against `docs/testing-checklist.md`. There is **no automated test rig** — the
   locked-down PowerPoint is the only test environment.

## Explicit exclusions (do not build)

- Any Excel↔PowerPoint or PowerPoint↔Word integration, table linking, or data
  import/export.
- Any icon, pictogram, stock-photo, template, or diagram **library**. (Generated-geometry
  micro-visuals like the Harvey balls, drawn from native shapes at insert time, are in
  scope — they are not a stored asset library.)
- Anything requiring a signing certificate, VSTO/COM registration, Office.js, admin
  rights, or a network dependency at runtime.
- Keyboard hooks / global shortcuts: VBA cannot register them. Users pin ribbon buttons
  to the Quick Access Toolbar for Alt+number access instead (documented in README).
