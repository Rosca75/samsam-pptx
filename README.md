# SamSam Tools — PowerPoint productivity add-in

A ribbon tab of everyday slide-editing power tools: precise sizing, shape arranging,
selection superpowers, deck-wide proofing-language presets, text and deck cleanup, and
deck statistics. Pure VBA — no installation rights, no internet, no signed certificate
needed.

> **Golden rule of every tool:** when a tool works "like a reference shape", the
> **FIRST shape you selected is the reference**. Select the shape you want to copy
> *from* first, then Ctrl/Shift-click the shapes you want to change.

---

## Installing the add-in (one time)

The add-in is a single file, `SamSamTools.ppam`, that must be **built on your own
machine** (see "Building" below, or ask a developer colleague for the build steps —
**never** accept the file by email/download/Teams: Windows marks downloaded macro files
and PowerPoint will refuse to run them, with no way to enable them).

1. Build the file (it lands in `%AppData%\Microsoft\AddIns` — paste that into the
   Explorer address bar to see it).
2. In PowerPoint: **File → Options → Add-ins**, at the bottom set **Manage:
   PowerPoint Add-ins** → **Go…** → **Add New…** → pick `SamSamTools.ppam` → Close.
3. A **SamSam** tab appears on the ribbon.

### The macro prompt (and how to remove it)

Company policy is "Disable all macros with notification". That means **once per
PowerPoint session** you may get an *Enable Macros* security prompt — click Enable and
the tab works until you close PowerPoint.

To make it load silently every time, add the AddIns folder as a **Trusted Location**:
File → Options → Trust Center → Trust Center Settings → Trusted Locations →
**Add new location…** → paste `%AppData%\Microsoft\AddIns` → OK. (If the button is
greyed out, policy blocks it — you'll live with the once-per-session prompt.)

### One recommended setting

Turn on **File → Options → Advanced → General → "Show add-in user interface errors"**.
If the ribbon ever has a problem it will then tell you (with a line number) instead of
silently showing nothing.

### Keyboard shortcuts

VBA add-ins cannot register real keyboard shortcuts. Instead, **right-click any SamSam
button → Add to Quick Access Toolbar**. Buttons on the QAT get automatic **Alt+1,
Alt+2, …** shortcuts — pin your favourites (e.g. Fixed gap, Select same, EN / FR) there.

---

## What's on the tab (quick tour)

- **Sizing** — size everything to the largest/smallest shape in the selection, and
  stretch an edge (left/right/top/bottom) of every shape to the reference's matching edge.
- **Arrange** — distribute with an exact gap in cm (horizontal / vertical), stack shapes
  with zero gap (horizontal / vertical), and lay a selection out in a grid of N columns.
- **Select** — select everything on the slide with the same fill / font / size / type as
  the first-selected shape.
- **Text** — set the proofing language for the WHOLE deck (incl. masters, layouts and
  notes), either by typing a code or with the one-click **EN** (English UK) and **FR**
  (French) preset buttons; remove double/trailing spaces; delete empty text boxes;
  margins on/off; word-wrap toggle.
- **Clean** — delete all speaker notes; delete unused slide masters/layouts; a shape
  inspector read-out (position/size in cm).
- **Audit** — deck statistics (slides, shapes, words, fonts in use, hidden and off-slide
  object census).

---

## Building the add-in (developer / power user)

You need: this repo on local disk, PowerPoint, and Python 3 (standard install — the
build uses only the standard library, no pip, no network).

### First-time setup of the working file

1. Open PowerPoint, new blank presentation, **Alt+F11** to open the VBA editor.
2. **File → Import File…** and import every `src/*.bas` module. Always use
   **File → Import File…** — never copy-paste a file's text into a module. The `src/`
   files are stored with **CRLF (Windows) line endings** (enforced by `.gitattributes`);
   the VBA editor's importer needs CRLF to read the module header. If a module ever
   imports oddly, your `src/` copies may be LF — re-checkout so the CRLF from
   `.gitattributes` applies (`del /Q src\*.bas` then `git checkout -- src`).
3. **Debug → Compile VBAProject**. No message appears on success — the menu item just
   greys out. That's a clean compile.
4. Save as `SamSam.pptm` (macro-enabled) somewhere in the repo folder — this is your
   editing workbench.

### The build loop (every time you change code or the ribbon XML)

1. Edit VBA in the editor (or re-import changed `.bas` files — **remove the old module
   first**, import doesn't overwrite). If you edited in the editor, **export** the
   changed modules back to `src/` (File → Export File…) so the repo stays the source of
   truth.
2. **File → Save As → PowerPoint Add-in (*.ppam)** → filename `SamSamTools_src.ppam`.
   ⚠️ PowerPoint will *silently ignore* whatever folder you pick and write the file to
   `%AppData%\Microsoft\AddIns`. That's normal; the build script expects it there.
3. **Close ALL PowerPoint windows.** (Two reasons: the registered add-in file is locked
   while PowerPoint runs, and PowerPoint caches the ribbon — an open window would keep
   showing the old tab.)
4. Double-click `build\build.bat`. It deletes the old `SamSamTools.ppam`, injects
   `ribbon\customUI14.xml` into a copy of `SamSamTools_src.ppam`, and reminds you to
   restart.
5. Reopen PowerPoint. (First time only: register the add-in — see Installing above.)

### Rules of the road

- `src/` is the **source of truth**. PowerPoint cannot import/export modules
  automatically here ("Trust access to the VBA project object model" is locked off), so
  the import/export round-trip is manual — do it every time, or the repo drifts.
- Ribbon icons: only use `imageMso` IDs listed in `ribbon/imageMso_verified.md`. One
  invalid ID kills the tab with a Custom UI Runtime Error (visible only if you enabled
  "Show add-in user interface errors").
- After ANY ribbon XML change: full PowerPoint restart, always.
- Testing is manual: work through `docs/testing-checklist.md` in the real environment.
- `git push` may fail through the corporate proxy. The repo is self-contained; fall back
  to Azure DevOps Repos or zip the folder and move it by approved channels.
