# SamSam Tools — PowerPoint productivity add-in

A ribbon tab full of everyday slide-editing power tools: precise sizing, alignment and
distribution, one-click theme colours that keep tracking the theme, font shortcuts,
selection superpowers, deck cleanup, Harvey balls, agenda slides, a sticky format
painter, and a deck colour audit. Pure VBA — no installation rights, no internet, no
signed certificate needed.

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
Alt+2, …** shortcuts — pin your favourites (Align, Match Size, theme Fill) there.

---

## What's on the tab (quick tour)

- **Sizing** — match width/height/both to the first-selected shape, size everything to
  the largest/smallest, set exact cm sizes, scale by %, stretch an edge to the
  reference's edge.
- **Position** — align left/centre/right/top/middle/bottom with a **toggle** that
  switches between *align to slide* and *align to selection* (= to the first-selected
  shape); centre on slide; nudge by a precise step you choose; set exact X/Y; **pick up
  position & size from one shape and apply it to shapes on any slide** (great for making
  every title sit identically); swap two shapes.
- **Arrange** — distribute with equal gaps, distribute with an exact gap in cm, stack
  shapes with zero gap, lay a selection out in a grid of N columns.
- **Font** — theme heading/body font, one-click sizes, grow/shrink across mixed
  selections, match the reference shape's font.
- **Theme colours** (the flagship) — live palettes of the *current* presentation's
  theme colours (plus lighter/darker rows) applied to **Fill**, **Line**, **Text**, or
  fill+line together. Colours are applied as *theme* colours, so if the deck's theme
  changes later, your shapes follow it. Click **Refresh palette** after switching to a
  presentation with a different theme.
- **Select** — select everything on the slide with the same fill / font / size / type as
  the first-selected shape; find shapes parked off the slide.
- **Text** — set the proofing language for the WHOLE deck (incl. masters and notes) in
  one go; remove double/trailing spaces; delete empty text boxes; margins on/off;
  word-wrap toggle.
- **Clean** — delete all speaker notes; delete unused slide masters/layouts; a shape
  inspector read-out (position/size in cm).
- **Insert** — Harvey balls (0–100 %) and RAG status dots drawn from native shapes in
  theme colours.
- **Guides** — build a margin/column guide grid; make guides from the selection; clear
  all guides.
- **Agenda** — generate section-divider slides from your PowerPoint sections (rerun to
  update; they're tagged so rebuilding replaces them cleanly).
- **Painter** — pick up a shape's full format and apply it anywhere, repeatedly; a
  **Sticky painter** toggle that formats every shape you click until you switch it off;
  copy corner radius / shape adjustments across shapes.
- **Audit** — deck statistics, and an **off-theme colour report** that lists every
  hard-coded colour that won't follow the theme (long reports are written onto a report
  slide at the end of the deck).

---

## Building the add-in (developer / power user)

You need: this repo on local disk, PowerPoint, and Python 3 (standard install — the
build uses only the standard library, no pip, no network).

### First-time setup of the working file

1. Open PowerPoint, new blank presentation, **Alt+F11** to open the VBA editor.
2. **File → Import File…** and import every `src/*.bas` file and `src/clsAppEvents.cls`.
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
