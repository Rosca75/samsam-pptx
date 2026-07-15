# Manual test checklist — SamSam Tools

There is no automated test rig: the locked-down PowerPoint is the only test
environment. Work through this list top-to-bottom after every build. Prep a scratch deck
with: 3+ slides, a couple of sections, mixed shapes (rectangle, rounded rectangle, oval,
text boxes, a table, a group), some text in FR **and** EN, one shape dragged fully off
the slide, and one shape filled with a hard-coded RGB colour.

## 0. Build & load (do this first, every time)

- [ ] `build.bat` runs to completion and prints the restart reminder.
- [ ] Fully restarted PowerPoint → **SamSam** tab appears.
- [ ] "Show add-in user interface errors" is ON and NO Custom UI Runtime Error appeared.
      If one appears, follow the triage in `ribbon/imageMso_verified.md`, then promote
      surviving IDs to `confirmed`.
- [ ] VBA editor → Debug → Compile VBAProject greys out (clean compile).
- [ ] Click any button with NOTHING selected → friendly MsgBox, not a VBA error.

## 1. Sizing

- [ ] Match width / height / size: select small shape FIRST, then two others → others take
      the first shape's dimension(s); top-left corners stay.
- [ ] Match size on a shape with **locked aspect ratio** → both dimensions still correct.
- [ ] All like largest / smallest (3 mixed shapes).
- [ ] Set exact size: width 5, height empty → widths 5 cm, heights unchanged.
- [ ] Scale by % → 50 then 200 returns roughly the original, centres unchanged.
- [ ] Stretch right edge to reference: target's left edge unchanged, right edge = ref's.

## 2. Position

- [ ] Toggle PRESSED (align to slide): Left/Right/Top/Bottom hit the slide edges;
      Centre/Middle centre on the slide.
- [ ] Toggle NOT pressed: with 2+ shapes, others align to the FIRST-selected shape.
      With 1 shape → clear "needs 2 shapes" message.
- [ ] Centre on slide H / V / both.
- [ ] Nudge in all 4 directions; set step to 0.5 → nudge distance visibly changes.
- [ ] Set exact X/Y: X 2 / Y empty → left edges at 2 cm, tops unchanged.
- [ ] Pick up pos+size on slide 1's title; Apply pos+size to slide 2's title → identical
      placement (check with Inspector).
- [ ] Apply with nothing picked up (fresh session) → clear message, no error.
- [ ] Swap: exactly 2 shapes swap centres; 3 shapes → clear "exactly two" message.

## 3. Arrange

- [ ] Distribute H with 3 shapes → equal gaps, outer two unmoved; V likewise.
- [ ] Distribute H with 2 shapes → clear "needs 3" message.
- [ ] Fixed gap 1 cm H and V → measure a gap with the Inspector (X₂ − X₁ − W₁ = 1 cm).
- [ ] Stack horizontally / vertically → shapes abut, no gap.
- [ ] Grid, 6 shapes, 3 columns, 0.2 gap → 2 rows × 3 cols anchored at the reference.

## 4. Font

- [ ] Heading/Body font on mixed text → font changes; after switching the deck THEME,
      the text follows the new theme fonts (theme linkage held).
- [ ] Size buttons 10–24 and exact size.
- [ ] Grow/Shrink on a text box with mixed sizes → each run steps ±2, differences kept.
- [ ] Match reference font: name/size/bold/italic/colour copied; a theme-coloured
      reference stays theme-linked on targets (change theme to confirm).

## 5. Theme colours (flagship)

- [ ] Open each gallery (Fill / Line / Text / Fill+Line): 10 columns × 6 rows of
      swatches showing the LIVE theme (compare with PowerPoint's own colour picker).
- [ ] Apply base accent to fill; then darker 25 % row → visibly darker same hue.
- [ ] **Theme-tracking proof:** colour two shapes from the gallery, switch the deck's
      theme/variant → both shapes recolour with the theme.
- [ ] Text gallery on selected text shape; Fill+Line paints both.
- [ ] Apply to a group and to a table → members/cells coloured, no error.
- [ ] Switch to a deck with a different theme → galleries show STALE colours until
      **Refresh palette** → then correct. (Known caching behaviour.)
- [ ] No fill / No line buttons.
- [ ] FALLBACK CHECK (only if galleries misbehave — blank swatches, runtime errors):
      swap the gallery XML for fixed buttons calling `ApplyThemeColorTagged` with tags
      like `fill:5` (see modColor header), rebuild, re-test.

## 6. Select

- [ ] Same fill: two accent-filled + one RGB-filled → only matching ones selected.
- [ ] Same font / same size / same shape type.
- [ ] Off-slide here: selects only the parked shape; slide with none → "No off-slide
      shapes" message.
- [ ] Off-slide report: finds parked shapes across the deck with slide numbers.

## 7. Text

- [ ] Deck language FR then EN-GB → spell-check squiggles change on slides AND notes;
      message reports frame count. Bad code ("XX") → clear error.
- [ ] Clean text: "a␣␣b␣␣␣" and trailing spaces removed; bold/colour formatting of the
      surrounding text SURVIVES (critical — the cleanup must not flatten runs).
- [ ] Delete empty text boxes: empty box goes, placeholder and filled boxes stay.
- [ ] Margins none/normal (check Format Shape → Text Options); word-wrap toggle.

## 8. Clean

- [ ] Remove notes: confirmation, then all notes empty, count reported. Cancel works.
- [ ] Unused masters: add an extra unused master/layout first → removed and reported;
      used ones survive.
- [ ] Inspector: values in cm match the Format pane (mind unit rounding); fill line
      says "theme-linked" vs "does NOT track the theme" correctly.

## 9. Insert (visuals)

- [ ] Harvey balls 0/25/50/75/100 at slide centre; 25/50/75 are ring+wedge groups;
      wedge starts at 12 o'clock, sweeps clockwise.
- [ ] Theme change → Harvey balls recolour (theme-linked); RAG dots do NOT (by design).

## 10. Guides

- [ ] **FIRST TEST — position convention:** build a grid with margin 1 cm, 1 column.
      If guides sit 1 cm from the slide EDGES → convention confirmed, tick this box.
      If they sit mirrored around the centre → fix `OffsetGuidePos` in modGuides
      (single adjustment point) and note it here.
- [ ] Grid 3 columns, 0.4 gutter → column boundary pairs correct.
- [ ] Guides from selection → 4 guides on the bounding box.
- [ ] Clear guides removes them all and reports the count.

## 11. Agenda

- [ ] Deck with 3 sections → Build creates one divider per section: title = section
      name, list shows all sections, current one bold/accent.
- [ ] Rename a section, Build again → dividers replaced (no duplicates, count stable).
- [ ] Remove dividers → deck back to original slide count.
- [ ] Deck without sections → clear message.

## 12. Painter (most fragile — test last, test hard)

- [ ] Pick up format from a styled shape; Apply to shapes on another slide → fill,
      line, effects, text formatting copied.
- [ ] Apply with nothing picked up → clear message.
- [ ] Sticky ON with a source selected → clicking other shapes reformats them on
      selection; works across slides.
- [ ] Sticky OFF stops it immediately; toggle state matches reality.
- [ ] Sticky ON with nothing selected → warning, toggle pops back out.
- [ ] Kill-switch: while sticky is ON, do something odd (click into a table cell,
      master view) → painter either applies sanely or SWITCHES ITSELF OFF with a
      message — never a repeating error loop.
- [ ] Match adjustments between two rounded rectangles; different-type shape reported
      as skipped. Corner radius 0.2 cm on mixed sizes → radii look identical.

## 13. Audit

- [ ] Deck statistics numbers are plausible (slides, shapes, words, fonts, hidden,
      off-slide).
- [ ] Off-theme report on the scratch deck → finds the RGB-filled shape with slide
      number and hex; a fully theme-linked deck reports clean.
- [ ] A colour set to the exact RGB of an accent (via the native picker's "More
      Colors") → flagged as "DETACHED".
- [ ] >25 findings → report slide at deck end; rerun replaces it; Remove report slide
      works.
- [ ] Select off-theme on this slide → offenders selected; recolour via gallery →
      rerun report → slide now clean.

## 14. Regression finale

- [ ] Compile still clean; no stray MsgBoxes during normal slide editing (events off).
- [ ] Close and reopen PowerPoint: add-in loads (macro prompt or trusted location),
      tab present, nudge step and align toggle back to defaults (session state resets
      by design).
