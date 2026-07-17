# Manual test checklist — SamSam Tools

There is no automated test rig: the locked-down PowerPoint is the only test
environment. Work through this list top-to-bottom after every build. Prep a scratch deck
with: 3+ slides, mixed shapes (rectangle, rounded rectangle, oval, text boxes, a table, a
group), and some text in FR **and** EN.

## 0. Build & load (do this first, every time)

- [ ] `build.bat` runs to completion and prints the restart reminder.
- [ ] **Fully close ALL PowerPoint windows, then reopen** (ribbon cache + file lock) →
      **SamSam** tab appears.
- [ ] "Show add-in user interface errors" is ON and **NO Custom UI Runtime Error dialog**
      appeared. If one appears, it names the offending line/column in `customUI14.xml` —
      follow the triage in `ribbon/imageMso_verified.md` (replace the suspect `imageMso`
      with `HappyFace`, rebuild, restart, repeat), then promote surviving IDs to
      `confirmed`.
- [ ] VBA editor → Debug → Compile VBAProject greys out (clean compile).
- [ ] Click any button with NOTHING selected → friendly MsgBox, not a VBA error.

## 1. Icons render (candidate imageMso — first-load check)

Every ID below is a **candidate** in `ribbon/imageMso_verified.md` — never yet seen in the
target PowerPoint. Confirm each control shows a real icon (not blank, no runtime error);
promote the survivors to `confirmed`.

- [ ] Sizing → More sizing menu: **All like largest** (`Consolidate`) and **All like
      smallest** (`SmartArtSmallerShape`) icons render.
- [ ] Sizing → More sizing menu: the four **stretch-edge** icons render — left (`FillLeft`),
      right (`FillRight`), top (`FillUp`), bottom (`FillDown`).
- [ ] Arrange: **Fixed gap horizontal** (`ObjectsAlignDistributeHorizontallyRe`) and
      **Fixed gap vertical** (`ObjectsAlignDistributeVerticallyRemo`) icons render.
      ⚠️ These two IDs look truncated — if either shows blank or errors, re-check the full
      ID in the Office RibbonX Editor Icon Gallery (do **not** guess the suffix).
- [ ] Arrange: **Stack horizontally** (`HorizontalSpacingDecrease`) and **Stack vertically**
      (`VerticalSpacingDecrease`) icons render.
- [ ] Arrange: **Arrange in grid** (`TableShowGridlines`) icon renders.
- [ ] Clean: **Inspector** (`RulesToCheck`) icon renders.
- [ ] Text: **EN** (`EnglishWritingAssistant`) and **FR** (`GetPowerQueryDataFromWeb`)
      icons render.

## 2. Sizing

- [ ] All like largest / smallest with 3 mixed shapes → every shape takes the
      largest/smallest (by area) shape's dimensions.
- [ ] Stretch right edge to reference (first shape = reference): target's left edge
      unchanged, right edge = reference's right edge. Repeat for left / top / bottom.

## 3. Arrange

- [ ] Fixed gap 1 cm H and V → measure a gap with the Inspector (X₂ − X₁ − W₁ = 1 cm).
- [ ] Stack horizontally / vertically → shapes abut, no gap.
- [ ] Grid, 6 shapes, 3 columns, 0.2 gap → 2 rows × 3 cols anchored at the reference.

## 4. Select

- [ ] Same fill: two accent-filled + one RGB-filled → only matching ones selected.
- [ ] Same font / same size / same shape type.

## 5. Text

- [ ] **EN button** → one click sets the WHOLE deck's proofing language to English (UK):
      spell-check squiggles change on slides AND notes; message reports the frame count.
- [ ] **FR button** → one click sets the WHOLE deck's proofing language to French (FR),
      same coverage and message.
- [ ] Deck language (interactive): FR then EN-GB via the InputBox → same effect as the
      preset buttons; a bad code ("XX") → clear error, no change.
- [ ] Clean text: "a␣␣b␣␣␣" and trailing spaces removed; bold/colour formatting of the
      surrounding text SURVIVES (critical — the cleanup must not flatten runs).
- [ ] Delete empty text boxes: empty box goes, placeholder and filled boxes stay.
- [ ] Margins none/normal (check Format Shape → Text Options); word-wrap toggle.

## 6. Clean

- [ ] Remove notes: confirmation, then all notes empty, count reported. Cancel works.
- [ ] Unused masters: add an extra unused master/layout first → removed and reported;
      used ones survive.
- [ ] Inspector: values in cm match the Format pane (mind unit rounding).

## 7. Audit

- [ ] Deck statistics numbers are plausible (slides, shapes, words, fonts, hidden,
      off-slide).

## 8. Regression finale

- [ ] Compile still clean; no stray MsgBoxes during normal slide editing.
- [ ] Close and reopen PowerPoint: add-in loads (macro prompt or trusted location),
      tab present, no Custom UI Runtime Error.
