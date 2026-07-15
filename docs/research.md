# Phase 0 Research — PowerPoint Productivity Add-in Feature Inventory

Research date: 2026-07-15. All feature sets verified against the vendors' live websites,
user guides, and support documentation (sources listed at the end) — not from memory.

## Scope and method

Seven reference products were catalogued in depth: **PPT Productivity**, **Power-user**,
**ToolsToo Pro**, **Efficient Elements for presentations**, **BrightSlide** (BrightCarbon,
free), **Nuts & Bolts**, and **Slidewise** (Neuxpower). A wider discovery sweep also covered
**Instrumenta** (iappyx, open source), **PPTools/THOR** (Steve Rindsberg), **PowerPointLabs**
(NUS), **YOUtools/YOUpresent**, **empower**, **QuickSlide**, **Templafy SlideProof**, and
free-floating power-user VBA macros from pptfaq.com, forums, and blogs.

Two findings frame everything below:

1. **Nuts & Bolts ships no add-in.** The site sells training only; its "tools" content
   recommends third parties (chiefly PPTools' free THOR hammer tool and Shyam Pillai's
   motion-path tools). Its recommendations are folded into the inventory where relevant.
2. **Instrumenta proves the whole category is buildable in pure VBA.** It is MIT-licensed,
   ~97 % VBA, with 270+ features spanning nearly everything the commercial suites do
   (github.com/iappyx/Instrumenta). It is a legal reference implementation for this project.

### Ratings

- **Value** (high/med/low): benefit to a heavy slide editor doing daily deck work.
- **Effort** (low/med/high): implementation cost **under our VBA constraints** — unsigned
  `.ppam`, ribbon-only UI (no task panes), no network, no programmatic VBA import, no
  keyboard hooks (VBA cannot register global shortcuts without Win32 hacks — every feature
  must be reachable from a ribbon button).

### Industry conventions worth adopting

- **Reference shape = first-selected** is the dominant convention (PPT Productivity,
  Power-user, PowerPointLabs). Efficient Elements uses last-selected ("Master shape");
  ToolsToo lets the user choose among four reference modes (first shape / picked-up shape /
  slide / user-defined region). We adopt **first-selected**, per the brief.
- **Align-to-slide vs align-to-selection** is exposed as an explicit, persistent toggle in
  BrightSlide and PPT Productivity — the model to copy.
- Every serious suite has a **"pick up / apply position+size" tool** (THOR's hammer, PPT
  Productivity's Alt+C/Alt+V, Power-user Smart Painter, BrightSlide Match Tools). It is
  arguably the single most-loved feature in the category.

---

## 1. Shape sizing

| Feature | Who ships it | Value | Effort |
|---|---|---|---|
| Match width / height / both to reference shape | All suites | high | low |
| Make all shapes size of largest / smallest in selection | ToolsToo, Instrumenta | high | low |
| Set exact width/height numerically (input box) | All (via panes) | high | low |
| Scale selection by percentage (e.g. 90 % / 110 % / custom) | PPT Productivity, EE "Magic Resizer" | high | low |
| Pick up position+size from one shape, apply to any shape on any slide (THOR pattern) | PPTools THOR, PPT Productivity, Power-user, BrightSlide | high | low-med |
| Stretch shape edge (L/R/T/B) to reference shape's edge, same-side or opposite-side | ToolsToo, EE, PPT Productivity | high | med |
| Fill gap: extend shape until it abuts an adjacent shape | PPT Productivity, EE "Fill/Dock" | med | med |
| Make shape square / circle (set height = width or vice versa) | ToolsToo "Scale to Height/Width" | med | low |
| Slice a shape into an N×M grid of mini-shapes in its footprint | PPT Productivity, EE, BrightSlide "Split & Align" | med | med |
| Multiply/replicate a shape into rows × columns with gaps | PPT Productivity, ToolsToo "Replicate" | med | med |
| Toggle lock-aspect-ratio in bulk | ToolsToo, Instrumenta | med | low |
| Fit shape to its text / toggle autofit modes in bulk | PPT Productivity, ToolsToo, Slidewise | med | low |
| Reset picture scale to 100 % / match picture scale or crop | ToolsToo | low | med |

## 2. Shape positioning

| Feature | Who ships it | Value | Effort |
|---|---|---|---|
| Align L/R/T/B/centre/middle with explicit align-to-slide vs align-to-selection toggle | BrightSlide, PPT Productivity, ToolsToo | high | low |
| Align to reference (first-selected) shape rather than selection bounds | PPT Productivity, EE, ToolsToo, BrightSlide | high | low |
| Centre on slide horizontally / vertically / both in one click | All | high | low |
| Swap positions of two shapes (centre-anchored; anchor selectable in ToolsToo/Power-user) | All suites | high | low |
| Swap variants: horizontal-only, vertical-only, also swap z-order | PPT Productivity, EE, BrightSlide | med | low |
| Nudge by a precise, user-set increment | (native arrow keys are grid-bound) | high | low |
| Set exact X/Y numerically | All (via panes) | high | low |
| Pick up position from one shape, apply to shapes on other slides (e.g. unify title positions) | THOR, Power-user, PPT Productivity | high | low |
| Straighten nearly-horizontal/vertical lines and connectors | PPT Productivity, Power-user, ToolsToo, BrightSlide | med | low |
| Copy an arrangement of N shapes onto another set of N shapes | PPT Productivity | med | med |
| Symmetry: mirror selection around slide centre axis | Power-user | low | med |
| Dock: move shapes in a direction until they touch the reference | EE | med | med |
| Golden-canon vertical placement inside a reference | EE | low | low |

## 3. Shape arrangement

| Feature | Who ships it | Value | Effort |
|---|---|---|---|
| Distribute horizontally / vertically with equal gaps | All (native is weaker) | high | low |
| Distribute with a user-specified fixed gap value | PPT Productivity, ToolsToo, BrightSlide, Instrumenta | high | low-med |
| Increase / decrease gaps incrementally (squeeze/expand spacing) | Power-user, PPT Productivity, ToolsToo | med | med |
| Remove gaps: stack/abut shapes with zero spacing (L/R/T/B) | Power-user "touch align", ToolsToo "Adjoin", EE "Stack" | high | low-med |
| Arrange selection into a grid of N columns (row/col gaps) | BrightSlide "Distribute to Grid", Power-user, ToolsToo | high | med |
| Pick up an existing gap and re-apply it elsewhere | ToolsToo | med | med |
| Group by row / group by column automatically | PPT Productivity, Instrumenta | med | med |
| Re-group (restore previous grouping) / add to group without breaking it | PPT Productivity, BrightSlide, ToolsToo | med | med-high |
| Distribute radially (equal angles around a point); duplicate-and-rotate | PowerPointLabs | low | med |
| Single out: one selected object per new slide | EE, PPT Productivity | low | low |
| Z-order tools beyond native (rearrange sequence dialog) | ToolsToo | low | med |

## 4. Fast font selection and sizing

| Feature | Who ships it | Value | Effort |
|---|---|---|---|
| Apply theme heading / body font to selection | Power-user, Slidewise (via replace), Instrumenta | high | low |
| One-click common size buttons (e.g. 10/12/14/18/24 pt) | (QAT patterns; various) | high | low |
| Increase/decrease font size stepwise across a mixed selection | PPT Productivity (2 pt steps), Instrumenta | high | low |
| Match font (name/size/colour/style) of reference shape | ToolsToo "Make Same Font", Power-user | high | low |
| Toggle text-box margins: none / normal (bulk) | PPT Productivity, EE, BrightSlide | high | low |
| Toggle word-wrap / autofit in bulk | PPT Productivity, Power-user, Slidewise | med | low |
| Merge multiple text boxes into one (selection order) | PPT Productivity, EE, BrightSlide, ToolsToo | med | med |
| Split text box into one shape per paragraph | Same four | med | med |
| Change/cycle case across selection | PPT Productivity, PPTools | med | low |
| Line/paragraph spacing steppers | BrightSlide (live sliders), Instrumenta | med | low |
| Autofit text to one line at current width | ToolsToo | low | med |
| Colour all bold text in one click (bold-as-highlight pattern) | PPT Productivity, Instrumenta | low | low |
| Saved paragraph/text styles applied by number | PPT Productivity, PPTools ShapeStyles | med | high |

## 5. Quick theme-colour selection and application (flagship)

| Feature | Who ships it | Value | Effort |
|---|---|---|---|
| Live theme palette in the ribbon (reads active colour scheme) | PPT Productivity palette toolbar, EE Color Bar | high | med |
| Apply theme colour to fill / line / text as separate targets | PPT Productivity, EE | high | low |
| Apply via `ObjectThemeColor` so colours track theme changes | (differentiator — most tools burn RGB) | high | low |
| Tints & shades of each theme colour (auto-generated rows) | PPT Productivity (3 lighter/2 darker) | med | med |
| Cycle a shape's fill through the accent colours by repeated click | PPT Productivity | med | low |
| Apply one colour to fill AND line together | PPT Productivity "set line with fill" | med | low |
| Convert theme colours → fixed RGB in selection (detach) and RGB → theme (re-attach) | EE | med | med |
| Report non-theme (off-palette) colours used in the deck | Slidewise, Power-user, PPT Productivity "Fix Colors" | high | med |
| Replace one colour deck-wide (fills/lines/text), incl. tint handling | Power-user, Slidewise, Instrumenta | high | med-high |
| Screen eyedropper (pick any pixel) | PPT Productivity, EE, BrightSlide | med | high (Win32 `GetPixel` declares; offline but fiddly) |
| Colour-contrast checker between two shapes (accessibility) | BrightSlide | low | med |

---

## Discovered — beyond the core

Standalone-VBA-viable features from outside the five core areas that earn consideration.
Grouped by theme; the ★ marks the candidates recommended for v1 or v1.1.

### Selection superpowers

| Feature | Origin | Value | Effort |
|---|---|---|---|
| ★ Select all shapes matching the reference: same fill / line colour / font / size / shape type | ToolsToo "Select Same", PPT Productivity, BrightSlide, YOUtools | high | low |
| ★ Select all off-slide shapes ("gutter junk" finder) | community macros | high | low |
| Hide selected / hide all but selected / show all hidden shapes | ToolsToo, EE | med | low |
| Invert selection on slide | community macros | med | low |
| Saved named selections re-selectable later | PPTools Selection Manager | low | med |

### Format painting with memory

| Feature | Origin | Value | Effort |
|---|---|---|---|
| ★ Pick up full format of a shape; apply to any shapes, repeatedly, across slides | Power-user Smart Painter, PPT Productivity painters, BrightSlide Multi-painter | high | med |
| Selective apply (fill only / line only / text format only / position only / size only) | Power-user, PowerPointLabs Sync Lab | high | med |
| Copy shape adjustment-handle values (corner radius, chevron angle) to same-type shapes | ToolsToo "Make Same Adjustments", PPT Productivity "Angles & Curves" | med | low |
| ★ Make same rounded-corner radius across selection | ToolsToo, EE, YOUtools | med | low |
| Swap the complete formats of two shapes | ToolsToo "Swap Formatting" | low | low |
| Make same rotation / snap rotation to increments | ToolsToo | low | low |

### Text hygiene and deck-wide text ops

| Feature | Origin | Value | Effort |
|---|---|---|---|
| ★ Set proofing language for the whole deck incl. masters/layouts/notes | ToolsToo, BrightSlide, EE, Slidewise, Instrumenta — universally shipped because natively painful | high | low |
| ★ Remove double spaces / trailing spaces / empty text boxes deck-wide | Power-user Cleaner, PPT Productivity proofing, community | med | low |
| Remove all character formatting (or strikethrough) from selection | Instrumenta | low | low |
| Swap the text of two shapes (formats and positions stay) | Power-user, Instrumenta | med | low |
| Deck-wide find & replace that also covers notes/masters/tables | community macros | med | med |
| Anonymise deck: replace all text with lorem ipsum | Instrumenta | low | low |

### Deck housekeeping / cleanup

| Feature | Origin | Value | Effort |
|---|---|---|---|
| ★ Remove all speaker notes (send-out prep) | ToolsToo, EE, Slidewise, community | high | low |
| ★ Delete unused slide masters / layouts | ToolsToo, Slidewise, Power-user, community | high | low-med |
| Unhide (or list) all hidden slides and hidden shapes | Slidewise, ToolsToo, Instrumenta | med | low |
| Remove all animations / transitions in one click | ToolsToo, EE, Power-user, BrightSlide | med | low |
| Remove empty placeholders / empty sections | ToolsToo, Power-user, Slidewise | med | low |
| Deck statistics: slides, shapes, words, fonts in use, hidden object census | ToolsToo, community | med | med |

### Insertable micro-visuals (drawn from native shapes — not an asset library)

| Feature | Origin | Value | Effort |
|---|---|---|---|
| ★ Harvey balls (0/25/50/75/100 % pie shapes, theme-coloured) | PPT Productivity, EE Smart Elements, Instrumenta | med | low |
| Traffic-light / RAG status indicators | PPT Productivity, EE, Instrumenta | med | low |
| DRAFT / CONFIDENTIAL stamp on all slides (tagged, cleanly removable) | Power-user, PPT Productivity, QuickSlide | med | low |
| Progress bar/pie per slide showing position in deck | Power-user | low | med |
| Star ratings, numbered badges | Instrumenta | low | low |

*(These are generated geometry, parameterised at insert time — they do not conflict with the
"no asset/object libraries" exclusion, which targets stored icon/template galleries.)*

### Guides, agenda, navigation

| Feature | Origin | Value | Effort |
|---|---|---|---|
| Build a precise guide grid (margins/columns/rows) programmatically | BrightSlide, YOUtools | med | low-med |
| Make guides from selection bounds; delete all guides | BrightSlide, ToolsToo | med | low |
| Agenda/section-divider slides generated from PowerPoint sections, updatable | ToolsToo, EE Agenda Wizard, PPT Productivity, PowerPointLabs | med | high |
| Send slide to appendix (move to end behind a divider) | PPT Productivity | low | med |
| Position/size read-out of the selected shape (inspector message) | pptfaq, Slidewise Inspector | med | low |

### Tables (pure-PowerPoint table UX gaps)

| Feature | Origin | Value | Effort |
|---|---|---|---|
| Equalise column widths / row heights across a table | ToolsToo, community | med | low-med |
| Move rows/columns; insert preserving widths | PPT Productivity, Instrumenta | med | med-high |
| Transpose table | ToolsToo, BrightSlide | low | high |

### Notable but rejected for scope

- **Multi-slide shape synchronisation** (edit once, propagate to tagged twins — Instrumenta):
  genuinely rare and clever, but state management via Tags is error-prone; revisit post-v1.
- **Deck linter with quick-fixes** (Templafy/empower category, proven by Instrumenta): the
  biggest "big" feature in the market, but it is a project of its own; revisit post-v1.
- **Screen-wide eyedropper**: works offline via Win32 `GetPixel`, but `Declare` statements and
  cursor hooks are exactly the kind of fragile code the locked-down environment punishes; the
  live theme palette removes most of the need.
- **Custom keyboard shortcuts**: VBA has no supported keyboard hook; every suite that offers
  them ships a compiled companion. We rely on ribbon + Quick Access Toolbar (users can pin any
  of our buttons to the QAT for Alt+n access — README will document this).
- **Live-preview sliders** (BrightSlide character-spacing): needs modeless UserForms with
  selection-change events; possible but fragile — steppers (buttons) achieve 90 % of it.
- **Excluded by the brief**: everything touching Excel/Word/Outlook, icon/template/asset
  libraries, cloud sync, AI features, font embedding and media internals (require reading the
  pptx ZIP — the reason Slidewise is a compiled add-in).

---

## Proposed v1 feature list (prioritised)

Everything below is ribbon-button-driven, pure object-model VBA, and respects the exclusions.
Order within each module is build order; the module order is the build order overall.
Convention throughout: **the first-selected shape is the reference**, and every operation
guards the selection with a clear message.

### Tier 1 — core (must ship)

**modSizing**
1. Match width / height / both to reference shape
2. Make all shapes the size of the largest / smallest in selection
3. Set exact width / height via input box (cm)
4. Scale selection by a percentage (input box; positions kept)

**modPositioning**
5. Align left / right / top / bottom / centre / middle — with a ribbon **toggle** for
   align-to-slide vs align-to-selection (selection mode aligns to the reference shape)
6. Centre on slide: horizontally, vertically, both
7. Nudge L/R/U/D by a precise increment (configurable step, input box to change)
8. Set exact X / Y via input box

**modArrange**
9. Distribute horizontally / vertically with equal gaps
10. Distribute with a user-specified fixed gap (H and V)
11. Arrange selection into a grid of N columns (prompt for N and gap)
12. Swap the positions of exactly two shapes (centre-anchored)

**modFont**
13. Apply theme heading font / body font to selection
14. One-click size buttons (10 / 12 / 14 / 18 / 24 pt) + set-exact-size input
15. Increase / decrease font size stepwise across mixed selections
16. Match font (name, size, bold/italic, colour) of reference shape

**modColor** (flagship)
17. Live theme palette: dynamic ribbon gallery reading the active scheme
    (`getItemCount`/`getItemImage`/`onAction`), one gallery each for **fill**, **line**,
    **text** targets; fallback design ready (fixed accent-button rows) if galleries misbehave
18. All application via `ObjectThemeColor` (msoThemeColorAccent1…6, dark/light 1–2) — never RGB
19. Apply to fill + line together ("paint whole shape")

### Tier 2 — discoveries that earned a place (ship in v1)

**modSizing / modPositioning extras**
20. **Pick up position & size / apply position & size** (THOR pattern) — the category's
    most-loved tool; also gives "make all titles sit identically across slides" for free
21. Stack/abut shapes with zero gap (left/top variants), as the natural sibling of
    fixed-gap distribute

**modSelect (new)**
22. Select all shapes on the slide with same fill / same font / same size / same type as
    the reference — multiplies the value of every other tool
23. Select all off-slide shapes (gutter junk finder)

**modText (new)**
24. Set proofing language for the whole deck (incl. masters and notes) — one input, huge pain
    removed (Luxembourg office reality: FR/DE/EN mixed decks)
25. Remove double/trailing spaces across the deck; delete empty text boxes
26. Toggle text-box margins none/normal on selection; toggle word-wrap

**modClean (new)**
27. Remove all speaker notes (with confirmation)
28. Delete unused slide masters/layouts (with report of what was removed)
29. Position/size read-out of the reference shape (inspector MsgBox, values in cm — doubles
    as a debugging aid)

### Tier 3 — deferred to v1.1 (documented, not built yet)

- Harvey balls / RAG status generators (theme-coloured native shapes)
- Make-same corner radius / adjustment values; stretch-edge-to-reference
- Guide-grid builder; make guides from selection
- Tints & shades rows in the colour galleries; off-theme colour report
- Full sticky format painter (needs Application-events plumbing)
- Agenda/section tools; deck statistics; multi-slide shape sync; deck linter

### Why this cut

- Tiers 1–2 are ~29 operations across 8 modules — enough to transform daily editing, small
  enough to build and manually test through the round-trip workflow in one iteration.
- Everything in Tier 1–2 is **low or low-med effort** except the colour gallery (med, and it
  has a designed fallback). Nothing depends on Win32 declares, application events, tags, or
  modeless forms — the fragile techniques are all quarantined in Tier 3.
- The three new modules (Select/Text/Clean) each cost little but cover the highest-value
  discoveries; they were chosen over flashier candidates (agenda tools, painters) because
  their failure modes are benign and their VBA is simple and testable.

---

## Sources

Read on 2026-07-15 (representative list; each product's feature/support pages were crawled
in depth):

- pptproductivity.com — /features and the create-faster / refine-easier feature pages
- efficient-elements.com — product pages + full ee4p handbook (PDF, 45 pp)
- powerusersoftwares.com — /features, /play-with-shapes, /smart-painter, /clean-presentation; support KB articles (alignment toolbar, Smart Painter, Cleaner, swap positions, progress tools)
- toolstoo.com — /toolset, /reference-shapes, /whats-new (4 pages)
- brightcarbon.com/brightslide — product page + full online user guide + feature blog posts
- nutsandboltsspeedtraining.com — /products, add-in round-up and THOR/motion-path tutorials
- neuxpower.com/slidewise-powerpoint-add-in + Neuxpower support centre (≈20 articles)
- github.com/iappyx/Instrumenta (MIT; full feature list)
- pptools.com (THOR, Selection Manager, StarterSet, ShapeStyles), rdpslides.com/pptfaq
- comp.nus.edu.sg/~pptlabs (PowerPointLabs docs), youpresent.co.uk (YOUtools)
- empowersuite.com, strategy-compass.com (QuickSlide), templafy.com (SlideProof checks)
- powerpointaddins.com directory; assorted community macro sources (brightcarbon blog,
  nullskull, avantixlearning, alexanderjarvis.com)
