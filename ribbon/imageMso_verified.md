# Vetted imageMso IDs for PowerPoint

`imageMso` icon IDs are **per-application** — an ID that works in Excel or Word may not
exist in PowerPoint. One invalid ID raises a **Custom UI Runtime Error** (only visible
when File → Options → Advanced → *Show add-in user interface errors* is enabled) and can
take down the whole tab.

**Rule: `customUI14.xml` may only use IDs from the table below.** When you want a new
icon: add it here first with status `candidate`, test it in the real PowerPoint, then
promote it to `confirmed` (or delete the row). If the tab ever fails with a Custom UI
Runtime Error, the fastest triage is to replace the suspect ID with `HappyFace`
(guaranteed everywhere) and re-test.

## Known-invalid (never use)

| ID | Note |
|---|---|
| `ShapeFill` | Empirically confirmed INVALID in PowerPoint (valid elsewhere) — the original discovery behind this file. |

## Vetted list

Status legend: **confirmed** = seen working in the target PowerPoint;
**high-confidence** = ID of a control that exists on PowerPoint's own ribbon or a
classic cross-app icon; verify on first load and promote to confirmed.

| ID | Used for | Status |
|---|---|---|
| `FillColorPicker` | colour galleries (fill) | confirmed |
| `HappyFace` | inspector, triage placeholder | high-confidence (classic cross-app) |
| `FontColorPicker` | colour galleries (line/text), audit | high-confidence (Home → Font Color) |
| `Bold` | heading font | high-confidence (Home tab) |
| `Italic` | body font | high-confidence (Home tab) |
| `GrowFont` | grow font | high-confidence (Home tab) |
| `ShrinkFont` | shrink font | high-confidence (Home tab) |
| `ChangeCaseGallery` | font size menu | high-confidence (Home tab) |
| `Copy` | match width/height, pick up | high-confidence (Home tab) |
| `Paste` | match size, apply format | high-confidence (Home tab) |
| `PasteSpecialDialog` | pick up/apply menu, sticky painter | high-confidence (Home tab) |
| `Delete` | remove notes / guides / dividers | high-confidence (cross-app) |
| `Refresh` | refresh palette, audit fix menu | high-confidence (cross-app) |
| `SelectionPane` | align toggle, select-same | high-confidence (Home → Select) |
| `ObjectsGroup` | arrange menu | high-confidence (drawing tools) |
| `ObjectBringToFront` | position menu | high-confidence (drawing tools) |
| `TableColumnsDistribute` | distribute H | high-confidence (table tools) |
| `TableRowsDistribute` | distribute V | high-confidence (table tools) |
| `ShapesInsertGallery` | sizing/insert/adjustment menus | high-confidence (Insert → Shapes) |
| `TextBoxInsert` | text box menu | high-confidence (Insert tab) |
| `SetLanguage` | deck language | high-confidence (Review tab) |
| `Spelling` | text cleanup menu | high-confidence (Review tab) |
| `FindDialog` | off-slide select | high-confidence (Home → Find) |
| `ReplaceDialog` | off-slide report | high-confidence (Home → Replace) |
| `ViewSlideMasterView` | unused masters | high-confidence (View tab) |
| `GridSettings` | guide buttons | high-confidence (grid & guides dialog) |
| `SlideNew` | agenda build | high-confidence (Home → New Slide) |
| `VisualBasic` | deck statistics | high-confidence (Developer tab) |

## First-load verification procedure

1. Ensure *Show add-in user interface errors* is ON.
2. Build + install the add-in, restart PowerPoint fully.
3. If a Custom UI Runtime Error dialog appears, it names the line/column in
   `customUI14.xml` — that line's `imageMso` is the offender. Replace it with
   `HappyFace`, rebuild, restart, repeat until clean.
4. Promote every ID that survived to `confirmed` in the table above.
