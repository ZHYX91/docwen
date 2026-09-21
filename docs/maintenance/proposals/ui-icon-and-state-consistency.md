# Implementation record: UI icon and visual-state consistency

Core operation/status assets and their consumers use the shared application SVG system. This record explains that migration and its validation boundaries; it does not claim a redesign of every screen or replace active design specifications.

## Goal and exclusions

Make application-owned navigation, operation icons and status indicators look and behave like one product. This means consistent visual weight, spacing, sizing, theme states and control semantics, not merely converting PNG files to SVG.

Preserve the current DocWen brand logo and the SVG-to-PNG/ICO generation chain. Do not return to Tkinter, replace the GUI framework, redraw every screen, or change conversion/admission behavior under a visual-cleanup label.

## Observed baseline

Review baseline: `93227f5d2a921e7168344903bccd1c3f0095f38d`.

- [Shared SVG loader](https://github.com/ZHYX91/docwen/blob/93227f5d2a921e7168344903bccd1c3f0095f38d/packages/apps/gui/src/docwen_gui/resources.py) already handles theme coloring and device-pixel-ratio rendering. Preserve and improve it instead of adding a second icon engine.
- [LocationButton](https://github.com/ZHYX91/docwen/blob/93227f5d2a921e7168344903bccd1c3f0095f38d/packages/apps/gui/src/docwen_gui/widgets/location_button.py) uses `QStyle.StandardPixmap.SP_DirOpenIcon`, while other application controls use bundled line SVGs.
- [Bundled icons](https://github.com/ZHYX91/docwen/tree/93227f5d2a921e7168344903bccd1c3f0095f38d/assets/icons) include a 24-unit line-style family and older path-style assets. Different coordinate systems alone do not establish a visible defect; inspect the rendered glyphs at intended sizes.
- [Assets](https://github.com/ZHYX91/docwen/tree/93227f5d2a921e7168344903bccd1c3f0095f38d/assets) retain older raster resources, and the packaging list still references some. Presence in the tree is not proof of runtime use or permission to delete.
- The SVG engine ignores icon mode/state in several drawing methods. Audit disabled/selected behavior in the real GUI before making specific visual-failure claims.

## Work packages

### A. Inventory and shared vocabulary

Trace application icon usage to actual consumers: navigation, locate/open, copy, remove, move/reorder, expand/collapse, settings/help, success/warning/error/skipped and empty states. Record each asset's purpose, source/license, logical display size, consumers and required states.

Distinguish Locate File from Open Folder and Remove from Delete where they are different actions. Consistency must not erase meaningful action differences. Prefer maintained existing assets over adding a new third-party icon dependency. Record attribution/license for any new external artwork.

### B. Common application-owned icon family

Use a coherent line weight, cap/join style, optical size and safe-area convention for small interface glyphs. Keep suitable existing icons. Replace the application location button's inconsistent stock graphic through the shared loader, while retaining its behavior, accessible label and visible fallback on load failure.

System-native file dialogs can retain native artwork. Brand art, empty-state illustration and data-format marks are different resource roles; do not force them all into a tiny monochrome operation glyph.

### C. Theme, font and state behavior

Use semantic design tokens for ordinary text, secondary text, action emphasis and warning/error/success states. Test light/dark theme changes without restarting. Never communicate a status by color alone.

Respect icon Normal/Disabled/Active/Selected and checked/unchecked roles where applicable. Disabled actions need a visibly distinct and semantically disabled state, not just a different image. Avoid hardcoding one global icon tint for every context.

Separate icon graphic size from hit-target size and control padding. Use logical sizes that cooperate with Qt device-pixel-ratio handling; do not multiply application font scaling and screen scaling twice. Preserve SVG rendering sharpness at fractional and high scale factors.

Do not regress existing per-format color distinctions (including DOC/DOCX and XLS/XLSX). Do not turn a non-blocking warning into a red failure card.

### D. Shared typography and geometry

Apply existing semantic roles consistently to filename/title, field label, regular content, helper text, path and compact status text. Keep a coherent spacing/radius/border scale. Changes should fix concrete inconsistent consumers rather than introduce arbitrary global font-size reductions.

Coordinate with the separate format-notice and settings-help plans. Those PRs own their feature-specific presentation and interactions; this work owns common icon/resource/state infrastructure. Do not build duplicate badges, tooltip engines or parallel style registries.

### E. Safe resource retirement

Before removing any raster or SVG file, establish that it is not used by source code, dynamic name lookup, packaging, tests, documentation or supported fallbacks. Update the inventory and packaging expectations together. Keep required application ICO/PNG derivatives and preserve deterministic generation/drift checks.

Do not delete an asset merely because a similarly named SVG exists. Do not change executable packaging or remove supported platform resources to make the directory appear tidy.

## Implementation and validation boundaries

[The asset inventory](../../../assets/icons/README.md) records the 31 role assets, their pinned Fluent source and license. Navigation, location, copy, delete, sorting, retry, template and software-priority controls use the shared loader. Status controls retain semantic colors and text; wrapped action buttons reserve icon space and draw disabled/RTL states correctly. Platform artwork remains only in missing-resource fallbacks and system-owned surfaces.

`test_icon_fidelity.py`, GUI control tests and resource/packaging checks cover the migrated assets and consumers. Unreferenced legacy operation/status PNGs were removed together with their packaging entries. Brand derivatives and the composite file-drop illustration keep their separate roles.

[Testing guidance](../../testing.md) governs native inspection. Validation records distinguish automated rendering/geometry tests from selected Windows light/dark, large-font, batch and action checks. Those checks do not establish every screen, all display scales, every font preset or every host combination; this record does not mark that broader matrix as passed.

## Integration order

Land behavioral fixes separately: body-placeholder restoration and YAML link-policy restoration. Common icon/state work can precede the feature-specific format badge and help UI, or those features can use the existing loader and later receive this resource update. Do not require an all-or-nothing visual mega-merge.

The implementation uses one icon engine and the existing GUI framework. Feature-specific format and help behavior remains owned by the corresponding components and regression tests.
