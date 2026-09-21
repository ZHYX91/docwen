# Implementation record: compact file-format warning

Single-file and batch presentation implement a separate actual-format notice. This record explains the design and regression coverage; release and host validation are recorded separately.

## User-visible outcome

Keep filename, format warning and location as separate pieces of information. Replace the long routine mismatch explanation with a small status badge such as:

```text
当前文件：实际是doc.docx
[warning icon] 实际格式：DOC
E:\conversion                         [locate]
```

The icon in the product should come from the shared SVG resource layer; the textual example is not a requirement to use a platform emoji. The badge is a status indicator, not a repair button. Use an appropriate warning foreground/background, a smaller semantic font role than the filename, and no red full-card treatment for a non-blocking mismatch. Do not rely on color alone.

Do not display internal processing-family explanations in this routine summary. Batch grouping and the conversion panel own category context. The panel architecture can share a conversion shell across categories; do not create extra panels merely to label a warning, and do not assume every category currently has a different shell.

## Baseline and historical evidence

历史对照：以下 v0.8.5 `src/docwen/` 路径仅用于重构前行为比较，不代表当前目录结构。

Review baseline: `93227f5d2a921e7168344903bccd1c3f0095f38d`.

- [v0.8.5 file-drop widget](https://github.com/ZHYX91/docwen/blob/v0.8.5/src/docwen/gui/components/file_drop.py) creates a distinct small warning-styled label.
- [Current input view model](https://github.com/ZHYX91/docwen/blob/93227f5d2a921e7168344903bccd1c3f0095f38d/packages/apps/gui/src/docwen_gui/view_models/input_area_vm.py) concatenates warning text into the selection message.
- [Current input widget](https://github.com/ZHYX91/docwen/blob/93227f5d2a921e7168344903bccd1c3f0095f38d/packages/apps/gui/src/docwen_gui/widgets/input_area.py) renders that message with a shared selection label.
- [Admission rules](https://github.com/ZHYX91/docwen/blob/93227f5d2a921e7168344903bccd1c3f0095f38d/packages/core/src/docwen_core/detection/_validation.py) distinguish allowed warnings, required explicit acceptance and blocking reasons.

## Implementation requirements

1. Project a typed format notice from existing immutable inspection facts: diagnostic code, declared format, detected format and admission decision. Do not parse a translated message or independently re-detect the file from its suffix in the widget.
2. Keep filename, warning summary and path independently styleable. Use the same short format-summary policy in single-file and batch presentation.
3. Only summarize the known format-mismatch warning. Do not replace arbitrary warnings, file-corruption messages, execution failures or other diagnostics with a generic format badge.
4. Keep complete localized diagnostic details accessible through the existing details/help mechanism. They must not be lost from logs or diagnostic state.
5. Preserve explicit confirmation for cross-family/unknown-extension cases where Core requires it. A concise confirmation can state the actual detected format and ask whether to continue without explaining internal family names. Do not bypass or silently pre-accept admission.
6. A blocked, unknown or unreadable file must show the real reason. Do not manufacture `Actual format: UNKNOWN` as a successful detection.
7. Keep batch classification based on the existing `workflow_category`. Keep task state (pending/processing/success/failure) separate from format-warning state so one does not overwrite the other.
8. Clear stale notices on input replacement, removal, reinspection, mode switches and drag-preview restoration. Preserve identity binding for asynchronous inspection results.
9. Keep filename/path rendering literal; angle brackets or rich-text-looking filenames must not become GUI markup. Maintain path copy and file-location actions.
10. Use theme/font tokens and adaptive geometry. Long translated labels, large-font settings and narrow windows must not clip the format value. Keep badge styling non-interactive unless an actual details action is implemented and named.
11. Introduce a presentation-specific compact translation rather than shortening shared Core/CLI diagnostics globally. Preserve locale coverage and explicit fallback behavior.

## Likely surfaces

- `packages/apps/gui/src/docwen_gui/file_admission_i18n.py`
- `packages/apps/gui/src/docwen_gui/view_models/input_area_vm.py`
- `packages/apps/gui/src/docwen_gui/widgets/input_area.py`
- `packages/apps/gui/src/docwen_gui/view_models/batch_list_vm.py`
- `packages/apps/gui/src/docwen_gui/widgets/batch_list.py`
- Existing notice/badge components, theme tokens, locales and relevant GUI tests.

## Acceptance matrix

| Scenario | Required outcome |
| --- | --- |
| DOC content named `.docx` | Filename retained; separate `实际格式：DOC` badge; permitted conversion remains available |
| Exact/equivalent format match | No mismatch badge or empty reserved warning row |
| Cross-family mismatch requiring consent | Compact notice plus existing explicit acceptance requirement |
| Corrupt/unsupported/unreadable input | Specific failure reason, not a reassuring generic format summary |
| Mismatch plus an unrelated warning | Format summary does not erase the unrelated warning |
| Batch mode | Existing detected-category group; notice does not replace per-file execution status |
| Replace/remove/reinspect current file | No stale notice from the previous input or stale async reply |
| Dark/light theme; all font presets; narrow width | Readable and visually distinct filename, badge and path |
| Long localized text and literal special characters | No clipping, rich-text injection or unintended elision of actual format |

## Implementation and validation boundaries

`render_file_format_notice` projects frozen inspection facts; `render_remaining_file_warnings` retains unrelated diagnostics. Single-file and batch widgets display the shared warning badge separately from filename and path. The asynchronous input-completion path resynchronizes the selected item so it cannot overwrite the compact notice with the old combined message.

Admission, single-input and batch regression tests cover real controller entry points, mismatch classification, preserved warnings and replacement behavior. Selected Windows native checks use actual DOC bytes with a misleading DOCX extension and exercise light/dark themes, large fonts and batch presentation. [Testing guidance](../../testing.md) governs host evidence; these selected checks do not certify every locale, font/display combination or platform.

Detection, consent, error reporting and transaction boundaries retain their existing contracts. The acceptance matrix above describes the maintained requirements, not a claim that every native combination has been executed.
