# Implementation proposal: consistent settings help

Status: **draft implementation plan; production code is unchanged**. This document is not a completion report and does not supersede current specifications.

## Reported problems

The formatting tooltip displays literal `\n` instead of line breaks and stretches across the window. The information icon is visually tiny, has an unwanted button-like box in the reported theme, and is separated from its field label by excessive space.

Review baseline: `93227f5d2a921e7168344903bccd1c3f0095f38d`.

## Evidence and correct historical classification

历史对照：以下 v0.8.5 `src/docwen/` 路径仅用于重构前行为比较，不代表当前目录结构。

- [Current Chinese locale](https://github.com/ZHYX91/docwen/blob/93227f5d2a921e7168344903bccd1c3f0095f38d/i18n/locales/zh_CN.toml) uses literal single-quoted strings containing `\n` for the body/heading formatting help.
- [v0.8.5 locale](https://github.com/ZHYX91/docwen/blob/v0.8.5/src/docwen/i18n/locales/zh_CN.toml) already contains this translation problem. It is an inherited bug, not proof of a PySide regression.
- [Old settings layout](https://github.com/ZHYX91/docwen/blob/v0.8.5/src/docwen/gui/settings/base_tab.py) places the information icon next to the label.
- [Current settings helper](https://github.com/ZHYX91/docwen/blob/93227f5d2a921e7168344903bccd1c3f0095f38d/packages/apps/gui/src/docwen_gui/widgets/settings/base_tab.py) uses an 18-by-18 help control, a 14-by-14 icon, `NoFocus`, and a tooltip without a connected click-help action.
- [Current FormRow](https://github.com/ZHYX91/docwen/blob/93227f5d2a921e7168344903bccd1c3f0095f38d/packages/apps/gui/src/docwen_gui/widgets/panel_card.py) stretches the label ahead of the suffix icon.
- [info.svg](https://github.com/ZHYX91/docwen/blob/93227f5d2a921e7168344903bccd1c3f0095f38d/assets/icons/info.svg) contains a circular information glyph, not the outer square control frame. Fixing the SVG alone cannot fix the control frame/layout.

## Translation correction

Fix malformed help strings at their locale source. Use actual multiline TOML strings or correctly escaped basic strings. Audit affected locales and formatting-help keys rather than repairing only the reported sentence.

Do **not** globally replace every literal `\n` in the translator or decode arbitrary escape sequences. Paths, regex examples and instructional escape syntax may intentionally contain those characters. Keep Markdown examples literal; do not accidentally render `**bold**` as formatting when the example is explaining source markup.

Recommended wording for the reported body-format help:

```text
应用格式：将 **粗体** 转为粗体效果。
保留标记：原样显示 **粗体**。
清理标记：去掉标记，显示文字并沿用模板样式。
```

The exact implementation must verify the wording against each formatting mode. Do not promise that clearing source markup removes the template's own fonts, sizes or paragraph formatting.

## Shared help component

1. Keep the icon next to the field name in both side-by-side and stacked layouts. Preserve the page's common field/control alignment without stretching the space between a short label and its help glyph.
2. Use the existing shared SVG loader. Coordinate asset choice with the icon-unification PR, but do not require an entirely new icon family to fix the help behavior.
3. Separate visual icon size from hit area. Suggested starting sizes are a 16–18 logical-unit glyph and a 28–32 logical-unit interaction target, adjusted by application font presets and measured text geometry. These are design starting points, not unvalidated constants to stamp everywhere.
4. Keep the resting help icon low emphasis and frameless; show clear hover/focus affordances. Do not make a visual status badge masquerade as an operable help button.
5. Provide keyboard focus and click/Enter/Space access to the same explanation; Escape closes an explicitly opened help surface and returns focus. Avoid modal dialogs for ordinary field help.
6. Use a concise accessible name derived from the field label, with explanatory text in the accessible description. Do not use the entire tooltip as the button's name. Do not introduce untranslated generic English button names in localized pages.
7. Constrain the help surface to a readable width and the current screen's available geometry. Preserve line breaks, wrap long prose, and handle large fonts, narrow screens and screen edges. Support long help content without off-screen loss.
8. Safely escape authored examples before using rich-text wrappers. Do not allow translated text or example values to inject arbitrary markup into a help surface.
9. Avoid repeating the same giant help text on every combo popup item. Distinguish field-level explanations from genuinely option-specific descriptions.
10. Where useful, show a small current-choice example, such as `[[target|label]] -> label`. It must reflect the actual implementation and not advertise YAML support until that restoration is implemented.

## Acceptance and regression checks

- [ ] Parse affected TOML resources and verify intended actual line breaks.
- [ ] Preserve intentionally literal `\n`, backslashes, paths, quotes and regex examples.
- [ ] Verify `**粗体**`, `<...>`, ampersands and bracket syntax display literally when instructional.
- [ ] Test help association and geometry in short/long labels, Chinese/English and a long-text locale.
- [ ] Test all font presets, light/dark themes and representative display scaling, including screen-edge placement.
- [ ] Test keyboard focus, Enter/Space activation, Escape dismissal and focus restoration.
- [ ] Check hover/click transitions do not flicker, immediately dismiss or strand a popup.
- [ ] Retain checkbox/combo behavior, wheel safety, scroll behavior and common row alignment.
- [ ] Inspect before/after screenshots in the actual Qt application, including Windows and a Linux desktop.
- [ ] Run applicable repository checks and report only tests actually executed.

## Scope and coordination

Likely changes: `widgets/settings/base_tab.py`, shared form/help components, locale files, help-related styles and GUI tests. Keep application branding, conversion semantics and admission rules out of this PR.

Mermaid PR #18 changes locale files, so source fixes must be reconciled rather than replacing its strings. The icon-unification work may share low-level resources; keep ownership clear to avoid duplicate tooltip or SVG engines.

Do not mark ready or merge until production changes and validation are complete. This initial commit contains only this plan.
