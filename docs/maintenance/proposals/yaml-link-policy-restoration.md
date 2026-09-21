# Implementation proposal: YAML ordinary-link policy parity

Status: **implementation in progress on this Draft PR**. Production code and focused tests are now present on the branch; this proposal does not override active specifications until validation and merge.

## Reported reproduction

The user selected Extract Text for both Wiki and Markdown non-embedded links. A YAML `抄送机关` list containing a quoted Wiki alias was nevertheless copied into the DOCX with brackets and the target name intact.

```yaml
---
抄送机关:
  - "[[某机关|喵喵]]"
  - 钱钱钱
参考网站: "[项目主页](https://example.com)"
公开方式: "[[免予公开]]"
份号: "001"
成文日期: 2025-07-01
---

正文 [[某机关|喵喵]]。
```

Under Extract Text, the first list item must become `喵喵`; the website field must become `项目主页`. Existing template text, separators, punctuation and styles remain owned by the template/field formatter. Quoted Wiki syntax is a string. An unquoted YAML value such as `[[免予公开]]` is a nested YAML list, not an implicit link; do not invent link semantics for arbitrary nested arrays.

## Historical evidence

历史对照：以下 `src/docwen/` 路径仅指重构前版本，不是当前源码入口。

Review baseline: `93227f5d2a921e7168344903bccd1c3f0095f38d`.

- [v0.8.5 reader](https://github.com/ZHYX91/docwen/blob/v0.8.5/src/docwen/converter/md2docx/core.py): `read_and_parse_md()` recursively processes YAML dictionaries, lists and string leaves using `process_markdown_links()`, before field-specific processors.
- [v0.8.1 reader](https://github.com/ZHYX91/docwen/blob/v0.8.1/src/docwen/converter/md2docx/core.py): the same YAML traversal already exists.
- [0.9.0 converter](https://github.com/ZHYX91/docwen/blob/0.9.0/packages/plugins/markdown/src/docwen_plugin_markdown/to_docx/converter.py): YAML is separated and field-processed, while only `md_body` receives the link policy.
- [Current field filler](https://github.com/ZHYX91/docwen/blob/93227f5d2a921e7168344903bccd1c3f0095f38d/packages/plugins/markdown/src/docwen_plugin_markdown/template_filler.py): YAML values are formatted as strings and inserted into text runs.

These sources establish lost YAML link-policy coverage, not proof that every old hyperlink/embed behavior was correct.

## Required behavior

| Ordinary-link mode | Wiki field value | Markdown field value |
| --- | --- | --- |
| `extract_text` | Alias when present; otherwise the displayed target text | Display label under the shared inline-link contract |
| `keep` | Literal source link syntax, no accidental activation | Literal source link syntax, no accidental activation |
| `remove` | Remove the link and its display text | Remove the link and its display text |
| `hyperlink` | A supported resolved target becomes a real DOCX hyperlink | A supported target becomes a real DOCX hyperlink |

Wiki and Markdown settings remain independent. Do not replace the user's existing choices or change shipped defaults as part of this fix. Extract Text and Remove must not require a local target to exist and must not trigger resource lookup.

## Architecture and scope

1. Parse YAML first; never regex-rewrite the entire YAML source. Preserve keys and scalar types, including dates, booleans, zero values and quoted leading-zero identifiers.
2. Apply request-scoped ordinary-link policy to the string values/list items that contribute to template output, including title fallbacks when those become visible. Preserve the source data instead of rewriting the source Markdown.
3. Use shared inline parsing/link policy, not a new handwritten Wiki regex in `抄送机关`. Resolve display semantics before list joining and field-specific formatting. Preserve rich link segments through special processors when hyperlink mode is selected.
4. DOCX hyperlink mode requires real relationship/run creation; inserting `[label](url)` into a `w:t` node is not implementation of that mode. Keep surrounding template run formatting, neighboring text, and unrelated existing hyperlinks intact.
5. Handle repeated/split-run placeholders, placeholders in supported template containers, ordinary lists, generic fields and enabled special field processors consistently. Never apply the formatter twice to newly inserted content.
6. A field emptied by Remove must follow the existing empty-field/conditional cleanup rules without leaving spurious delimiters or deleting unrelated template content.
7. Omit filesystem/resource work for unused metadata. Respect existing declared-input boundaries and target-scheme safeguards. Keep independent resolved-v4 behavior isolated.
8. Non-embedded links only: do not turn this fix into automatic whole-document/image embedding inside names, dates or other YAML fields. Protect embed syntax from being consumed as an ordinary link; keep existing literal behavior pending an explicit embedding contract.
9. Do not silently standardize Wiki and Markdown missing-file behavior in this PR. That policy remains a separate product decision. Document current downgrade behavior and do not show a fake clickable link for an unsupported target.

## Implementation surfaces

- `packages/plugins/markdown/src/docwen_plugin_markdown/to_docx/converter.py`
- `packages/plugins/markdown/src/docwen_plugin_markdown/template_filler.py`
- `packages/plugins/markdown/src/docwen_plugin_markdown/field_registry.py`
- `packages/plugins/markdown/src/docwen_plugin_markdown/field_processors/gongwen.py`
- Shared ordinary-link and DOCX inline-rendering utilities.
- Link-tab help text and focused conversion/package tests.

## Acceptance and completion

- [ ] Implement all four ordinary-link modes without changing existing defaults.
- [ ] Test the reported quoted alias in `抄送机关` and a generic list field; assert visible result, not only helper output.
- [ ] Test Wiki/Markdown mode combinations independently in YAML and body.
- [ ] Verify missing targets do not matter in Extract Text/Remove; mock lookup to fail if called.
- [ ] Assert real hyperlink relationships only in Hyperlink mode and correct fallback for unsupported targets.
- [ ] Test dates, `false`, `0`, `"001"`, nested list structure, inline-code literals and escaped bracket syntax.
- [ ] Test existing template runs, repeated/split placeholders and empty-field cleanup.
- [ ] Test metadata-only templates with no body placeholder, coordinating with the body-placement restoration.
- [ ] Reopen DOCX packages and inspect representative documents in Word/WPS.
- [ ] Update documentation only after behavior is implemented; record exact executed tests and tested commit.

PR #18 also modifies the converter and some locale files. Rebase/reconcile implementation without overwriting its Mermaid behavior. Keep this PR in draft until implemented and locally validated; no automatic merge.
