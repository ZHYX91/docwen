# Implementation record: explicit template body placement

The authored-Markdown converter implements explicit template body placement. This record explains the design and regression coverage; [the current specification](../../specs/markdown-compatibility.md) defines supported behavior.

## User-approved outcome

Markdown body content belongs in the selected DOCX template only when a supported body placeholder is present. No placeholder means metadata-only template filling, not implicit body append and not a conversion error. The default template must continue to contain a supported body placeholder, so ordinary conversion without selecting a custom template remains usable.

The user explicitly retains the current default DOCX-to-Markdown heading rule: only headings recognized from Word styles/outline are eligible for heading-number cleanup. This PR must not alter that rule.

## Historical evidence and baseline

历史对照：以下 v0.8.5 `src/docwen/` 路径仅用于重构前行为比较，不代表当前目录结构。

Review baseline: `93227f5d2a921e7168344903bccd1c3f0095f38d`.

- [v0.8.5 document processor](https://github.com/ZHYX91/docwen/blob/v0.8.5/src/docwen/converter/md2docx/processors/docx_processor.py): `process_main_content()` inserts only after finding a body placeholder; otherwise it returns `False` and records that the Markdown body was not inserted.
- [0.9.0 template filler](https://github.com/ZHYX91/docwen/blob/0.9.0/packages/plugins/markdown/src/docwen_plugin_markdown/template_filler.py): a missing placeholder is documented as append-at-end and invokes `_append_rendered_batch()`.
- [Current converter at review baseline](https://github.com/ZHYX91/docwen/blob/93227f5d2a921e7168344903bccd1c3f0095f38d/packages/plugins/markdown/src/docwen_plugin_markdown/to_docx/converter.py): renders into the destination document before filling the template.

This establishes a behavioral difference between released versions, not the exact introducing commit and not a defect in PySide itself.

## Implementation requirements

1. Resolve the selected template and its body-placement decision before producing destination body nodes. Reuse the maintained localized body-placeholder aliases; do not hardcode only the Chinese spelling.
2. For an absent body placeholder, skip destination-body rendering and body-owned side effects. Preserve YAML extraction, selected-template metadata substitution, title fallbacks, special field processors and conditional empty-field cleanup.
3. Do not implement the fix by merely deleting `_append_rendered_batch()`: the current renderer has already attached nodes by that point. Do not render and then indiscriminately delete template content or all package relationships.
4. Audit body-owned image relationships, notes, numbering, captions, semantic anchors, bibliography and renderer finalization. Omitted body content must not leave newly generated orphan media, note bodies, dangling references or stale semantic session records. Preserve resources already owned by the template. Distinguish an explicit template bibliography placeholder from bibliography output owned only by an omitted body.
5. Keep source bytes unchanged. Do not rewrite the user's Markdown or custom template.
6. Keep the separate resolved-v4 port isolated. Audit its explicit placement contract before making changes; do not route it through the historical authored-Markdown reader or weaken its validation. Document any intentionally separate behavior.
7. Separate missing placeholders from present-but-unsupported placements. Do not silently interpret a table-cell, textbox or embedded-in-text marker as a valid paragraph-body anchor. An invalid authored placement needs a specific diagnostic; do not promise expansion of placement support in this fix.
8. Remove/update tests and documentation that explicitly require implicit appending. Preserve their unrelated ordering, terminal section-property and hyperlink coverage by moving that coverage to a valid-placeholder fixture.

A missing placeholder in an intentionally metadata-only template should not require a modal confirmation. Any informational diagnostic must be non-blocking and must not imply file corruption.

## Acceptance fixtures

| Fixture | Expected outcome |
| --- | --- |
| Custom template containing only `{{title}}` and other metadata fields | Fields are filled; nonempty Markdown body is absent |
| Same template, body containing images, notes, headings and captions | No destination body materialization or new body-only package parts |
| Supported localized body aliases | Body appears exactly at the resolved anchor; template prefix/suffix remain ordered |
| Default template with no explicit custom selection | Ordinary Markdown body conversion still works |
| Empty body with a valid marker | No visible marker remains under the maintained empty-body contract |
| Existing template media, headers, footers and numbering | Existing content and relationships survive metadata-only filling |
| Marker split across runs | Follows supported placeholder matching without flattening surrounding formatting |
| Marker in unsupported context / duplicate markers | Explicitly tested maintained behavior, not accidental end append |
| Reopened generated DOCX | Valid package, terminal `w:sectPr` ordering and no dangling body-generated references |

## Likely implementation surfaces

- `packages/plugins/markdown/src/docwen_plugin_markdown/to_docx/converter.py`
- `packages/plugins/markdown/src/docwen_plugin_markdown/template_filler.py`
- `packages/plugins/markdown/src/docwen_plugin_markdown/template_utils.py`
- `packages/plugins/markdown/tests/test_managed_style_rendering.py`
- Additional focused converter/artifact regressions and user-facing template documentation.

## Implementation and validation boundaries

The converter determines body placement before body rendering. It independently validates existing template notes, and only creates the body renderer, semantic session, new notes and numbering when a valid body marker exists. An omitted body must not launch a Mermaid renderer solely for content that will not be output. YAML field filling remains independent of that decision.

`test_md_to_docx_01.py` and `test_docx_notes_v4.py` cover metadata-only output, valid split-run markers, invalid/duplicate placement, default body behavior, existing note graph preservation and invalid note graph rejection. These exercise conversion and saved packages rather than source-text projections alone.

Body markers are supported only as a unique standalone paragraph in the main document. Table, textbox, header/footer and inline placements fail explicitly; so do markers hidden inside unsupported inline content controls or split within the alias by structural breaks. A marker found by XML scanning must also be recognizable by the body-placement finder, preventing silent metadata-only output. Expanding placement support is not part of this restoration. The separate resolved-v4 route keeps its own documented placement contract.

[Testing guidance](../../testing.md) governs actual host checks. PR validation records identify the source and selected real-document checks; package tests do not certify every Word/WPS/LibreOffice version or every original fixture.
