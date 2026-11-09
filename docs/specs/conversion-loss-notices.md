# Conversion loss notices

`conversion.loss_comments` (General settings: “Mark content not fully preserved with comments”) defaults to `true`. It controls annotations in generated DOCX and XLSX only. Disabling it does not suppress warning diagnostics. Generated heading-number review comments are a separate opt-in, default-off setting.

A request-local `ConversionLoss` records the reason, actual output change, source/output location and readable original content when available. Identical records are deduplicated without combining distinct positions. Warnings and completion summaries use the existing result diagnostics, including GUI details, CLI JSON and Machine completion; no extra report artifact is generated. Generated notice text follows the task output locale, not the GUI language; technical location labels may still be in English.

## Current producers and output behavior

- Markdown → DOCX: unsupported equations retain a short visible marker at the original position, with a Word comment containing the authored expression and reason. Successfully converted formulas are unchanged. Unsupported mathematical font variants are rejected rather than silently normalized.
- Markdown note equations → DOCX: readable LaTeX is retained in the note; its first body reference receives the loss comment. Repeated references do not duplicate comments. No unsupported direct annotation of a footnote/endnote part is attempted.
- DOCX → Markdown: unsupported native equations retain a visible marker and a diagnostic pointing to the source part/occurrence. Raw Office XML is not inserted into the user document. Unreadable referenced notes retain visible placeholder definitions so a reference is never left dangling.
- Markdown → XLSX: equations retained as text, protected-cell skipped input, and unsupported merge structures receive cell-specific notices. Existing comment text, author, width and height are retained when appending a conversion explanation.
- Markdown → CSV: loss diagnostics identify the affected cell or merge range; notices are not added to data cells or extra rows/columns.

An explicitly chosen text/discard policy is not itself a conversion failure. Existing warnings for other conversion features remain in their existing diagnostic channels; this is not a claim that every unsupported feature in every format can be detected.

Comments never replace existing user comments. A comment-write failure leaves the content-loss warning in conversion details and adds an annotation failure warning. A completion message claims that comments were added only when every recorded loss was annotated. Counts refer to content-loss occurrences, not additional comment-write errors.

## Review boundaries

This feature depends on the content repairs in PRs #145 and #148 → #149. It does not replace those fixes with warnings and does not introduce legacy callback/config compatibility. Windows file replacement, cancellation and temporary-directory ownership remain separate concerns in #151 → #152/#153.
