# Each line as a Word paragraph

The optional setting **每行独立成段 / Each line as a paragraph** belongs to Settings → Markdown syntax. It is **off by default** and is not a legacy mode.

Source-native Markdown→DOCX interprets ordinary top-level prose soft line breaks as separate Word paragraphs when enabled. The same intermediate DOCX is used by DOC/RTF/ODT/PDF bridge targets. Explicit Markdown hard breaks remain inside a paragraph. Lists, blockquotes, headings, code, tables, captions and authenticated semantic-owned paragraphs keep their own structure.

The source Markdown is never rewritten. Conversion options override the saved `conversion.md_to_docx.each_line_paragraph` boolean. The strict source-only resolved-v4 port remains unchanged.

Acceptance: three plain adjacent lines → three Word `w:p` when enabled, one `w:p` with breaks when disabled; formatting retained; no semantic anchors moved; test explicit hard breaks, nested blocks and paragraph spacing.
