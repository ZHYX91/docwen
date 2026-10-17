# Raw Markdown HTML in DOCX exports

Common raw HTML blocks no longer disappear silently when Markdown is rendered into Word. DocWen extracts visible, local text from paragraphs, divisions and simple inline containers into Word paragraphs.

This is a safe **text projection**, not a web renderer:
- HTML scripts, styles and embedded document contents never execute or load.
- Remote resources are never requested. Image alt text is retained if available.
- Complex table/layout semantics and unsupported elements generate `MD2DOCX-HTML-FLATTENED` diagnostic; visible cell text remains accessible.
- HTML comments and non-visible script/style text do not become document body content.

No HTML parsing is allowed to change authenticated Word semantics or Markdown source bytes.
