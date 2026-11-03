# Heading numbering

## Markdown to DOCX

The heading-numbering output mode is the single authority for newly generated
numbers. **Text** inserts a literal label; **DOCX native** uses Word multilevel
numbering. The mode description is available from the setting's information
button. Adding numbering remains a separate switch.

Review comments can be enabled independently in either mode. Text-mode comments
anchor the generated number. Native-numbering labels are not text runs, so their
comments anchor the first visible character of the heading instead. Existing
handwritten numbers are not marked as newly generated numbers.

Literal numbers suppress inherited Word numbering to avoid duplicate labels.
When a template's numbering definition supplies heading indentation, that
indentation is copied to the output paragraph before suppressing its numbering.
The template's shared styles and numbering definitions remain unchanged.

### Existing configurations

The retired `template_fill.gongwen_numbering_text` switch could previously force
text output even when the saved mode was `word_native`. The first configuration
load after this change migrates that conflict to `text`, preserving the previous
actual output. It records `text.heading_numbering_migration_version = 1` alongside
the migrated renderer. Saving a new renderer choice also records this version.
Later output-mode choices are not migrated again, including after a restart.
Fresh profiles keep the default text mode without writing files on initial load. Explicit conversion
options take precedence over the migrated default. The legacy switch remains
readable for compatibility but no longer overrides the mode or disables comments.

## DOCX to Markdown

In document settings, **Mark new numbers with zero-width characters** is optional
and off by default. It is available when adding heading numbers. The persisted
setting is `document.export_numbering_markers`; the conversion option is
`export_numbering_markers`. An explicit option overrides the saved setting.

When enabled, newly generated heading numbers are wrapped in paired U+2060
WORD JOINER characters, compatible with Number Suite:

```text
# <U+2060>1.<U+2060> Heading
```

The characters are invisible in normal display. The complete number is inside
the pair, and its original separator follows the closing marker. Existing
numbering is not retroactively tagged. Cleanup recognizes complete, nonempty,
single-line leading marker pairs with meaningful heading text after them. A
single unmatched marker protects the heading from automatic cleanup or additional
numbering. Repeated add-only processing does not duplicate an existing marked
number. Removing and then adding numbering intentionally replaces it.

Authenticated resolved-v4 DOCX recovery remains on its existing recovery route.
These markers do not change its contract or recovered numbering semantics.
