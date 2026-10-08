# Word note metadata schema

This is the transitive declaration subset for `rPr`, `sdtPr`, `sdtEndPr`,
`smartTagPr`, and `customXmlPr` from ECMA-376 Part 4, fifth edition (2016),
Transitional XML schemas. Source URLs, member hashes, entry types and extraction
method are recorded in `provenance.json`.

Declarations retain their original definitions and order. Unused imports and
declarations are omitted; five global elements expose the original complex types
to the validator. No document/body schema is included. Nested format revisions
and structured control properties remain valid metadata; fields, drawings and
body text are not valid substitutes for those properties.

The document plugin loads both files through package resources. Its resolver
accepts only these local names and never follows document-supplied paths or URLs.
Both frozen GUI and CLI builds collect this package's data. These schemas validate
metadata structure, not bookmark identity, visible payload, note relationships,
the whole DOCX package, or Office rendering.
