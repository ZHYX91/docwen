# Controlled clipboard provider fixtures

These are **derived regression inputs**, not untouched native captures or
product/Office acceptance evidence. The private original capture archive and
each original provider payload have SHA-256 provenance in `provenance.json`.
CI uses only these checked-in files and never depends on a maintainer path.

The controlled Word and WPS Writer sources contain alpha, a blue image in a
nested table, and repeated alpha. Their original document XML, image identifiers,
relationship targets and PNG bytes remain unchanged; unrelated package parts,
document properties and package timestamps were removed. Image relationships
were retained in a minimal deterministic package. The Word CFB wrapper was
rebuilt with complete metadata and an unpadded logical Package tail. WPS's
489-byte image-record stream is unchanged and contains only two controlled PNGs.

The provider contract tests check order, nesting, physical pixel dimensions,
encoded-byte identity, drawing extents and duplicate resource reuse. They also
exercise complete FAT/DIFAT/mini-stream controls and malformed container rejection.
They do not establish GUI paste, final output or native host acceptance.

`rich-derived.html` and `rich-derived.txt` are compact **synthesized** inputs
based on the observed source shape. They retain the provider's literal paragraph
contexts and nested cell positions, and deliberately include HTML source
wrapping, NBSP padding, a void `col`, and a non-PNG inline representation. They
do not preserve original HTML or plain-format bytes. Rich ingress tests use
these files with the separately identified provider fixtures and exercise the
public frozen-rich adapter through actual output conversion.
