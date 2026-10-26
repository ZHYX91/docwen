# Web archives and HTML image references

This is the source conversion contract for HTML/MHTML/ENEX to Markdown. Native clipboard/provider, Office, candidate packages and platform execution are separate validation layers.

## MHTML root and URI context

Follow [RFC 2387 section 3.2](https://www.rfc-editor.org/rfc/rfc2387.html#section-3.2) when selecting a `multipart/related` root: an explicit `start` must identify exactly one direct child by Content-ID; otherwise the first child is the root. A declared media type must match that child. A missing, ambiguous or unsupported root fails with a bounded conversion error. Do not select another frame merely because it contains HTML. A root `multipart/alternative` selects its last supported HTML representation; nested related roots retain the same selection rules.

Keep MIME parentage when resolving Content-Location. Each part uses its own enclosing MIME context, independently of the selected HTML page's URI. The HTML page's location and HTML base element resolve references inside that HTML. This follows [RFC 2557 sections 5 and 8.2](https://www.rfc-editor.org/rfc/rfc2557.html#section-5); the outer-location example in section 9.3 illustrates why a resource part must not borrow a sibling HTML page's directory. When no absolute base exists, use the standard's synthetic `thismessage:/` namespace. Read-only Content-Base support remains a conservative fallback permitted by section 12.

Only resources in the same or surrounding related scope may satisfy the selected HTML. Preserve host, query and path case in identity; do not percent-decode URI identity into filesystem paths. CID references match Content-ID only, not a Content-Location whose text starts with `cid:`. Output-safe filenames remain separate from resource identity.

Exact URI/CID matching has priority. Compatibility fallback is limited to a bare relative filename without query or fragment and one uniquely named resource, with case preserved. A missing absolute URI, CID or qualified path cannot discard its identity through basename fallback. An actually referenced ambiguous identity or basename fails rather than selecting arbitrary image bytes; unused conflicting resources do not block unrelated content.

## Local HTML base and missing resources

Resolve the source HTML file URI, then its first base element with an href, then each image source, before converting a file URI into a platform path. This includes relative bases, explicit empty bases, absolute file bases, spaces and Unicode. Preserve POSIX roots and Windows drive/UNC policy. Network image URLs remain links and are not downloaded.

Companion-folder fallback applies only to ordinary relative file references without an explicit base. It cannot replace the identity chosen by an explicit base, even when a similarly named sibling or companion image exists. Missing, unreadable or unmappable local resources retain their original links and add one `HTML-LOCAL-IMAGE-UNAVAILABLE` warning with the total occurrence count and at most three image positions. Diagnostic text never repeats unbounded paths or source values. Deliberately omitted images do not claim to have been packaged.

## Authored text and generated image links

HTML, MHTML and ENEX image projection must preserve body/code text that resembles internal placeholders. Use a per-conversion namespace checked against authored text, decoded HTML/ENML text and relevant attributes. Restore generated placeholders in one pass so a generated link cannot be processed by a later replacement. Omit/disabled-image modes remove the image occurrence, not authored text. ENEX resource validation derives references from actual en-media placeholders rather than scanning ordinary author text for hash-shaped tokens.

The Layout HTML helper uses the same pure Core namespace/restoration functions. Its final helper output is covered by regression tests; there is currently no production caller of that helper, so these tests do not establish a PDF/Office route or native host result.

## Regression boundaries

`test_web_archive_review_boundaries.py` executes real MIME/HTML/ENEX inputs through Runtime and Finalizer, checking selected body text, exact image bytes, parent-relative identities, nested roots, bounded errors, unique aliases, explicit-base decoys, missing-resource warnings and body/code/omit preservation. Existing `test_mhtml_resource_identity.py` retains host/query/path/CID, URI, source-byte and image-order regressions. Core tests cover namespace collisions and nonrecursive replacement; Layout tests exercise its final HTML-to-Markdown helper.

Structured clipboard XLSX overflow tests additionally preserve an existing same-name, previously published workbook byte-for-byte under a frozen publication identity. Source/template/input hashes and original failed runs retain their own evidence status; source regressions do not certify a package or physical host.
