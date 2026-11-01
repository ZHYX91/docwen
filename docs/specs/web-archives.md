# Web archives and HTML image references

This is the source conversion contract for HTML/MHTML/ENEX to Markdown. Native clipboard/provider, Office, candidate packages and platform execution are separate validation layers.

## MHTML root and URI context

Follow [RFC 2387 section 3.2](https://www.rfc-editor.org/rfc/rfc2387.html#section-3.2) when selecting a `multipart/related` root: an explicit `start` must identify exactly one direct child by Content-ID; otherwise the first child is the root. A declared media type must match that child. A missing, ambiguous or unsupported root fails with a bounded conversion error. Do not select another frame merely because it contains HTML. A root `multipart/alternative` recursively skips valid but unsupported representations to select its last supported HTML form. Missing/ambiguous start targets, malformed roots and declared type mismatches remain fatal; they are not swallowed as unsupported representations.

Keep MIME parentage when resolving Content-Location. Each part uses its own enclosing MIME context, independently of the selected HTML page's URI. The HTML page's location and HTML base element resolve references inside that HTML. This follows [RFC 2557 sections 5 and 8.2](https://www.rfc-editor.org/rfc/rfc2557.html#section-5); the outer-location example in section 9.3 illustrates why a resource part must not borrow a sibling HTML page's directory. When no absolute base exists, use the standard's synthetic `thismessage:/` namespace. Read-only Content-Base support remains a conservative fallback permitted by section 12.

Synthetic and absolute bases follow the same [RFC 3986 section 5.2.2](https://www.rfc-editor.org/rfc/rfc3986#section-5.2.2) query rules. An empty-path reference without its own query inherits the base query; an explicit empty `?` clears it and remains distinct from an absent query. Relative paths use their own query. Fragments are preserved during resolution and removed only when comparing resource identities, so a fragment-only reference cannot silently select an unqueried sibling image.

Relative path merging preserves ordinary empty segments: `images//p.png` is distinct from `images/p.png`. Only complete literal `.` and `..` segments in a merged relative path are resolved; percent-encoded segments and query/fragment text are not paths to normalize. An empty-path reference inherits the exact base path, including a rootless synthetic path such as `thismessage:g?red`. Explicit scheme/authority resource identities are not subjected to additional path canonicalization.

相对路径合并保留普通空段：`images//p.png` 与 `images/p.png` 是不同身份。仅处理合并后相对路径中完整的字面 `.`／`..` 段，不规范化百分号转义段或 query／fragment。空路径引用原样继承 base 路径，包括 `thismessage:g?red` 这类无前导斜杠的 synthetic base；显式 scheme／authority 资源身份不额外归一化路径。

A reference declaring its own scheme remains absolute, including when that scheme matches its MIME parent's. Recomposition retains absent versus explicitly empty authority and rootless versus rooted paths: `http:g`, `http:/g` and `http:///g` do not become one resource. A real authority matching the internal synthetic carrier's hostname also keeps its identity. This does not add path canonicalization or network retrieval.

Only resources in the same or surrounding related scope may satisfy the selected HTML. Preserve host, query and path case in identity; do not percent-decode URI identity into filesystem paths. CID references match Content-ID only, not a Content-Location whose text starts with `cid:`. Raw Content-ID headers retain literal identifiers; only the HTML cid URL is decoded, exactly once, following [RFC 2392 section 2](https://www.rfc-editor.org/rfc/rfc2392.html#section-2). A literal `%2F` header differs from `/`. Output-safe filenames remain separate from resource identity.

Exact URI/CID matching has priority. Compatibility fallback is limited to a bare relative filename without query or fragment and one uniquely named resource, with case preserved. A missing absolute URI, CID or qualified path cannot discard its identity through basename fallback. An actually referenced ambiguous identity or basename fails rather than selecting arbitrary image bytes; unused conflicting resources do not block unrelated content.

An explicit empty query or fragment still qualifies a reference: missing `p.png?` or `p.png#` cannot silently use a similarly named resource from another directory. Exact matching still ignores fragment text as described above.

## Local HTML base and missing resources

Resolve the source HTML file URI, then its first base element with an href, then each image source, before converting a file URI into a platform path. This includes relative bases, explicit empty bases, absolute file bases, spaces and Unicode. Preserve POSIX roots and Windows drive/UNC policy. A root-relative reference already supplying `/C:/` or another drive keeps its own drive; only a drive-less reference inherits the base drive. Network image URLs remain links and are not downloaded.

Companion-folder fallback applies only to ordinary relative file references without an explicit base. It cannot replace the identity chosen by an explicit base, even when a similarly named sibling or companion image exists. Missing, unreadable or unmappable local resources retain their original links when image links are enabled, and add one `HTML-LOCAL-IMAGE-UNAVAILABLE` warning with the total occurrence count and at most three image positions. Diagnostic text never repeats unbounded paths or source values. Omit without OCR skips resource reads and missing-image warnings. OCR still receives needed bytes, and unavailable-input diagnostics accurately describe omitted links and OCR failure. Deliberately omitted images do not claim to have been packaged.

## Authored text and generated image links

HTML, MHTML and ENEX image projection must preserve body/code text that resembles internal placeholders. Use a per-conversion namespace checked against authored text, decoded HTML/ENML text and relevant attributes. Restore generated placeholders in one pass so a generated link cannot be processed by a later replacement. Omit/disabled-image modes remove the image occurrence, not authored text. ENEX resource validation derives references only from parsed en-media elements and actual hash attributes. Comments, CDATA and attribute lookalikes cannot register dependencies. ENML uses an Expat-backed TreeBuilder with external entity resolution and parameter-entity parsing disabled; internal entity declarations are rejected before expansion. Standard external ENML doctype identifiers are not fetched. Common named HTML characters are normalized outside comments/CDATA, while literal code is preserved. Actual missing resources still fail, including in omit mode.

ENEX keeps each note's hash-to-resource binding until the complete merged document has been built. Images, non-image attachments and OCR Markdown sidecars allocate names from one request-owned inventory, including names allocated by earlier notes. Different bytes with the same source filename remain distinguishable in final links; repeated occurrences of the same bytes may share a resource. Finalizer relocation does not infer note ownership from filenames.

ENEX 在合并文档完成前保留每条 note 的 hash→资源绑定。图片、普通附件和 OCR Markdown 侧车共享本请求的命名库存，包含前面 note 已分配的名称。同源文件名但不同字节必须在最终链接中保持可区分；相同字节的重复出现可合法复用资源。Finalizer 不通过文件名猜测 note 归属。

The Layout HTML helper uses the same pure Core namespace/restoration functions. Its final helper output is covered by regression tests; there is currently no production caller of that helper, so these tests do not establish a PDF/Office route or native host result.

## Regression boundaries

`test_web_archive_review_boundaries.py` executes real MIME/HTML/ENEX inputs through Runtime and Finalizer, checking selected body text, exact image bytes, parent-relative identities, nested roots, bounded errors, unique aliases, explicit-base decoys, missing-resource warnings and body/code/omit preservation. Existing `test_mhtml_resource_identity.py` retains host/query/path/CID, URI, source-byte and image-order regressions. Core tests cover namespace collisions and nonrecursive replacement; Layout tests exercise its final HTML-to-Markdown helper.

`test_mhtml_query_resolution.py` checks synthetic/absolute base pairs through actual Runtime outputs, with the correct PNG alone and alongside a conflicting sibling. Its fragment, query replacement/clearing, empty delimiter and relative-path cases check the emitted image bytes rather than accepting an identity-helper result alone.

Structured clipboard XLSX overflow tests additionally preserve an existing same-name, previously published workbook byte-for-byte under a frozen publication identity. Final XLSX save checks include already instantiated retained template/copied-sheet text without expanding sparse ranges or retyping formulas. Source/template/input hashes and original failed runs retain their own evidence status; source regressions do not certify a package or physical host.
