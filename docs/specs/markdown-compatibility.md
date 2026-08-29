# Markdown compatibility / Markdown 兼容

DocWen supports a documented Markdown subset for conversion to and from office formats. Route options freeze syntax, link, image, numbering and template behavior per request.

DocWen 为 Office 双向转换支持一组明确的 Markdown 子集。语法、链接、图片、编号和模板行为在每个请求开始前冻结。

## Supported structures / 支持结构

- headings, paragraphs, emphasis, code, quotes and horizontal rules;
- ordered and unordered lists with nested continuation handling;
- tables, including the documented merge markers;
- links, wiki links, embedded resources and request-scoped link policies;
- YAML front matter and template field projection;
- inline/block formulas and supported note forms;
- local images, optional Base64 output and OCR sidecars where the route declares them.
- authored `Figure:`, `Table:`, `Equation:`, and `Code:` declarations with distinct DOCX paragraph styles, plus
  structure-owned Heading and caption semantic targets;
- one exclusive `{{ bibliography }}` template paragraph when an already-presented typed bibliography resource is supplied.

ATX Heading levels 1..6 follow CommonMark; levels 7..9 are DocWen extensions that map one-to-one to Word
`Heading 7`..`Heading 9`. Ten or more leading `#` markers remain visible paragraph text. DOCX import resolves both
direct `outlineLvl=0..8` values and outline levels inherited through paragraph styles.

## Rules / 规则

- Parsers must not fetch remote content implicitly.
- Missing or unsafe local resources produce diagnostics instead of traversal.
- Source syntax and target rendering are compared semantically unless exact text is the contract.
- Unsupported constructs remain visible as text or emit a warning; they must not disappear silently.
- Configuration and per-request options have one precedence chain and no process-global fallback after admission.

## Obsidian extension interoperability / Obsidian 扩展互通

The round-trip contract is: an unchanged DOCX recovers the authenticated authored Markdown byte-for-byte; after a
DOCX edit, DocWen preserves the represented semantics and writes canonical Markdown. Exact recovery is never claimed
from visual similarity alone. It requires an authentic resolved-numbering v4 projection plus the adjacent owned
`<document>.docwen` artifact. That artifact has media type
`application/vnd.docwen.round-trip-sidecar+zip`, schema `docwen.round_trip_sidecar.v1`, and exactly four ZIP members
in order: `authored-source.md`, `neutral-document.json`, `numbering-export-plan.json`, and `manifest.json`. The
canonical manifest binds the exact DOCX byte count/SHA-256 and the byte count, SHA-256, and strict media type of each
of the three payloads. The producer fixes member order, timestamps, permissions, comments, and storage encoding for
deterministic bytes. The reader rejects extra or duplicate members, absolute or parent paths, links, encryption,
non-canonical metadata, and oversized or unsafe compression. The whole sidecar is additionally integrity-pinned by
Artifact Bundle v2. A missing, foreign, invalid, stale, linked, structurally unsafe, oversized, or DOCX-mismatched
sidecar falls back to canonical semantic recovery with an explicit diagnostic. UTF-8 BOM, CRLF/LF choice, blank
lines, and final-newline presence are therefore retained only on the authenticated unchanged path. A Word edit
disables byte-exact recovery before the old source can be considered; it does not prevent authenticated semantic
normalization when the semantic carriers still prove.

往返合同为：DOCX 未修改时，恢复经过认证的 Markdown 原文字节；在 Word 中修改后，保留可表达的语义并
输出规范化 Markdown。系统不会因视觉相似就声称精确恢复。精确路径要求 DOCX 中有效的 resolved-numbering
v4 投影，以及相邻且归属明确的单文件 `<document>.docwen` ZIP artifact。它的媒体类型为
`application/vnd.docwen.round-trip-sidecar+zip`、schema 为 `docwen.round_trip_sidecar.v1`，并按固定顺序仅含
`authored-source.md`、`neutral-document.json`、`numbering-export-plan.json` 与 `manifest.json`。manifest 绑定
DOCX 及三份数据的字节数和 SHA-256；Artifact Bundle v2 另行绑定整个 sidecar。sidecar 缺失、外来、损坏、
过期、不安全、过大或与 DOCX 不匹配时，系统会带明确诊断退回语义规范化。Word 编辑会先关闭逐字恢复，
不会套用旧源码；只要语义载体仍通过认证，就继续规范化恢复。因此 BOM、换行风格、空行和文末换行只在
经过认证的未修改路径中逐字节保留。

DocWen accepts the Structural Tables pipe-table dialect in addition to ordinary GFM tables:

- consecutive equal-width rows before the delimiter are column-header rows;
- one adjacent `||` inside the delimiter marks the columns to its left as row headers and adds no column;
- an exact `<` merges left and an exact `^` merges up; `\<` and `\^` are literal cell text;
- escaped pipes and pipes inside code spans do not split cells; and
- invalid widths or structures remain visible source text instead of being guessed.

DOCX export maps these roles and merge rectangles to native table semantics. DOCX import emits the canonical
Structural Tables spelling when native table metadata requires multiple column-header rows or row-header columns;
ordinary tables remain ordinary GFM. Number Suite interoperation is consumer-neutral: an Obsidian adapter supplies
authenticated heading/caption/reference facts and their effective displayed counters in DocWen's resolved document
and exact-two numbering plan. DocWen does not import Number Suite code, scan a Vault, or infer numbers from visible
prefixes.

除普通 GFM 表格外，DocWen 还接受 Structural Tables 管道表格语法：分隔行前连续且等宽的行是多行列表头；
分隔行内唯一相邻的 `||` 标记其左侧为行表头且不增加列；严格匹配的 `<` 向左合并、`^` 向上合并，`\<` 与
`\^` 表示字面量；转义管道和代码跨度中的管道不会切分单元格；无效宽度或结构保持为可见源码而不猜测。
DOCX 导出把这些角色与矩形合并映射为原生表格语义；导入在原生表格元数据要求多行列表头或行表头时输出
规范 Structural Tables 写法，普通表格仍输出普通 GFM。Number Suite 互通通过消费者无关数据完成：Obsidian
适配器提供经过认证的标题、题注、引用和实际显示计数；DocWen 不导入 Number Suite 代码、不扫描 Vault，
也不从可见前缀猜测编号。

## Anchors and semantic targets / 锚点与语义目标

DocWen's canonical Markdown is self-contained and has no Pandoc dependency. Pandoc-style `{#id}` attributes are not
canonical input or output. A stable Markdown block anchor is the Obsidian-compatible suffix `^id`; its ID is 1 to
128 ASCII characters from `[A-Za-z0-9-]`. All IDs share one document-local, case-sensitive namespace. ID spelling
does not encode kind: `fig-`, `tbl-`, `eq-`, and `code-` are neither reserved nor interpreted specially. IDs
containing `.`, `_`, or `/`, empty IDs, and IDs longer than 128 characters are invalid. IDs are generated only when
a user operation requires one. An external record, file, workspace, or node identity is never projected as a
Markdown anchor merely because it exists.

Placement follows the [Obsidian block-link rules](https://obsidian.md/help/links). Only a simple paragraph or a
paragraph containing exactly one Markdown image/Obsidian image embed uses one space plus `^id` at the end of its
content line. An individual list item may likewise carry the ID on that item. A table, block quote, callout, whole
list, `$$` display-math block, or fenced code block uses a following anchor-only line at the same CommonMark
container path and sibling depth. Canonical output places one blank line between the complete block and the marker;
the parser accepts zero or more blank lines, including at EOF, because blank lines do not create an intervening
block. A non-empty intervening block or container/depth change breaks attachment and produces a dangling-anchor
diagnostic. An ID is never appended to the closing `$$` or closing code fence. The parser records the complete
list/quote/callout container path. A caption declaration retains its own inline `^id`; an object's distinct ordinary
anchor remains attached to the object.

An ordinary anchor identifies only its Markdown block:

```md
![[image.png]] ^image-system
```

It participates in ordinary Markdown navigation or embedding through `[[Page#^image-system]]` or
`![[Page#^image-system]]`. It is not a Figure, is not addressable as a semantic cross-reference, and never enters a
caption counter or Word bookmark/`SEQ`/`REF` projection. The same rule applies to ordinary anchors on paragraphs,
tables, equations, fenced-code blocks, lists, block quotes, and callouts. DocWen preserves this source navigation for
Markdown round-trip; it does not manufacture a Word-native target where the source has no semantic owner.

Every object in the closed source-oracle `anchors` array additionally requires `container_path`. It is an ordered
outermost-to-innermost array of closed source-only segments. A segment has exactly `block_kind` and `block_range`;
`block_kind` is one of `list`, `list_item`, `block_quote`, or `callout`, and `block_range` is exactly
`{start,end}` in the projection source's zero-based Unicode-code-point, exclusive-end coordinate system. A top-level
anchor has `container_path=[]`. The path contains only strict structural container ancestors in this dialect, never
the anchor's own block, a file/Workspace path, a consumer ID, or a renderer ordinal. If the anchor's own `block_kind`
is one of those four container kinds, its `block_range` is the complete structural-container range, including every
continuation block; the identical container range is used in every descendant's corresponding path segment. An
inline multi-paragraph `list_item` anchor therefore never contracts its owner range to the first source line.

For topology, an anchor's *owner path* is a comparison-only sequence: its stored `container_path` pairs followed by
one terminal owner coordinate made from that anchor record's own `block_kind` and `block_range`. The terminal
coordinate is not another serialized `container_path` segment, and its `block_kind` uses the complete ordinary-anchor
kind enum rather than the four container-segment kinds. Ordinary anchor P is the direct source parent of ordinary
anchor C exactly when P's owner path is the longest proper prefix of C's owner path among ordinary anchors in the
authenticated source.
Source ranges along a path are laminar and may be strictly contained or equal; equal ranges do not collapse distinct
structural levels. Crossing ranges, a self segment, an unknown segment kind, or a parent inferred only from range,
block kind, or output order fails closed. This source-only path is how the neutral runtime derives nesting before any
DOCX element exists.

DOCX recovery has one exact resource-less image carrier: `![image omitted]()`. The ASCII-lowercase alt text
`image omitted` and empty destination are both fixed; no other empty-destination image is this carrier. When an
authenticated DOCX Figure or ordinary image anchor owns an image but `preserve_resources=false`, recovery emits this
carrier in the original image-owner position instead of coupling resource preservation to source validity. A Figure
declaration binds the carrier as its image object, while an ordinary image ID remains on the same carrier line, for
example `![image omitted]() ^image-system`. The lossless authored-source parser also recognizes the exact token as an
`image` block without performing a filesystem, Workspace, or network lookup. Close spellings remain ordinary authored
Markdown and do not receive recovery semantics.

The carrier asserts only that an authenticated image owner was recovered without an exported resource. It is not a
resource locator, fragment locator, OCR sidecar link, or invitation to reconstruct bytes. With `main_md`, any OCR
presentation follows the carrier and cannot take ownership of its ID. With `image_md`, the primary Markdown retains
the carrier and all Figure/ordinary-anchor ownership; the sidecar contains OCR presentation only. Moving the carrier,
declaration, or ordinary ID into the sidecar is an owner change and is forbidden. Each successful recovery that emits
one or more carriers adds the artifact-bound warning `DOCX2MD-IMAGE-OWNER-RESOURCE-OMITTED`; because the warning
describes output recovery rather than an authored-source defect, it has no Markdown `range`, `fixes`, or source
evidence claim.

A captioned semantic object exists when one exact declaration has one unique adjacent captionable object. The
semantic declaration kind (`Figure`, `Table`, `Equation`, or `Code`) and the carrier's native structure are
independent: any declaration kind may own an image paragraph, native table, block equation, or fenced code block.
The declaration may be above or below the carrier in authored input, with zero or one blank line between them; two
or more blank lines break ownership. A carrier on both sides of one declaration, or declarations on both sides of
one carrier, is ambiguous and fails closed. Canonical DOCX recovery always writes the declaration above the carrier.

Its ID is optional, except that Equation and Code with empty caption text must have an ID. The ID may be the trailing
token on the declaration line or an ID-only line immediately after it; a blank or intervening line is not the same
target. In `resolved_document.v1`, the authenticated target range remains the short physical declaration line and
the adapter separately authenticates the immediately following ID line against the typed `target_id` and complete
source hash. Only a declaration that carries an ID is addressable by `@[[...#^id]]`, has a Word bookmark, or receives
an entry in the semantic-target map:

```md
Figure: System architecture ^system-architecture

![[image.png]]
```

The four closed kinds are:

| Declaration | IR kind | Caption text and ID rule |
|---|---|---|
| `Figure:` | `figure` | non-empty text; type-neutral ID optional |
| `Table:` | `table` | non-empty text; type-neutral ID optional |
| `Equation:` | `equation` | text may be empty; an empty-text declaration requires a type-neutral ID |
| `Code:` | `code_block` | text may be empty; an empty-text declaration requires a type-neutral ID |

`Equation: ^energy` and `Code: ^snippet` are therefore valid; bare `Equation:` and bare `Code:` are not. Figure and
Table always require visible caption text. A caption without an ID is still styled and numbered,
but is not an addressable target. The declaration binds the complete declaration-plus-object pair.
`@[[#^system-architecture]]`, for example, is a semantic cross-reference, while
`[[Page#^system-architecture]]` is ordinary navigation whose source and Word landing position is the declaration
line. If the raw object also needs its own navigation target, it may carry a different ordinary ID; DocWen does not
generate that second ID by default and the declaration ID must not be repeated on the object.

A Heading with an inline or immediately following ID-only `^id` is also an addressable semantic target. A Heading
without an ID remains a heading and may be numbered, but is not addressable through the stable semantic-reference grammar. Paragraph and raw-object
anchors never become semantic targets merely because their spelling resembles a historical typed prefix.

After stripping only an already-parsed CommonMark container prefix, the declaration matcher recognizes exactly the
four ASCII-case-insensitive words `Figure`, `Table`, `Equation`, and `Code`, immediately followed by the literal
colon. It does not use a prefix match: `Figures:`, `Figure :`, a bare `:`, `Listing:`, and `List:` are ordinary source,
not declarations. A normal lossless save preserves the authored keyword spelling; newly created or explicitly
canonicalized declarations write exactly `Figure:`, `Table:`, `Equation:`, or `Code:`. `listing`, `lst-`, a `list`
semantic kind, and `list-` are not accepted as a second canonical vocabulary.

### Links, cross-references, and citations / 链接、交叉引用与引文

The parser preserves five observable source constructs as different IR variants:

- `[[Page#^id]]` is an ordinary navigation link;
- `![[Page#^id]]` is an ordinary embed;
- `@[[#^id]]` and `@[[Page#^id]]` are stable-ID semantic cross-references;
- `@[[#Heading]]` and `@[[Page#Parent#Heading]]` are soft Heading-path semantic cross-references; and
- narrative `@citation-key` and parenthetical `[@first; @second]` are citations.

The soft form contains one or more non-empty authored Heading segments and never resolves to a caption or ordinary
anchor. Within the supplied document, DocWen resolves it only when the exact path selects one Heading: zero matches
is `missing`, multiple matches is `ambiguous`. A cross-document soft selector stays `external_unresolved` until its
external owner supplies one neutral resolution record; DocWen never chooses the first Heading or scans another file.

Semantic cross-references allow an optional Alias, for example `@[[Page#^id|Short title]]` or
`@[[Page#Parent#Heading|Short title]]`. When the resolved numbering/export plan supplies a materializable number,
presentation contains that derived number and then the authored Alias; Alias never replaces the number or rewrites
the target. If numbering is disabled or a Heading-level template is empty, the same semantic reference fails as
`docwen.markdown.cross_reference.unnumbered_target`; it does not display Alias without a number.
Target kind comes from the resolved Heading or bound caption declaration, never from the ID or reference spelling.
A semantic reference to an ordinary anchor fails as `docwen.markdown.cross_reference.non_semantic_target`; it does
not silently degrade to an ordinary WikiLink. `@[[...]]`, ordinary WikiLinks/embeds, and citations are recognized
before generic link preprocessing and retain authored-source ranges. They are not recognized in code spans/blocks,
raw URLs, or HTML attributes.

A citation key is 1 to 128 ASCII characters matching `[A-Za-z0-9][A-Za-z0-9_-]*` and is case-sensitive. The `@`
before a narrative citation must not be immediately preceded by an ASCII letter, digit, or one of `._%+-`, so an
email local-part is not parsed as a citation. The authored key is a mutable lookup key, not a stable bibliography,
node, or provider record identity.

### Consumer-neutral resolution boundary / 消费者无关解析边界

DocWen's direct Markdown parser resolves only targets contained in the supplied document. It does not scan sibling
files, interpret a Workspace page name, resolve a consumer's WikiLink, choose a bibliography scope, or search for a
citation key. The source projection keeps authored syntax separate from resolution facts. A semantic-reference
record has `selector_kind=stable_id|heading_path`; optional authored `page_locator`; exactly one of stable
`target_id` or a non-empty ordered `heading_path`; optional authored `alias`; and `resolution_status` exactly
`resolved`, `missing`, `ambiguous`, `non_semantic`, `unnumbered`, or `external_unresolved`. Resolved document
identity, resolved kind, and cached number remain separate optional facts. A resolved soft selector always has
`resolved_kind=heading`:

This semantic-target rule is separate from direct raw-file image materialization. For CLI/GUI conversion of a raw
Markdown file, a short image basename may resolve only beside that file, in its sibling same-name directory, or in
configured source-local convention directories (`.`, `assets`, `images`, and `attachments` by default). The resolver
supports spaces but performs no recursive, Workspace, Vault, parent-tree, CWD, or drive search. Absolute or traversal
search directories cannot enlarge this boundary. Explicit authored absolute and relative paths remain exact locators.
Machine `neutral_document` requests do not use this fallback at all: their authenticated resource bytes and Unicode
occurrence ranges are complete, and DocWen never consults the source filesystem for them.

- a semantic target records explicit `kind`, optional `id`, source form, title/caption content, and optional derived
  number;
- a cross-reference records selector kind/payload, authored page locator and Alias, resolved document identity,
  resolved kind, resolution status, and cached number as separate fields; and
- each citation item records authored key, optional resolved stable record identity, lossless token/range, and
  presentation data separately. A citation occurrence also records `narrative` or `parenthetical` form.

An adapter that owns an external resolver must lower its result into these consumer-neutral records before invoking
the neutral DocWen conversion API. DocWen validates the supplied input identity and content SHA-256 but neither
replays nor imports the external resolver. Page paths, workspace/node IDs, private domain types, and resolver
instructions are forbidden neutral inputs. Machine options never carry semantic context as opaque strings. If a
Machine route needs external resolution data, that data requires a separately versioned typed resource and complete
Machine/schema/fixture/consumer re-freeze; no such undeclared input is accepted by this baseline.

### Derived numbering authority / 派生编号权威

Only Structured Numbering creates numbers. A number-like prefix inside authored Heading text is opaque title text:
`## 2.3 标题` has the complete title `2.3 标题`. Normal parse, Live/Read presentation, cross-reference resolution,
export, and import never split, delete, hide, migrate, or reuse `2.3`, `第二章`, `一、`, Roman numerals, or localized
lookalikes as a number. When automatic Heading numbering is enabled, the derived prefix and the complete authored
title may both be visible. Figure/Table/Equation/Code declarations likewise have no authored-number source; a
derived number is never written into declaration content.

For each of `heading`, `figure`, `table`, `equation`, and `code_block`, enabled means that the upstream numbering
owner supplies a materializable derived number; disabled means none exists. An empty selected template disables that
Heading level. There is no hidden-display mode that retains a semantic number. Toggling any state or profile causes
zero Markdown changes. Ordinary WikiLinks continue to navigate an unnumbered target, while semantic `@[[...]]`
fails with the unnumbered-target diagnostic.

DocWen consumes a provider-neutral resolved document plus resolved numbering/export plan under the separate
[Resolved structured numbering and export plan](structured-numbering-phases.md). The upstream semantic provider owns profile selection,
counter scope/reset/format/label/chapter inclusion and WikiLink resolution. DocWen neither parses private consumer models nor guesses a
counter from authored text. The source-authoring controls `remove_numbering`, `add_numbering`, `numbering_scheme`, and
regex Heading cleanup belong only to capabilities that declare them. They are rejected on the resolved-plan route
and are not inputs to that Conversion Port.

Provider-neutral authored Markdown Heading targets use ATX levels 1..9; levels 7..9 are DocWen extensions to
CommonMark. Heading target materialization, caption restart binding, and chapter-number binding accept the same 1..9
range and map to Word `Heading 1`..`Heading 9`.

### Fenced code serialization / 围栏代码序列化

A CommonMark fenced code block always remains one code block. Fence-looking lines inside its literal body do not
create nested blocks, nested named styles, or a second `Code Block`; only an external `Code:` declaration may create
the separate `Code Block Caption` paragraph. The body itself uses the existing `Code Block` style.

When a lossless Markdown source fence is available and remains safe for the current literal body, a normal save
preserves its authored fence character, length, indentation/container prefix, info string, and closing spelling. It
does not normalize a safe authored backtick fence to tildes, or vice versa. When a new fence must be serialized, or
an authored fence would conflict after an edit, the writer computes both candidates independently:

- opener and closer use the same character, either backtick or tilde;
- the length is at least three and strictly greater than every potentially closing run of that same character in
  the literal body at the relevant CommonMark container depth; and
- candidate selection is deterministic: choose the shorter safe fence, break equal length in favor of backticks,
  and retain that choice for both opener and closer.

Literal runs of the other fence character do not increase a candidate's length. If the body contains both fence
characters, the same comparison still chooses one safe candidate; the body bytes remain unchanged. List and
blockquote container markers/indentation are applied outside the chosen fence and may not cause an inner literal
run to be reparsed as a sibling fence. Markdown→DOCX→Markdown must return exactly one fenced code block with the
same literal body and, where a lossless source form was retained, the same safe authored fence.

### Fenced source occurrence authority / 围栏源码出现权威

Every authored fenced block also produces one required `fenced_sources` record in
`docwen.markdown_semantics.v3`, including unanchored blocks, captioned Code objects, ordinary-anchored blocks, and
specialized Mermaid/query/view forms. This occurrence carrier does not change the source IR kind: a specialized
block remains `fenced_block`, while the DOCX ordinary-anchor projection remains the closed `code_block` kind. It is
not an ordinal match inferred after AST normalization. The parser binds the complete Unicode-code-point source range
before generic Markdown or link processing and authenticates the exact opener, info suffix, logical body, per-line
container prefixes, closer, and EOLs.

Each record is closed and has this canonical member order:
`tag, source_sha256, source_start, source_end, identity_sha256, block_sha256, body_sha256, fence_character,
opening_length, opening_prefix_b64, info_b64, opening_eol, body_prefix_count, body_prefixes_b64, closing_state,
closing_length, closing_prefix_b64, closing_suffix_b64, closing_eol`. Decimal values are canonical non-negative
integers. Hashes are lowercase SHA-256 over UTF-8. The five `*_b64` values are strict RFC 4648 base64 of exact UTF-8
text except `body_prefixes_b64`, which encodes the NUL-separated exact prefix for each logical body line;
`body_prefix_count` is authoritative. EOL values are only empty, LF, or CRLF. `closing_state` is `present` or
`omitted_eof`; the latter requires zero closing length and empty closing prefix, suffix, and EOL, so recovery never
manufactures a closer at EOF.

`block_sha256` authenticates the complete authored slice `[source_start,source_end)` and `body_sha256` authenticates
the exact de-containerized logical body, including its mixed EOLs and final-EOL state. The identity preimage is exact
UTF-8
`docwen-fenced-source-map-v1\0<source_sha256>\0<source_start>\0<source_end>\0<block_sha256>\0<body_sha256>`.
`identity_sha256` is its SHA-256 and `tag` is `docwen-fenced-source-v1:` plus its first 32 lowercase hex characters.
Records sort by `(source_start,source_end,tag)`, may not overlap, share one source SHA within a document, and have
unique tags. Decoded small fields are bounded to 16,384 bytes; body-prefix payload to 1,048,576 bytes; prefix count
and either fence length to 65,536; the map to 16,384 records and 16,777,216 decoded payload bytes. Non-canonical
base64, decimal, EOL, prefix, ordering, range, hash, or bound fails closed.

DOCX uses a separate custom XML map with namespace
`https://docwen.dev/schema/document-fenced-source-map/v1`, root
`documentFencedSourceMap version="1"`, and one empty `fencedSource` element per record using exactly that attribute
order. It uses the same independently allocated `itemN.xml`/`itemPropsN.xml`/relationship/content-type topology,
canonical byte framing, deterministic properties UUID, and collision rules as every owned map. Every record has
exactly one inline `w:sdt` tagged with its `tag` inside the Code Block paragraph; that SDT contains the complete
visible payload and no run outside it. It may nest under an ordinary-anchor SDT and/or addressable Code target SDT.
The fenced map and inline SDT contain no source ID, target ID, anchor ID, bookmark name, target kind, or consumer
identity and create zero bookmark, `SEQ`, or `REF` facts by themselves.

Import proves closed package topology, canonical map bytes, one-to-one tag ownership, exact non-overlapping source
inventory, complete visible body topology/hash, and reconstruction of the complete authored block. Missing,
duplicate, unmapped, swapped, overlapping, partially wrapped, moved, host-stripped, or tampered data fails closed;
the reader does not fall back to AST ordinals or synthesize a fence. Direct/top-level, blockquote, and list-container
forms must round-trip their exact prefixes and LF/CRLF/omitted-EOF state. If a full container form cannot satisfy this
authority, conversion is unsupported rather than weakened.

### Diagnostics and single-document rewrites / 诊断与单文档重写

Every source diagnostic uses the Machine evidence identity `docwen.machine.diagnostic_evidence.v1` (schema ID
`urn:docwen:schema:machine-diagnostic-evidence:v1`). The diagnostic object adds the optional members
`evidence_schema`, `source`, `range`, `related_ranges`, and `fixes`. If any one is present, `evidence_schema`, `source`,
and `range` are all required; the identity value is exact, `related_ranges` contains at most 16 ranges, `fixes`
contains at most 8 fixes, and every nested object is closed. Diagnostics without these members retain the base
Machine diagnostic shape and make no source-coordinate or fix claim.

`source` contains the exact `input_id`, lowercase 64-hex content SHA-256, `encoding="utf-8"`,
`coordinate_system="unicode_code_point"`, `offset_base=0`, and `range_end="exclusive"`. Primary and related ranges
are `{start,end}` pairs into the authenticated authored source. A primary range is non-empty; an ID-token range
includes its leading `^`, and a semantic cross-reference range includes its leading `@`. Missing required Figure or
Table content, or a missing ID on empty Equation/Code, selects the non-empty declaration keyword without its colon.
Only a fix edit insertion may be zero-width.

Fixes are closed objects `{fix_id,edits}`. `edits` contains 1..16 ordered, non-overlapping objects
`{range,replacement}`; `replacement` is at most 4096 Unicode code points and every range uses the diagnostic's same
authenticated source coordinates. The consumer must recheck `input_id`, source SHA-256, range bounds, non-overlap,
and applicability before staging all edits, then reparse and commit all or none. Diagnostics for this semantic
series use the following exact codes:

| Diagnostic code | Condition and primary range | Permitted fix-it ID |
|---|---|---|
| `docwen.markdown.anchor.dangling` | an anchor-only line has no attachable preceding block in the same container/depth; range is that `^id` | none; DocWen does not search farther backward |
| `docwen.markdown.anchor.duplicate` | an ID already has an owner in the shared namespace; primary range is the later `^id`, with the first owner as related range | `docwen.markdown.fix.rename_anchor` when the caller supplies one conflict-free replacement |
| `docwen.markdown.anchor.invalid_id` | a `^id` has an empty suffix, illegal character, invalid placement, or exceeds 128 characters; range is the complete candidate token | `docwen.markdown.fix.rename_anchor` only with one preselected valid replacement |
| `docwen.markdown.caption.content_required` | Figure/Table has empty trimmed caption content; primary range is the non-empty declaration keyword | none; DocWen does not invent prose |
| `docwen.markdown.caption.object_mismatch` | no unique adjacent captionable object exists within zero or one blank line, including two-sided ambiguity or an intervening block/container boundary; primary range identifies the declaration/boundary | none; DocWen does not guess ownership |
| `docwen.markdown.caption.empty_equation_target_required` | Equation or Code has both empty trimmed caption content and no ID; primary range is the declaration keyword | `docwen.markdown.fix.add_semantic_id`, whose insertion edit alone has a zero-width range at declaration end |
| `docwen.markdown.cross_reference.missing` | the supplied locator has no target; range is the complete `@[[...]]` | `docwen.markdown.fix.add_semantic_id` only when one intended ID-less semantic object is already uniquely selected by the caller |
| `docwen.markdown.cross_reference.ambiguous` | a same-document soft Heading path selects more than one Heading; range is the complete `@[[...]]` and related ranges identify all matching Heading declarations | none; the caller must select a full path or establish a stable ID |
| `docwen.markdown.cross_reference.non_semantic_target` | the resolved ID belongs only to an ordinary block anchor; range is the complete `@[[...]]` | `docwen.markdown.fix.move_anchor_to_declaration` only for one adjacent, matching, ID-less declaration; otherwise none |
| `docwen.markdown.cross_reference.unnumbered_target` | the resolved semantic target has no materializable number under the supplied numbering plan; range is the complete `@[[...]]` | none; it does not degrade to an ordinary link |
| `docwen.markdown.cross_reference.alias_stale` | Alias differs from current resolved title/caption; warning over the complete Alias | none; an explicit custom title remains valid |

All codes above have severity `error` except `alias_stale`, which is `warning`. There is no generic “caption ID
missing” diagnostic: Figure/Table/Equation/Code captions may omit an ID under the rules above. A fix is offered only
when the declaration, object, replacement ID, and supplied in-document reference set are unambiguous. A DocWen
rewrite plan is limited to one supplied Markdown document. It may update a declaration, Heading, ordinary anchor,
and matching `@[[...#^id]]`/`[[...#^id]]`/`![[...#^id]]` occurrences in that document. It does not scan a Workspace,
resolve `Page`, or rename other files; an external multi-document transaction owner must apply its own resolver and
journal. Neither layer may convert an ordinary object anchor into a semantic target by inference or generate anchors
for unrelated blocks.

The provider mapping is exact and one-to-one:
`interop.cross_reference.unnumbered_target` maps only to
`docwen.markdown.cross_reference.unnumbered_target`, preserving error severity, authenticated reference range,
resolved target identity/kind, and the absence of a fix; the reverse adapter maps it back to that sole interop code.
Plan admission instead uses the non-source codes `docwen.numbering_export_plan.missing`,
`docwen.numbering_export_plan.invalid`, and `docwen.numbering_export_plan.unsupported_materialization`. Those codes
have no Markdown range/fix and must never be coerced into `unnumbered_target`: disabled is a valid plan state, while
missing, malformed, contradictory, or non-portable plan input is not.

An adapter mapping is total rather than string passthrough. It maps the eleven diagnostic and three fix suffixes above
one by one, preserves severity, source identity/hash, primary/related ranges and edit preconditions, and has no
default branch. Consumer-only resolver diagnostics do not enter this table. An unknown code/fix or invalid evidence
envelope fails the structured handoff instead of being coerced or dropped. Coordinate conversion, when needed by a
consumer, happens only after evidence validation; DocWen never reports UTF-8 byte offsets, UTF-16 units, inferred
line/column locations, normalized text, or inclusive ends.

### DOCX projection / DOCX 投影

An addressable semantic target uses the deterministic 40-character Word bookmark name `DW_T_<digest>`, where `<digest>`
is the first 35 lowercase hexadecimal characters of SHA-256 over the exact UTF-8 byte sequence
`docwen-target-map-v1\0<kind>\0<source-id>`. `<kind>` is exactly `heading`, `figure`, `table`, `equation`, or
`code_block`; the
separators are single NUL bytes and the source ID is not normalized. Before rendering, DocWen compares complete
64-hex hashes, source IDs, exact bookmark names, and Word's case-insensitive bookmark namespace. A full or truncated
digest collision, pre-existing bookmark conflict, duplicate mapping, or inconsistent metadata fails before package
commit; no suffix guessing is allowed.

For an enabled simple caption target, the bookmark is in the caption paragraph. Its range begins immediately before
that caption's `SEQ` complex field and ends immediately after the field end, so it encloses the field and its cached
numeric result only—not the localized label, punctuation, caption text, or object. An enabled chapter caption instead
encloses the complete `STYLEREF` field, chapter separator run, and `SEQ` field as one contiguous bookmark; both cached
field results plus the separator exactly equal the plan's derived number. An ID-bearing disabled caption
retains navigation through the same deterministic bookmark name, but that bookmark is zero-width at the start of the
caption paragraph before all visible runs and there is no `SEQ`. The four field counter names are exactly `Figure`,
`Table`, `Equation`, and `Code`. A Heading target binds its bookmark to the heading paragraph; when enabled, it also
binds the plan-supplied list/heading numbering. A numbered
Heading cross-reference uses the Word equivalent of `REF DW_T_<digest> \n \h` so only the heading number is
materialized. Other semantic cross-references use `REF DW_T_<digest> \h` with a cached numeric result. Alias rich
runs are appended outside the REF field. An unnumbered semantic reference fails before rendering and creates no
`REF`. Ordinary navigation to a semantic ID lands in the declaration/Heading paragraph regardless of numbering
enablement. A caption or Heading without an ID may still be numbered, but creates no bookmark or target-map entry.

DOCX caption/object order is fixed by the declaration's semantic kind and is not selected by a profile or Export
Style. A Figure-labelled carrier is immediately followed by its `DocWenFigureCaption` paragraph. A
`DocWenTableCaption`, `DocWenEquationCaption`, or `DocWenCodeBlockCaption` paragraph is immediately followed by its
one native carrier. Semantic kind and carrier structure are independent: the carrier may be an image paragraph,
native table, block equation, or fenced code block; Figure plus a native multi-image table therefore remains a
Figure-labelled native table rather than becoming a Table caption or rasterized image. A fenced Code carrier's
consecutive Word paragraphs count as one logical object. Direct adjacency is measured between logical blocks; an
ordinary-anchor SDT that wraps the carrier is the same carrier, not an intervening block. DOCX-to-Markdown always
restores the canonical source order of declaration followed by carrier.

An ID-less caption is recovered only from that fixed direct adjacency and the one exact request-resolved managed
caption style authenticated by the caption-style binding map below. When enabled, it additionally requires exactly
one closed simple-`SEQ` or chapter-`STYLEREF`+separator+`SEQ` materialization and matching cached result under the
resolved numbering/export plan. A disabled ID-less declaration requires the independent occurrence authority below;
style plus adjacency alone is never proof of a declaration. An ID-less caption has no
`docwen-target-v1:` pairing SDT, target-map or reference-occurrence-map record, hidden/generated ID, bookmark,
hyperlink target, or `REF`. Import fails closed when the semantic-kind-defined DOCX order is reversed, a logical block
intervenes, the native carrier proof is absent, the style or `SEQ` kind/count/cached number is inconsistent, or one
caption/carrier participates in multiple pairing claims, including claims from both sides. It never guesses a pair,
coerces the carrier structure to the semantic kind, or invents an ID.

#### Caption-style binding map / 题注样式绑定表

Whenever a request renders at least one Figure/Table/Equation/Code caption, including an ID-less caption, it emits
one independent closed custom XML item in namespace
`https://docwen.dev/schema/document-caption-style-binding-map/v1`. This item authenticates request-local style
identity only. It is not a caption/object pairing claim, semantic target, source anchor, hidden ID, bookmark, or
authorization to infer a target. No such item is emitted when the request renders no caption.

The exact root is
`<documentCaptionStyleBindingMap xmlns="https://docwen.dev/schema/document-caption-style-binding-map/v1" version="1">`.
It has exactly four empty `binding` children in this fixed order: `figure_caption`, `table_caption`,
`equation_caption`, `code_block_caption`. Every child has attributes in exact order
`semantic_key,resolved_style_id,visible_name`; there are no other elements, attributes, text, or tails. The four
semantic keys, case-insensitively unique resolved style IDs, and case-insensitively unique managed visible names are
closed. XML byte framing, escaping, `itemN.xml`/`itemPropsN.xml`/relationships/content-types topology, and
deterministic UUID use the same rules as the target map, with this caption-style namespace in the UUID preimage.

On import, each binding must select exactly one `w:style` in `word/styles.xml` by exact `w:styleId`; that style must
be a paragraph style, contain exactly one direct `w:name` equal to `visible_name`, and contain no `w:aliases`.
Preserving an unrelated conflicting user style may leave the same visible name elsewhere; only the collision-free
resolved style ID is identity. Every addressable and ID-less caption paragraph must contain exactly one direct
`w:pPr/w:pStyle`, and its value must equal the binding selected by the caption's structural kind. Prefix matching,
the canonical requested style ID after it was displaced by a conflict, localized-name guessing, `SEQ`-kind
guessing, and cross-kind fallback are forbidden. A missing/extra/reordered map record, changed bytes/UUID/topology,
missing/duplicate/wrong-type/wrong-name/aliased style, absent/duplicate/mismatched `pStyle`, or style-map stripping
fails closed. The renderer binds the same complete four-entry request-local style table before it records any
caption and proves the reopened map and `styles.xml` before artifact registration.

#### Disabled ID-less caption occurrence authority / 禁用且无 ID 题注出现权威

Every disabled Figure/Table/Equation/Code declaration without an ID has exactly one independent block-level
`w:sdt` tagged `docwen-numbering-occurrence-v1:<digest32>`. Its `w:sdtContent` contains exactly the caption paragraph
and one authenticated native carrier in the semantic-kind-defined physical order above, with no third block. A
fenced Code carrier's paragraphs count as one carrier; an ordinary-anchor SDT may wrap the carrier and still occupies
that one slot. This
wrapper is source-recovery pairing only. It creates no target, hidden ID, bookmark, `SEQ`, `REF`, or number.

The full digest is lowercase SHA-256 over exact UTF-8
`docwen-numbering-occurrence-map-v1\0<source_sha256>\0<source_start>\0<source_end>\0<kind>\0false\0\0\0<plan_sha256>`;
`digest32` is its first 32 hexadecimal characters. Ranges use the authenticated source's zero-based Unicode-code-
point, exclusive-end system and cover the complete declaration-plus-object pair. `kind` is exactly
`figure|table|equation|code_block`; the two consecutive empty fields are the canonical null encodings for target ID
and derived number.

All records live in one custom XML item with namespace
`https://docwen.dev/schema/document-numbering-occurrence-map/v1`. The exact root is
`documentNumberingOccurrenceMap` with attributes in order `version,plan_sha256`, where version is `1` and the plan
digest matches both required input envelopes. Its only children are empty `occurrence` elements with attributes in
exact order
`tag,source_sha256,source_start,source_end,kind,enabled,target_id,derived_number,plan_sha256,sha256`.
`enabled` is exactly `false`; `target_id` and `derived_number` are exactly empty strings; `sha256` is the full digest.
Records sort by integer `(source_start,source_end)` then kind/tag and may not overlap or duplicate a range/tag. The map
is emitted if and only if at least one such occurrence exists and is never emitted empty. The zero values in the
following canonical item line are field-shape placeholders; a real writer recomputes every digest:

```xml
<documentNumberingOccurrenceMap xmlns="https://docwen.dev/schema/document-numbering-occurrence-map/v1" version="1" plan_sha256="0000000000000000000000000000000000000000000000000000000000000000"><occurrence tag="docwen-numbering-occurrence-v1:00000000000000000000000000000000" source_sha256="0000000000000000000000000000000000000000000000000000000000000000" source_start="0" source_end="20" kind="table" enabled="false" target_id="" derived_number="" plan_sha256="0000000000000000000000000000000000000000000000000000000000000000" sha256="0000000000000000000000000000000000000000000000000000000000000000"/></documentNumberingOccurrenceMap>
```

The map owns an independently allocated canonical custom-XML trio:
`/customXml/itemN.xml`, `/customXml/itemPropsN.xml`, and `/customXml/_rels/itemN.xml.rels`, plus one document
relationship and the two content-type Overrides. It uses the target map's exact relationship types, `rId`/`itemN`
collision rules, lexical Override insertion, UTF-8/no-BOM declaration-one-LF-single-root-one-final-LF framing, and
UUIDv5 formula, substituting this namespace as the sole `ds:schemaRef/@ds:uri` and UUID namespace preimage. No trio
part or relationship may be shared with another owned map.

Import recomputes every digest, proves one-to-one record/SDT ownership, exact plan/source hashes and ranges, exact
managed caption style, exact kind and physical two-block order, and absence of bookmark/`SEQ`/`REF`. Missing, extra,
enabled/non-null, overlapping, reordered, wrong-kind, wrong-plan, host-stripped, partially wrapped, reversed-pair,
map-byte/UUID/relationship/content-type, or visible-content tamper fails closed. It never falls back to style,
adjacency, paragraph order, or a historical numbering-cleanup heuristic.

#### REF-based semantic-reference occurrences / 基于 REF 的语义引用出现位置

Every semantic-reference occurrence rendered with a `REF` field has a separate reversible source-recovery record.
This includes a stable-ID selector and a soft Heading-path selector that resolves to an ID-bearing Heading. The
occurrence is wrapped in exactly one inline `w:sdt` tagged `docwen-ref-occurrence-v1:<digest32>`. The SDT contains
exactly one `REF` field followed by any authored Alias rich runs; the Alias is outside `REF` but inside this SDT, so
the visible presentation remains number plus Alias. This occurrence wrapper is not a target, bookmark, second ID, or
authorization to rewrite Markdown.

`digest32` is the first 32 lowercase hexadecimal characters of SHA-256 over the exact UTF-8 byte sequence
`docwen-ref-occurrence-map-v1\0<source-sha256>\0<start>\0<end>\0<authored-token>\0<resolved-bookmark-name>\0<cached-number>`.
`source-sha256` is the complete lowercase 64-hex source digest; `start` and `end` are decimal zero-based Unicode
code-point offsets with an exclusive end. `authored-token` is the exact range from the leading `@` through the
closing `]]`, including the authored page locator, stable-ID or Heading-path selector, and optional Alias.
`resolved-bookmark-name` is the exact `DW_T_` bookmark referenced by the field, and `cached-number` is the exact
cached numeric result inside that `REF` field. None of these values is normalized.

The records live in one separate custom XML item with namespace
`https://docwen.dev/schema/document-reference-occurrence-map/v1`. The item is emitted if and only if at least one
REF-based semantic-reference occurrence exists. Its closed root is
`<documentReferenceOccurrenceMap xmlns="https://docwen.dev/schema/document-reference-occurrence-map/v1" version="1">`;
its only children are empty `referenceOccurrence` records with attributes in exact order
`tag,source_sha256,source_start,source_end,authored_token,resolved_bookmark_name,cached_number`. Records sort by integer
`(source_start,source_end)` and then `tag`. Unknown elements, attributes, text, order, or duplicate tags/ranges are
invalid. XML escaping applies to attribute values.

This separate item uses the target map's collision-free `itemN.xml`/`itemPropsN.xml`/relationships/content-types
topology and exact byte framing. Its one `ds:schemaRef/@ds:uri` is the reference-occurrence namespace above, and its
deterministic item UUID uses that namespace in the target-map UUID formula. It never enters the semantic target map
or the ID-less soft-reference map. Import recomputes the complete digest and verifies one matching inline SDT, exact
source identity/range/token, exact `REF` bookmark/instruction/cached result, and exact visible number-plus-Alias runs.
Missing, duplicate, overlapping, changed, host-stripped, or cross-linked SDT/map data fails closed; the importer does
not flatten the occurrence or infer a selector from the target.

When a soft Heading reference resolves uniquely to an ID-less Heading, DocWen must not write an ID back to Markdown,
invent a bookmark, add a target-map entry, or emit a `REF`/hyperlink field. The DOCX occurrence is static cached
number text followed by any authored Alias, wrapped in exactly one inline non-target SDT. Its tag is
`docwen-soft-ref-v1:<digest32>`, where `digest32` is the first 32 lowercase hex characters of SHA-256 over exact UTF-8
`docwen-soft-ref-map-v1\0<source-sha256>\0<start>\0<end>\0<authored-token>`; offsets are decimal zero-based Unicode
code-point offsets with an exclusive end, and the authored token includes its leading `@` and optional Alias.

The complete token is stored outside target metadata in one closed custom XML item with namespace
`https://docwen.dev/schema/document-soft-reference-map/v1`. Its root is
`<documentSoftReferenceMap xmlns="https://docwen.dev/schema/document-soft-reference-map/v1" version="1">`; each
empty `softReference` record has attributes in exact order `tag,source_sha256,source_start,source_end,authored_token,cached_number`,
and records sort by integer `(source_start,source_end)` then tag. XML escaping applies. This item uses the same
collision-free `itemN.xml`/`itemPropsN.xml`/relationships/content-types allocation, byte framing, deterministic item
UUID, and one-schemaRef rules as the target map, but its schemaRef URI is this soft-reference namespace. It is a
source-recovery map, never a target/bookmark map.

Import verifies the complete digest, unique inline SDT, authenticated source coordinates/token, and exact static
cached-number text before restoring the authored semantic token. Missing, duplicate, changed, or host-stripped
SDT/map data fails closed. If export cannot emit this exact reversible projection, or the Heading has no materialized
number, conversion fails with the structured cross-reference diagnostic; it never degrades the token to plain text.
An ID-bearing Heading continues to use the addressable target bookmark/REF projection above even when selected by a
soft path.

#### Resolved citation item and occurrence authority / 已解析引文条目与出现权威

A resolved Citation keeps three identities separate: the authored `citation_key` is a mutable lookup key,
`record_id` plus `record_sha256` is the provider's stable/versioned record identity, and the Word-safe tag below is
an opaque package-local address. DocWen never truncates a record identity into a field, treats the key as identity,
queries a citation database, or accepts a third input. The provider-supplied `presentation` and complete
`cached_result` remain authoritative; DocWen does not run CSL or recompute either string.

For each unique `(record_id,record_sha256,presentation)` in the authenticated source document, let
`presentation_sha256` be lowercase SHA-256 of exact UTF-8 `presentation`. The item full digest is lowercase SHA-256
of exact UTF-8
`docwen-citation-item-map-v1\0<source_sha256>\0<record_id>\0<record_sha256>\0<presentation_sha256>`.
Its Word field tag is exactly `DWCIT_<digest32>`, where `digest32` is the first 32 lowercase hexadecimal characters;
the complete tag is therefore exactly 38 ASCII characters matching `DWCIT_[0-9a-f]{32}`. One `record_id` may occur
under different authored keys, versions, or occurrence-specific presentations. Only an exact three-field tuple is
deduplicated; every distinct tuple receives its own authenticated item record and opaque tag, without changing the
stable `record_id`. A different full preimage with the same full digest, the same truncated Word tag, or an existing
non-owned CITATION tag is a collision and fails before package mutation.

Item authority lives in one custom XML item with namespace
`https://docwen.dev/schema/document-citation-item-map/v1`. Its exact root is
`documentCitationItemMap` with attributes in order `version,source_sha256`, where version is `1`. Its only children
are empty `item` records with attributes in exact order
`word_tag,record_id,record_sha256,presentation_base64,presentation_sha256,sha256`; `sha256` is the full item digest
and `presentation_base64` is canonical padded RFC 4648 base64 of exact UTF-8 presentation with no whitespace.
Records are unique and sort by Unicode code-point `word_tag`. The item map is emitted if and only if at least one
resolved Citation exists and is never emitted empty.

For each Citation item reference, `item_ref_sha256` is lowercase SHA-256 of exact UTF-8
`docwen-citation-item-ref-v1\0<citation_key>\0<word_tag>\0<item_sha256>`. Let `cached_result_sha256` be lowercase
SHA-256 of exact UTF-8 `cached_result`, and let `<item_ref_sha256_csv>` be the ordered lowercase full item-reference
digests joined by a single comma with no spaces. The Citation occurrence full digest is lowercase SHA-256 of exact
UTF-8
`docwen-citation-occurrence-map-v1\0<source_sha256>\0<source_start>\0<source_end>\0<source_slice_sha256>\0<form>\0<cluster_id>\0<cached_result_sha256>\0<item_ref_sha256_csv>`.
Its inline SDT tag is `docwen-citation-occurrence-v1:<digest32>` and its bookmark name is `_DWC_<digest35>`, using
the first 32 and 35 lowercase hexadecimal characters respectively. The bookmark is exactly 40 ASCII characters;
neither physical name encodes or replaces `cluster_id`.

Occurrence authority lives in a separate custom XML item with namespace
`https://docwen.dev/schema/document-citation-occurrence-map/v1`. Its exact root is
`documentCitationOccurrenceMap` with attributes in order `version,source_sha256`, where version is `1`. Its
`citationOccurrence` children sort by integer `(source_start,source_end)` then tag, do not overlap, and have
attributes in exact order
`tag,bookmark_name,source_sha256,source_start,source_end,source_slice_sha256,authored_token_base64,form,cluster_id,cached_result_base64,cached_result_sha256,sha256`.
Each occurrence has 1..64 empty `itemRef` children in authored order with attributes in exact order
`citation_key,word_tag,item_sha256,sha256`; the last value is `item_ref_sha256`. The two base64 attributes are
canonical padded RFC 4648 encodings of exact UTF-8 strings. There are no other elements, attributes, text, or tails.
If and only if the document contains at least one resolved Citation, it has exactly one item map and exactly one
non-empty occurrence map. Every item record is referenced by at least one `itemRef`, and every `itemRef` resolves to
exactly one item record with the same `word_tag` and full `item_sha256`; orphan, dangling, or unused records fail
closed.

Every occurrence is one inline direct-paragraph `w:sdt` with the exact occurrence tag. Its `w:sdtContent` contains
only the complete `_DWC_` bookmark and one locked, clean complex field: begin `w:fldChar` has
`w:fldCharType="begin"` and `w:fldLock="true"` with no `w:dirty`; then exact `xml:space="preserve"` instruction
` CITATION <word_tag_1> \m <word_tag_2> ... `, separator, the exact provider `cached_result`, field end, and bookmark
end. A one-item occurrence omits every `\m`. The SDT, bookmark, field, occurrence record, and ordered itemRefs are
one-to-one. The bookmark/SDT is source recovery only, not a bibliography record identity, semantic target, or
authorization to rewrite the authored Citation.

Both maps independently allocate the target map's canonical `itemN.xml`/`itemPropsN.xml`/relationships/content-
types trio. Each uses its own namespace as the sole `ds:schemaRef/@ds:uri` and UUID preimage, plus the same
UTF-8/no-BOM declaration-one-LF-single-root-one-final-LF framing and lexical allocation rules; no trio, relationship,
UUID, or item number is shared. Before mutation, DocWen compares all full preimages/digests and rejects full or
truncated item-tag, SDT-tag, bookmark-name, case-insensitive existing bookmark, record-version/presentation map
mismatch, source-range, or custom-part collision.

Import accepts a Citation only when both canonical maps and their OPC trios, the exact inline SDT/bookmark/locked
CITATION field, every mapped Word tag and ordered `\m`, provider cache, source hash/range/token, keys, record IDs,
record hashes, and presentations agree. It restores the authenticated authored token and neutral identities; it
does not derive identity from the Word tag or cached display. Missing, extra, reordered, duplicated, unlocked/dirty,
tampered, host-stripped, cross-linked, or collision-injected evidence fails closed. Headless XML and round-trip gates
are separate from Word, WPS, and LibreOffice Update Fields/open-save-reopen observations; host parity is required
before this lowering may enter a packaged candidate.

Every addressable semantic target also has exactly one outer block-level `w:sdt`. Its `w:tag/@w:val` is
`docwen-target-v1:<pair-digest>`, where `<pair-digest>` is the first 32 lowercase hexadecimal characters of the same
complete target digest used for its `DW_T_` bookmark. A Heading target SDT contains exactly one heading paragraph. A
caption target SDT contains exactly one caption paragraph and exactly one authenticated native carrier, in the
semantic-kind-defined physical order frozen above, with direct adjacency and no third block. A fenced code carrier's
Word paragraphs collectively count as that one carrier. If the carrier has a distinct ordinary anchor, that anchor's
SDT may occupy the carrier slot inside this outer SDT and still counts as one carrier. The outer
SDT is required reversible internal pairing, not a source anchor, public target, or optional extra marker; it never
emits a Markdown token. Import rejects a missing, duplicate, mistagged, wrong-kind, or wrong-cardinality pairing SDT.

An ordinary `^id` never receives a Word bookmark or a `SEQ`/`REF` field. For reversible Markdown recovery, DocWen
wraps the corresponding block in one block-level `w:sdt`. Its opaque tag is
`docwen-anchor-v1:<digest>`, where `<digest>` is the first 32 lowercase hexadecimal characters of SHA-256 over exact
UTF-8 `docwen-anchor-map-v1\0anchor\0<source-id>`. The tag does not reveal the source ID. The wrapper must contain
exactly one corresponding block and may not be used as a Word hyperlink target. Import verifies unique tags, exact
single-block containment, the complete hash, source ID, tag, and block kind against the custom XML map; missing,
duplicated, moved, or changed data fails closed rather than guessing an anchor location.

#### Nested ordinary-anchor topology map / 嵌套普通锚点拓扑表

The closed four attributes of each ordinary-anchor record in `document-target-map/v1` remain
`block_kind,source_id,tag,sha256`, and each ordinary-anchor SDT remains tag-only. Nested ordinary anchors therefore use
one independent consumer-neutral custom XML item; this does not revise or alias the target-map identity. Its namespace
is `https://docwen.dev/schema/document-anchor-topology-map/v1`, its root is
`documentAnchorTopologyMap version="1"`, and its only children are empty `edge` records with attributes in exact order
`child_tag,parent_tag,sha256`.

An edge records only the direct ordinary-anchor parent derived from the source owner-path rule above. Its `sha256` is
the lowercase SHA-256 over exact UTF-8
`docwen-anchor-topology-edge-v1\0<child_tag>\0<parent_tag>`. Edges sort by `(child_tag,parent_tag)` Unicode code-point
order. Each child appears at most once, neither endpoint may equal the other, and the complete graph must be an
acyclic forest. Both tags must select exactly one ordinary `docwen-anchor-v1:` record in the same
`document-target-map/v1` and exactly one block SDT in `word/document.xml`; `docwen-target-v1:` and
`docwen-fenced-source-v1:` tags are never topology vertices. The map contains no source ID, Markdown path, consumer
identity, visible text, bookmark, or target kind.

The exact canonical item body is:

```xml
<documentAnchorTopologyMap xmlns="https://docwen.dev/schema/document-anchor-topology-map/v1" version="1"><edge child_tag="docwen-anchor-v1:11111111111111111111111111111111" parent_tag="docwen-anchor-v1:22222222222222222222222222222222" sha256="28a486c7939e34bd8d6654ec694c0a7fdbf3f1af2aceb37d76db22d6b01124de"/></documentAnchorTopologyMap>
```

This item is emitted if and only if the source projection has at least one direct ordinary-anchor parent edge. Two or
more unrelated ordinary anchors, nesting under only a semantic-target SDT, or an ordinary anchor containing only a
fenced-source carrier emits no topology item, relationship, property part, or content-type override. When emitted,
the item independently allocates the lowest collision-free `itemN.xml` trio and uses the target map's exact document
relationship, item-properties relationship, content types, XML declaration/LF/final-LF byte framing, and
`UUIDv5(NAMESPACE_URL, UTF-8(map_namespace + "\0" + sha256(itemN.xml-bytes).lowercase_hex))` formula, substituting the
topology namespace as both `map_namespace` and the sole `ds:schemaRef/@ds:uri`. Existing target-map bytes and UUID do
not change.

Every edge has one and only one physical nesting observation: the parent ordinary-anchor SDT is outside the child,
and the child is its nearest ordinary-anchor descendant in the corresponding `w:sdtContent`. No unrecorded ordinary
anchor SDT may intervene. Recursively flattened logical body-element ranges must be laminar: the child's ordered range
is either a proper contiguous subset of the parent's (*strict*) or the same ordered range (*equal*). Equal physical
ranges are legal only with the authenticated source owner path, edge, and outer-parent/inner-child XML order; block
kind or an equal element set never chooses direction. Missing or extra edges/maps, a second parent, cycle, unknown
endpoint, swapped tags, reversed wrappers, partial overlap, changed edge/hash/order/bytes/UUID/relationship, or a map
stripped by a host fails closed.

For this quote containing only one fenced block and its inner anchor, both ordinary wrappers cover the same visible
Code Block paragraph even though source ownership is nested. The exact physical order is outer quote ordinary SDT,
inner `code_block` ordinary SDT, inline fenced-source carrier SDT, then visible payload; recovery reports ordinary
groups inner-to-outer and restores the authenticated source prefixes:

````md
> ```mermaid
> graph TD
> ```
>
> ^inner-fence

^outer-quote
````

The inverse container orders are also source-authoritative, not a hard-coded kind ranking:

````md
> - one
> - two
>
> ^inner-list

^outer-quote
````

````md
- > quoted
  > continued

  ^inner-quote

^outer-list
````

### Reversible target-map package parts / 可逆 target-map 包部件

Semantic targets and ordinary-anchor SDTs share one consumer-neutral custom XML map. No map part is emitted when both
sets are empty. Otherwise, the writer allocates the lowest
positive `N` for which none of `/customXml/itemN.xml`, `/customXml/itemPropsN.xml`, and
`/customXml/_rels/itemN.xml.rels` exists. It then emits all of these package facts:

- `/word/_rels/document.xml.rels` has one relationship with the lowest unused numeric `rIdK`, from
  `/word/document.xml` to
  `../customXml/itemN.xml`, with type
  `http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml`;
- `/customXml/_rels/itemN.xml.rels` has exactly one relationship, `Id="rId1"`, target `itemPropsN.xml`, and type
  `http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps`;
- `[Content_Types].xml` has an `Override` for `/customXml/itemN.xml` with content type `application/xml` and an
  `Override` for `/customXml/itemPropsN.xml` with content type
  `application/vnd.openxmlformats-officedocument.customXmlProperties+xml`; and
- `/customXml/itemPropsN.xml` is UTF-8 XML with root `ds:datastoreItem`, namespace
  `http://schemas.openxmlformats.org/officeDocument/2006/customXml`, one deterministic braced UUID in
  `ds:itemID`, and exactly one `ds:schemaRefs/ds:schemaRef` whose `ds:uri` is the target-map namespace below.

The newly owned relationship, properties, and content-type records have these exact structures; `K`, `N`, and
`{DETERMINISTIC-UPPERCASE-BRACED-GUID}` are the only placeholders. `TargetMode` is absent. Existing unrelated records
retain their relative order; the new document relationship is inserted without reordering them, the two new
Overrides are inserted in `PartName` lexical order without reordering unrelated entries, and the item relationship
has no other siblings.

```xml
<!-- in /word/_rels/document.xml.rels -->
<Relationship Id="rIdK" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/itemN.xml"/>

<!-- complete /customXml/_rels/itemN.xml.rels -->
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemPropsN.xml"/></Relationships>

<!-- owned records in [Content_Types].xml -->
<Override PartName="/customXml/itemN.xml" ContentType="application/xml"/>
<Override PartName="/customXml/itemPropsN.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>

<!-- complete /customXml/itemPropsN.xml -->
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="{DETERMINISTIC-UPPERCASE-BRACED-GUID}"><ds:schemaRefs><ds:schemaRef ds:uri="https://docwen.dev/schema/document-target-map/v1"/></ds:schemaRefs></ds:datastoreItem>
```

The two new Overrides are inserted in `PartName` lexical order without reordering unrelated entries. Each of the
three complete new XML parts—`itemN.xml`, `itemPropsN.xml`, and `_rels/itemN.xml.rels`—has the same exact byte framing:
UTF-8 without BOM; the exact declaration `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` as the first
line; one LF; its complete root element on one line with no comments or indentation; one final LF; then immediate
EOF with no additional byte. A conflicting part, relationship, property part, or override fails before mutation.
The deterministic UUID is `UUIDv5(NAMESPACE_URL, UTF-8(map_namespace + "\0" +
sha256(itemN.xml-bytes).lowercase_hex))`, where the digest is the complete 64-character lowercase hexadecimal digest.
It is serialized as uppercase hexadecimal in braced RFC 4122 `8-4-4-4-12` form and must not collide with another
`ds:itemID` in the package.

`/customXml/itemN.xml` is UTF-8 without a BOM, begins with the exact XML declaration
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`, has no comments or insignificant indentation, and uses
this closed element order and attribute order (XML escaping applies to attribute values):

```xml
<documentSemanticMap xmlns="https://docwen.dev/schema/document-target-map/v1" version="1"><targets><target kind="figure" source_id="example" bookmark_name="DW_T_00000000000000000000000000000000000" sha256="0000000000000000000000000000000000000000000000000000000000000000"/></targets><anchors><anchor block_kind="image" source_id="image-example" tag="docwen-anchor-v1:00000000000000000000000000000000" sha256="0000000000000000000000000000000000000000000000000000000000000000"/></anchors></documentSemanticMap>
```

Both `targets` and `anchors` are always present; an empty container is serialized exactly as
`<targets></targets>` or `<anchors></anchors>`, never as a self-closing element. Target entries sort by
`(kind, source_id)` Unicode code-point order; anchor entries sort by
`(block_kind, source_id)` Unicode code-point order. Legal target kinds are exactly `heading`, `figure`, `table`,
`equation`, and `code_block`. `block_kind` is the closed DOCX anchor projection of the parser block kind, not a path
or consumer type.
No part contains a Workspace path, external node/record ID, private source text, or a second ordinary-anchor locator.
Ordinary-anchor `block_kind` values are exactly `paragraph`, `image`, `table`, `equation`, `code_block`, `list`, `list_item`, `block_quote`, and
`callout`. Every ordinary anchor on a CommonMark fenced block uses `code_block`, including Mermaid, query, view, and
any other info-string specialization; `fenced_block` is never serialized in this DOCX map. Readers reject unknown
elements/attributes, order violations, duplicate relationships, wrong content
types, mapping entries without the corresponding bookmark and semantic outer SDT or ordinary-anchor SDT, and
bookmarks/owned SDTs without a mapping entry. The semantic outer tag is recomputed from the target entry's complete
`sha256`; it is not a second map attribute. DOCX-to-Markdown recreates the
Heading/declaration-line ID for a semantic target and a valid Obsidian placement for an ordinary anchor of that block kind. A
normal Markdown-to-Markdown lossless save, by contrast, preserves the authored anchor placement and whitespace. Word, WPS Writer,
and LibreOffice Writer must each preserve the bookmark or SDT, complete part trio, relationship, and content-type
facts through save/readback before that host is accepted as reversible.

## Semantic-oracle identity / 语义 oracle 身份

The sole candidate-eligible DocWen source-oracle identity for this grammar is `docwen.markdown_semantics.v3`.
Its projection schema ID is `urn:docwen:schema:markdown-semantics:v3`; its diagnostic schema identity is
`docwen.markdown_diagnostics.v3` with schema ID `urn:docwen:schema:markdown-diagnostics:v3`. The immutable
manifest is `contracts/oracles/docwen.markdown_semantics.v3/manifest.json`, with projection and diagnostic
schemas below that oracle's `schemas/` directory. These DocWen identities are independent of every consumer's IR
version and do not alias or synchronize another project's major number.

`{#id}`, `Listing:`/`listing`/`lst-`, `List:`/`list-`, and resolver-aware WikiLink citations are not canonical v3
input. A bare `@fig-legacy` is a v3 Citation whose key is `fig-legacy`; ID-like spelling never changes that lexical
ownership. DocWen exposes no migration mode or alternate legacy grammar.

This semantic series retains the Machine Protocol v1 identity and requires Artifact Bundle v2. Its optional source-backed
diagnostic evidence and physical-route options are governed by the current schema and conformance set. A release
candidate must be rebuilt whenever the parser, writer, IR projection,
style registry, bookmark mapping, corpus, or schema changes.

## Image export / 图片导出

Layout, presentation and markup routes use the shared Markdown image-export contract only where their capabilities
declare it. DOCX-to-Markdown supports image modes `file`, `base64`, `embed`, and `omit` plus OCR placement
`image_md`/`main_md`. Physical PDF/OFD/XPS supports only `image_mode=file`, and TIFF has no `image_mode`; those
physical routes expose no OCR-placement option because their page fragments and resources use the Bundle matrix.
The public booleans are `recognize_text` and `preserve_resources`; old public boolean names are rejected. A route
advertises only values it consumes.

The visible Export settings own the global file-to-Markdown image and OCR-placement defaults (`file` and `main_md`).
A request-level CLI/GUI option overrides them; category TOML sections do not provide alternate defaults. The `kind`
argument to `get_markdown_export_modes` identifies the route and does not change precedence.

Layout、presentation 与 markup 路由在 manifest 声明后使用共享的 Markdown 图片导出契约。图片模式为 `file`、`base64`、`embed`、`omit`，OCR 放置模式为 `image_md` 或 `main_md`；每条路由只公开 converter 实际消费的值。

固定版式与多帧 TIFF 的正式交接语义见
[Physical-page OCR and artifact relations](physical-page-ocr.md)。物理页数不得由提取图片数或旧
`image_md` sidecar 数量推断。

可见的“导出”设置页唯一拥有文件→Markdown 的全局图片与 OCR 放置默认（`file`、`main_md`），请求级 CLI/GUI 选项优先覆盖。各类别 TOML 不接受同名默认键；`get_markdown_export_modes` 的 `kind` 只标识 route，不参与优先级选择。

## Regression entrypoints / 回归入口

- `tests/golden/test_md_to_docx_old_baseline.py`
- `packages/plugins/markdown/tests/test_md_to_docx_*.py`
- `packages/plugins/markdown/tests/test_md_to_docx_formatting.py`
- `packages/plugins/markdown/tests/test_md_to_spreadsheet_*.py`
- `packages/core/tests/test_links_markdown_orchestrator.py`
