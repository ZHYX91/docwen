# Machine Protocol v1 and Artifact Bundle v2 / 机器协议与产物包

DocWen's public process boundary is JSON-RPC 2.0 over Content-Length framed stdio. It returns a
consumer-neutral artifact graph and does not expose internal route objects, argv, runtime manifests, or a
consumer's storage model. The normative source is [`contracts/`](../../contracts/README.md).

DocWen 的公开进程边界采用 JSON-RPC 2.0 + Content-Length stdio，返回消费者无关的产物图；内部 route
对象、argv、Runtime Manifest 以及消费者存储模型均不属于 wire contract。权威定义位于
[`contracts/`](../../contracts/README.md)。

## Current authority / 当前权威

`docwen.machine.v1`, `docwen.artifact_bundle.v2`, and the schemas and fixtures under `contracts/` are
the current wire authority. The in-process CLI/GUI and external `DocWenCLI serve --stdio` share one
plan-first Application Service; the CLI JSON presentation envelope and private provider paths do not enter this wire.

`docwen.machine.v1`、`docwen.artifact_bundle.v2` 以及 `contracts/` 下的 schema 与 fixture 是当前
wire 权威。进程内 CLI/GUI 与外部 `DocWenCLI serve --stdio` 共用同一个 plan-first Application
Service；CLI JSON 展示 envelope 与私有 provider 路径不进入该 wire。

## Boundaries / 边界

- DocWen owns `docwen.machine.v1`, `docwen.artifact_bundle.v2`, capability discovery, task lifecycle, and
  conformance fixtures.
- GUI, Assistant, OpenClaw, and external consumers use the same Machine Protocol.
- A consumer maps `document`, `fragment`, `resource`, entry roles, and relations into its own domain only after
  path and integrity validation.
- Provider-specific capability IDs and local locators stop at the adapter boundary.
- Bundle relations are semantic facts, never Page/Node/Workspace placement commands.
- Markdown→DOCX accepts no independent bibliography or `citation_style` slot. Its exact two inputs are the neutral
  document and numbering/export plan; any already-presented closed bibliography is embedded and authenticated inside
  that neutral document. DocWen renders typed runs but does not run CSL.
- Physical-page relations carry closed one-based page/source semantics while relation ordinals remain zero-based;
  `page_nodes` is a consumer-owned import strategy, not a Bundle command.
- Fixed-layout PDF/OFD/XPS and multi-frame TIFF Markdown capabilities declare a many-artifact shape with
  `relation_payloads=[page_fragment,page_resource]`; this advertises closed Bundle facts only and carries no
  consumer placement or Node instruction.
- Semantic Markdown parsing directly resolves only targets in the supplied document. Workspace/page lookup and
  citation-record lookup belong to an external resolver, which must lower resolved facts into the neutral DocWen
  boundary before conversion. Machine `options` cannot carry resolver state, source text, or a hidden metadata bag.
  Any resolver resource requires its own typed schema, fixtures, capability slot, and consumer contract.
- Markdown→DOCX structured numbering accepts a provider-neutral resolved document plus a separately versioned,
  closed resolved numbering/export plan. The upstream owner chooses profiles, counters, enablement, scope, reset,
  format, label, chapter inclusion, and separator. DocWen validates one source identity/hash and materializes only
  resolved Heading list/`numbering.xml`, caption `SEQ`, bookmark/`REF`/cached-result, and managed-style facts.
  Consumer Workspace/Node data, WikiLink resolver instructions, counter rules, and an opaque plan string are forbidden.
- The v4 resolved-plan path does not expose or emulate the legacy source-authoring controls `remove_numbering`,
  `add_numbering`, `numbering_scheme`, or regex Heading cleanup. Those controls mutate/interpret authored Markdown
  and are not inputs to the resolved-plan Conversion Port.

DocWen 拥有两份契约及其发现、任务和验收语义。消费者在完成路径与完整性校验后自行映射；Bundle
中的关系只是语义事实，不能解释为 Page、Node 或 Workspace 放置指令。完整样式/参考文献合同见
[`templates-and-styles.md`](templates-and-styles.md)，物理页合同见
[`physical-page-ocr.md`](physical-page-ocr.md)。

## Resolved-numbering inputs / 已解析编号输入

The v4 plan-aware Markdown→DOCX capability requires exactly two input resources; neither is optional or repeatable:

| Role | Media type | Cardinality | Schema identity |
|---|---|---:|---|
| `neutral_document` | `application/vnd.docwen.resolved-document+json` | exactly 1 | `docwen.resolved_document.v1` / `urn:docwen:schema:resolved-document:v1` |
| `numbering_export_plan` | `application/vnd.docwen.numbering-export-plan+json` | exactly 1 | `docwen.numbering_export_plan.v1` / `urn:docwen:schema:numbering-export-plan:v1` |

The successful output is one atomic many-artifact Bundle: the preferred DOCX `document` plus exactly one `resource`
with media type `application/vnd.docwen.round-trip-sidecar+zip`. The resource has one
`resource_of(role=manifest, ordinal=0)` relation to the DOCX and suggested name `<DOCX suggested name>.docwen`.
`docwen.round_trip_sidecar.v1` is DocWen-owned; consumers do not recreate it from private inputs. Its ZIP inventory is
exactly `authored-source.md`, `neutral-document.json`, `numbering-export-plan.json`, `manifest.json` in that order. The
manifest binds the exact DOCX and all three payload identities. A consumer publishes or moves the DOCX and sidecar as
a pair, preserving the adjacent basename; Artifact Bundle locators and SHA-256 values remain the authority while the
files are in request staging.

成功输出是一个原子多 artifact Bundle：首选 DOCX `document`，以及唯一的
`application/vnd.docwen.round-trip-sidecar+zip` `resource`。它通过
`resource_of(role=manifest, ordinal=0)` 归属于 DOCX，建议文件名为 `<DOCX 建议文件名>.docwen`。sidecar 的
四成员、manifest、DOCX/数据哈希全部由 DocWen 生成；消费者只校验并成对发布，不自行用私有数据重建。

Both are strict UTF-8 JSON, at most 8 MiB each, reject duplicate keys/non-finite numbers, and are closed at every
object. Each envelope requires exactly `$schema,schema,input_id,source_sha256,plan_sha256` plus its schema-owned
payload (`document` or `plan`). The identity strings equal the table above; both resources carry the same `input_id`,
`source_sha256`, and `plan_sha256`. `input_id` is 1..256 characters and matches byte-for-string; `source_sha256` is
the same lowercase digest of the authenticated authored Markdown.
`plan_sha256` is the lowercase SHA-256 of RFC 8785 canonical UTF-8 bytes of the closed `plan` member only, excluding
both envelopes, so it is acyclic. The producer writes that same digest into both envelopes; DocWen recomputes it and
validates all three pointers before reading a target. Bundle/resource file SHA-256 remains a separate byte identity.

The `document` payload is exactly
`authored_markdown,targets,references,resource_occurrences,citations,resources`. It contains provider-neutral authored
content, semantic target occurrences and already-resolved facts, never a private Workspace path/Node object. The closed
dependency records are normative in
[Resolved document dependencies](structured-numbering-phases.md#resolved-document-dependencies--已解析文档依赖).
The closed `plan` materialization grammar is normative in
[Resolved structured numbering and export plan](structured-numbering-phases.md#closed-portable-materialization--闭合可移植物化形式).
Because `authored_markdown` uses CommonMark ATX levels 1..6 plus DocWen ATX extension levels 7..9, a resolved Heading
target and every used `heading_list`, caption restart, or caption chapter binding has level 1..9. All nine levels map
one-to-one to Word Heading styles and outline levels.
Missing `numbering_export_plan` emits
`docwen.numbering_export_plan.missing`; malformed schema, duplicate/mismatched pointer, stale hash, or contradictory
target emits `docwen.numbering_export_plan.invalid`; a valid but non-portable Word form emits
`docwen.numbering_export_plan.unsupported_materialization`. These admission diagnostics have no Markdown source
range/fix and publish no partial artifact. They are distinct from a valid disabled target.
Missing or invalid `neutral_document` independently uses `docwen.resolved_document.missing` or
`docwen.resolved_document.invalid`; it is never disguised as a numbering state either.

The sole provider diagnostic mapping for that valid state is
`interop.cross_reference.unnumbered_target` ↔
`docwen.markdown.cross_reference.unnumbered_target`. It is one-to-one in both directions and preserves severity,
source evidence, target identity/kind, and no-fix status; neither side maps missing/invalid/unsupported plan errors to
an unnumbered target.

## OCR-capable option schemas / OCR 能力选项 schema

The final v4 DOCX-to-Markdown capability has exactly these seven properties: `recognize_text` (boolean, default `false`),
`preserve_resources` (boolean, default `true`), `ocr_language` (the eight-value language enum, default `auto`),
`image_mode` (`file|base64|embed|omit`, default `file`), `ocr_placement` (`image_md|main_md`, default `main_md`),
`image_link_style` (`wiki_embed|wiki_link|markdown_embed|markdown_link`, default `wiki_embed`),
and `table_merge_strategy` (`fill|empty|marker`, default `fill`). Its schema has `required=[]` and
`additionalProperties=false`. The three source-authoring numbering properties are rejected on this capability because
the closed Machine/capability schema exposes them only on other declared capabilities. The schema, manifest,
fixtures, source/package capability output, and consumers change atomically.

DOCX is a `document_with_resources` route, not a physical-page route. Let `K` be the number of preserved embedded
image resources and `R` the number of non-empty OCR results produced from eligible embedded images. The two public
booleans remain independent. With the default `ocr_placement=main_md`, OCR text is written into the primary
Markdown and creates no fragment; with explicit `ocr_placement=image_md`, each of the `R` results is one Markdown
`fragment` with one `fragment_of/ocr_text` relation to the primary. `preserve_resources=true` independently emits
exactly `K` image `resource` artifacts, each with one `resource_of/image` relation to the primary; `false` emits
none. In every combination there is exactly one `primary` entry and no page fragment, page relation, or consumer
Node instruction. `image_mode` controls Markdown presentation only and does not override either boolean or whether
an exported image is a Bundle resource.

With the default `image_mode=file`, authenticated DOCX Figures and ordinary image anchors keep their image owner in
the primary Markdown in all six observable combinations (`recognize_text=false` with either preservation value, plus
`recognize_text=true` with either preservation value under each OCR placement). With preservation on, that owner is
the exported image reference. With preservation off, it is the exact resource-less carrier `![image omitted]()`. A
different presentation mode remains independent of resource preservation, but if it cannot retain the authenticated
image owner the task fails closed; it cannot coerce either boolean, replace the owner with a comment, or silently move
the owner. Under `main_md`, a
non-empty OCR result follows that primary owner and produces no fragment. Under `image_md`, the primary owner is never
replaced by an OCR-sidecar embed: the sidecar is OCR-only and contributes one `fragment_of/ocr_text` relation. Thus an
OCR fragment can never inherit a Figure declaration or ordinary image ID. Every result containing at least one such
carrier includes one primary-artifact-bound warning with code `DOCX2MD-IMAGE-OWNER-RESOURCE-OMITTED`. This artifact
warning does not set `evidence_schema`, `source`, `range`, `related_ranges`, or `fixes`.

The two common boolean controls are independently negotiated by capability discovery. PDF/OFD/XPS-to-Markdown each have
the exact property set `recognize_text`, `preserve_resources`, `ocr_language`, `image_mode`, and `render_dpi`;
TIFF-to-Markdown has exactly `recognize_text`, `preserve_resources`, and `ocr_language`. Every route has
`required=[]` and `additionalProperties=false`. `recognize_text` is boolean with default `false` and
`preserve_resources` is boolean with default `true`. The full enum, bounds, defaults, and exact four Bundle matrices
are normative in [Physical-page OCR](physical-page-ocr.md#closed-machine-options--闭合机器选项). Public
`to_md_enable_ocr` and `to_md_keep_images` are rejected; unrelated route-specific controls are not removed by the
rename.

## Source diagnostic evidence / 源码诊断证据

Machine diagnostics retain the required base fields `severity`, `code`, and `message` and optional `artifact_id`.
Source-backed diagnostics may additionally use `docwen.machine.diagnostic_evidence.v1`, schema ID
`urn:docwen:schema:machine-diagnostic-evidence:v1`, through the direct optional fields `evidence_schema`, `source`,
`range`, `related_ranges`, and `fixes`. If any evidence field is present, `evidence_schema`, `source`, and `range` are
all required. Nested objects are closed; `related_ranges` has at most 16 entries, `fixes` at most 8, each fix has
1..16 ordered non-overlapping edits, and each replacement has at most 4096 Unicode code points.

`source` is exactly `{input_id,sha256,encoding,coordinate_system,offset_base,range_end}` with lowercase SHA-256,
`encoding=utf-8`, `coordinate_system=unicode_code_point`, `offset_base=0`, and `range_end=exclusive`. A range/edit
range is exactly `{start,end}` with non-negative integer endpoints; the primary range is non-empty, while an
insertion edit may be zero-width. A fix is exactly `{fix_id,edits}` and an edit exactly `{range,replacement}`.
Consumers verify the authenticated input hash and all bounds before applying all edits atomically. This wire evidence
is distinct from a source oracle, packaged observation, round-trip result, or host observation; one layer cannot be
cited as another. The exact Markdown codes and coordinate rules are in
[Markdown compatibility](markdown-compatibility.md#diagnostics-and-single-document-rewrites--诊断与单文档重写).

Fenced-source preservation is not a Machine diagnostic extension and does not add resolver context or a consumer
identity to task options. Its source observation is the required `fenced_sources` array in the closed
`docwen.markdown_semantics.v3` projection: each record points to the authenticated input SHA-256 and exact Unicode-
code-point range and authenticates complete authored-block and logical-body hashes. Its packaged observation is the
separate canonical `document-fenced-source-map/v1` custom XML part plus one complete inline carrier SDT per record.
Its round-trip observation is the recovered primary Markdown bytes, including container prefixes, LF/CRLF, closer
spelling, and omitted-EOF state. Source-oracle records cannot prove package bytes; package inspection cannot prove
returned Markdown; neither is `task/completed.params.diagnostics` wire evidence.

A candidate record for this feature declares exactly one of `source_oracle`, `packaged`, `headless_ooxml`,
`roundtrip`, `word_host`, `wps_host`, or `libreoffice_host` and binds the relevant source projection or oracle-
manifest identity, package artifact identity, and evidence path/bytes/SHA-256. The candidate receipt binds those
records to the final spec and implementation commit/tree and packaged executable identity. It must not replace the
fenced-source source hash/range with a consumer page/node ID, infer the record from paragraph order, or cite an
unpackaged source fixture as wire/package evidence.

Nested ordinary-anchor topology is layered the same way but remains independent of fenced-source occurrence evidence.
Its source-oracle observation is each anchor's required closed `container_path` plus authenticated source identity and
range. Its packaged/headless-OOXML observation is the separate canonical
`document-anchor-topology-map/v1`, the unchanged `document-target-map/v1` ordinary-anchor records, and the exact
outer-parent/inner-child tag-only block-SDT nesting. Its round-trip observation is recovered Markdown with both IDs at
their authenticated container paths. None of these is a Machine option, Bundle relation, consumer hierarchy, or
`task/completed.params.diagnostics` wire observation. Candidate evidence must bind the source-oracle projection and
manifest identity separately from package bytes and returned Markdown; an equal visible block range, block kind, or
consumer parent/child record is not evidence of the DocWen topology edge.

Resolved numbering evidence is also layered. `source_oracle` binds exact Markdown bytes, the fresh semantic/plan
schema and manifest identity, five-kind enable/disable plus Heading-template-empty expectations, and zero Markdown
diff. `packaged` proves the installed converter consumes that exact closed plan. `headless_ooxml` proves Heading
`numbering.xml`/list bindings, caption `SEQ`, navigation bookmarks, `REF`, cached results, managed styles, and their
required absence. `roundtrip` proves neutral extraction and canonical Markdown bytes. `word_host`, `wps_host`, and
`libreoffice_host` independently prove open/save/reopen. Upstream-provider write/read observations remain provider evidence;
they cannot be relabeled as Machine wire, DocWen package, OOXML, round-trip, or host evidence. Every record and the
candidate receipt bind the final spec/implementation commit and tree, plan/schema/corpus manifest identity, and exact
package manifest/binary digest.

## Required gates / 必须门禁

- `python tools/validate_contracts.py` passes offline and rejects every declared negative fixture.
- Source and packaged servers emit identical initialization/capability shapes and terminal semantics.
- A success is valid only after every Bundle locator stays under the request-owned staging root and every
  `size_bytes`/`sha256` matches the bytes.
- Accepted tasks produce one terminal notification; failure/cancellation never publish a partial Bundle.
- Accepted tasks emit progress with strictly increasing per-task sequence numbers. Machine v1 retains
  `phase=conversion` for convert, validate, render, transform, and merge. Only bounded integer percentage facts are
  projected through closed server-owned messages; duplicate, regressing, boolean, non-finite, or late Runtime values
  are ignored. Paths, locators, document text, sheet names, and diagnostic locations never enter progress phase or
  message. Success emits server-owned 100 percent before the terminal; failure/cancellation do not fabricate it, and
  the terminal uses the next sequence.
- Capability discovery exposes output cardinality, artifact kinds, possible relations, options schema, dependency
  availability, and limitations without requiring a conversion.

The 0.9.0 package and the 2.0.0 consumer releases record the hosted release evidence for this contract. Any schema,
capability, consumer, adapter, or packaging-input change invalidates the affected result and requires the corresponding
gate to rerun.

0.9.0 包与两个 2.0.0 消费者 Release 已记录本合同的托管发布证据。schema、capability、消费者、adapter
或打包输入变化后，受影响结果立即失效并须重跑相应门。
