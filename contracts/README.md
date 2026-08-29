# DocWen wire contracts

`contracts/` is the single authoritative source for DocWen-owned process-boundary contracts. Machine Protocol v1,
Artifact Bundle v2, and Proofread Report v2 are the only accepted current wire identities. Schemas, fixtures,
consumers, packaged contracts, and release hashes must be updated atomically. `docwen.markdown_semantics.v3` is the
authored Markdown grammar/range/topology/fenced/YAML/citation projection authority, but is not by itself a numbering,
resolved-plan, candidate, package, or release authority. The contracts remain intentionally independent of the CLI JSON presentation envelope, internal
`ArtifactManifest`, runtime `OutputManifest`, GUI models, and any consumer's storage model.

The schemas and fixtures are normative. Source and packaged `DocWenCLI serve --stdio` implementations are
validated against them; formal release-candidate evidence remains a separate gate.

## Contract identities and ownership

- Machine protocol: `docwen.machine.v1` / schema ID `urn:docwen:schema:machine-protocol:v1`.
- Artifact bundle: `docwen.artifact_bundle.v2` / schema ID `urn:docwen:schema:artifact-bundle:v2`.
- Proofread report: `docwen.proofread_report.v2` / schema ID `urn:docwen:schema:proofread-report:v2`.
- Presented bibliography resource: `docwen.semantic_bibliography.v1` / schema ID
  `urn:docwen:schema:semantic-bibliography:v1`.
- Resolved document: `docwen.resolved_document.v1` / schema ID
  `urn:docwen:schema:resolved-document:v1`.
- Numbering export plan: `docwen.numbering_export_plan.v1` / schema ID
  `urn:docwen:schema:numbering-export-plan:v1`.
- Source diagnostic evidence: `docwen.machine.diagnostic_evidence.v1` / schema ID
  `urn:docwen:schema:machine-diagnostic-evidence:v1`.
- Markdown semantic source oracle: `docwen.markdown_semantics.v3` / schema ID
  `urn:docwen:schema:markdown-semantics:v3`; its diagnostic oracle is `docwen.markdown_diagnostics.v3` / schema ID
  `urn:docwen:schema:markdown-diagnostics:v3`. These are source/round-trip oracle identities, not a new Machine input.
- DocWen owns these schemas/oracles, fixtures, versioning, capability discovery, transport, task lifecycle, and product
  errors. There is no third shared-contract repository.
- A consumer owns the mapping from the exchange objects into its own domain. The Bundle never contains
  provider-specific Workspace, Node Tree, Page, vault, encryption, or private consumer-model semantics.

The first implemented protocol has no compatibility obligation to CLI protocol 3. Incompatible machine wire changes
require a new Machine Protocol major; incompatible Bundle shape or graph changes require a new Artifact Bundle major.
Additive optional capability data may use a protocol minor only after an explicit version decision and new
conformance fixtures.

## Transport

Machine Protocol v1 uses JSON-RPC 2.0 over the child process's stdin/stdout. Each message is UTF-8 JSON without
a byte-order mark and is framed exactly as:

```text
Content-Length: <decimal UTF-8 byte count>\r\n
\r\n
<JSON object bytes>
```

There is exactly one canonical `Content-Length` header, its value is in the range `1..16777216`, and line endings
are CRLF. Artifact bytes never travel through JSON-RPC. stderr is for human/runtime logs and must not contain
protocol frames. EOF terminates an idle server; the client must first cancel or await any accepted task.

## Machine lifecycle

The v1 request methods are `initialize`, `capability/list`, `health/check`, `file/inspect`, `resource/list`,
`gui/status`, `gui/activate`, `gui/open`, `task/plan`, `task/execute`, and `task/cancel`.
The server emits `task/progress` and exactly one of `task/completed`, `task/failed`, or `task/cancelled` for every
accepted task.

1. `initialize` fixes protocol `1.0`, feature support, the method set, Bundle schema, and concurrency.
2. `capability/list` returns provider-specific capability IDs, typed input slots, capability-specific closed option schemas,
   dependency availability, and the expected artifact graph shape. Each `input_shape` declares unique roles and
   rejects undeclared roles. Ordinary capabilities have a required `source`; the v4 Markdown-to-DOCX capability
   instead has the exact required pair `neutral_document` + `numbering_export_plan`. A consumer maps these IDs at its boundary;
   they are not consumer-domain IDs.
3. `task/plan` binds immutable typed input fingerprints, an existing empty request-owned local staging root, and
   normalized options. It returns a `plan_id`, effective options, output shape, warnings, limitations, and any
   confirmation requirement without starting conversion. Calling `task/execute` is the client's confirmation.
4. `task/execute` accepts one plan and returns a `task_id`; progress and terminal notifications then use a
   strictly increasing per-task `sequence`.
5. `task/cancel` acknowledges the request. Cancellation is complete only after `task/cancelled`; a race may
   instead end in `task/completed` or `task/failed`, but never a second terminal notification.

JSON-RPC `error` is reserved for framing, parsing, invalid request/parameters, unknown method, and server-level
protocol faults. A conversion failure after task acceptance uses `task/failed` with a stable product error
`category` and `code`. A successful terminal notification contains one complete Bundle. Failed and cancelled
tasks do not publish partial Bundles.

## OCR-capable Machine options

DOCX-to-Markdown publishes this exact closed property set:

| Property | Type and exact constraints | Default |
|---|---|---|
| `recognize_text` | boolean | `false` |
| `preserve_resources` | boolean | `true` |
| `ocr_language` | string enum `auto`, `chinese`, `chinese_cht`, `english`, `japanese`, `korean`, `latin`, `cyrillic` | `auto` |
| `image_mode` | string enum `file`, `base64`, `embed`, `omit` | `file` |
| `ocr_placement` | string enum `image_md`, `main_md` | `main_md` |
| `image_link_style` | string enum `wiki_embed`, `wiki_link`, `markdown_embed`, `markdown_link` | `wiki_embed` |
| `table_merge_strategy` | string enum `fill`, `empty`, `marker` | `fill` |
| `remove_numbering` | boolean | `true` |
| `add_numbering` | boolean | `false` |
| `numbering_scheme` | string with `x-docwen-resource-kind=numbering-schemes` | `gongwen_standard` |

This table belongs only to the isolated legacy DOCX-to-Markdown route. It is not a structured-numbering authority,
cannot supply the v4 Markdown-to-DOCX plan, and is excluded from current v4 numbering evidence. Its schema has
`type=object`, `required=[]`, and `additionalProperties=false`; `numbering_scheme` is validated against the
discovered numbering registry rather than hard-coded as an enum.

DOCX uses the `document_with_resources` profile and never adopts physical-page P/K semantics. Let `K` be preserved
embedded image resources and `R` non-empty OCR results. Every result has exactly one primary document entry.
`preserve_resources=true` adds exactly K image resources and K `resource_of/image` relations to the primary;
`false` adds none. When `recognize_text=true`, the default `ocr_placement=main_md` places OCR text in the primary and
adds no fragments, while explicit `image_md` adds exactly R Markdown fragments and R `fragment_of/ocr_text`
relations. Recognition off adds no OCR fragment. No DOCX combination creates `ocr_page`, `page_fragment`,
`page_resource`, or a consumer Node instruction. `image_mode` changes Markdown presentation, not these two public
controls or the Bundle kind of an exported image.

The public PDF/OFD/XPS-to-Markdown schemas each contain exactly these five properties; the public TIFF-to-Markdown
schema contains exactly the first three. All four route schemas have `type=object`, `required=[]`, and
`additionalProperties=false`.

| Property | Type and exact constraints | Default | PDF/OFD/XPS | TIFF |
|---|---|---|---|---|
| `recognize_text` | boolean | `false` | yes | yes |
| `preserve_resources` | boolean | `true` | yes | yes |
| `ocr_language` | string enum `auto`, `chinese`, `chinese_cht`, `english`, `japanese`, `korean`, `latin`, `cyrillic` | `auto` | yes | yes |
| `image_mode` | string enum containing only `file` | `file` | yes | no |
| `render_dpi` | integer, minimum 72, maximum 600 | `200` | yes | no |

`recognize_text` and `preserve_resources` are independent. Public `to_md_enable_ocr` and `to_md_keep_images` are
undeclared and rejected; a private producer adapter may translate the new fields after validation. With P positive
physical pages/frames and K non-negative preserved images, the exact physical-route Bundle matrix is:

| Recognition | Resources | Artifact cardinality | Entry cardinality | Structural relation cardinality |
|---|---|---|---|---|
| off | off | 1 primary Markdown document | exactly one primary preferred entry at ordinal 0 | 0 |
| off | on | 1 document + K image resources | the same one entry | K `resource_of` to the primary |
| on | off | 1 document + P page fragments | the same one entry | P `fragment_of/ocr_page` to the primary |
| on | on | 1 document + P page fragments + K image resources | the same one entry | P `fragment_of/ocr_page` + K `resource_of` |

An image sidecar is always a resource and never an entry, document, fragment, or implicit Node. A proven resource
targets its page fragment when recognition created one; otherwise it targets the primary. An unproven resource also
targets the primary, omits `page_resource`, and has the required artifact-bound diagnostic. The complete relation
payload rules are in [`physical-page-ocr.md`](../docs/specs/physical-page-ocr.md).

## Artifact Bundle graph

An artifact is a delivered, integrity-bound output. Logs, diagnostics, temporary files, caches, and internal
intermediates are not artifacts unless a capability explicitly exports their bytes as a user-consumable output.

Artifact `kind` has exactly three values:

- `document`: a self-contained document representation, including an attachment that can be consumed on its own;
- `fragment`: content that normally needs composition with another document, such as an OCR page or section;
- `resource`: non-document bytes such as images, CSV/TSV outputs, retained originals, or previews.

`entries` identify the graph's consumer-visible starting points and deterministic order. Their contextual role is
one of `primary`, `supplementary`, `ocr_page`, `section`, `worksheet`, `image`, or `original`. At most one entry
may be `preferred`; zero is valid when outputs are peers. Preference is a presentation/default-output hint, not a
storage instruction.

Relations point from the subject artifact to its context/source:

- `attachment_of`: document -> document, role `attachment`;
- `fragment_of`: fragment -> document, role `ocr_page`, `ocr_text`, `section`, or `worksheet`;
- `resource_of`: resource -> document/fragment, role `image`, `original`, `preview`, or `worksheet`;
- `derived_from`: any artifact -> any source artifact, role `source` or `original`.

Attachment and fragment relations require zero-based `ordinal`. A `fragment_of`/`ocr_page` relation also requires
closed `page_fragment` semantics: `fragment_kind=page`, one-based `page_index`, one-based `source_page`, positive
`page_count`, and a closed OCR status. Its ordinal equals `page_index - 1`; siblings under one owner form the exact
one-based sequences `1..page_count` without gaps or duplicates. A `resource_of` image/original/preview relation may
carry `page_resource.source_page`. Proven page resources target the matching page fragment when it exists; an
unresolved resource targets the primary document, omits page semantics, and receives an artifact-bound diagnostic.
No page field is inferred from file names or resource count. The normative invariants are documented in
[`physical-page-ocr.md`](../docs/specs/physical-page-ocr.md).

Artifact IDs and locators are unique. Locators
are normalized relative POSIX paths under the `task/plan` staging root; absolute paths, drive prefixes,
backslashes, empty/dot segments, traversal, and duplicate targets are invalid. Every artifact has a byte size and
lowercase SHA-256 digest. Consumers must resolve paths beneath the exact request-owned staging root, reject link
or reparse escapes, verify size/hash, and import/copy bytes before that staging root is released.

Machine `local_path` values use absolute native paths at the process boundary. Output `staging_policy` is fixed to
`require_empty`: the root must already exist, be owned by this request, contain no entries, and not be a link or
reparse point. Replace, rename, and destination overwrite decisions belong to the caller's commit layer, not the
DocWen provider process.

Input `logical_path` values are request-virtual, case-sensitive normalized relative POSIX keys. They are nonempty,
globally unique within a plan, and reject native separators, absolute paths, drive prefixes, URIs, NUL, and empty,
dot, or traversal segments. Input `kind` is `document` or `resource`; `source` may use either, while
`linked_resource`, `bibliography`, `citation_style`, and `numbering_export_plan` require `resource`;
`neutral_document` requires `document`. A plan validates in this stable order:
duplicate input ID, invalid logical path, duplicate logical path, undeclared role, slot kind, slot media type, then
slot cardinality.

`convert.markdown.to_docx` accepts exactly two inputs and no others: one `neutral_document` document with media type
`application/vnd.docwen.resolved-document+json`, and one `numbering_export_plan` resource with media type
`application/vnd.docwen.numbering-export-plan+json`. The plan never appears in options. The neutral document embeds
authenticated raster resources, an optional closed semantic bibliography, and already-resolved citation occurrences;
there is no independent `linked_resource`, `bibliography`, or `citation_style` slot and no source-relative fallback.
Both JSON files are strict UTF-8, duplicate-key-free, closed, and at most 8 MiB. Embedded decoded resources total at
most 6,000,000 bytes. Every linked raster is non-empty, hash/media/content checked and bound by a complete-token
Unicode source range plus opaque resource ID. The exact dependency and numbering grammars are frozen in
[`structured-numbering-phases.md`](../docs/specs/structured-numbering-phases.md).

A successful conversion returns one atomic two-artifact Bundle: the preferred DOCX `document` and one
`application/vnd.docwen.round-trip-sidecar+zip` `resource`, related by
`resource_of(role=manifest, ordinal=0)`. DocWen alone produces the adjacent `<name>.docx.docwen` artifact. Its
`docwen.round_trip_sidecar.v1` manifest authenticates the exact DOCX plus `authored-source.md`,
`neutral-document.json`, and `numbering-export-plan.json`; consumers publish the pair and must not reconstruct the
sidecar from private inputs. Missing or mismatched sidecars disable byte-exact source recovery without disabling
authenticated semantic normalization.

The `remove_numbering`, `add_numbering`, `numbering_scheme`, and
`heading_numbering_render_mode` controls are available only on separately identified capabilities
that still declare them. They are undeclared and rejected on `convert.markdown.to_docx`; the v4 provider consumes
only the validated closed plan and never derives a plan from those options.

The Markdown source oracle distinguishes ordinary `[[...]]`/`![[...]]`, numbered semantic
`@[[#^id]]`/`@[[Page#^id]]`, and citations `@citation-key`/`[@key; @key]`. Direct Markdown conversion resolves only
same-document targets. Cross-document targets and citation records must already be resolved by an external owner and
lowered to DocWen's neutral resolved-reference/citation boundary; DocWen neither embeds that resolver nor scans a
Workspace. Machine options cannot smuggle source text, page indexes, citation indexes, or a generic semantic
context. Any typed resolver resource requires its own schema, capability slot, fixtures, and consumer contract.

Source-backed Machine diagnostics may add the exact optional fields `evidence_schema`, `source`, `range`,
`related_ranges`, and `fixes`. If any is present, `evidence_schema=docwen.machine.diagnostic_evidence.v1`, `source`,
and `range` are required. `source` authenticates `input_id` and SHA-256 and freezes UTF-8 Unicode-code-point,
zero-based, exclusive-end coordinates. Closed ranges are `{start,end}`; fixes are `{fix_id,edits}` and edits are
`{range,replacement}`. Related ranges (maximum 16), fixes (maximum 8), edits per fix (1..16), and replacement text
(maximum 4096 code points) are bounded. The exact applicability and atomic-edit rules are documented in
[`markdown-compatibility.md`](../docs/specs/markdown-compatibility.md).

The graph must be acyclic, all references must resolve, a structurally owned artifact has exactly one owner, an
entry cannot also be structurally owned, ordered siblings cannot reuse an ordinal, and every artifact belongs to
a connected component containing an entry.

## Multi-artifact mappings

These mappings describe DocWen output meaning, not a consumer's import plan:

| Current output family | Bundle mapping |
|---|---|
| Government document body plus Markdown attachments | body `document` entry; each attachment `document attachment_of` body |
| Physical-page OCR | capability output shape advertises `relation_payloads=[page_fragment,page_resource]`; primary Markdown is the only entry; every recognition-enabled physical page/frame is a `fragment fragment_of` with `ocr_page`, ordinal, and complete `page_fragment`; preserved images are resource-only sidecars with proven `page_resource` or an artifact-bound unresolved-page diagnostic |
| `image_md` OCR companions | main Markdown is `document`; companion Markdown is `fragment` with `ocr_text`; retained image is `resource`, and provenance may add `derived_from` |
| DOCX/PPTX/HTML/MHTML/EPUB/XLSX extracted images | each image is `resource resource_of` its owning document/fragment, with optional ordinal |
| Resolved Markdown to DOCX | preferred DOCX `document`; one adjacent `.docwen` `resource resource_of(role=manifest, ordinal=0)` |
| XLSX multi-sheet Markdown | each self-contained sheet Markdown may be a `document` entry with role `worksheet`; a composition-only sheet is `fragment` |
| XLSX/ODS/ET to per-sheet CSV/TSV chain | each CSV/TSV is a `resource` entry with role and ordinal `worksheet`; one entry is preferred |
| PDF split outputs | independently usable PDFs are `document` entries with role `section` and deterministic ordinal |

Consumers may merge fragments, create separate documents, attach resources, or choose another materialization.
Those choices are outside this contract.

## Conformance

`conformance-manifest.json` inventories every schema and positive/negative fixture. Schema-invalid examples cover
closed vocabularies and required terminal payloads. Schema-valid semantic-invalid Bundle fixtures cover every
portable graph, ownership, ordering, locator, and entry invariant enforced by the offline and typed validators.
Lifecycle traces cover request/response correlation, task acceptance, cancellation, progress, Bundle/task
identity, and exactly-one-terminal honesty. Framing fixtures cover canonical headers, byte lengths, the message
size limit, and JSON-object payloads. A mismatched framing `expected_messages` oracle is fixture-harness
corruption rather than an invalid protocol frame and is therefore covered only by a validator unit test.

Run the offline gate from the repository root:

```powershell
uv run --extra test python tools/validate_contracts.py
uv run --extra test pytest tests/test_repo/test_machine_protocol_v1_contracts.py
```

Any schema, vocabulary, lifecycle, framing, locator, integrity, or fixture expectation change must update this
README and the conformance set in the same change.

Candidate receipts bind the final semantic/diagnostic identities; exact clean DocWen and provider commit/tree plus
their spec-baseline commit/tree; candidate ID; package manifest and executable relative path/bytes/SHA-256/version;
the complete closed options schema and four-row Bundle matrix; and every evidence artifact directly as a normalized
candidate-relative path, byte count, and lowercase SHA-256. Evidence records declare exactly one layer:
`source_oracle`, `machine_wire`, `packaged`, `roundtrip`, `headless_ooxml`, `word_host`, `wps_host`, or
`libreoffice_host`. A pointer from one layer is never proof of another. All identities and content hashes are inputs;
changing any one invalidates the candidate and derived receipt. Consumer identities may be bound by an external
handoff receipt, but private absolute paths are not canonical source-contract data.

The receipt never contains its own path, byte count, or digest, and it does not bind a post-receipt `candidate.json`
or outer evidence manifest. Generation is acyclic: source/package/evidence bytes and their package manifest are
fixed first; the receipt binds those inputs; only after the receipt bytes are final do `candidate.json`, the outer
evidence manifest, and any external handoff independently bind the receipt's normalized relative path, byte count,
and lowercase SHA-256. Those outer records may also repeat source/package identities for convenience, but no digest
edge points back from the receipt to an object that already hashes the receipt. A receipt self-hash or mutual-hash
cycle is invalid.
