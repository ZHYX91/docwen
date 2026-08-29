# Golden regression suite / Golden 回归套件

Golden fixtures capture stable semantic projections or artifact properties that are expensive to express as isolated unit assertions. They validate current behavior; they are not a repository for development history.

Golden fixture 保存稳定的语义投影或产物属性，用于覆盖不适合拆成单元断言的行为。它们只验证当前行为，不保存开发过程。

## Locations / 位置

- `tests/fixtures/golden/`: JSON semantic and artifact projections.
- `tests/fixtures/files/`: source documents used by focused tests.
- `tests/golden/`: golden test entry points.
- package-local test fixture directories for plugin-owned cases.

## Update policy / 更新规则

1. Explain the behavior or oracle change before updating a fixture.
2. Keep source inputs immutable and record provenance when redistribution permits it.
3. Prefer semantic projections over unstable binary or pixel equality.
4. Use rendered comparisons only when layout is the contract and the renderer is fixed.
5. Review fixture diffs; never regenerate baselines merely to make a failure green.
6. Run the owning focused tests and the affected release/canonical gate.

Candidate-blocking corpora additionally include:

- exactly 43 managed styles in every locale and exactly 47 physical styles in a four-style mother document; canonical
  built-in names, four distinct caption styles, zero aliases, request-local non-destructive conflict mapping, and
  renderer references through the resolved map. Footnote/endnote fixtures start with colliding positive IDs, reserved
  separator IDs, rich multi-block bodies, hyperlinks/drawings, and existing relationships; they prove atomic new ID
  allocation and independent Word/WPS/LibreOffice save-reopen preservation rather than treating style presence as
  note fidelity;
- the `docwen.markdown_semantics.v3` / `docwen.markdown_diagnostics.v3` source oracles remain immutable.
  Because their target shape requires a non-empty number, the resolved-plan route uses a separate closed
  semantic/plan schema identity that can represent enabled and disabled targets without editing v3 in place. The
  shared `[A-Za-z0-9-]{1,128}` ID carries no type. Only its owning structure supplies meaning: Heading and valid
  Figure/Table/Equation/Code declarations are semantic targets; paragraph, list item, raw image/table/equation/code,
  quote, callout, list, or fenced-block anchors are navigation/embed targets only. A raw object anchor never creates
  a numbered target even if its spelling resembles a historical prefix;
- inline anchors for simple paragraphs, individual list items, and single image embeds; standalone post-block
  anchors for whole lists, block quotes, callouts, tables, display math, fenced code, Mermaid, query, and view blocks.
  The marker binds the immediately preceding complete block at the same container depth and is never paragraph text.
  Every DOCX-map record for a fenced block uses ordinary `block_kind=code_block`, including Mermaid, query, and view;
  no `fenced_block` value is accepted on that boundary.
  Fixtures cover zero/one/multiple separators, EOF, no predecessor, an intervening block, depth mismatch, duplicate
  ownership, and closing-fence/closing-`$$` suffix rejection;
- every source-oracle anchor has a required closed `container_path` of outer-to-inner
  `block_kind,block_range` source segments; top-level is empty and no segment contains a consumer path or ID. Fixtures
  prove the longest-proper-prefix direct-parent rule for a multi-block strict subset, a multi-paragraph inline list
  item whose complete owner range is inside its whole-list anchor, a quote-only Mermaid fence whose inner/outer
  ordinary SDTs cover the same visible paragraph,
  a list inside a quote, and a quote inside a list. The exact IDs are `inner-fence`/`outer-quote`,
  `inner-list`/`outer-quote`, and `inner-quote`/`outer-list`. Two disjoint top-level paragraph anchors are the exact
  negative authority: they emit no topology map;
- every nested ordinary-anchor DOCX has one separate closed `document-anchor-topology-map/v1` custom XML item and one
  direct `edge` per child, with exact `child_tag,parent_tag,sha256` attributes, canonical edge digest preimage,
  deterministic ordering/bytes/UUID, and an independently allocated item/itemProps/relationship/content-type trio.
  The existing `document-target-map/v1` bytes and tag-only ordinary SDTs remain unchanged. Fixtures prove a unique
  acyclic forest, at most one parent per child, ordinary-anchor-only endpoints, parent-outside-child physical nesting,
  both strict and equal flattened body-element ranges, and inner-to-outer recovery. Missing/extra edge or map, two
  parents, cycle, target/fenced endpoint, swapped tags, reversed wrappers, partial overlap, digest/byte/UUID/topology
  tamper, and host stripping all fail closed;
- separate observable projections for ordinary `[[Page#^id]]`/`![[Page#^id]]`, same/cross-document stable-ID
  `@[[#^id]]`/`@[[Page#^id]]`, same/cross-document soft Heading
  `@[[#Heading]]`/`@[[Page#Parent#Heading]]`, aliased numbered references, bare `@citation-key`, and grouped
  `[@key; @key]` citations. Soft fixtures cover a unique same-document match, missing and ambiguous fail-closed,
  authored cross-document `external_unresolved`, and externally supplied neutral resolution; they never select an
  ordinary anchor or caption. A numbered Alias retains the materialized number plus Alias and exercises the stale
  warning; an Alias on an unnumbered target fails closed and is not displayed alone.
  External target/citation records bind their own identity/hash. No fixture uses Machine options as resolver context
  or claims that DocWen scanned a Workspace;
- separately hash-addressed invalid fixtures for dot, slash, underscore, more than 128 characters, duplicate ID, and
  declaration/object kind mismatch, plus dangling anchor, missing semantic target, ordinary-anchor-as-semantic-ref,
  unnumbered target, empty Figure/Table/Code caption, and empty ID-less Equation. Every diagnostic observation binds
  the exact `docwen.machine.diagnostic_evidence.v1` source identity/hash, Unicode-code-point exclusive range,
  related ranges, fix ID, ordered edits, and unchanged source on rejection;
- one resolved-numbering matrix for all five kinds (`heading`, `figure`, `table`, `equation`, `code_block`) with
  enable on/off, plus Heading-level template empty. Source fixtures include `## 2.3 标题` and prove that its complete
  authored title survives when a derived prefix is also visible; no fixture/parser recognizes `2.3`, `第二章`, `一、`,
  or localized lookalikes as a number. Toggling each kind on→off→on produces a byte-identical Markdown SHA-256.
  Enabled cases resolve semantic `@[[...]]`; disabled/template-empty cases emit
  `docwen.markdown.cross_reference.unnumbered_target`; ordinary WikiLinks keep navigation in every row;
- exact-one `neutral_document` and `numbering_export_plan` inputs bind the current schema/URN/media identities and the
  same `input_id,source_sha256,plan_sha256`, with the plan digest recomputed over RFC 8785 canonical plan-member bytes.
  Missing plan, invalid/pointer-mismatched plan, and valid-but-unsupported materialization use three distinct
  diagnostics and never become disabled/unnumbered. Corpus proves the exact one-to-one
  `interop.cross_reference.unnumbered_target` ↔ `docwen.markdown.cross_reference.unnumbered_target` mapping;
- exact-two resource preservation uses only range-bound locator→opaque-ID occurrences and embedded, byte/hash/media-
  checked payloads; bibliography and resolved citation records are closed; missing/tampered/unused/oversize records
  fail before staging. Figure/image-owner, YAML, fenced, cross-reference, citation and bibliography fixtures prove no
  source-relative read, third input, or options-side payload is used;
- citation physical fixtures preserve `reference-record:98` as stable identity while its authored key remains a
  lookup key; inspect the exact item/item-ref/occurrence digest preimages, `DWCIT_<digest32>` and `_DWC_<digest35>`
  addresses, two independent canonical custom-XML trios, locked clean `CITATION`/`\m` fields, provider cache,
  inline SDT/bookmark containment, and full round-trip. Negatives cover full/truncated digest collision, record-
  version/presentation map mismatch, missing/extra/reordered maps/refs, source range/token/hash, key/record substitution,
  unlock/dirty/cache/instruction/OPC/UUID/host-strip tamper, and prove Word/WPS/LibreOffice separately from headless;
- Heading materialization covers the closed nine-format token grammar, definitions/instances, levels, starts,
  restart boundaries, exact `abstractNum`/`num`/`numPr` output, and deterministic comparison to the plan value without
  inventing a cached Heading run. Caption materialization covers `continue`, `reset_to_start`, and
  `restart_by_heading_level`; exact no-action/`\r N`/`\s N` SEQ lowering; all six independent
  `simple_seq|chapter_seq` × action combinations; separately resolved chapter-display and counter-restart Heading
  levels/styles (including an unequal-level case); simple and composite bookmark/cached results; `start_value != 1`;
  independently supplied chapter/sequence caches including a chapter value that contains the same separator; two
  independent scopes; and Word/WPS/LibreOffice Update Fields. Unsupported cross-format references, heading
  restart start other than 1, reset/scope, or field switch fail closed;
- exact `DW_T_` plus 35-hex target bookmark derivation and synthetic full/truncated digest collision rejection. An
  enabled caption bookmark contains only its `SEQ` field and cached number; an ID-bearing disabled caption uses the
  same deterministic name as a zero-width navigation bookmark at paragraph start and has no `SEQ`/`REF`. An
  ID-bearing Heading target uses its heading-number REF
  projection. Alias runs remain outside `REF`. Every addressable target has one `docwen-target-v1:` SDT and reversible
  custom XML metadata. A caption or Heading without an ID may still be numbered but creates no bookmark/target map;
  an ordinary anchor has zero bookmark, hyperlink target, `SEQ`, and `REF` and round-trips only through its opaque
  anchor SDT/map. Wrong/missing/duplicate tag, identity, hash, relationship, content type, nesting, kind, cardinality,
  and ordering all fail closed;
- every stable-ID `REF` occurrence and soft Heading-path occurrence resolved to an ID-bearing Heading has exactly one
  `docwen-ref-occurrence-v1:` inline SDT and a record in the separate closed
  `document-reference-occurrence-map/v1` custom XML item. Fixtures inspect the exact digest preimage, authenticated
  Unicode source range and complete authored token including Alias, resolved bookmark, cached number, field/Alias
  containment, item/itemProps/relationships/content types, deterministic UUID, ordering, and rejection of missing,
  duplicate, overlapping, tampered, host-stripped, or cross-linked metadata. The occurrence never becomes a target,
  second ID, or target-map entry;
- a uniquely resolved soft reference to an ID-less Heading emits static cached number text plus Alias, zero ID
  writeback, bookmark, target-map entry, and REF/hyperlink field. It round-trips through exactly one
  `docwen-soft-ref-v1:` inline SDT and the separate closed document-soft-reference custom XML map. Fixtures inspect
  digest preimage, Unicode source range/token, cached text, item/itemProps/relationships/content types, duplicate/
  missing/tampered metadata, and Word/WPS/LibreOffice preservation; inability to preserve it fails conversion rather
  than flattening the semantic token;
- single-document revision/hash rewrite plans update only authenticated local `@[[...#^id]]`, `[[...#^id]]`, and
  `![[...#^id]]` occurrences, reject stale/overlap/ambiguous plans without partial write, and never claim Workspace
  rename coverage. Provider fixtures prove a total diagnostic/fix mapping without string fallthrough or coordinate
  loss;
- Figure/Table/Equation/Code positives with and without IDs, enabled and disabled ID-less non-empty Heading/caption,
  non-empty ID-less Equation/Code, and empty Equation/Code with ID. Enabled DOCX recovery proves exact caption style,
  matching single `SEQ` and cached number, and independent preservation of the caption's semantic kind and the
  carrier's native image/table/equation/code structure. Physical order follows semantic kind: Figure
  carrier→caption versus Table/Equation/Code caption→carrier, including cross-type pairs. A disabled declaration
  requires authenticated plan/occurrence authority;
  style plus adjacency alone is rejected. An ID-less caption has no target pairing SDT/map, hidden ID, bookmark, or
  `REF`; fixtures reject reversed/opposite-side order, an intervening block, missing/unproved native carriers,
  style/`SEQ` mismatch, and multiple claims.
  A target SDT exists only when addressable, and Markdown emits no pairing token.
  Markdown→DOCX exact Heading `numbering.xml`/list and caption `SEQ`/bookmark/`REF`/cached-result positives and
  disabled absences pass headless round-trip before separate Word, WPS, and LibreOffice host observations. Reverse
  fixtures separate a number only from authenticated Word list/field semantics; ambiguous visible prefixes remain
  authored text plus diagnostic and semantic numbering is never written into Markdown;
- every disabled ID-less caption has one canonical `document-numbering-occurrence-map/v1` record and one exact
  two-block `docwen-numbering-occurrence-v1:` SDT, bound to source hash/range, kind, false enabled state, empty
  target/derived values, and plan SHA. Fixtures inspect the digest preimage, closed attribute order, canonical
  custom-XML trio/UUID/relationship/content types, physical caption/object order, and zero target/bookmark/SEQ/REF;
  missing/extra/non-null/reordered/overlap/wrong-plan/wrapper/map/host tamper fails without style/adjacency fallback;
- any caption-bearing DOCX has one independent closed
  `document-caption-style-binding-map/v1` item with exactly four ordered
  `semantic_key,resolved_style_id,visible_name` records. Corpus evidence binds its canonical bytes, deterministic
  UUID, item/itemProps/relationship/content-type topology, exact unique paragraph style ID, exact direct name, zero
  aliases, and the exact direct `pStyle` on every addressable and ID-less caption. A real conflicting template proves
  the user style is preserved, the request-local collision-free ID survives save/reopen, and recovery uses that ID.
  Negatives cover a forged requested-ID prefix, missing map/style/`pStyle`, wrong type/name, alias, cross-kind style,
  duplicate resolved ID/style ID, non-closed/reordered records, changed bytes/UUID/topology, and map stripping. This
  map is style identity only and must never be counted as caption pairing, a target, an anchor, or a hidden ID;
- CommonMark fenced-code round trips prove that inner fence-looking lines remain literal and exactly one Code Block:
  backtick outer with literal tilde runs, a four-backtick outer containing a triple-backtick line, bodies containing
  both fence characters, list and blockquote containers, and MD→DOCX→MD. New serialization uses matching opener and
  closer characters, length at least three and longer than every conflicting same-character body run, shortest-safe
  deterministic selection with backtick tie-break; an existing safe authored fence is preserved. Code body uses
  only `Code Block`; only an external `Code:` declaration may use `Code Block Caption`. The source oracle freezes one
  exact `fenced_sources` record for every rust/Mermaid/query/view occurrence, including direct mixed-EOL, blockquote
  CRLF, list prefixes, and omitted closer at EOF. Package fixtures prove the separate
  `document-fenced-source-map/v1`, complete inline `docwen-fenced-source-v1:` payload SDT, canonical field/attribute
  order, bounded RFC 4648 framing, and exact source/range/block/body hashes. Caption target plus ordinary raw-code
  anchor retains all three independent layers; the carrier alone yields zero bookmark/`SEQ`/`REF`. Negatives cover
  non-canonical base64/decimal/EOL, changed prefix/info/body/closer, swapped or overlapping occurrences, duplicate or
  unmapped tags, partial payload wrapping, map/topology/content-type/relationship tamper, and visible body tamper;
- normal parse/convert/save never recognizes `{#id}`, typed-prefix declaration rules, Listing/listing/lst-, or
  List/list- as current grammar. Bare `@fig-legacy` is a Citation with key `fig-legacy`; DocWen exposes no alternate
  migration grammar;
- DOCX, PDF, OFD, XPS, and TIFF closed Machine option schemas: both common booleans with defaults, all retained
  route-specific controls, `required=[]`, `additionalProperties=false`, and rejection of old public boolean names.
  DOCX exercises all four boolean combinations in both `main_md` and `image_md`: always one primary entry, K image
  resources iff preservation is on, no OCR fragment for `main_md`, and R `fragment_of/ocr_text` fragments for the R
  non-empty results only when recognition plus `image_md` is selected. It asserts zero physical-page payloads and
  that image presentation mode cannot turn a resource into a fragment or couple the two booleans. With the default
  `image_mode=file`, Figure and ordinary image-anchor fixtures additionally close the six distinct observable
  combinations: `preserve_resources=false`
  retains the exact `![image omitted]()` owner plus artifact-bound
  `DOCX2MD-IMAGE-OWNER-RESOURCE-OMITTED`; `main_md` OCR follows that owner; and `image_md` keeps the owner only in the
  primary while its sidecar is OCR-only. Negative assertions reject close carrier spellings, any carrier/ID/declaration
  copied into an OCR fragment, any sidecar embed replacing the primary owner, any resource when preservation is off,
  and any Markdown range/fix/source-evidence claim on the artifact warning. Other presentation modes prove that they
  either retain the authenticated image owner or fail closed without coercing either boolean or emitting a partial
  Bundle.
  Fixed-layout `P=4,K=5` and four-frame TIFF run all four recognition/resource combinations. They assert exactly one
  primary entry, P fragments iff recognition is on, K resource-only sidecars iff preservation is on, exact relation
  cardinality/ownership, blank/failure fragments, unresolved-resource diagnostics, and no duplicate OCR in the
  primary Markdown.

Page-count or ownership expectations must be explicit fixture facts. A golden generator may not derive them from
extracted file count, file names, or the implementation under test.

Machine Protocol v1 and Artifact Bundle v2 are frozen atomically with this semantic series; current fixtures are
immutable acceptance inputs. Packaged gates run the v3 corpus against installed CLI/resources and inspect
final OOXML. Each observation declares exactly one layer: `source_oracle`, `machine_wire`, `packaged`, `roundtrip`,
`headless_ooxml`, `word_host`, `wps_host`, or `libreoffice_host`. Source expectations are never wire observations,
and headless OOXML never substitutes for a desktop host gate.

The immutable candidate receipt binds final semantic and diagnostic identities; clean DocWen and provider
commit/tree plus each spec-baseline commit/tree; candidate ID; package manifest and executable relative path, byte
count, lowercase SHA-256, and product version; the complete options contract and four-row Bundle matrix; and every
evidence artifact directly by case ID, layer, normalized candidate-relative path, bytes, and SHA-256. Cross-reference,
citation, diagnostic, package, and round-trip records point to their exact source/manifest identity rather than a
mutable directory. Any input, source identity, binary, schema, fixture, or manifest change invalidates the candidate
and every derived receipt.

Receipt hashing is acyclic. The receipt has no self path/bytes/SHA field and does not hash `candidate.json` or the
outer evidence manifest. After the receipt is serialized, `candidate.json`, the outer evidence manifest, and the
external handoff each bind its normalized relative path, exact bytes, and lowercase SHA-256. The receipt must not
point back to any of those post-receipt records; self-hash and mutual-hash cycles are invalid fixtures.

User data and private documents must not be committed. Derived fixtures must be minimal and scrubbed.
