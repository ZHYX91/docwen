# Resolved structured numbering and export plan / 已解析结构化编号与导出计划

Status: `CURRENT_NORMATIVE_CONTRACT`.

DocWen accepts only a provider-neutral resolved document plus a resolved numbering/export plan. The upstream semantic
provider selects the profile and resolves counters. DocWen materializes the supplied result into DOCX; it is not a
second numbering-rule authority.

## Number authority and source invariance / 编号权威与源码不变性

Only Structured Numbering produces a derived number. Text that an author types in a Heading remains indivisible
authored title text. For example, the title of `## 2.3 标题` is exactly `2.3 标题`; `2.3`, `第二章`, `一、`, Roman
numerals, and localized lookalikes are never parsed, removed, hidden, migrated, or reused as a number by normal
parse, Live/Read presentation, cross-reference resolution, export, or import. If Heading numbering is enabled, a
derived prefix and that complete authored title may both be visible.

Figure, Table, Equation, and Code captions likewise have no authored Markdown number source. Their declaration text
is caption content, while their number is derived. A number is never written into the `Figure:`, `Table:`,
`Equation:`, or `Code:` declaration content. Enabling, disabling, or changing a numbering profile produces a
byte-for-byte zero Markdown diff. There is no v4 cleanup command and no authored/manual-number source.

The five closed kinds are `heading`, `figure`, `table`, `equation`, and `code_block`. Each kind has one effective
numbering state:

- enabled means the upstream owner supplies a materializable derived number for the target; a consumer may display it in
  editing and reading views, use it for TOC and semantic cross-references, and ask DocWen to materialize it on
  export;
- disabled means the target has no materializable number. A Heading level whose selected template is empty is
  disabled for that target even when the Heading kind is otherwise enabled; and
- v1 has no separate state in which a number is hidden in the views but remains available to cross-references or
  export. A semantic `@[[...]]` that resolves to a disabled target displays its Alias or current target title
  with an empty `cached_number`; it never invents a number. An ordinary `[[...]]` remains a navigation link.

The upstream semantic provider owns the five independent counters and every rule that
produces their resolved values: enablement, format, localized label, start/reset, document-wide versus Heading
scope, chapter-number inclusion, and separator. Font, position, and spacing belong to Export Style. DocWen receives
only the resolved result and never interprets a private consumer combination scope, chooses a profile, resets a counter from a
Heading title, or derives a value from a WikiLink.

## Provider-neutral Conversion Port / Provider 中立转换端口

One request snapshot binds the resolved document and resolved numbering/export plan to the same immutable input
identity and content SHA-256. The plan enumerates every semantic target occurrence in document order, its closed kind,
Heading level where applicable, effective enabled state, exact derived number when enabled, and exact portable
materialization. The neutral document carries already-resolved reference facts whose cached number must equal the
target's plan value. The document contains authored content and structural identities; the plan contains presentation facts.
Workspace paths, Node IDs, consumer types, resolver instructions, counter rules, and consumer objects are forbidden.

The two required input resources and their schemas are frozen in
[Machine Protocol v1](machine-protocol-v1.md#resolved-numbering-inputs--已解析编号输入). Both envelopes carry the
same `input_id`, `source_sha256`, and `plan_sha256`. `plan_sha256` is the lowercase SHA-256 of the RFC 8785 canonical
UTF-8 bytes of the plan resource's closed `plan` member only, so the pointer is not a self-hash. The neutral-document
envelope points to that digest; the plan envelope repeats it and is accepted only after recomputation.

For every target, `enabled=true` requires a non-empty exact `derived_number` and one closed materialization record.
`enabled=false` requires JSON null for both. A missing plan is
`docwen.numbering_export_plan.missing`; malformed JSON/schema, pointer/hash mismatch, missing/duplicate target,
kind/Heading-level mismatch, or enabled/derived/materialization contradiction is
`docwen.numbering_export_plan.invalid`. These are admission failures with no Markdown range or fix. They are never
treated as a valid disabled plan. A syntactically valid plan that requests a form outside the portable set below is
`docwen.numbering_export_plan.unsupported_materialization`. All three fail before package mutation; DocWen does not
repair or recompute the plan.

The closed `docwen.markdown_semantics.v3` source projection requires a non-empty `number` on every target and is a
separate source oracle, not authority for this contract's disabled state. The resolved-plan route uses its own closed
semantic/plan schema identities, corpus, and capability. Machine options may not carry the plan as an opaque JSON
string or private metadata bag.

### Resolved document dependencies / 已解析文档依赖

The exact-two port does not authorize a fallback to a source-relative file, Workspace path, citation database, or a
third Machine input. The closed `document` member has exactly
`authored_markdown,targets,references,resource_occurrences,citations,resources`. Empty arrays are explicit.

Each `resource_occurrences[]` record has exactly
`source_start,source_end,source_slice_sha256,authored_token,authored_locator,resource_id`. Its Unicode-code-point range
selects the complete Markdown image/embed token, `authored_token` equals that slice, its digest authenticates the
slice, and the separately parsed locator points directly to one `linked_resource` record by opaque ID. Records sort
by `(source_start,source_end,resource_id)`, do not overlap, and are the only authority for binding an authored image
occurrence; DocWen performs no base-path, global locator replacement, or filesystem resolution. The same locator at
different ranges may resolve to different resource IDs, and multiple occurrences may share one resource ID.

Each `resources[]` record has exactly
`resource_id,role,media_type,size_bytes,sha256,content_base64`, sorts uniquely by resource ID, and uses canonical
RFC 4648 base64. `role` is `linked_resource|bibliography`. Linked resources use the closed image media set
`image/png|image/jpeg|image/gif|image/bmp|image/webp`, are non-empty, pass media magic/dimension/content validation,
and each is referenced by at least one occurrence. SVG is not a v1 linked-resource media type because external
dependencies and host-safe rendering are not frozen. At most one bibliography resource exists and its media type is exactly
`application/vnd.docwen.semantic-bibliography+json`; its decoded bytes must independently pass the closed
`docwen.semantic_bibliography.v1` parser. Every decoded payload is single-read, then checked against its declared byte
count and lowercase SHA-256 before any request-owned staging write.

Each `citations[]` occurrence has exactly
`source_start,source_end,source_slice_sha256,authored_token,form,cluster_id,items,cached_result`. Records are sorted and
non-overlapping; the range selects the complete authored token and authenticates it. `form` is
`narrative|parenthetical`, cluster IDs are unique portable IDs, and `items` contains 1..64 ordered closed records with
exactly `citation_key,record_id,record_sha256,presentation`. Citation keys use
`[A-Za-z0-9][A-Za-z0-9_-]{0,127}`, keys do not repeat in one cluster, record identities/presentations are non-empty,
and record digests are lowercase SHA-256. `cached_result` is the non-empty provider-presented cluster result; DocWen
does not run a citation resolver or reinterpret a key as target identity.

DOCX lowers those arbitrary provider identities through the independent closed item/occurrence authority in
[Resolved citation item and occurrence authority](markdown-compatibility.md#resolved-citation-item-and-occurrence-authority--已解析引文条目与出现权威).
The physical `DWCIT_` tag and `_DWC_` bookmark are collision-checked package addresses only; the custom maps retain
the complete `record_id`, `record_sha256`, presentation, key, cluster, source range/token, and provider cache. Thus a
valid identity such as `reference-record:98` is never rejected, truncated, or replaced by its authored key.

Decoded resources total at most 6,000,000 bytes and the entire neutral-document envelope remains at most 8 MiB.
Non-canonical base64, unsupported media, missing/unused/duplicate resource pointers, invalid bibliography/citation
records, either limit, or any attempt to use an external path is `docwen.resolved_document.invalid` before package
mutation. The plugin decodes only authenticated records into request-owned staging and never reads beside the neutral
JSON or authored locator.

### Closed portable materialization / 闭合可移植物化形式

The closed `plan` member has exactly `heading_definitions`, `heading_instances`, and `targets`. IDs match
`[A-Za-z][A-Za-z0-9-]{0,63}` and are unique. Definitions and instances are ordered by first target use; targets are
ordered by `(source_start,source_end,kind)`. Every target record has exactly
`source_start,source_end,kind,enabled,target_id,derived_number,materialization`; ranges are zero-based Unicode-code-
point offsets with exclusive end into the authenticated source, `target_id` is a valid source ID or JSON null, and
the five kinds are closed as above.

Each Heading definition has `definition_id` and 1..9 increasing Word `levels`. A level has exactly
`level,start,number_format,display,suffix,restart_after_level`: `level` is 1..9, `start` is 1..2147483647, `suffix`
is `nothing|space|tab`, and `restart_after_level` is null or a lower level. `number_format` is exactly one of
`chinese_lower`, `chinese_upper`, `arabic_half`, `arabic_full`, `arabic_circled`, `letter_upper`, `letter_lower`,
`roman_upper`, or `roman_lower`, mapped respectively to OOXML `chineseCounting`, `chineseCountingThousand`,
`decimal`, `decimalFullWidth`, `decimalEnclosedCircleChinese`, `upperLetter`, `lowerLetter`, `upperRoman`, and
`lowerRoman`.

`display` is 1..19 closed segments, each exactly `{"counter":{"level":N,"number_format":"..."}}` or
`{"literal":"..."}`. A literal is 1..32 XML-1.0 Unicode code points with no control character or `%`. Counter
levels are no greater than the containing level, each referenced level occurs at most once, the containing level
occurs exactly once, and every reference repeats that referenced level's single declared format. DocWen translates
counter segments to `%N` and literals byte-for-text into `w:lvlText`; conflicting formats, forward references,
unknown segments, or lossy translation are unsupported rather than approximated.

Each Heading instance has exactly `instance_id,definition_id,starts`; `starts` is an increasing closed array of
`{"level":N,"value":V}` overrides. A reset starts a new instance; DocWen never infers one from title text or scope.
Definitions become `w:abstractNum` in first-use order, instances become `w:num`, and IDs use the lowest unused legal
integer without rewriting pre-existing numbering. Each enabled Heading target's materialization is exactly
`{"type":"heading_list","definition_id":"...","instance_id":"...","level":N}` where `N` is 1..9 and exactly
matches the provider-neutral authored Markdown Heading target. DocWen writes the resolved
Heading style, `w:numPr/w:ilvl=N-1`, and that instance's `w:numId`. Deterministic counter simulation over target order
must equal `derived_number`; Heading paragraphs contain no separate cached-number run.

The 1..9 definition/instance grammar preserves Word's complete Heading/list model. Authored Markdown uses CommonMark
levels 1..6 plus DocWen ATX extension levels 7..9, so every used `heading_list` binding and every caption
restart/chapter binding are restricted to 1..9 and can bind to the corresponding Word Heading level.

Each definition emits `w:multiLevelType="multilevel"`. Level `start`, mapped `number_format`, translated `display`,
`suffix`, resolved Heading style, and `restart_after_level` become exact `w:start`, `w:numFmt`, `w:lvlText`, `w:suff`,
`w:pStyle`, and `w:lvlRestart`; null restart serializes `w:val="0"`, otherwise the exact lower one-based level.
Instance start entries become `w:lvlOverride w:ilvl="N-1"/w:startOverride w:val="V"`. Missing/extra level XML or a
different effective number is invalid, not an accepted host normalization.

Every enabled caption materialization has exactly
`type,counter,number_format,sequence_action,start_value,restart_heading_level,restart_heading_style,chapter_heading_level,chapter_heading_style,chapter_separator,chapter_cached_number,sequence_cached_number,localized_label,label_separator`.
`type` is `simple_seq|chapter_seq`; `counter` is the kind-bound `Figure|Table|Equation|Code`; and `sequence_action`
is `continue|reset_to_start|restart_by_heading_level`. Caption `number_format` is the portable field-switch subset
`arabic_half|letter_upper|letter_lower|roman_upper|roman_lower`, mapped to
`ARABIC|ALPHABETIC|alphabetic|ROMAN|roman`. Separators are 1..8 XML-1.0 code points without controls.
`localized_label` is 1..64 XML-1.0 code points and `label_separator` is 0..8; both are visible outside the target
bookmark and cannot change the counter result.

`type` and `sequence_action` are independent dimensions. `simple_seq` has null
`chapter_heading_level,chapter_heading_style,chapter_separator`; `chapter_seq` requires a 1..9
`chapter_heading_level`, the matching `chapter_heading_style=heading_N`, and a non-null separator. The chapter fields
select the visible `STYLEREF` component only. They never imply a counter restart. `simple_seq` requires
`chapter_cached_number=null` and a non-empty `sequence_cached_number` exactly equal to the target's `derived_number`.
`chapter_seq` requires non-empty chapter and sequence cached numbers whose exact concatenation
`chapter_cached_number + chapter_separator + sequence_cached_number` equals `derived_number`. DocWen never splits a
composite number on the separator or recomputes either cached component from the total.

For `continue`, `start_value,restart_heading_level,restart_heading_style` are null and the exact local field is
` SEQ <counter> \* <switch> `. For `reset_to_start`, `start_value` is 1..2147483647, both restart-Heading members are
null, and the exact local field is ` SEQ <counter> \r <start_value> \* <switch> `. For
`restart_by_heading_level`, `start_value` is exactly `1`, `restart_heading_level` is 1..9, and
`restart_heading_style` is exactly `heading_N` for that restart level; the exact local field is
` SEQ <counter> \s N \* <switch> `. A heading restart with any start other than 1 is not representable by this portable
form and fails closed. The restart Heading level/style and chapter Heading level/style are separately resolved facts
and may differ; DocWen must not infer either from the other.

For `chapter_seq`, the exact sequence is the complex field
` STYLEREF "<resolved-chapter-heading-N-name>" \n `, the literal chapter separator, then the action-selected local
`SEQ` field above. For `simple_seq`, only that local `SEQ` field is emitted. Thus all six closed combinations of
`simple_seq|chapter_seq` and the three actions are representable. The STYLEREF and restart-style names are exact
direct `w:name` values selected by their respective resolved Heading-N style bindings, never localized guesses or
authored titles.

For simple captions the target bookmark encloses the complete `SEQ` field. For chapter captions it encloses the
complete `STYLEREF` field, separator run, and `SEQ` field as one contiguous composite. Every field has the exact cached
result supplied by `chapter_cached_number` or `sequence_cached_number`, and their visible concatenation equals
`derived_number`; localized label and label separator stay outside the
bookmark. A caption `REF` cached result equals the complete bookmarked derived number. A plan requiring another
field switch, arbitrary static number text, a non-representable reset/scope, or a format that differs for one reused
counter is unsupported and cannot be flattened. Corpus includes `start_value != 1`, two independent sequence scopes,
all three actions, and Update Fields in Word/WPS/LibreOffice so cached success cannot hide a wrong live field.

## DOCX materialization / DOCX 物化

DocWen implements the physical layer independently of any particular DOCX provider:

- enabled Headings use real `word/numbering.xml` abstract-number/number instances and paragraph list semantics that
  deterministically produce the plan value; disabled or template-empty Headings receive no effective
  numbering binding from DocWen;
- enabled Figure/Table/Equation/Code captions use the existing managed caption styles and matching `SEQ` complex
  or closed `STYLEREF`+separator+`SEQ` fields with exact cached results. Disabled captions retain their declaration
  content and managed style but have no
  `SEQ` field or derived number;
- an ID-bearing target retains its deterministic `DW_T_` navigation bookmark and target metadata regardless of
  enablement. An enabled simple caption bookmark encloses its `SEQ`; an enabled chapter caption bookmark encloses the
  complete `STYLEREF`+separator+`SEQ` composite. For an unnumbered
  caption it is a zero-width bookmark at the start of the caption paragraph, before all visible label/title runs.
  A Heading bookmark remains bound to its Heading paragraph;
- a resolved numbered Heading reference uses the Word equivalent of `REF <bookmark> \n \h`; other numbered target
  references use `REF <bookmark> \h`. Cached results exactly equal the resolved plan and Alias remains outside the
  field. An unnumbered semantic reference is rejected before rendering and creates no `REF` occurrence; and
- style completion remains the 43-style request-local non-destructive operation. Counter rules do not move into the
  style registry, and Export Style never changes numbering identity or value.

Disabled ID-less captions round-trip through the independent closed
`document-numbering-occurrence-map/v1` and physical two-block occurrence SDT frozen by
[Markdown compatibility](markdown-compatibility.md#disabled-id-less-caption-occurrence-authority--禁用且无-id-题注出现权威).
They are not inferred from style/adjacency and receive no target, bookmark, field, or hidden ID.

DocWen never writes a derived number into Markdown, rewrites a Heading, interprets a WikiLink, or runs an upstream
resolver. Any provider consumes the same Conversion Port, export plan, corpus, and physical acceptance contract.

## DOCX-to-neutral extraction / DOCX 到中立语义提取

Import separates a number from visible text only when the package proves Word semantics: Heading list numbering
must be backed by valid `numbering.xml` plus paragraph numbering bindings, and caption/reference values must be
backed by valid `SEQ`/`REF`, bookmark, cached-result, and DocWen-owned map/SDT relationships where this specification
requires them. Valid proof produces neutral numbering facts; the calling provider or consumer decides the final
Markdown representation, and semantic numbering is never written into a Heading or caption declaration.

A visible prefix without sufficient Word proof remains authored text and produces a stable ambiguous-number-prefix
diagnostic. Import must not guess from punctuation, locale, glyph style, indentation, or resemblance to a configured
profile. Thus a plain DOCX paragraph visibly beginning `2.3 标题` remains `2.3 标题` unless authentic Word list
semantics prove a separate number.

### Independent document recovery / 独立文档回转

Markdown→DOCX produces one DOCX artifact. DOCX→Markdown reconstructs current document content from the package's
text, styles, list/SEQ/REF semantics and supported semantic carriers. No original-source replay or adjacent
`.docwen` file is generated or consumed. The original Markdown and numbering plan are conversion inputs only;
round-trip tests may retain them as independent expected data, never as reverse-converter input.

Semantic carriers describe document structures, references and literal code; they do not embed an original-document
backup. Ordinary body edits are extracted from the current DOCX. Invalid semantic carriers still fail explicitly.
There is no missing-source-snapshot diagnostic and no byte-identical Markdown restoration promise.

Unnumbered resolved targets accept `cached_number=""`. Their references display the authored Alias when present,
otherwise the current target title (or the kind label for an empty title). They use a validated inline text carrier
without inventing a list number, SEQ value, or REF field. The authored reference token remains recoverable from that
carrier when output reference syntax is enabled.

## Required matrix / 必须矩阵

| Kind/case | Provider presentation | Semantic `@[[...]]` | Ordinary `[[...]]` | Markdown diff | DOCX physical result |
|---|---|---|---|---|---|
| Heading enabled | derived number visible | resolved number | navigates | zero | exact numbering.xml/list semantics; no Heading cached-number run |
| Heading disabled | no derived number | current title/Alias | navigates | zero | no effective numbering |
| Heading-level template empty | no derived number | current title/Alias | navigates | zero | no effective numbering |
| Figure enabled | derived number visible | resolved number | navigates | zero | Figure `SEQ` + cached result |
| Figure disabled | no derived number | current title/Alias | navigates | zero | caption style, no `SEQ` |
| Table enabled | derived number visible | resolved number | navigates | zero | Table `SEQ` + cached result |
| Table disabled | no derived number | current title/Alias | navigates | zero | caption style, no `SEQ` |
| Equation enabled | derived number visible | resolved number | navigates | zero | Equation `SEQ` + cached result |
| Equation disabled | no derived number | current title/Alias | navigates | zero | caption style, no `SEQ` |
| Code enabled | derived number visible | resolved number | navigates | zero | Code `SEQ` + cached result |
| Code disabled | no derived number | current title/Alias | navigates | zero | caption style, no `SEQ` |
| `## 2.3 标题`, Heading enabled | derived prefix plus complete `2.3 标题` | derived number | navigates | zero | list number plus unchanged authored title |

Additional fixtures toggle every kind on→off→on and prove byte-identical Markdown; distinguish ordinary WikiLinks
from semantic references; cover Alias with a materialized number and fail-closed Alias on an unnumbered target; and
prove ambiguous visible prefixes remain authored text on DOCX import. MD→DOCX fixtures inspect exact Heading
numbering/list semantics, caption `SEQ`, target bookmarks, `REF` instructions, cached results, managed styles, and
absence for disabled/template-empty cases. DOCX→neutral→Markdown fixtures prove semantic numbering does not become
source text.

Add/remove/regex-cleanup controls belong only to the capabilities that declare them. They are rejected on the
resolved-plan route, do not define manual-number semantics, and are excluded by its negative capability gate.

## Evidence layers / 证据分层

Evidence is recorded without substitution:

1. `source_oracle` binds the authenticated Markdown bytes, resolved-plan bytes/schema identity, expected neutral
   projection, zero-diff assertion, and each positive/negative matrix case.
2. `packaged` proves the installed converter consumes that same closed Conversion Port and rejects stale, missing,
   private, or contradictory plan data.
3. `headless_ooxml` inspects exact `numbering.xml`, paragraph numbering bindings, `SEQ`, bookmarks, `REF`, cached
   results, styles, custom maps/SDTs, and their required absence.
4. `roundtrip` binds DOCX-to-neutral and final canonical Markdown bytes separately from source expectations.
5. `word_host`, `wps_host`, and `libreoffice_host` each prove open/save/reopen preservation of the exact candidate;
   none is replaced by headless XML or another host.
6. Provider presentation/readback evidence is provider evidence and is bound by identity, but never represented as DocWen source,
   package, XML, or host evidence.

The candidate receipt binds every layer to the final spec commit/tree, implementation commit/tree, plan/schema and
corpus manifest identities, package manifest and binary digest. A source oracle is not packaged or wire evidence,
and a host screenshot is not proof of Markdown round-trip.
