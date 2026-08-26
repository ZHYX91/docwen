# Physical-page OCR and artifact relations / 物理分页 OCR 与产物关系

This specification defines DocWen's provider-neutral handoff for fixed-layout documents and multi-frame TIFF.
DocWen reports artifact facts; it never emits private Workspace/Node models or consumer placement instructions.

本文冻结固定版式文档和多帧 TIFF 的消费者无关交接语义。DocWen 只报告产物事实，不输出私有 Workspace/Node
模型或消费者放置指令。

## Current authority / 当前权威

Machine `docwen.machine.v1` and protocol `1.0` require `docwen.artifact_bundle.v2`. Bundle v1 is
rejected. The relation schema, conformance fixtures, all consumers, packaged contracts, and release hashes change
atomically. A consumer that does not understand the closed relation fields must fail closed.

## Strategies and ownership / 策略与所有权

The consuming import domain has exactly two values:

```text
page_nodes
single_document
```

An importing consumer may default these physical-page facts to `page_nodes`, preview them, and change that strategy.
DocWen does not advertise an import default, name nodes, or write a workspace. Import materialization policy belongs
to the consuming domain and cannot branch on provider names.

## Physical pages, frames, and resources / 物理页、帧与资源

Let `P` be the positive number of physical pages/frames and `K` the non-negative number of preserved image resources.
`P` and `K` are independent.
Extracted image count, file names, output order, or diagnostic text must never be used to infer physical page count.

| `recognize_text` | `preserve_resources` | Artifacts | `entries` | Structural relations |
|---|---|---|---|---|
| `false` | `false` | 1 primary Markdown document | exactly 1: `role=primary`, `ordinal=0`, `preferred=true` | none |
| `false` | `true` | 1 primary Markdown document + exactly K image resources | the same single primary entry | exactly K `resource_of` relations to the primary document |
| `true` | `false` | 1 primary Markdown document + exactly P page fragments | the same single primary entry | exactly P `fragment_of/ocr_page` relations to the primary document |
| `true` | `true` | 1 primary Markdown document + exactly P page fragments + exactly K image resources | the same single primary entry | exactly P `fragment_of/ocr_page` plus exactly K `resource_of` relations |

Page fragments and image sidecars are never Bundle entries. A preserved image is always `kind=resource`, never a
Markdown fragment, document, entry, or implicit consumer Node. Enabling resource preservation without OCR therefore
still creates exactly one Markdown artifact: the primary document. The relation target records only proven artifact
ownership as described below; it does not change entry cardinality.

- PDF/OFD/XPS is enumerated by physical page after format admission. TIFF is enumerated frame by frame; the original
  multi-frame path is never passed as one OCR image attempt.
- OCR `success`, `no_text`, and every operational failure all produce a page fragment. A later failure does not stop
  remaining pages.
- Temporary rendered page images used only for OCR are not exported when `preserve_resources=false`.
- The primary document carries source/navigation context and does not duplicate the same page OCR text.
- `recognize_text=false` never creates page fragments or auxiliary page Markdown because images happened to be
  preserved.

## Closed Machine options / 闭合机器选项

Capability discovery is the public authority for the accepted option names. The fixed-layout PDF/OFD/XPS
Markdown routes each publish the same exact closed object:

```json
{
  "type": "object",
  "properties": {
    "recognize_text": {"type": "boolean", "default": false},
    "preserve_resources": {"type": "boolean", "default": true},
    "ocr_language": {
      "type": "string",
      "enum": ["auto", "chinese", "chinese_cht", "english", "japanese", "korean", "latin", "cyrillic"],
      "default": "auto"
    },
    "image_mode": {"type": "string", "enum": ["file"], "default": "file"},
    "render_dpi": {"type": "integer", "minimum": 72, "maximum": 600, "default": 200}
  },
  "required": [],
  "additionalProperties": false
}
```

The multi-frame TIFF-to-Markdown route publishes this exact closed object; its OCR default is deliberately the same
as the fixed-layout routes:

```json
{
  "type": "object",
  "properties": {
    "recognize_text": {"type": "boolean", "default": false},
    "preserve_resources": {"type": "boolean", "default": true},
    "ocr_language": {
      "type": "string",
      "enum": ["auto", "chinese", "chinese_cht", "english", "japanese", "korean", "latin", "cyrillic"],
      "default": "auto"
    }
  },
  "required": [],
  "additionalProperties": false
}
```

`recognize_text` and `preserve_resources` are negotiated independently; neither implies the other. The public keys
`to_md_enable_ocr` and `to_md_keep_images` are undeclared and rejected. A producer may translate the new names to a
legacy internal call only behind the capability boundary; such an adapter does not advertise or accept the old names.

## Closed relation semantics / 闭合关系语义

Page numbers in these payloads are **one-based**. Existing relation `ordinal` remains zero-based and, for a page
fragment, must equal `page_index - 1`.

A `fragment_of` relation with `role=ocr_page` contains exactly:

```json
{
  "page_fragment": {
    "fragment_kind": "page",
    "page_index": 1,
    "page_count": 4,
    "ocr_status": "success",
    "source_page": 1
  }
}
```

`ocr_status` is closed to `success`, `no_text`, `input_missing`, `unavailable`, `model_missing`,
`initialization_failed`, and `recognition_failed`.

A `resource_of` relation with role `image`, `original`, or `preview` may contain:

```json
{"page_resource": {"source_page": 1}}
```

Rules:

1. `page_fragment` is required only for `fragment_of/ocr_page` and forbidden elsewhere. Its source is the fragment
   and target is the primary document.
2. For one owner document, every page fragment has the same `page_count`; there are exactly P fragments; both
   `page_index` and `source_page` form the complete sequence `1..P`; ordinals form `0..P-1`.
3. When page ownership is proven and OCR page fragments exist, an exported image is `resource_of` its page fragment
   and must carry `page_resource`; its `source_page` equals that fragment's `source_page`.
4. When page ownership is proven but OCR is off, an image is `resource_of` the primary document and may still carry
   `page_resource`.
5. When page ownership cannot be proven, the resource belongs to the primary document, omits `page_resource`, and
   produces `resource_page_unresolved` bound to that resource's `artifact_id`. DocWen never guesses.
6. `derived_from` is emitted only when the source is itself a delivered Bundle artifact. Shared origin is not enough
   to invent a provenance edge.

The Bundle does not contain `page_nodes`, consumer basenames, parent IDs, private consumer fields, or a generic metadata bag.

## Diagnostics and validation order / 诊断与校验顺序

`ConversionDiagnostic` and Machine diagnostics carry optional `artifact_id`; a referenced artifact must exist in
the same Bundle. Stable page-semantic errors are evaluated before file I/O:

1. JSON/schema/envelope/task/producer and limits;
2. artifact/entry basic shape and relation references;
3. relation kind/role/owner/ordinal rules;
4. missing or unexpected page payload;
5. page ranges and ordinal/index agreement;
6. owner page-count agreement, duplicate/gap/complete sequence;
7. resource/page owner and source-page agreement;
8. cycles/orphans;
9. only then locator, link/reparse, size, and SHA-256 verification.

The stable semantic codes are `missing_page_fragment_semantics`, `unexpected_page_semantics`,
`invalid_page_range`, `page_ordinal_mismatch`, `page_count_mismatch`, `duplicate_page_index`,
`incomplete_page_sequence`, `page_source_mismatch`, `resource_page_mismatch`, and
`dangling_diagnostic_artifact`. The page-specific duplicate check owns `fragment_of/ocr_page` ordinal collisions so
that a repeated page is reported as `duplicate_page_index`; the generic `duplicate_relation_ordinal` check remains
authoritative for every other ordered relation. `page_source_mismatch` compares the complete per-owner
`source_page` sequence with `1..P`; it does not add an unstated per-item equality rule beyond that frozen sequence.

On a successful terminal result, every diagnostic `artifact_id` must name an artifact in the same Bundle. Failed or
cancelled results have no Bundle and therefore must not emit artifact-bound diagnostics.

Before commit, the Application also cross-checks producer evidence that is intentionally not serialized into the
Bundle: the primary artifact's positive physical-page count and boolean recognition/resource modes must agree with the
accepted Machine options and typed page relations; `recognize_text=false` permits no page fragment;
`preserve_resources=false` permits no image resource; and a recognition-off
resource's proven `source_page` cannot exceed the primary count. An unresolved delivered image must have exactly one
artifact-bound `resource_page_unresolved` diagnostic. A diagnostic-coverage mismatch fails with the internal stable code
`resource_page_diagnostic_mismatch`; this code diagnoses a broken producer boundary and is not a new Bundle field.
An accepted-option mismatch fails with the internal stable code `physical_page_option_mismatch` before commit.

## Capability and compatibility / 能力与兼容

Machine capability discovery must explicitly advertise fixed-layout PDF/OFD/XPS-to-Markdown and multi-frame
TIFF-to-Markdown only after their packaged paths implement this contract. Their output shape declares a primary
document, optional page fragments, optional resources, and `relation_payloads=[page_fragment,page_resource]`. It contains no
consumer import strategy or private placement hint.

The physical-page runtime route does not accept `image_md`/`main_md` placement options. Consumer placement is
outside the producer contract and cannot change the exact entry/relationship matrix above.

## Acceptance corpus / 验收语料

The canonical fixed-layout corpus deliberately uses `P=4`, `K=5`: success, blank/no-text, forced OCR failure, and
success after failure, plus one resource whose page cannot be proven. All four boolean combinations must produce
the exact table counts above, preserve the four statuses, omit duplicated OCR from the primary document, and bind or
diagnose all five resources without guessing.

A four-frame TIFF repeats success/no-text/failure/success. OCR receives four independently materialized frame
inputs; failure does not suppress frame four. Image export emits zero or four frame resources independently of the
four page fragments.

The same fixtures must pass source schema/core/Application/Machine, Windows and Ubuntu packaged Machine, a neutral
consumer import plan and transaction, and generic downstream consumers. OFD and XPS each require a real
multi-page packaged path in addition to PDF and TIFF.
