# Conversion result-directory output

Status: normative for new DocWen producers.

## Scope and identity

Conversions from Markdown to documents or spreadsheets publish a single final deliverable directly in the selected parent. Multiple final deliverables (for example worksheet CSV files), or an explicitly requested audit sidecar, publish one result directory. Internal intermediates and logs do not force grouping. All new conversions to Markdown retain a result directory, even for one Markdown output. A conversion freezes one timezone-aware creation instant and the admitted source identity before intermediate formats are produced.

Root and ordinary primary filename stem:

`{sanitized-source-stem}_YYYYMMDD_HHMMSS_from{CanonicalSourceFormat}`

Examples use `fromMd`, `fromDocx`, `fromWps` or `fromRtf` according to this conversion's admitted input, never a private JSON or DOCX intermediate. Every related output shares that timestamp. A source filename is not parsed heuristically for provenance. A later conversion uses the actual admitted input name; no adjacent JSON, hidden sidecar or persistent index is consulted for provenance.

## Layout

Every Markdown file has the same basename as its containing document node. A single generated DOCX or XLSX uses the shared conversion identity as its filename, without an enclosing result folder:

```text
项目记录_20260907_180000_fromMd.docx
```

A single CSV is also a direct file. Multiple CSV outputs share one result directory; each puts the worksheet name before the shared timestamp:

```text
项目记录_20260907_180000_fromMd/
  项目记录_人员明细_20260907_180000_fromMd.csv
  项目记录_部门_20260907_180000_fromMd.csv
```

GUI Markdown-to-spreadsheet conversion uses an XLSX template. Each worksheet is filled before conversion to the selected spreadsheet format; CSV publishes the worksheets as separate files. Explicit Machine table-export capabilities remain separate from that GUI workflow.

Gongwen produces one combined attachment node, without an attachment number or title in its name:

```text
通知_20260907_180000_fromDocx/
  通知_20260907_180000_fromDocx.md
  正文图片.png
  通知_附件_20260907_180000_fromDocx/
    通知_附件_20260907_180000_fromDocx.md
    附件图片.png
```

Attachment titles remain in the content. Other auxiliary Markdown outputs are child nodes. Images and other linked resources are placed with their owning document, and Markdown links are rewritten relative to their final logical paths. Attachment resources stay inside the attachment subtree; copying that subtree keeps its declared resources available. A byte source used by both the main document and an attachment has separate resource artifact identities and copies, one per owner. Link resolution is scoped to the referring document's resource ownership, so equal resource names in different document directories do not conflict. Original resource bytes are retained without rewriting their contents. Link relocation applies only to known local artifact targets: URI-scheme targets plus scheme-relative or UNC network targets remain byte-for-byte unchanged even when malformed or when their basename matches a local artifact; relocation does not require a target to pass complete URL validation. The raw target is split at literal query and fragment markers before matching. Only the local path portion is percent-decoded for artifact lookup, the authored query and fragment spelling is preserved verbatim, and structural delimiters in the new local relative path are percent-escaped. This keeps encoded filename characters such as `%23` and `%3A` distinct from URI structure and avoids decoding `%29` / `%22` into Markdown or HTML delimiters.

Relocation matches a producer-declared path identity exactly after separator normalization and one path-only percent decode. A missing qualified path never falls back to its basename. A bare basename is eligible only when it names one declared artifact in the document's ownership scope; an actually referenced ambiguous alias rejects publication. Unused conflicts do not reject unrelated content. The same Core full-source ownership and link grammar protects YAML, comments, code examples and container-scoped fences. Real Markdown/Wiki destinations and actual HTML `src`/`href` attributes may relocate; code samples, raw-text HTML blocks and attribute lookalikes stay literal. BOM and original line endings are retained.

重定位在路径分隔符规范化及仅路径部分的一次百分号解码后，精确匹配生产者声明的资源身份。未匹配的限定路径不得降成 basename；裸 basename 仅可指向唯一声明产物，实际引用的歧义别名拒绝发布，未被引用的冲突不影响无关正文。使用 Core 同一完整源码所有权及链接语法保护 YAML、注释、代码示例和容器围栏；只改真实 Markdown/Wiki 目标和 HTML `src`/`href` 属性。HTML 原始文字块、属性内伪链接、BOM 与原换行保持不变。

## Publication and collisions

For grouped output, the complete root is prepared in a temporary sibling directory and committed in one atomic no-replace directory rename. An existing directory is never deleted or replaced, including one created by a competing writer after the collision check. Failure or cancellation before publication exposes no partial result directory. The root collision policy is evaluated once:

- `error`: reject an existing root;
- `rename`: choose a new root suffix and rebase logical paths, preserving matching root/primary basenames;
- `overwrite`: invalid for result-directory conversions, rejected before publication (and by CLI preflight); it remains available for explicitly targeted individual files;
- `skip`: reuse only a complete regular directory whose paths and byte hashes match the newly prepared result. Missing or extra paths, changed bytes, links and junctions fail without changing the existing tree. A matching result reports reuse and zero newly written bytes.

Within a grouped result, independent filenames are never renamed separately to resolve a publication conflict. Direct single-file results retain the existing file-level error/rename/overwrite/skip policies and protection against replacing an unrelated input; their final logical basename follows any collision rename. CSV worksheet names retain their source/worksheet/timestamp identity within the chosen root.

## API and CLI contract

`OutputPolicy.output_dir` and CLI `--output-dir DIR` select the publication parent. CLI `--output PATH` may name a single document or spreadsheet result exactly, including a Markdown source. New Markdown outputs, multiple CSV files and explicitly audited grouped outputs require the parent-directory form. The finalizer decides from final deliverables rather than assuming one output from an input extension; if a template yields several CSV files, an exact-file request is rejected without publishing any of them. Explicit in-place Markdown transformations may use the input path itself; these update operations are separate from new conversion results.

GUI and Machine consumers receive the same producer-defined logical layout. Single-file bundles expose a flat logical path; grouped bundles expose the complete relative directory. Consumers preserve that declared layout in the user-selected parent. Consumers list business outputs, choose the explicitly preferred output for file location, and exclude layout manifests or extracted image resources from document counts.

## Artifact Bundle v3

`docwen.artifact_bundle.v3` carries each artifact's display basename in `suggested_name` and its stable relative location in `logical_path`. Result directories use `layout_schema=docwen.document_node.v1`; other bundles use `docwen.artifact_layout.v1`.

The runtime assigns final byte counts and SHA-256 in its in-memory artifact descriptions after link relocation. Bundle commit rechecks those identities and refuses modified bytes. The layout schema describes paths; it does not require an on-disk node manifest. Ordinary conversion emits only business documents, fragments and requested resources, never `docwen-node.json`.

The separately enabled audit export (`output.manifest.save_to_output`) remains supported. For Markdown-origin conversion, enabling it makes the audit an explicitly requested accompanying deliverable and therefore retains grouped publication even with one business file. Its redacted `manifest.json` is prepared in the same directory transaction, receives a collision-free name and integrity identity, and maps to a supplementary Bundle resource entry. An audit preparation failure publishes nothing. A temporary cleanup failure after successful reuse returns success with a visible warning, not a failed conversion that invites another write.

Readers accept only Bundle v3 and preserve validated relative paths. This contract contains no knowledge-base-specific storage concepts.


Producer-declared original embedded resources (including archive attachments) and native PPTX chart workbooks map to the existing `resource_of/original` Bundle role. OCR sidecars retain their distinct fragment role. Generic unknown auxiliaries still fail closed; filenames alone do not grant semantics. This mapping does not add a Machine capability or change the Bundle schema.

生产者明确声明的原始嵌入资源（含归档附件）及原生 PPTX 图表 workbook 使用既有 `resource_of/original` Bundle 角色；OCR 侧车保持独立 fragment 角色。未知 auxiliary 继续拒绝，不凭文件名授予语义；此映射不新增 Machine 能力或修改 Bundle schema。


## Producer ownership contract

`ArtifactManifest.ownership` is an optional, serialized Core contract available before publication.
It carries `relation_type` (`attachment_of`, `fragment_of`, or `resource_of`), `owner_artifact_id`,
role, zero-based ordinal where ordered, and an optional source-relative child label. Attachment/fragment owners must be documents; resources may belong to documents or fragments. Physical-page OCR continues to use its richer page-semantics contract; the ownership shorthand cannot declare an `ocr_page` without that data. The runtime
validates identities, owner references, roles, ordinals and acyclic ownership before publishing;
the application maps the same declaration to existing Bundle v3 relations. No Bundle schema
version changes. Legacy producers without this optional field retain their existing explicit
route-family mappings; unknown auxiliary outputs still do not acquire semantics from filenames.

Gongwen recognition owns which paragraphs belong to its combined attachment and declares the
label `附件`. The common layout and Bundle mapper do not inspect `gongwen_attachment` or parse
attachment names. Main rendering excludes attachment-owned paragraphs including their images;
structural main-document paragraphs still retain their seal/signature images. Attachment and
main images use the same image rendering mechanism. Resource ownership comes from those source
paragraphs, never from a filename guess or a regex scan of the rendered Markdown.

The final-output grouping decision lives in the runtime finalizer. Task orchestration and private
post-processing stages pass the admitted conversion identity and explicit user policy, and do not
force a final directory merely because an intermediate stage needed one. An explicitly requested
`group_outputs` still selects grouped publication.
