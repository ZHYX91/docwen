# Conversion result-directory output

Status: normative for new DocWen producers.

## Scope and identity

Conversions from Markdown to documents or spreadsheets and all conversions to Markdown publish one result directory per input, even for one business output. A conversion freezes one timezone-aware creation instant and the admitted source identity before intermediate formats are produced.

Root and ordinary primary filename stem:

`{sanitized-source-stem}_YYYYMMDD_HHMMSS_from{CanonicalSourceFormat}`

Examples use `fromMd`, `fromDocx`, `fromWps` or `fromRtf` according to this conversion's admitted input, never a private JSON or DOCX intermediate. Every related output shares that timestamp. A source filename is not parsed heuristically for provenance. A later conversion uses the actual admitted input name; no adjacent JSON, hidden sidecar or persistent index is consulted for provenance.

## Layout

Every Markdown file has the same basename as its containing document node. A generated DOCX or XLSX also uses its result root's basename:

```text
项目记录_20260907_180000_fromMd/
  项目记录_20260907_180000_fromMd.docx
```

Spreadsheet CSV outputs put the worksheet name before the shared timestamp:

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
  通知_附件_20260907_180000_fromDocx/
    通知_附件_20260907_180000_fromDocx.md
```

Attachment titles remain in the content. Other auxiliary Markdown outputs are child nodes. Images and other linked resources retain safe names within the root, and Markdown links are rewritten relative to their final logical paths.

## Publication and collisions

The complete root is prepared in a temporary sibling directory and committed in one atomic no-replace directory rename. An existing directory is never deleted or replaced, including one created by a competing writer after the collision check. Failure or cancellation before publication exposes no partial result directory. The root collision policy is evaluated once:

- `error`: reject an existing root;
- `rename`: choose a new root suffix and rebase logical paths, preserving matching root/primary basenames;
- `overwrite`: invalid for result-directory conversions, rejected before publication (and by CLI preflight); it remains available for explicitly targeted individual files;
- `skip`: reuse only a complete regular directory whose paths and byte hashes match the newly prepared result. Missing or extra paths, changed bytes, links and junctions fail without changing the existing tree. A matching result reports reuse and zero newly written bytes.

Independent filenames are never renamed separately to resolve a publication conflict. CSV worksheet names retain their source/worksheet/timestamp identity within the chosen root.

## API and CLI contract

`OutputPolicy.output_dir` and CLI `--output-dir DIR` select the publication parent. All conversions from or to Markdown require this directory form; CLI `--output PATH` remains an exact file target for other conversions. Explicit in-place Markdown transformations may use the input path itself; these update operations are separate from new conversion results.

GUI and Assistant use the same producer-defined directory layout. Assistant preserves the complete logical directory in the user-selected parent and rejects existing result roots. Consumers list business outputs, choose the explicitly preferred output for file location, and exclude layout manifests or extracted image resources from document counts.

## Artifact Bundle v3

`docwen.artifact_bundle.v3` carries each artifact's display basename in `suggested_name` and its stable relative location in `logical_path`. Result directories use `layout_schema=docwen.document_node.v1`; other bundles use `docwen.artifact_layout.v1`.

The runtime assigns final byte counts and SHA-256 in its in-memory artifact descriptions after link relocation. Bundle commit rechecks those identities and refuses modified bytes. The layout schema describes paths; it does not require an on-disk node manifest. Ordinary conversion emits only business documents, fragments and requested resources, never `docwen-node.json`.

The separately enabled audit export (`output.manifest.save_to_output`) remains supported. Its redacted `manifest.json` is prepared in the same directory transaction, receives a collision-free name and integrity identity, and maps to a supplementary Bundle resource entry. An audit preparation failure publishes nothing. A temporary cleanup failure after successful reuse returns success with a visible warning, not a failed conversion that invites another write.

Readers accept only Bundle v3 and preserve validated relative paths. This contract contains no knowledge-base-specific storage concepts.
