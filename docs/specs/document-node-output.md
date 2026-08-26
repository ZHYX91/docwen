# Markdown document-node output

Status: normative for new DocWen producers.

## Identity

A conversion owns one timezone-aware creation instant. Its root name is:

`{sanitized-source-stem}_YYYYMMDD_HHMMSS_from{CanonicalSourceFormat}`

Every artifact from that conversion uses the same identity. Collision suffixes are applied to the root as a whole,
never independently to individual files.

## Layout

Every published Markdown file is a document node whose directory and Markdown basename match:

```text
root_name/
  root_name.md
  docwen-node.json
  resource.png
  attachment_name/
    attachment_name.md
```

The primary Markdown is the root node. Auxiliary Markdown outputs, including Gongwen attachments, are child nodes.
Binary resources stay in the owning root and Markdown links are rewritten relative to their final logical paths.

## Publication and collisions

DocWen builds the complete root in a temporary sibling directory and commits it with one directory rename. Failed or
cancelled work cannot expose a partial document node.

The collision policy is evaluated once for the root:

- `error`: reject an existing root;
- `rename`: choose a new root suffix and rebase every logical path;
- `overwrite`: replace only a root carrying a valid DocWen node manifest, with backup/restore protection;
- `skip`: reuse only a node whose recorded source hash matches the current source.

## API and CLI contract

`OutputPolicy.output_dir` is the publication parent for Markdown. `output_path` remains an exact target for non-Markdown
artifacts. An explicit in-place Markdown transform may use `output_path` equal to its input path; it is an update, not a
new document-node publication.

For `docwen convert --to md`, `--output PATH` therefore names the publication parent directory. GUI conversions already
select an output directory and use the same runtime contract.

## Artifact Bundle v2

New producers emit `docwen.artifact_bundle.v2`. Each artifact carries both:

- `suggested_name`: a display/download basename fallback;
- `logical_path`: the stable relative path inside the committed bundle.

`layout_schema` is `docwen.document_node.v1` for node layouts and `docwen.artifact_layout.v1` for other bundles.
Readers accept only `docwen.artifact_bundle.v2`; missing layout or logical-path facts are protocol errors.

DocWen deliberately does not encode any knowledge-base-specific concepts such as Obsidian folder notes. Consumers can
map the neutral document-node and relation graph into their own storage model.
