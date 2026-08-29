# Templates and styles / 模板与样式

This specification is normative for every Markdown-to-DOCX conversion. Template discovery and selection remain
owned by `TemplateRegistry`; style completion happens only in the request-owned output document and never mutates a
shipped or user template in place.

本文是每次 Markdown→DOCX 转换的正式样式合同。模板发现和选择仍由 `TemplateRegistry` 负责；完整样式
只注入请求拥有的输出文档，不原地修改随包模板或用户模板。

## Managed identity / 受管身份

Every managed style has one identity chain:

```text
semantic key -> stable OOXML styleId -> visible w:name
```

- A built-in Word style uses its canonical English `styleId` and canonical English `w:name`; Word/WPS localize its
  presentation. If a Word/WPS template already represents the same built-in identity and type under a host-specific
  numeric or localized internal `styleId`, DocWen reuses that actual ID through the resolved map; it does not delete
  and recreate the style merely to force the canonical ID. A newly injected built-in uses the canonical ID/name.
- A DocWen style uses the stable `DocWen...` styleId below and a locale-specific `w:name` from the exact output
  locale. All eleven shipped locales are required.
- `w:aliases` is never emitted. Import-time aliases are an internal recognition aid, not output identity.
- An existing style with the same semantic identity and OOXML type is preserved, including user formatting. If a
  stable DocWen `styleId` or visible name is occupied by a different identity or OOXML type, DocWen preserves the
  user's style and allocates a request-local collision-free `styleId` for the managed identity. The resolved
  semantic-key-to-styleId map is the only map used by renderers and is recorded with a stable conflict diagnostic.
  DocWen never deletes, renames, retypes, or rewrites the conflicting user style.

Word's `Normal`, `DefaultParagraphFont`, `TableNormal`, and `Title` are template base styles and are not part of the
43 managed entries. Starting from a document that contains only those four base styles, a complete output therefore
contains exactly 47 physical styles. A custom template may contain more.

## Complete registry / 完整注册表

### Built-in styles (16)

| Semantic key | styleId | OOXML type | Canonical w:name |
|---|---|---|---|
| `heading_1` ... `heading_9` | `Heading1` ... `Heading9` | paragraph | `heading 1` ... `heading 9` |
| `footnote_text` | `FootnoteText` | paragraph | `footnote text` |
| `footnote_reference` | `FootnoteReference` | character | `footnote reference` |
| `endnote_text` | `EndnoteText` | paragraph | `endnote text` |
| `endnote_reference` | `EndnoteReference` | character | `endnote reference` |
| `caption` | `Caption` | paragraph | `caption` |
| `bibliography` | `Bibliography` | paragraph | `Bibliography` |
| `hyperlink` | `Hyperlink` | character | `Hyperlink` |

### DocWen styles (27)

| Semantic key | Stable styleId | Type | English w:name | 简体中文 w:name | Based on |
|---|---|---|---|---|---|
| `body_paragraph` | `DocWenBodyParagraph` | paragraph | Body Paragraph | 正文段落 | `Normal` |
| `image_paragraph` | `DocWenImageParagraph` | paragraph | Image Paragraph | 图片段落 | `Normal` |
| `code_block` | `DocWenCodeBlock` | paragraph | Code Block | 代码块 | `Normal` |
| `inline_code` | `DocWenInlineCode` | character | Inline Code | 行内代码 | `DefaultParagraphFont` |
| `formula_block` | `DocWenFormulaBlock` | paragraph | Formula Block | 公式块 | `Normal` |
| `inline_formula` | `DocWenInlineFormula` | character | Inline Formula | 行内公式 | `DefaultParagraphFont` |
| `list_block` | `DocWenListBlock` | paragraph | List Block | 列表块 | `Normal` |
| `horizontal_rule_1` ... `horizontal_rule_3` | `DocWenHorizontalRule1` ... `DocWenHorizontalRule3` | paragraph | Horizontal Rule 1 ... 3 | 分隔线 1 ... 3 | `Normal` |
| `table_content` | `DocWenTableContent` | paragraph | Table Content | 表格内容 | `Normal` |
| `table_header` | `DocWenTableHeader` | paragraph | Table Header | 表格表头 | `Normal` |
| `three_line_table` | `DocWenThreeLineTable` | table | Three Line Table | 三线表 | `TableNormal` |
| `table_grid` | `DocWenTableGrid` | table | Table Grid | 网格表 | `TableNormal` |
| `quote_1` ... `quote_9` | `DocWenQuote1` ... `DocWenQuote9` | paragraph | Quote 1 ... 9 | 引用 1 ... 9 | `Normal` |
| `figure_caption` | `DocWenFigureCaption` | paragraph | Figure Caption | 图题 | `Caption` |
| `table_caption` | `DocWenTableCaption` | paragraph | Table Caption | 表题 | `Caption` |
| `equation_caption` | `DocWenEquationCaption` | paragraph | Equation Caption | 公式题注 | `Caption` |
| `code_block_caption` | `DocWenCodeBlockCaption` | paragraph | Code Block Caption | 代码块题注 | `Caption` |

The other nine locale names are stored in their corresponding `i18n/locales/*.toml` `[styles]` registry. Missing
or blank names are contract errors; falling back to another language would change visible output identity.

`FollowedHyperlink`, `In-text Citation`, `Cross Reference`, `Bibliography Heading`, and caption-label character
styles are deliberately outside this registry. Visit state, fields/anchors, citation presentation, and bibliography
headings have separate owners.

## Completion and rendering / 注入与渲染

Before any Markdown renderer writes content, it completes all 43 styles for the exact request locale. It does not
limit completion to styles observed in the source. Every resulting `w:pStyle`, `w:rStyle`, and `w:tblStyle` reference
must resolve to a style of the correct type.

- Figure/Table/Equation/Code authored declarations use their four dedicated paragraph styles. Renaming
  `listing_caption` to `code_block_caption` changes one identity in place: the registry remains exactly 43 entries,
  and no `List Caption` style or semantic key is reserved.
- `Equation:` and `Code:` mean explicit authored declarations; each may have caption text, an `^id`, or both. Their
  empty-text forms require an ID. Figure and Table declarations require non-empty visible caption text, but
  their IDs are optional. Every valid declaration receives its caption style. It receives a derived field number
  only when the provider-neutral resolved numbering/export plan enables that kind and supplies the exact number;
  disabled captions retain authored content/style and contain no `SEQ`. Only a declaration with an ID is
  addressable. Caption semantic kind and native carrier structure are independent. DOCX order follows the semantic
  kind: a Figure-labelled carrier directly precedes `DocWenFigureCaption`, while Table-, Equation-, and Code-labelled
  captions directly precede their carrier, including cross-type pairs such as Figure plus a native table. The order
  is not profile-configurable; canonical Markdown recovery writes every declaration above its carrier.
- A declaration ID and its structure-owned numbered target semantics follow
  [Markdown compatibility](markdown-compatibility.md). When enabled, its target bookmark belongs to the caption
  paragraph and encloses either the simple `SEQ` field or the complete chapter `STYLEREF`+separator+`SEQ` composite,
  with exact cached result. When disabled, an ID-bearing caption instead has a zero-width navigation bookmark at
  paragraph start and no number/`SEQ`/`REF`. The required `docwen-target-v1:` outer SDT binds exactly that styled caption
  paragraph and one authenticated native carrier; it is internal pairing, not an authored marker or second target. A
  generic object/block anchor is styled as its underlying object and retained by an opaque block-level SDT plus
  custom XML mapping; it receives no bookmark, hyperlink target, caption style, `SEQ`, or `REF` semantics.
- An enabled ID-less caption is recovered only from the fixed direct physical adjacency, its exact managed style,
  exactly one closed simple-`SEQ` or chapter-`STYLEREF`+separator+`SEQ` materialization, and exact cached result. A
  disabled ID-less caption instead requires the independent `document-numbering-occurrence-map/v1` plus exact
  physical two-block occurrence SDT; it has no field, number, target, bookmark, or hidden ID. Style/adjacency alone,
  reversed/opposite-side placement, an intervening block, missing/unproved carrier, mismatched field/occurrence map, or
  multiple pairing claims fail closed.
- The exact caption style means the request-resolved style ID, not a requested-ID prefix. If any caption is rendered,
  a separate `document-caption-style-binding-map/v1` custom XML item persists exactly four ordered bindings
  (`figure_caption`, `table_caption`, `equation_caption`, `code_block_caption`) with each semantic key, resolved
  style ID, and visible name. This is style identity only, never pairing/target/ID metadata. On save/reopen, every
  bound style must be one exact paragraph style with one exact direct name and no aliases, and every caption must
  carry the matching exact direct `pStyle`. A non-destructive collision may therefore use, for example,
  `DocWenTableCaptionDocWen1`; the preserved conflicting `DocWenTableCaption` never regains authority through prefix
  matching or its visible name.
- A soft semantic reference that uniquely selects an ID-less Heading uses static cached number text in the required
  non-target soft-reference SDT/map. It never auto-writes an ID or creates a bookmark/REF. The Alias retains the
  surrounding paragraph's character formatting; no `Cross Reference` named style is injected. If that metadata
  cannot be emitted and recovered exactly, conversion fails rather than flattening the authored token.
- Ordinary Markdown links use the built-in `Hyperlink` character style while retaining explicit rich-run
  formatting. Each hyperlink run has exactly one `rStyle=Hyperlink`; code, highlight, superscript, and subscript
  labels use direct run properties so they do not create competing character-style references. Code properties
  are projected from the request-owned resolved `DocWenInlineCode` style; configured defaults are used only when
  no compatible style exists. Until clickable
  OMML is a separately frozen interoperability contract, inline math inside a link is preserved visibly as
  `$...$` text rather than silently flattened or emitted as an ambiguous hyperlink/math structure.
- Table header paragraphs use `DocWenTableHeader`; table body paragraphs use `DocWenTableContent`.
- Existing compatible custom formatting wins over DocWen defaults. Defaults are applied only when a style is
  newly created.

## Existing note parts / 既有注释部件

Markdown note syntax is frozen as follows: `[^id]` and `[^footnote:id]` are footnotes, while
`[^endnote:id]` is an endnote. The retired `[^endnote-id]` spelling is rejected; DOCX → Markdown always emits the
canonical colon form. Footnotes and endnotes are numbered independently from 1 in first-reference
order, and repeated references reuse the first number. Missing definitions, duplicate definitions, and collisions
between default and explicit footnote IDs fail closed. Continuation content indented by two spaces or a tab remains
inside the definition and is not reinterpreted as a heading, caption, or semantic reference. Normalization is scoped
to the conversion request and never rewrites the source Markdown or its authored IDs.

The four built-in note styles (`FootnoteText`, `FootnoteReference`, `EndnoteText`, and `EndnoteReference`) are always
completed, even when the source contains no new note. Their presence is not evidence that note content survived.

Before adding a footnote or endnote, DocWen reads the complete existing note part, its relationships, and its content
type. It validates every existing note ID, including Word's reserved separator IDs, rejects illegal or duplicate IDs,
and allocates the lowest unused positive ID independently in the footnote and endnote domains. The body entry and the
main-document reference run are written atomically with the same ID. A failure cannot leave only one side behind.

Existing rich note bodies are preserved as OOXML, including blocks, runs, hyperlinks, drawings, and every reachable
note-part relationship. DocWen neither flattens the body to text nor deletes, renumbers, or recreates existing notes.
Relationship IDs are allocated against the complete existing part, not only relationships created during the current
request. Word, WPS, and LibreOffice must each open, save, reopen, and retain the note body, reference, relationship,
and ID graph; a headless check that only sees the four styles is not a host-fidelity gate.

## Bibliography placeholder / 参考文献占位符

`{{ bibliography }}` must be the sole visible content of one direct main-document paragraph. A bibliography input
is already presented data, not raw CSL: DocWen does not select a citation style or run a CSL processor.

The closed bibliography payload uses media type
`application/vnd.docwen.semantic-bibliography+json` and contains ordered entries. Each entry has a portable
`item_id` and one or more typed runs; each run contains non-empty `text` and optional `bold`, `italic`, and absolute
`http`/`https` `href`. On the active v4 `convert.markdown.to_docx` route it is not an independent Machine input: at
most one payload is embedded as a `role=bibliography` resource inside the exact-one `neutral_document`, alongside
the exact-one `numbering_export_plan`. An optional independent `role=bibliography` input belongs only to separately
identified legacy/other capabilities that still declare that slot. `citation_style` remains undeclared and is
rejected until it participates in an implemented citation-processing contract.

The resource is strict UTF-8 JSON, at most 8 MiB, rejects duplicate keys and non-finite numbers, and has this exact
closed shape (all unshown properties are forbidden):

```json
{
  "schema": "docwen.semantic_bibliography.v1",
  "entries": [
    {
      "item_id": "smith2025",
      "runs": [
        {"text": "Smith, A. "},
        {"text": "Neutral documents", "italic": true, "href": "https://example.org/neutral-documents"},
        {"text": "."}
      ]
    }
  ]
}
```

An explicit `{{ bibliography }}` paragraph is authoritative. Because the eleven shipped 0.9 templates predate this
typed resource, a non-empty resource may synthesize that paragraph only immediately after one unique direct-body
body marker from the frozen shipped-locale alias set (`body`, `Inhalt`, `cuerpo`, `Corps`, `本文`, `본문`, `Corpo`,
`Текст`, `Nội dung`, `正文`) in the request-owned document. If neither marker is unique, conversion fails; no
template filename, ID, or hash may be used to guess placement. An absent or empty resource with no bibliography
marker is a no-op. An explicit marker with an absent or empty resource is removed by the bibliography owner.

Replacement is atomic and follows these rules:

1. Validate the complete resource, all item IDs/runs/links, and the unique direct-body placeholder before changing
   the DOCX tree.
2. Empty input removes only the placeholder. A `sectPr` on that paragraph must be safely transferred to the
   immediately preceding direct-body paragraph; otherwise conversion fails rather than losing section semantics.
3. Non-empty input creates one paragraph per entry, deep-copies the placeholder's complete `pPr`, and renders the
   typed bold/italic/hyperlink runs. A copied `sectPr` appears only on the final entry.
4. A custom placeholder style is preserved. Only an unstyled generic fallback may use built-in `Bibliography`.
5. The preceding “References”/“参考文献” heading is outside the fragment and is never restyled or replaced.

## Resolved structured numbering boundary / 已解析结构化编号边界

The normative boundary is [Resolved structured numbering and export plan](structured-numbering-phases.md). The
upstream semantic provider owns the selected profile and the five independent
enable/counter/format/label/start/reset/scope/chapter/separator rules. DocWen accepts the already-resolved document and plan, materializes enabled Heading `numbering.xml`/list
semantics, enabled caption simple `SEQ` or chapter `STYLEREF`+separator+`SEQ`, navigation bookmarks, numbered `REF`
cached results, and this managed style registry, and never becomes the rule owner.
Font, position, and spacing remain Export Style concerns.

No title-text or language heuristic may infer a number or phase. `## 2.3 标题` keeps the indivisible authored title
`2.3 标题`; if Heading numbering is enabled, the derived prefix and that title coexist. Disabled numbering and an
empty Heading-level template both mean no materializable number and semantic references fail with
`docwen.markdown.cross_reference.unnumbered_target`; ordinary links still navigate. Changing enablement/profile is a
zero-Markdown-diff operation. Source-mutating add/remove/regex-cleanup controls are not part of this port.

## Acceptance / 验收

- blank four-style mother template, all eleven shipped templates, and a user template missing managed styles;
- exactly 43 managed identities, exactly 47 total styles for the blank mother template, correct types and zero
  aliases;
- localized custom names and canonical built-in names/styleIds;
- deterministic conflict mapping and diagnostics without deleting or rewriting user formatting;
- each caption, table header/body, inline formula/code, link, and bibliography run references a real style;
- enabled semantic caption bookmark/field/`REF` plus its required two-object outer SDT, disabled ID-bearing
  zero-width navigation bookmark, disabled ID-less occurrence map/SDT, and ordinary-anchor SDT/target-map
  projections, survive save/reopen; the separate closed four-entry caption-style binding map authenticates exact
  resolved styles for both addressable and ID-less captions, including request-local conflict IDs; Markdown emits
  no extra pairing token, ID-less captions create no target bookmark; enabled ID-less captions use proved fields and
  disabled ID-less captions use only their independent pairing authority. All retain fixed physical order; ordinary
  anchors create no bookmark or
  field, 128-character IDs round-trip, and deterministic
  bookmark/tag collisions fail;
- five-kind enable on/off plus Heading-level-template-empty: enabled Headings use exact `numbering.xml`/list
  semantics and enabled captions use matching simple/chapter field forms; disabled cases retain styles/navigation
  but create no number or
  `REF`. The authored Heading `2.3 标题` is never split or cleaned, and toggling the plan yields byte-identical
  Markdown;
- DOCX-to-neutral separates numbering only from proved list/`SEQ`/`REF`/bookmark semantics. An ambiguous visible
  prefix remains authored text plus diagnostic, and canonical Markdown never receives a materialized number;
- bibliography empty/non-empty, custom pPr, `sectPr`, bold/italic/link, save/reopen and heading independence;
- pre-existing footnote/endnote ID collisions, rich bodies, hyperlinks/drawings/relationships, atomic reference/body
  allocation, and unchanged reserved separator entries;
- source and packaged conversion, headless OOXML inspection, round-trip, then separately bound Word, WPS, and
  LibreOffice open/save/reopen.

Template-generation scripts are maintenance tools, not runtime evidence. Any registry, localization, placeholder,
anchor, bookmark, or rendering change requires focused OOXML tests, packaged-resource verification, and the Word,
WPS, and LibreOffice three-viewer open/save/reopen gate.
