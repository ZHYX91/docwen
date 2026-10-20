# Settings information architecture / 设置归属复查

Reviewed against `02f46e742a582bacb3efc324c3f09f5310a79ded` (PR #115).
This change reorganizes presentation, not conversion contracts or configuration.

## Navigation / 导航

| Group / 分组 | Pages / 页面 |
| --- | --- |
| Basics / 基本设置 | 通用、文件保存、模板、转换软件、日志 |
| Input types / 输入类型 | 传入文本、传入文档、传入表格、传入图片、传入版式文件、传入其他文件 |
| Content / 内容处理 | Markdown 语法、链接、OCR 与图片、校对 |
| Specialized / 专项优化 | 公文、发票 |

There are still 17 pages. Stable page keys, cross-page links, configuration keys,
Apply/Cancel and per-page reset ownership are unchanged. Group headings are not
pages or focus stops. Empty groups remain hidden.

“OCR 与图片” covers recognition and Markdown image/OCR presentation. Input pages
still own whether OCR is enabled by default for their category. Image format
conversion and compression remain on Incoming images. The description explains
this distinction rather than introducing a second copy of any setting.

Sidebar width still follows native text/font metrics. Shortening the label is the
first fix; a hard width cap plus a custom eliding button is not added without a
separate need. Existing wide/short, narrow, large-font and RTL navigation tests
continue to apply. This avoids solving one long title by clipping other locales.

## Markdown policy is not a Word preference / Markdown 规则不属于 Word 专页

The Markdown page presents, in order:

1. Generate Markdown: optional output extensions and preferred output syntax.
2. Read Markdown: optional input recognition, independent of output choices.
3. Word-specific mappings and generation: page/section separators, paragraphs
   and Mermaid rendering, explicitly scoped to this conversion family.

Output conventions are source-independent. Documents, spreadsheets, images,
fixed-layout files and specialized optimizers should use the same convention for
an equivalent semantic structure. There are three different cases:

- A source does not contain the structure: no output is invented to satisfy a switch.
- Extraction has not recovered the structure: enabling a dialect cannot create
  missing recognition or reconstruction capabilities.
- A renderer has the structure but ignores the selected output rule: that is an
  implementation gap, not a reason to duplicate or permanently narrow the setting.

Valid Markdown, the selected dialect, and the preferred spelling of syntax are
separate guarantees. A valid table does not prove compliance with Structural
Tables settings; sharing image-link policy does not prove use of inline syntax
choices or all four extension switches. Output choices must not narrow input
recognition of otherwise supported syntax.

## Current coverage / 当前覆盖（代码复查，不代表全部路线实测）

| Path | Evidence at review baseline | Limitation |
| --- | --- | --- |
| Standard Word → Markdown | `docwen_plugin_document/to_markdown/converter.py` resolves output extensions; `request_policy.py` projects output syntax from the request snapshot | Full fidelity still depends on extracted semantics; this does not include the separate Gongwen optimizer |
| Spreadsheet → Markdown | `docwen_plugin_spreadsheet/to_markdown/converter.py` reads the same output-extension policy and uses `structural_tables` to select table rendering | Reading the policy object does not imply implementation of every extension or inline style |
| Image → Markdown | `docwen_plugin_image/to_markdown/converter.py` uses `resolve_markdown_request_policy` for image links, placement and OCR presentation | No evidence here of complete inline-style / extension-policy coverage; OCR recognition remains a separate capability |
| PDF/OFD/XPS → Markdown | `docwen_plugin_layout/to_markdown/converter.py` uses shared export semantics alongside extraction / third-party rendering | Do not advertise complete dialect normalization; physical pages are not automatically Word section breaks |
| Gongwen → Markdown | `docwen_plugin_optimizer_gongwen/rendering/markdown_renderer.py` has its own structure rendering and shared export inputs | Must be audited separately from standard Word conversion; not assumed to inherit every syntax choice |
| Invoice → Markdown | `docwen_plugin_optimizer_invoice_cn/invoice_cn/markdown_writer.py` builds YAML and a plain table | Plain output is not proof that all global syntax/extension switches are consumed |

Presentation and markup helpers were also inspected: `docwen_plugin_presentation/pptx_md/request_policy.py` projects shared image/OCR export settings; `docwen_plugin_markup/markdown_resources.py` applies shared policy to resources. Neither observation establishes complete body syntax or extension conformance for PPTX, HTML, EPUB or ENEX. Support must be established per route and per rule, not inferred from importing a shared helper or returning a `.md` file.

The interface keeps the common settings in one place and explicitly describes
current coverage. It does not add per-source copies, locale-dependent hiding,
a generic plugin-settings engine, configuration migration, or compatibility keys.

## Verification boundary / 验收边界

GUI tests cover the four groups, all existing page identities, localizations,
source-independent Markdown card order, independent input/output switches and
unchanged per-category OCR defaults. Existing reset, lifecycle, accessibility,
Mermaid, i18n and Markdown extension integration tests remain required.

This is not an output-engine rewrite. Existing coverage gaps are recorded above,
not converted into permanent contracts or tests that require missing behavior.
When a route gains support, verify real generated Markdown for applicable shared
structures, explicit dialect-off fallback and loss diagnostics. Do not add a
post-processing regex pass that rewrites arbitrary Markdown or raw OCR text.
