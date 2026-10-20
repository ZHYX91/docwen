# Configuration / 配置

DocWen ships immutable base TOML files under `configs/` and writes user overrides to the selected profile's configuration directory. Runtime consumers receive typed or read-only snapshots; plugins do not write configuration directly.

DocWen 在 `configs/` 提供不可变基础 TOML，并把用户覆盖写入所选资料的配置目录。运行时消费者接收类型化或只读快照；插件不得直接写配置。

`DOCWEN_DATA_DIR` selects a whole profile containing `configs/`, `templates/`, template state and default logs. `DOCWEN_CONFIG_DIR` overrides only the exact configuration directory. Archive, MSIX and source defaults, logging overrides, and explicit copy/import behavior are defined in [User profile paths](specs/user-profile.md). All entry points keep the startup selection until exit.

`DOCWEN_DATA_DIR` 选择包含 `configs/`、`templates/`、模板状态和默认日志的一整份资料；`DOCWEN_CONFIG_DIR` 只覆盖精确的配置目录。便携包、MSIX、源码默认路径、日志覆盖和明确复制/导入方式见[用户资料路径](specs/user-profile.md)。所有入口均在启动时确定路径，运行中不切换资料。

## Ownership / 所有权

- `docwen_runtime.config.registry` declares config files, reset groups and protected files.
- `ConfigLoader` merges base and user values and owns public read/write/reset operations.
- TOML writes use a same-directory staged transaction, process coordination and recovery journal.
- GUI and CLI call the injected configuration port; they do not maintain a second persisted state.

- `docwen_runtime.config.registry` 声明配置文件、重置组和保护文件。
- `ConfigLoader` 合并基础值与用户值，并拥有公开读、写、重置操作。
- TOML 写入采用同目录暂存事务、进程协调和恢复日志。
- GUI/CLI 通过注入的配置端口操作，不维护第二份持久化状态。

## Rules / 规则

1. Add new keys to the owning base TOML and typed model together.
2. Preserve extension-owned unknown user keys verbatim; the loader never renames, migrates or backfills user keys.
3. Reset deletes or rewrites only registered user-owned values and retains protected files.
4. Request-affecting values are frozen before external or concurrent work starts.
5. Secrets and sensitive paths must not be echoed in diagnostics.

配置文件与消费方的完整映射可通过 `docwen inspect --json` 和 registry 源码核对。

The packaged base contains the files declared by the registry, including `configs/text.toml`, `configs/ocr.toml`, `configs/template_fill.toml`, `configs/numbering/add.toml` and `configs/numbering/cleanup.toml`. Chinese official-document field formatting is built into `docwen_plugin_markdown.gongwen_fields` and called directly by `docwen_plugin_markdown.template_filler`. Configuration does not load Python field processors.

随包基础配置包含 registry 声明的全部文件，包括 `configs/text.toml`、`configs/ocr.toml`、`configs/template_fill.toml`、`configs/numbering/add.toml` 与 `configs/numbering/cleanup.toml`。公文字段格式化内置于 `docwen_plugin_markdown.gongwen_fields`，由 `docwen_plugin_markdown.template_filler` 直接调用。配置不再加载 Python 字段处理器。

## Shared settings / 共用设置

General settings keep interface language (`gui.language.locale`) separate from task output language (`conversion.output_locale`). The latter controls generated labels and document styles; explicit task locale takes precedence. Existing profiles freeze their previously saved locale once, while new profiles default to `zh_CN`. See [Display and task languages](specs/task-language.md).

常规设置分别管理界面语言 `gui.language.locale` 与任务输出语言 `conversion.output_locale`。后者控制生成标签和文档样式，显式任务参数优先。旧配置首次升级时冻结原先保存的语言，新配置默认 `zh_CN`；详见[语言职责](specs/task-language.md)。

OCR language is shared by every input format through `ocr.language` in `ocr.toml`, owned by the Markdown resources settings page. Word and spreadsheet template filling share `template_fill.list_separator` in `template_fill.toml`, owned by incoming text settings. Empty separators and intentional spaces are preserved. Document YAML/frontmatter is unchanged.

OCR 语言由所有传入格式共用，通过 `ocr.toml` 中的 `ocr.language` 设置，归属“Markdown 资源”。Word 与表格模板填充共用 `template_fill.toml` 中的 `template_fill.list_separator`，归属“传入文本”；空字符串和有意输入的空格原样保留。文档 YAML/frontmatter 不变。

The removed keys `image.ocr_language` and `conversion.md_to_docx.list_separator` are rejected by validation; there are no aliases or automatic migrations. Other keys retain functional namespaces even when their settings page changes. Resetting an input page preserves converter priorities; the conversion software page owns all eight priority lists. Numbering and proofreading editor saves remain immediate transactions, independent of unrelated settings drafts.

旧键 `image.ocr_language` 和 `conversion.md_to_docx.list_separator` 不再通过验证，不提供别名或自动迁移。其他键即使调整设置页归属，也保留功能命名空间。重置传入页保留软件优先级；八组优先级统一归属“转换软件”。编号和校对编辑器保存仍是立即生效的事务，不丢弃其他设置草稿。

## Chinese official-document fields / 公文字段

`template_fill.gongwen_fields_enabled` is a boolean, enabled by default. It formats the fixed Chinese official-document fields in Word templates independently of interface or output language. Both Markdown input routes use the same implementation. It does not select document-to-Markdown recognition. The incoming Markdown settings page owns this switch and its reset behavior.

`template_fill.gongwen_fields_enabled` 是布尔值，默认开启。它处理 Word 模板中的固定中文公文字段，不随界面或输出语言改变；两条 Markdown 输入路线使用同一套实现。它不控制文档转 Markdown 的公文识别。该开关及其恢复默认行为归属“传入文本”设置页。

| Field / 字段 | Enabled behavior / 开启后的行为 |
| --- | --- |
| `附件说明` | One paragraph per attachment, with a prefix, consistent numbering and hanging indentation. 每项附件独立成段，使用统一前缀、编号和悬挂缩进。 |
| `抄送机关` | Join recipients with `，`. 使用中文逗号连接接收单位。 |
| `成文日期` / `印发日期` | Format recognized dates as Chinese dates; publication dates end with `印发`. Unrecognized text is retained. 可识别日期使用中文年月日，印发日期增加“印发”，无法识别的原文保持不变。 |
| `附注` | Remove one existing outer pair of parentheses; the template owns the displayed parentheses. 去掉内容已有的一层外括号，由模板控制显示的括号。 |
| Empty official fields / 空公文字段 | Remove only eligible dedicated paragraphs or rows, preserving other nonempty fields and unrelated content. 仅清理符合条件的专用段落或表格行，保护其他有值字段和无关内容。 |

`{{附件说明}}` must occupy an entire paragraph in the main document body or a table cell, including nested tables. Each original occurrence expands separately. Inline occurrences, headers, footers and text boxes produce a template error when field formatting is enabled. Empty attachments keep table cells structurally valid. Inserted YAML values are literal content and are never interpreted again as template placeholders.

`{{附件说明}}` 须独占正文或表格单元格中的一个段落，支持嵌套表格和多处独立占位段。开启字段格式化时，行内、页眉、页脚和文本框中的附件占位符会触发模板错误。空附件不会破坏单元格结构；插入的 YAML 值始终作为内容，不会被再次解释为模板占位符。

When disabled, these fields use ordinary template filling, including `template_fill.list_separator`; official-document container cleanup is disabled. The original YAML data is not rewritten. Removed field-processor configuration has no aliases, migration or dynamic module fallback.

关闭后，这些字段按普通模板规则填充，列表遵守 `template_fill.list_separator`，不执行公文专用容器清理。原始 YAML 数据不被改写。已删除的字段处理器配置不提供别名、迁移或动态模块回退。
