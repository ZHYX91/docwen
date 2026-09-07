# Configuration / 配置

DocWen ships immutable base TOML files under `configs/` and writes user overrides to the platform user-config directory. Runtime consumers receive typed or read-only snapshots; plugins do not write configuration directly.

DocWen 在 `configs/` 提供不可变基础 TOML，并把用户覆盖写入平台用户配置目录。运行时消费者接收类型化或只读快照；插件不得直接写配置。

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

The packaged base contains the 25 files declared by the registry, including `configs/text.toml`, `configs/ocr.toml`, `configs/template_fill.toml`, `configs/numbering/add.toml`, `configs/numbering/cleanup.toml` and `configs/field_processors.toml`. The `gongwen` field processor resolves to `docwen_plugin_markdown.field_processors.gongwen.process_yaml`; its placeholder rules are consumed by `docwen_plugin_markdown.template_filler`.

随包基础配置完整包含 registry 声明的 25 个文件，包括 `configs/text.toml`、`configs/ocr.toml`、`configs/template_fill.toml`、`configs/numbering/add.toml`、`configs/numbering/cleanup.toml` 与 `configs/field_processors.toml`。`gongwen` 字段处理器解析到 `docwen_plugin_markdown.field_processors.gongwen.process_yaml`，其占位符规则由 `docwen_plugin_markdown.template_filler` 消费。

## Shared settings / 共用设置

OCR language is shared by every input format through `ocr.language` in `ocr.toml`, owned by the Markdown resources settings page. Word and spreadsheet template filling share `template_fill.list_separator` in `template_fill.toml`, owned by incoming text settings. Empty separators and intentional spaces are preserved. Document YAML/frontmatter is unchanged.

OCR 语言由所有传入格式共用，通过 `ocr.toml` 中的 `ocr.language` 设置，归属“Markdown 资源”。Word 与表格模板填充共用 `template_fill.toml` 中的 `template_fill.list_separator`，归属“传入文本”；空字符串和有意输入的空格原样保留。文档 YAML/frontmatter 不变。

The removed keys `image.ocr_language` and `conversion.md_to_docx.list_separator` are rejected by validation; there are no aliases or automatic migrations. Other keys retain functional namespaces even when their settings page changes. Resetting an input page preserves converter priorities; the conversion software page owns all eight priority lists. Numbering and proofreading editor saves remain immediate transactions, independent of unrelated settings drafts.

旧键 `image.ocr_language` 和 `conversion.md_to_docx.list_separator` 不再通过验证，不提供别名或自动迁移。其他键即使调整设置页归属，也保留功能命名空间。重置传入页保留软件优先级；八组优先级统一归属“转换软件”。编号和校对编辑器保存仍是立即生效的事务，不丢弃其他设置草稿。
