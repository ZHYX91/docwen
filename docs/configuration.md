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

`template_fill.gongwen_fields_enabled` is a boolean, enabled by default. It formats the fixed Chinese official-document fields in Word templates independently of interface or output language. Both Markdown input routes use the same implementation. It does not select document-to-Markdown recognition. The Official documents settings page owns this switch and its reset behavior; the generic list separator remains on Incoming text.

`template_fill.gongwen_fields_enabled` 是布尔值，默认开启。它处理 Word 模板中的固定中文公文字段，不随界面或输出语言改变；两条 Markdown 输入路线使用同一套实现。它不控制文档转 Markdown 的公文识别。该开关及其恢复默认行为归属“公文”设置页；通用列表分隔符仍归属“传入文本”。

| Field / 字段 | Enabled behavior / 开启后的行为 |
| --- | --- |
| `附件说明` | One paragraph per attachment, with a prefix, consistent numbering and hanging indentation. 每项附件独立成段，使用统一前缀、编号和悬挂缩进。 |
| `签发人` / `主送机关` / `抄送机关` | Preserve authored strings. Flat lists use `、`; multiple nonempty items get a Word comment asking the user to confirm signer layout or recipient punctuation and order. 已写好的字符串原样保留；列表使用顿号，多项列表在填充值上添加批注，请用户确认签发人排版或机关之间的标点与顺序。 |
| `成文日期` / `印发日期` | Format recognized dates as Chinese dates; publication dates end with `印发`. Unrecognized text is retained. 可识别日期使用中文年月日，印发日期增加“印发”，无法识别的原文保持不变。 |
| `附注` | Remove one existing outer pair of parentheses; the template owns the displayed parentheses. 去掉内容已有的一层外括号，由模板控制显示的括号。 |
| Empty official fields / 空公文字段 | Remove only eligible dedicated paragraphs or rows, preserving other nonempty fields and unrelated content. 仅清理符合条件的专用段落或表格行，保护其他有值字段和无关内容。 |

`{{附件说明}}` must occupy an entire paragraph in the main document body or a table cell, including nested tables. Each original occurrence expands separately. Inline occurrences, headers, footers and text boxes produce a template error when field formatting is enabled. Empty attachments keep table cells structurally valid. Inserted YAML values are literal content and are never interpreted again as template placeholders.

Required signer/recipient comments are supported in the main body and table cells. A multi-item field in an
unsupported Word part fails explicitly. Single-item lists and authored strings receive no automatic review comment;
turning field formatting off restores generic list filling. The converter does not move document numbers or arrange
signer names into columns.

Document-to-Markdown Gongwen extraction retains copy-recipient expressions as strings, including their punctuation
and parentheses. An explicit document-number/signature label admits unusual or Latin signer names without guessing
from surnames. Primary-recipient continuation consumes only clear collective address text; uncertain preceding
paragraphs remain in the body and receive a review warning.

`{{附件说明}}` 须独占正文或表格单元格中的一个段落，支持嵌套表格和多处独立占位段。开启字段格式化时，行内、页眉、页脚和文本框中的附件占位符会触发模板错误。空附件不会破坏单元格结构；插入的 YAML 值始终作为内容，不会被再次解释为模板占位符。

When disabled, these fields use ordinary template filling, including `template_fill.list_separator`; official-document container cleanup is disabled. The original YAML data is not rewritten. Removed field-processor configuration has no aliases, migration or dynamic module fallback.

关闭后，这些字段按普通模板规则填充，列表遵守 `template_fill.list_separator`，不执行公文专用容器清理。原始 YAML 数据不被改写。已删除的字段处理器配置不提供别名、迁移或动态模块回退。

## Optimization defaults / 专项优化默认选择

The Official documents and Invoices pages form a separate navigation group. Defaults use one resource ID per input category, with `"none"` as the shipped value. A default initializes the operation panel when an input is selected; it does not classify file contents or modify the ordinary conversion route. Changing the selection for a task does not save it as a new default.

“公文”和“发票”位于独立的导航分组。每类输入仅保存一个资源 ID，出厂值均为 `"none"`。默认值用于传入文件后初始化操作面板，不自动判断文件内容，也不改变普通转换入口。单次任务中的选择不会回写为设置默认值。

| Input / 输入类别 | Persisted key / 持久化键 | Built-in resource / 内置资源 | Settings page / 设置页 |
| --- | --- | --- | --- |
| Documents / 文档 | `document.to_md_optimization` | `gongwen` | Official documents / 公文 |
| PDF and OFD / 版式文件 | `layout.to_md_optimization` | `invoice_cn` | Invoices / 发票 |
| Images / 图片 | `image.to_md_optimization` | `invoice_cn` | Invoices / 发票 |

The value is an exact, nonempty resource ID or `"none"`. Availability comes from the existing capability catalog and may differ by format or installed conversion software. Unavailable choices remain saved but are not selected for execution; users can still clear their page's saved default. Resource IDs and internal action names remain separate. The removed enable/type pair has no alias, dual write or automatic migration.

取值为精确的非空资源 ID 或 `"none"`。可用性由现有能力目录决定，可能随文件格式和已安装的转换软件变化。不可用选择继续保留，但不会用于执行；用户仍可取消该页已保存的默认。资源 ID 与内部动作名称分开管理。已移除的 enable/type 配置对没有别名、双写或自动迁移。

Resetting Official documents restores field formatting and clears the document default only while its persisted choice is `gongwen`. Resetting Invoices independently clears layout and image defaults only while they are `invoice_cn`. These comparisons read persisted values inside the same configuration transaction and process lock as the reset. Other optimization choices, common input settings, and unrelated unsaved drafts are preserved. Concurrent external writes refresh the Cancel baseline without overwriting another page's draft. GUI and `docwen config reset gongwen --yes` / `docwen config reset invoice --yes` share the same reset plan.

恢复“公文”默认会恢复字段格式化开关，并仅在持久化的文档默认为 `gongwen` 时取消该选择。恢复“发票”默认分别检查版式与图片，仅取消值为 `invoice_cn` 的选择。条件判断在重置使用的同一配置事务和进程锁内读取持久化值。其他优化选择、通用输入设置和其他页未保存草稿得到保留；并发外部写入会刷新取消操作的基线，不覆盖其他页的草稿。GUI 与 `docwen config reset gongwen --yes` / `docwen config reset invoice --yes` 共用同一重置计划。

## Settings navigation and Markdown scope

Settings are grouped into Basics, Input types, Content and Specialized. The OCR and images page owns shared recognition and Markdown image output; input pages still own per-category OCR defaults. Markdown output rules remain source-independent, with separate input recognition and explicitly scoped Word mappings. Current conversion coverage is described in the page help and the [settings architecture review](maintenance/proposals/settings-information-architecture.md). Page identities, config keys and reset ownership are unchanged.

设置按基本设置、输入类型、内容处理、专项优化分组。“OCR 与图片”管理共享识别与 Markdown 图文输出，各输入页仍管理本类文件的 OCR 默认开关。Markdown 输出规则不按来源拆成重复配置；读取规则独立，Word 专用映射另行标注。当前覆盖范围见页面说明与上述设计复查，页面标识、配置键和重置归属不变。
