# GUI behavior / GUI 行为

The GUI renders one application/runtime state through view models. Widgets own presentation and interaction wiring; they do not become configuration or task-truth sources.

GUI 通过 ViewModel 呈现唯一的 application/runtime 状态。Widget 负责展示和交互接线，不成为配置或任务真相源。

## Main behavior / 主要行为

- Single and batch input modes support file dialog, drag/drop, filtering and stable ordering.
- Available conversion panels and actions derive from selected files and route capabilities.
- Batch rows expose truthful pending, processing, completed, partial, failed and cancelled states.
- Admitted input files cannot be removed until their worker finishes. Completed operation records own failure details and retry intent independently of editable batch rows; retry restores removed inputs through normal admission.
- Mixed batches use the intersection of routes for every admitted source format. XLSX protection analysis runs in the background with bounded file-identity caching; only the latest selection can publish results. Pending or unknown protection blocks ODS delivery. Empty protection elements are not enabled locks, and execution still revalidates inputs.
- Cancellation becomes reachable immediately, is idempotent and does not publish incomplete output.
- Settings use persisted, draft and preview layers; Apply/OK/Cancel and reset operations preserve their ownership boundaries.
- Light, dark and system themes, semantic typography presets and high-DPI geometry remain supported.
- Form rows reflow after font, style, text and geometry changes. System theme preference survives font changes and responds to Qt system color-scheme events.
- Keyboard shortcuts are suppressed while text-editing controls own focus.
- A second launch forwards files to the existing instance through IPC.

Single-file mode retains exactly one validated current input across file picking, dropping, second launches and Assistant GUI requests. A batch keeps its entire visible list; switching a multi-file batch back to single mode confirms the retained selection. Busy inputs cannot be replaced, and rejected external requests report failure to their caller. Ordinary input receipt is not a task-history event. Assistant CLI processing does not change GUI selection.

单文件模式在添加、拖入、二次启动和 Assistant GUI 传入时均只保留一个验证通过的当前文件。批量模式保留完整可见清单；多文件批次切回单文件时确认保留项。运行中不替换输入，外部请求被拒绝时向调用方报告失败。普通接收不进入活动记录；Assistant CLI 处理不改变 GUI 选择。

## Settings ownership / 设置归属

The fourteen pages follow this order: general, incoming text, incoming documents, incoming spreadsheets, incoming images, incoming layout files, incoming other files, proofreading, Markdown syntax, links, Markdown resources, conversion software, file saving, logging. Incoming text owns Markdown content formatting, heading merging, table styling and template filling; incoming documents own DOCX content retention. Markdown syntax owns only syntax and extension policies. Global OCR language belongs to Markdown resources, and all eight converter priority lists belong to conversion software. Numbering editor changes refresh both incoming text and document pages without resetting unrelated drafts.

十四页按“常规 → 传入文本 → 传入文档 → 传入表格 → 传入图片 → 传入版式文件 → 传入其他文件 → 校对 → Markdown 语法 → 链接 → Markdown 资源 → 转换软件 → 文件保存 → 日志”排列。传入文本负责 Markdown 内容格式、标题合并、表格样式和模板填充；传入文档负责 DOCX 内容保留。Markdown 语法只管理语法与扩展策略。全局 OCR 语言归入 Markdown 资源，八组转换器优先级归入转换软件。编号编辑器改动同步刷新传入文本和文档页，并保留其他草稿。

The four proofreading rule sets expose edit, import and export together. The paired punctuation editor names opening and closing symbols. Imports validate TOML and preview additions, conflicts and retained entries before the user chooses merge or replacement. Invalid merges are unavailable; cancellation or a changed source preserves existing data. Saves compare preview source inside the configuration transaction. Exports use the complete effective editable source, including comments, independently of editor filters.

四组校对规则同时提供编辑、导入和导出；成对标点编辑器明确区分左侧和右侧符号。导入先验证 TOML，再预览新增、冲突和保留项，由用户选择合并或替换；无效合并不可选。取消或原配置变化时保留已有数据，保存事务内部再次核对预览原文。导出完整有效配置及注释，不受编辑器筛选影响。

## Accessibility and feedback / 可访问性与反馈

Visible controls require localized labels or accessible names. Errors, warnings and confirmations use the shared feedback layer. Terminal summaries, history and retained artifacts must agree with runtime truth.

Activity records open directly in one modeless window with search, status/operation filters, sorting and an integrated detail pane. Per-file warnings and skip reasons remain available independently of the bounded notification feed. Selection remains readable in both themes. The main result card uses its terminal state as its title, omits redundant single-file counts, and links to this same activity window; failures give the entry a warning colour.

## Visual system / 视觉规范

- Shared cards use a left-aligned, theme-aware header band and 16 logical-pixel body padding. Main workflow and settings cards share this primitive.
- Regular controls have a 36 logical-pixel minimum height; execution buttons have a 40 logical-pixel minimum height and fill their card's content width. Short text buttons have an 80 logical-pixel minimum width. Font metrics may increase these dimensions.
- Related controls use an 8-pixel horizontal gap, functional groups and form rows use 12 pixels, and cards use 16 pixels. Qt handles system DPI scaling; these tokens are not scaled a second time.
- Checkbox labels follow the indicator on its right. Form fields align within a settings page; narrow rows and long localized text reflow without clipping. Proofreading choices use one column when two do not fit.
- Markdown target selection occupies a separate form row above the full-width generation action. The action names its target format. Batch conversion buttons show the same category count used by request dispatch; merge buttons show their own matching-input count.
- The workspace previews committed output-directory settings. Template selection has a persistent check marker and a visible selected name. Idle feedback shows a compact ready-state card header; its body expands when work starts and remains open for results. Existing activity and non-default output hints remain reachable while idle.
- Light and dark themes use identical geometry, with readable hover, pressed, disabled and keyboard-focus states.
- Theme-independent design tokens and shared control geometry own dimensions and spacing. Theme colour changes reuse the same complete button box model; new themes do not redeclare padding or minimum sizes.

## Regression / 回归

Widget/view-model tests cover deterministic behavior. GUI smoke, screenshots and physical desktop interaction cover the final host-dependent surface. Current reference screenshots are stored under `docs/assets/screenshots/`.
