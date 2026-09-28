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

### Clipboard Markdown input / 剪贴板 Markdown 输入

The input header keeps Single File and Batch as a vertical choice group on the left. Add remains the primary action; Paste is a normal action; Clear is visually quiet and separated from the other actions. The action group stays horizontal when it fits. When translated labels, large text or a narrow viewport do not fit, the mode and action groups reflow to separate rows; at still smaller widths the action group may stack rather than hide, truncate or replace labels with icons.

Paste is an explicit snapshot operation. A button click or Ctrl+V while the input area (and not an unrelated text editor) owns focus reads the clipboard's current plain text once and writes those exact Unicode characters as UTF-8 Markdown in a DocWen-owned profile directory. No background listener exists. YAML front matter, fenced code and meaningful leading/trailing whitespace are preserved. Path-shaped strings, URLs and HTML are ordinary Markdown text, not commands. Empty or whitespace-only text creates no input, and non-text clipboard data reports an explicit message. A later clipboard change cannot mutate an admitted snapshot.

The backing filename is opaque and is not a user source path. The visible identity is a localized “Clipboard Markdown N.md” label with a bounded, control-safe plain-text preview. Clipboard snapshots never enter Recent Files and do not expose an Open source location action. They have no source resource directory: relative resources are not guessed from the process working directory or the managed snapshot directory, and remote resources retain the existing link/network policy.

Single-file paste replaces the current input only after the normal Core admission succeeds. Batch paste adds one independently visible snapshot per user action. UI selection owns snapshots while they remain editable; the execution supervisor independently retains them through worker shutdown; failed task history retains only snapshots needed for retry. Successful, skipped and cancelled terminal records release their snapshot history ownership. Clearing history or the input list removes snapshots once no live owner remains. Retry always reuses the original snapshot and fails explicitly if that snapshot is no longer available; it never rereads the clipboard.

A synthetic clipboard source has no user source directory for the default source-output policy. Before conversion starts, GUI therefore asks for a persistent output parent. Cancelling that chooser does not start conversion and does not modify saved output preferences. A valid configured custom output directory is reused without prompting. In a mixed ordinary-file/clipboard batch, only clipboard inputs receive per-input output-directory overrides; ordinary files retain same-as-source placement.

输入区左侧用纵向单选显示“单文件/批量”，右侧保留“添加/粘贴/清空”完整文字；“添加”为主操作，“粘贴”为普通操作，“清空”为弱化操作并与前两者留出间隔。空间不足时模式组与操作组分行；更窄或大字号/长翻译场景继续重排，而不是把文字截断或退化成纯图标。

“粘贴”是显式快照操作：仅在用户点击按钮，或输入区（而非其他文本编辑框）拥有焦点时按 Ctrl+V，才读取一次当前纯文本并将完全相同的字符按 UTF-8 Markdown 写入 DocWen profile 所有的受管目录。不会后台监听剪贴板；YAML front matter、代码块及有意义的首尾空白原样保留；看似文件路径、URL 或 HTML 的文本仍只是 Markdown 内容。空值/纯空白不创建输入，非文本剪贴板给出明确提示；之后的剪贴板变化不能改变已接纳快照。

底层文件名是内部不透明身份，不作为用户源路径。界面显示本地化“剪贴板 Markdown N.md”及有界、控制字符安全的纯文本预览；不会加入“最近文件”，活动记录也不提供源位置入口。合成输入没有来源资源目录，不从 cwd 或受管目录猜测相对资源；远程资源继续遵循既有链接/网络策略。

单文件粘贴仅在正常 Core 准入成功后替换当前输入；批量每次用户操作追加一个独立快照。可编辑列表、活动 worker 与失败重试历史分别持有快照：worker 完整结束前不会被清理，失败重试只保留所需原快照；成功、跳过、取消以及历史清理会释放无主快照。重试必须复用原快照，原快照缺失时明确失败，不读取新剪贴板冒充旧输入。

合成剪贴板输入在默认 source 输出策略下没有用户源目录，因此执行前必须选择持久输出父目录；取消选择不启动转换，也不更改已保存输出偏好。已有有效 custom 输出目录直接沿用。普通文件与剪贴板混合批量时，只为剪贴板项设置逐输入输出父目录，普通文件仍按 same-as-source 规则发布。

Execution owns independent input metadata and option snapshots. Confirming a detected format applies only
to the facts shown, and updates the live list only while those facts still match. The worker rechecks the
input before conversion. Cancellation and reservation release use the controller that started the task;
shutdown rejects late starts and retains live workers until they finish. Result and activity state are
committed before output navigation and completion notifications.

执行持有独立的输入元数据与参数快照。格式确认仅针对展示的检测事实，当前列表仍匹配时才同步确认记录；后台执行前继续重验输入。
取消与保留资源释放使用启动任务时的控制器。关闭后拒绝迟到启动，并保留仍运行的线程直至结束；结果与活动状态先提交，再打开结果或发出完成通知。

## Settings ownership / 设置归属

The fifteen pages follow this order: general, incoming text, incoming documents, incoming spreadsheets, incoming images, incoming layout files, incoming other files, proofreading, Markdown syntax, templates, links, Markdown resources, conversion software, file saving, logging. Incoming text owns Markdown content formatting, heading merging, table styling and template filling; incoming documents own DOCX content retention. Markdown syntax owns only syntax and extension policies. Global OCR language belongs to Markdown resources, and all eight converter priority lists belong to conversion software. Numbering editor changes refresh both incoming text and document pages without resetting unrelated drafts.

十五页按“常规 → 传入文本 → 传入文档 → 传入表格 → 传入图片 → 传入版式文件 → 传入其他文件 → 校对 → Markdown 语法 → 模板 → 链接 → Markdown 资源 → 转换软件 → 文件保存 → 日志”排列。传入文本负责 Markdown 内容格式、标题合并、表格样式和模板填充；传入文档负责 DOCX 内容保留。Markdown 语法只管理语法与扩展策略。全局 OCR 语言归入 Markdown 资源，八组转换器优先级归入转换软件。编号编辑器改动同步刷新传入文本和文档页，并保留其他草稿。

The four proofreading rule sets expose edit, import and export together. The paired punctuation editor names opening and closing symbols. Imports validate TOML and preview additions, conflicts and retained entries before the user chooses merge or replacement. Invalid merges are unavailable; cancellation or a changed source preserves existing data. Saves compare preview source inside the configuration transaction. Exports use the complete effective editable source, including comments, independently of editor filters.

四组校对规则同时提供编辑、导入和导出；成对标点编辑器明确区分左侧和右侧符号。导入先验证 TOML，再预览新增、冲突和保留项，由用户选择合并或替换；无效合并不可选。取消或原配置变化时保留已有数据，保存事务内部再次核对预览原文。导出完整有效配置及注释，不受编辑器筛选影响。

## Accessibility and feedback / 可访问性与反馈

Visible controls require localized labels or accessible names. Errors, warnings and confirmations use the shared feedback layer. Terminal summaries, history and retained artifacts must agree with runtime truth.

If route preflight rejects a new request, the main feedback shows that rejection and clears the previous task's result shortcut. Earlier operations remain available in activity records; a rejected request is recorded as not started, without inventing a failed worker or a new output.

新请求在路线预检中被拒绝时，主反馈显示本次拒绝原因，并清除旧任务的输出快捷入口。先前操作仍保留在活动记录；被拒请求记为未开始，不伪造执行失败或新产物。

Activity records open directly in one modeless window with search, status/operation filters, sorting and an integrated detail pane. Per-file warnings and skip reasons remain available independently of the bounded notification feed. Selection remains readable in both themes. The main result card uses its terminal state as its title, omits redundant single-file counts, and links to this same activity window; failures give the entry a warning colour.

Activity records and copyable feedback separate local details from redacted diagnostics. The diagnostic preview and Copy diagnostics use the same finite summary: status, reviewed error/exception categories, output/warning counts and an explicitly reported recoverability flag when available. They exclude document content, paths, identifiers, arbitrary error codes, raw errors, tracebacks, commands and configuration. Unknown categories remain unknown. Reading local details is independent of sharing a summary; nothing is copied automatically. Copying keeps the diagnostic dialog open, and viewing diagnostics does not resolve or cancel a pending recovery choice.

活动记录与可复制反馈将本地详情和脱敏诊断分开展示。诊断预览和“复制诊断”使用同一份有限摘要：状态、已审核的错误/异常类别、输出与警告数量，以及实际返回的可恢复标记（若有）。摘要不含正文、路径、标识符、任意错误码、原始错误、堆栈、命令或配置；未知类别保持未知。本地详情供当前用户阅读，不自动复制。复制后诊断窗口保持打开，查看诊断不会确认或取消待处理的恢复选择。

## Visual system / 视觉规范

The initial main-window client size is 476×720 Qt logical pixels, with a 420×560 minimum. Saved user geometry takes precedence. Content scrolls independently of the fixed bottom bar when the window is short; available-screen fitting remains authoritative on smaller work areas.

- Shared cards use a left-aligned, theme-aware header band and 12 logical-pixel body padding. Main workflow and settings cards share this primitive.
- Regular controls have a 32 logical-pixel minimum height; execution buttons have a 36 logical-pixel minimum height and fill their card's content width. Short text buttons have an 80 logical-pixel minimum width. Font metrics may increase these dimensions.
- Related controls and form rows use an 8-pixel gap; cards use 12 pixels. These are 100% design values. Application zoom scales them once in Qt logical coordinates; Qt independently handles physical screen DPI.
- Checkbox labels follow the indicator on its right. Form fields align within a settings page; narrow rows and long localized text reflow without clipping. Proofreading choices use one column when two do not fit.
- Markdown target selection occupies a separate form row above the full-width generation action. The action names its target format. Batch conversion buttons show the same category count used by request dispatch; merge buttons show their own matching-input count.
- The workspace previews committed output-directory settings. Template selection has a persistent check marker and a visible selected name. Idle feedback shows a compact ready-state card header; its body expands when work starts and remains open for results. Existing activity and non-default output hints remain reachable while idle.
- Light and dark themes use identical geometry, with readable hover, pressed, disabled and keyboard-focus states.
- Theme-independent design tokens and shared control geometry own dimensions and spacing. Theme colour changes reuse the same complete button box model; new themes do not redeclare padding or minimum sizes.

## Text size and interface scale / 字号与界面缩放

General → Appearance provides three text sizes (Small / Standard / Large) and five interface scales (90 / 100 / 110 / 125 / 150%). Standard body text is 10.5 pt (14 logical pixels at 96 DPI); Small and Large use 9 pt and 13 pt. Semantic headings keep their hierarchy.

“通用 → 界面外观”提供小、标准、大三档字号，以及 90 / 100 / 110 / 125 / 150% 五档界面缩放。字号只调整文字层级；界面缩放同时调整文字、图标、控件、边距与间距。系统 200% 配合应用 90% 的名义物理比例为 180%，由 Qt 直接绘制文字与矢量资源，不缩放整张窗口位图。像素取整和屏幕字体渲染仍可能产生细微差别。

Changes preview immediately. Apply or OK saves them; Cancel restores the last saved appearance. Reset uses the existing confirmed reset transaction. Existing and newly opened controls use the same scale. Settings navigation becomes a page selector when the sidebar and content cannot fit. Scroll areas keep long forms and the main workflow reachable.

`styles/design_tokens.py` owns baseline sizes; `styles/ui_scale.py` owns content scaling and baseline metric bindings; `font_utils.py` owns text presets; `ThemeManager` applies the assembled stylesheet and application font. Font-measured layout dimensions are not scaled a second time. `window_geometry.py` owns Qt logical window positions, sizes, restoration and screen fitting, independently of content zoom.

`gui.appearance.scale_percent` is the sole content-scale setting. The loader removes retired `gui.dpi.ui_scale` and `gui.dpi.enable_dpi_scaling` preferences without interpreting them as the new zoom. It maps the old extra-large font preference to Large once, preserves unrelated preferences and does not write defaults into sparse user files. Runtime controls accept only the current presets.

Windows native title bars follow the resolved Light / Dark / System theme through `ThemeManager` and Qt’s application color-scheme hint. Existing windows update on preview and rollback; Qt applies the current scheme to newly shown or recreated native window handles. The operating system retains the title bar, caption buttons, system menu and geometry behavior. There is no separate caption-color preference.

Shared checkboxes preserve Qt keyboard/accessibility behavior and draw their indicators as vectors. Brand art uses `assets/icon.svg`; action icons use the licensed SVG collection under `assets/icons`. Runtime icons render for the current paint device. Windows package builds generate theme-aware target-size variants and a `resources.pri` index before MakeAppx packaging; asset/index checks do not replace installed Start menu verification.

## Regression / 回归

Widget/view-model and Runtime integration tests cover deterministic clipboard, ownership, routing and publication behavior. GUI smoke, screenshots and physical desktop interaction remain the authority for final host-dependent layout and clipboard interaction. In particular, minimum-width, Large text, 150% application scale, long translations, Light/Dark themes, Tab order and physical clipboard conversion require Computer Use or equivalent desktop acceptance; automated Qt tests do not claim that pass. Current reference screenshots are stored under `docs/assets/screenshots/`.
