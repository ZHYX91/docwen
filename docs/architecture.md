# Architecture / 架构

DocWen uses a layered workspace. Dependencies point inward toward stable contracts; composition belongs to the bundle package.

DocWen 采用分层 workspace。依赖指向稳定的内层契约，最终组合由 bundle 包负责。

```text
core <- plugins
  ^       ^
  |       |
runtime  application
   \       /
     bundle
     /   \
   cli   gui
```

## Layers / 分层

- `docwen_core`: data contracts, shared parsing, OCR, links, formula and pure services.
- `docwen_application`: request admission, pre-conversion orchestration and use-case control.
- `docwen_runtime`: plugin registry, route resolution, workspaces, configuration, security, IPC and output finalization.
- `docwen_plugin_*`: format-specific route implementations; plugins depend on core contracts, not application or runtime internals.
- `docwen_cli` and `docwen_gui`: presentation and interaction layers.
- `docwen_bundle`: composition root, entry points and adapter wiring.

## Execution flow / 执行主链

1. CLI or GUI builds a request from user input and a configuration snapshot.
2. Application admits the request, owns protective input copies and coordinates optional pre-conversion.
3. Runtime resolves a manifest route and invokes the owning plugin.
4. The plugin returns typed results and artifact declarations.
5. Runtime finalizes outputs transactionally and emits one truthful terminal result.

## Boundaries / 边界

GUI orchestration (`execution_coordinator`) uses one route/admit/reserve/start path for single, batch
and aggregate tasks. It binds progress to the active request and refuses late starts after shutdown
or reentrant starts during confirmation. Request construction (`execution_requests`) owns input and
option snapshots, template binding,
output policy and redacted retry context. `execution_admission` owns frozen inspection validation and
acceptance records; the window only presents the confirmation. Acceptance is copied to live list state
only if it still represents the same inspection. `qt_bridge.execution_supervisor` owns input reservations,
native worker lifetime and each task's original cancellation controller. `qt_bridge.execution` executes
the frozen request. `execution_presenter` commits terminal task history, list status and result navigation
before emitting output-open and desktop-notification signals. Terminal result projection and native
worker cleanup are separate: neither cancellation
nor a reported startup error permits destroying a running worker. The window wires these collaborators
and presents their signals without inheriting their state or importing runtime internals.

GUI 编排模块统一单文件、批量与合并的路线/接纳/保留/启动流程，将进度绑定到实际请求，并拒绝关闭后的迟到启动及确认期间的重入。
请求构造模块负责输入与参数快照、模板绑定、输出策略及脱敏重试上下文；准入模块负责冻结检测事实与确认记录，
窗口仅展示确认。只有检测事实仍相同，才将确认同步到当前文件列表。执行监督对象拥有输入保留、原始控制器的取消句柄
与线程生命周期，后台线程只执行冻结请求。结果呈现模块先提交任务历史、列表状态与导航目标，再发出打开结果和桌面通知信号。
结果展示与线程清理分别处理；取消或启动报错不能销毁仍运行的线程。

`clipboard_paste` owns capture, rich-content projection and submission through the existing
input ViewModel. The window's managed store keeps its inspection, execution and history
leases; the paste handler owns no second file list or output writer. `execution_summary`
projects only the normalized request's options. The presenter displays that snapshot when
the task starts and displays merge counts from the completed result. Template recents are
session-only views of started tasks and never write the persisted default template.

`clipboard_paste` 负责捕获、富内容投影及向既有输入 ViewModel 提交。窗口管理的 Store 继续持有检测、
执行和历史租约；粘贴处理器不维护第二份文件列表，也不写转换输出。`execution_summary` 只投影规范化请求
中的选项，呈现器在启动时展示该快照，并从完成结果显示汇总计数。最近使用模板仅是当前会话已启动任务的视图，
不改写持久化的默认模板。

Word title/subtitle extraction and localized YAML front matter live in
`docwen_plugin_document.to_markdown.metadata`; the converter retains policy selection,
body rendering and artifact ownership. Document metadata and invoice writers share
`docwen_core.yaml_tools.format_yaml_scalar` for string-safe scalar encoding; each writer
retains its own normalization and empty-field policy. Spreadsheet coverage and decision evidence live in
`table_merger.semantics`; workbook loading reuses the CSV/XLSX formula-cache check and stops
before producing an artifact if a needed result is unknown. See the [merge guide](user-guides/table-merge.md).

Word 标题/副标题提取及本地化 YAML 前言由 `docwen_plugin_document.to_markdown.metadata` 负责；
转换器继续拥有策略选择、正文渲染及产物。文档元数据与发票写入器共用 Core 的
`yaml_tools.format_yaml_scalar` 保留标量字符串，各自保留原有规范化和空字段规则。
表格覆盖关系与处理证据由 `table_merger.semantics` 负责；
工作簿读取复用 CSV/XLSX 的公式缓存检查，所需结果未知时在生成产物前停止。参见[汇总指南](user-guides/table-merge.md)。

Plugin registration keeps last-registration-wins for a repeated ID. The complete next
route index is built before replacing registry state, so removed routes cannot resolve
to the new plugin. Registration order among different IDs is unchanged.

重复插件 ID 保持后注册替换前注册的策略。先构造完整的新路由索引，再替换注册状态，避免已移除的路由
指向新插件；不同 ID 之间的注册顺序不变。

Application conversion planning separates typed requests and results (`conversion_contracts`), capability
bindings (`conversion_capabilities`), option validation (`conversion_options`), runtime request construction
(`conversion_requests`), and the plan/accept/execute/cancel lifecycle (`conversion_service`). Plans own copies
of their input options and public result data; mutating a returned discovery or plan object cannot alter a
later execution. Input integrity and active capabilities are checked again at acceptance.

`conversion_routes` composes the same preconversion chain used by the controller and resolves every step
against the runtime catalog. `optimization_selection` requires all selected inputs to have an available
chain and intersects their route options. GUI localization and settings remain presentation concerns;
missing or unavailable Office bridge routes cannot be hidden by an available final DOCX optimizer.

应用层分别维护请求/结果类型、能力绑定、参数校验、运行请求构造与计划/接纳/执行/取消生命周期。
计划持有自己的参数和结果快照，外部修改发现结果或计划对象不会改变执行；接纳时仍重验输入完整性和当前能力。
转换路线计划与控制器使用同一前置转换链，逐步核对真实运行时路线。优化选择要求每个输入的完整链可用，
批量选项取交集；GUI 只负责显示与交互，不能因最终 DOCX 优化器可用而掩盖前置 Office 转换不可用。

- Core does not import application, runtime, apps, plugins or bundle.
- Plugins do not import application/runtime/apps/bundle.
- GUI and CLI do not deep-import plugin implementations or runtime internals.
- Network policy, configuration persistence and output placement each have one owner.

Output producers declare attachment and resource ownership before publication. Runtime uses these
relations to place child documents and their resources; Application projects the same relations into
Artifact Bundle instead of reconstructing Gongwen-specific ownership. Shared image bytes may back
separate resource artifacts, each owned by the document whose portable directory contains it.
Final delivery shape belongs to Runtime: Markdown-to-document/spreadsheet conversions publish a single
deliverable as a file; multiple deliverables or an explicitly requested audit sidecar form a result
directory. Conversions to Markdown keep their document directory. Private intermediate
directories in a conversion/proofreading pipeline do not dictate the final public layout.

产物生产者在发布前声明附件与资源归属，Runtime 据此布局，Application 将同一关系投影到 Bundle，
不再分别推断公文私有标记。正文和附件共用图片时可复用源字节，但各自产生有明确归属的资源产物，
使附件目录可独立携带。最终交付形态由 Runtime 决定，流水线的私有中间目录不决定用户输出目录。

DOCX numbering reviews, field reviews and proofreading share the Core comment append mechanism.
It reserves existing package-wide identities and preserves template-owned comments and styles;
callers retain business text, insertion provenance and exact span selection. The helper imports no
python-docx implementation, keeping document-library integration at the caller boundary.
- GUI view models own presentation state; widgets render or forward interaction and do not become persistence sources.

These rules are enforced by Import Linter, repository guards and Pyright.

这些规则由 Import Linter、仓库门禁和 Pyright 持续检查。

## Network protection and local control / 网络防护与本机控制

### English

DocWen does not provide file upload, telemetry, online downloads or a network service. Its supported GUI/CLI entry points install and verify a **CPython audit hook** before importing the product components, then enforce it throughout their execution. Packaged applications activate the same guard from a PyInstaller runtime hook and keep it active until process exit. If installation or verification fails, startup stops before normal application work begins; execution does not continue without the guard.

While active, the guard blocks **all DNS/name resolution** and the audited `AF_INET/AF_INET6` operations `bind`, `connect`, `connect_ex`, `sendto` and `sendmsg` from in-process Python dependencies. This includes local IP destinations. Non-IP local communication through **Windows named pipes** and **Unix-domain sockets** remains available.

This is protection against accidental dependency network access, **not an operating-system sandbox**. Native code and separately launched processes, including Word/Office, WPS, LibreOffice and the **dedicated Office helper**, are outside its coverage and follow their own and the operating system's network policies. The guard is not automatically inherited as protection for those processes. CPython also does not emit an audit event for an ordinary payload write on an already-connected or inherited socket; those writes are outside the guarantee. Supported entry points activate the guard before product imports and do not intentionally establish such sockets.

For local desktop control, the Core runtime uses a Windows named pipe on Windows and an `AF_UNIX` socket on Linux/macOS. A file lock establishes single-instance ownership; files do not carry control commands. These transport capabilities do not by themselves establish which operating systems a companion plugin supports.

### 中文

DocWen 不提供文件上传、遥测、在线下载或网络服务功能。受支持的 GUI/CLI 入口在导入产品组件前安装并验证 **CPython 审计钩子**，随后在整个执行期间启用防护。打包程序通过 PyInstaller 运行时钩子提前启用相同防护，并保持到进程退出。如果安装或验证失败，程序会在开始正常工作前停止，不会在缺少防护的情况下继续运行。

防护启用时，会阻止进程内 Python 依赖的**所有 DNS/名称解析**，以及 `AF_INET/AF_INET6` 的 `bind`、`connect`、`connect_ex`、`sendto`、`sendmsg` 受审计操作；本机 IP 地址也在限制范围内。**Windows 命名管道**和 **Unix 域套接字**等不使用 IP 的本机通信仍然可用。

这项防护用于防止依赖意外联网，**不是操作系统级沙箱**。原生代码，以及单独启动的 Word/Office、WPS、LibreOffice、**专用 Office 辅助进程（dedicated Office helper）**等进程不在覆盖范围内，仍遵循自身和操作系统的网络策略；这些进程不会自动继承这项防护。CPython 对已连接或继承的套接字进行普通数据写入时，也不会触发相应的审计事件，因此这类写入不在保证范围内。受支持的入口会在导入产品组件前启用防护，不会主动建立此类套接字。

本机桌面控制由 Core runtime 在 Windows 上使用命名管道，在 Linux/macOS 上使用 `AF_UNIX` 套接字。文件锁只负责确认单实例归属，文件不用于传送控制命令。这些底层通信能力不等于配套插件的平台支持范围。

Implementation references / 实现参考：

- [Network guard / 网络防护](../packages/runtime/src/docwen_runtime/security/network.py)
- [GUI entry point / 图形界面入口](../packages/bundle/src/docwen_bundle/gui_entry.py) and [CLI entry point / 命令行入口](../packages/bundle/src/docwen_bundle/cli_entry.py)
- [Packaged startup hook / 打包启动钩子](../packages/bundle/src/docwen_bundle/pyi_runtime_egress_guard.py)
- [Local desktop control / 本机桌面控制](../packages/runtime/src/docwen_runtime/ipc/single_instance.py)

## External Office macro policy / 外部 Office 宏策略

The Office bridge opens its private input copy only after confirming ownership of the COM process
and setting and reading back `AutomationSecurity=3` (ForceDisable). When Word has no application
window handle, a fixed, generated RTF probe supplies a document window for process identification;
the probe is closed before the actual input opens. An unowned server receives no application-setting
changes or Quit call. Failure to enforce the policy rejects that backend.

Excel's ForceDisable does not cover Excel 4.0 macros. Before COM opens a spreadsheet, a bounded
inspection rejects macro sheets, encryption and unrecognized or ambiguous framing. BIFF5/8 checks
compare declared sheet types with actual substreams; OOXML checks content types and relationships.
CSV must decode as text without binary controls or ambiguous SYLK/XML/RTF prefixes. Inspection is
limited to a 128 MiB container, 64 MiB BIFF/CSV, 10000 ZIP entries and 4 MiB per XML / 16 MiB total XML.
These conservative refusals can reduce COM compatibility; they do not certify arbitrary Office files
as safe. Other configured backends remain eligible only through their own policy gate.

Every disposable LibreOffice profile sets `DisableMacrosExecution=true` and retains applicable font
substitution. Failure to prepare the profile prevents launching LibreOffice. Exhausted backend
attempts retain the specific macro-policy diagnostic. This policy does not extend the Python network
guard to external applications or replace an operating-system sandbox. Installed Word, Excel, WPS and
LibreOffice behavior, including VBA/XLM suppression and ordinary conversion, requires separate host
acceptance; source tests prove inspection and call ordering only.

Office 桥接仅在确认 COM 进程归属、写入并读回 `AutomationSecurity=3` 后打开私有输入副本。
Word 没有应用窗口句柄时先打开固定生成的无宏 RTF 探针以确认 PID，关闭探针后才打开实际输入；
不更改或退出未归属的应用实例。无法设置禁宏策略时拒绝该后端。

Excel 4.0 宏不在 ForceDisable 的保证内，因此 COM 前还检查有界 BIFF5/8、OOXML 和 CSV 内容；
宏表、加密、未知或歧义结构均拒绝。限制为容器 128 MiB、BIFF/CSV 64 MiB、ZIP 10000 项、
单份 XML 4 MiB／总 XML 16 MiB；严格拒绝可能降低 COM 兼容性，不构成任意文档安全认证。
其他后端只能通过各自安全策略继续尝试。临时 LibreOffice 资料明确禁用宏执行并保留字体替换；
资料准备失败不启动进程。真实宿主的宏抑制与普通转换需单独验收，源码替身只证明检查和调用顺序。


### Linux external Office process ownership / Linux 外部 Office 进程归属

Linux LibreOffice calls use a short-lived spawned supervisor that enables child-subreaper ownership before launching the selected executable. The long-running GUI/CLI process is not a subreaper. The supervisor is the sole waitpid authority; launcher exit alone is insufficient. Procfs only discovers its direct children, kernel wait-child checks authorize each pidfd signal, and success/cleanup require ECHILD, a matching receipt and a normally exited, joined supervisor. Cooperative cancellation and the public execution timeout use bounded TERM, then KILL and reap. The existing 90-second default remains unchanged; supervisor startup has a separate finite bound.

Linux LibreOffice 调用先启动短生命周期监督进程并建立 child-subreaper 所有权，再运行所选可执行文件；长期 GUI／CLI 主进程不设置 subreaper。监督进程独占 waitpid 回收，launcher 退出不代表整条调用结束。procfs 仅发现直属孩子，每个 pidfd 信号还须通过内核 wait-child 关系检查；只有 ECHILD、匹配回执及监督进程正常退出且已 join，才允许成功或清理。协作取消与公开执行超时均有界 TERM、KILL、回收；原默认 90 秒不变，监督启动另有有限预算。

Child discovery prefers task `children` files. An unavailable interface is distinct from a disappearing secondary thread; when needed, a finite procfs scan selects only status records whose `PPid` is this supervisor in the procfs mount namespace. `NSpid` maps candidates into the caller namespace before the kernel wait-child/pidfd checks. Both paths are discovery only; scan exhaustion cannot authorize a signal or cleanup and never enables process-group fallback.

发现优先读取任务 `children` 文件，区分接口不可用与次线程已退出；必要时按有限预算扫描 procfs，只选取 `PPid` 为该监督者的状态记录。procfs 挂载命名空间中的候选先经 `NSpid` 映射，再进入内核 wait-child／pidfd 校验。两条路径都只用于发现；扫描预算耗尽不能授权发信号或清理，也不启用进程组 fallback。

Linux private Office inputs and backend outputs belong to an independent process-temp workspace, so Runtime task cleanup cannot erase bytes still potentially used by an unconfirmed native child. Confirmed results are copied through the existing bounded freshness-checked copy into same-volume publication staging. Missing exit/reap proof fails the bridge, stops backend fallback and preserves the profile, private input and backend output directory. Their paths remain in the plugin-consumed failure message. This is a local Office lifecycle contract; it adds no cross-repository parent-death protocol and does not claim recovery from D-state or the simultaneous forced death of every supervisor. Windows/macOS retain their existing process paths and need separate native acceptance.

Linux 的 Office 私有输入和后端输出置于独立的进程临时工作区；Runtime 清理本任务时不会删除未确认退出的原生孩子仍可能使用的字节。确认结束后，复用有界、带来源新鲜度核验的复制，把结果放入最终输出同卷的提交暂存。缺少退出／回收证明时，桥接失败、停止后端 fallback，并保留 profile、私有输入及后端输出目录，实际插件消费的错误消息包含这些路径。这是 Office 局部生命周期合同，不增加跨仓 parent-death 协议，不宣称已解决 D-state 或全部监督进程同时被强杀；Windows／macOS 保持既有进程路径，原生验收另行执行。
