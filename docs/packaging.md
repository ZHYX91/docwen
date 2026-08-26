# Packaging / 打包

DocWen 0.9 publishes one Windows x64 package and two Ubuntu 24.04 x64 packages built from the bundle composition root. The Windows archive contains both GUI and CLI; Ubuntu has GUI+CLI and CLI-only archives. Package verification runs against the produced directory, not the source tree. macOS has no 0.9 release asset, and its current capability contract reports conversion, validation, numbering, merge, and split as unavailable.

DocWen 0.9 正式发布一个 Windows x64 完整包和两个 Ubuntu 24.04 x64 包；Windows 包同时包含 GUI 与 CLI，Ubuntu 分为 GUI+CLI 和仅 CLI 两个压缩包。打包验证面向实际产物目录执行，而不是用源码态结果代替。macOS 没有 0.9 正式附件，当前 capability 契约明确把转换、校对、编号、合并和拆分报告为不可用。

| Platform / 平台 | 0.9 distribution status / 发行状态 | Evidence boundary / 证据边界 |
| --- | --- | --- |
| Windows x64 | Published in 0.9.0 / 已随 0.9.0 发布 | Exact packaged candidate plus automated and manual Windows acceptance / 精确打包候选及 Windows 自动与人工验收 |
| Ubuntu 24.04 x64 | Published in 0.9.0 / 已随 0.9.0 发布 | Exact manifest-bound package, post-extract automation and Ubuntu desktop acceptance / 精确清单绑定包、解压后自动化及 Ubuntu 桌面验收 |
| macOS x64/arm64 | No 0.9 release asset / 无 0.9 正式附件 | Source CI and opt-in packaging experiment only; primary document operations are unavailable / 仅源码 CI 与手动打包实验，主要文档操作不可用 |

## Required contents / 必需内容

- `DocWen.exe` or `DocWenCLI.exe` and PyInstaller `_internal` content.
- `configs/`, `templates/`, `models/`, locale files and application assets.
- The complete `pymupdf-layout` distribution resource manifest under `_internal/pymupdf/layout/resources`, including its ONNX models and YAML descriptors.
- License, third-party notices and the supported runtime metadata.

PyInstaller collection targets are import-package names, not necessarily distribution names. In particular, the `pymupdf-layout` distribution installs the `pymupdf.layout` package. The build preflight rejects missing or non-package collection targets, and post-package verification derives the expected Layout resource paths from the installed distribution instead of pinning one model-version filename.

PyInstaller 收集参数使用可导入包名，不一定等于发行包名；`pymupdf-layout` 发行包实际安装为 `pymupdf.layout`。构建前置检查会拒绝不存在或不是包的收集目标；产物检查则从当前已安装发行包清单推导全部 Layout 资源路径，要求 ONNX 模型与 YAML 描述文件完整、非空，而不是只写死某一个模型版本文件名。

`pymupdf-layout` 1.27.2.2 ships the same seven resource paths on every platform covered by the locked resource manifests, but its three YAML descriptors use CRLF bytes in the Windows wheel and LF bytes in the Linux and macOS wheels. DocWen therefore pins two complete raw-byte manifests: one for `win32`, and one shared by `linux` and `darwin`. This resource-integrity coverage is independent of the public package-support boundary. Unknown platforms fail closed with `unsupported_resource_platform`; verification does not normalize line endings, accept alternate hashes, or learn trusted bytes from the installed package. Dependency upgrades must audit every locked platform wheel and update both manifests explicitly.

`pymupdf-layout` 1.27.2.2 在锁定资源清单覆盖的平台上提供相同的七条资源路径，但三个 YAML 描述文件在 Windows wheel 中使用 CRLF 字节，在 Linux 与 macOS wheel 中使用 LF 字节。因此 DocWen 固定两套完整的原始字节清单：`win32` 一套，`linux` 与 `darwin` 共用一套。该资源完整性覆盖与正式发布平台边界相互独立。未知平台以 `unsupported_resource_platform` 失败关闭；校验过程不规范化换行、不接受备选哈希，也不从已安装包自举可信字节。升级依赖时必须审计锁文件中的每个平台 wheel，并显式更新两套清单。

The Windows production builder limits PyInstaller's DLL search path to the clean project environment, the manifest-verified CPython base directory, and the Windows system directories. Host-selected `api-ms-win-*` forwarders and `ucrtbase.dll` are forbidden in the payload; supported Windows supplies those system components. The four packaged MSVC runtime files are instead replaced from the locked `pikepdf` wheel before the frozen payload allowlist is checked.

Windows 生产构建器把 PyInstaller 的 DLL 搜索路径限制为干净项目环境、清单验证过的 CPython 基础目录和 Windows 系统目录。载荷禁止包含由宿主偶然选中的 `api-ms-win-*` 转发 DLL 与 `ucrtbase.dll`，这些系统组件由受支持的 Windows 提供；随包发布的四个 MSVC 运行库文件则在冻结载荷白名单检查前统一替换为锁定 `pikepdf` wheel 中的副本。

## Release gates / 发布门禁

Repository settings must enable Immutable Releases and an active no-update/no-delete ruleset for numeric `x.y.z`
tags before a version tag is pushed. 仓库设置必须在推送版本标签前启用 Immutable Releases，并启用禁止更新或
删除数字 `x.y.z` 标签的 ruleset。

The hosted GitHub Release workflow enforces this deterministic baseline before publishing one Windows archive and two Ubuntu 24.04 x64 archives:

1. Source repository release tests on Windows, Linux, and macOS.
2. Windows package resource/layout verification.
3. Packaged Windows CLI baseline doctor, conversion, and JSON checks.
4. Packaged Windows GUI settings-page construction smoke.
5. Manifest-bound deterministic Ubuntu archive construction with fixed `DocWen-0.9.0-linux-x64.tar.gz` and `DocWenCLI-0.9.0-linux-x64.tar.gz` names.
6. CLI and GUI verification against fresh directories extracted from those exact Ubuntu archives.
7. Byte-for-byte comparison of two clean builds per release platform, followed by Release SHA-256 generation over all three archives.

The exact local Windows candidate used for a release decision has a wider gate: source/package CLI parity; OCR and successful-warning CLI paths; GUI startup, settings, notification/OCR, IPC, and successful-warning paths; plus Office, presentation, and SmartDoc routes on a machine where their external dependencies are actually configured. These environment-sensitive checks are recorded separately and are not falsely attributed to the hosted workflow. Visible notification presentation, target-device rendering, and manual UI inspection remain manual evidence.

托管的 GitHub Release 工作流在发布一个 Windows 压缩包和两个 Ubuntu 24.04 x64 压缩包前执行以下确定性基线：

1. 在 Windows、Linux 和 macOS 上运行源码态 release 测试。
2. 验证 Windows 打包资源与 Layout 资源。
3. 验证 Windows 打包 CLI 的基础 doctor、转换和 JSON 路径。
4. 验证 Windows 打包 GUI 的设置页构造。
5. 按受版本控制的生产清单确定性构造 Ubuntu 压缩包，并固定为 `DocWen-0.9.0-linux-x64.tar.gz` 与 `DocWenCLI-0.9.0-linux-x64.tar.gz`。
6. 从这两个精确 Ubuntu 压缩包解压到全新目录后，再分别验证 CLI 和 GUI。
7. 对每个发布平台执行两次干净构建并逐字节比较，再为三个压缩包生成 Release SHA-256。

用于发版决策的精确本地 Windows 候选还要通过更宽门禁：源码/打包 CLI 一致性、CLI OCR 与成功警告、GUI 启动、设置、通知/OCR、IPC 与成功警告，以及在外部依赖已真实配置的机器上验证 Office、演示文稿和 SmartDoc 路径。这些环境相关结果单独记录，不能冒充托管工作流已经执行。通知中心可见性、目标设备渲染和人工 UI 检查仍属于人工证据。

Signing and publication are separate release operations. A successfully built unsigned package must not be described as signed or published.

The Ubuntu archives are generated only by `scripts/release/linux_archive.py` under `release/linux-production-manifest.v1.json`. The contract fixes the top-level directory, entry order, owner, timestamp, modes, gzip header, generated `manifest.json` and payload `SHA256SUMS.txt`. It permits only manifest-declared relative symlinks to internal regular files and rejects absolute, escaping, dangling, directory-target and cyclic links. The archive helper verifies the completed bytes before publishing them without replacement. Hosted post-extract smoke is package evidence; visible desktop behavior and target-host integration remain separate acceptance evidence.

Ubuntu 压缩包只能由 `scripts/release/linux_archive.py` 按 `release/linux-production-manifest.v1.json` 生成。该契约固定顶层目录、条目顺序、属主、时间戳、权限、gzip 头、内置 `manifest.json` 和载荷 `SHA256SUMS.txt`；仅允许清单声明的、指向包内普通文件的相对符号链接，并拒绝绝对、越界、悬空、目录目标和循环链接。helper 会先校验完整压缩包字节，再以不覆盖已有目标的方式发布。托管的解压后 smoke 属于发布物证据；可见桌面行为与目标宿主集成仍需单独验收。

## Packaged GUI smoke boundaries / GUI 打包冒烟边界

The default GUI smoke uses `DOCWEN_GUI_TEST_AUTOCLOSE_MS` and disables IPC so it cannot prove the single-instance lock, second-launch file delivery, window activation, or semantic settings control. Use the optional `--ipc-smoke` for that boundary, with no DocWen GUI already running. It first requires the packaged `info --json` result to expose the exact protocol-3 `gui.settings` v1 contract (`runtime_check_required`, cold start, and only the `proofread` section). It then invokes packaged `DocWenCLI gui open-settings --section proofread` as the operation that cold-starts the GUI, verifies the runtime `open_settings`/`proofread` handshake, opens the same section again to prove singleton reuse, and forwards a file to that same primary instance. The verifier only terminates a failed test process after it has captured the exact PID created inside this run; a pre-existing GUI causes a fail-closed result and is never cleaned up. Deadline race and budget semantics remain source-level adversarial evidence: tests prove that an expired request cannot open settings later and that an explicit deadline is not shortened by the default 15-second queue budget. The packaged IPC gate does not inject delays into a release binary and therefore does not claim those two timing cases as package-level evidence. `--notification-smoke` and `--ocr-smoke` exercise the packaged application paths but do not prove that a Windows notification was visibly presented or that every OCR model/language works on the target device.

默认 GUI smoke 使用 `DOCWEN_GUI_TEST_AUTOCLOSE_MS` 并禁用 IPC，因此不证明单实例锁、二次启动文件投递或窗口激活，也不证明语义设置控制；该边界使用可选 `--ipc-smoke`，且运行前不能已有 DocWen GUI。该门先要求打包 `info --json` 精确声明 protocol 3 的 `gui.settings` v1 合同（`runtime_check_required`、允许冷启动且仅支持 `proofread`），随后直接以打包 `DocWenCLI gui open-settings --section proofread` 冷启动 GUI，核对运行时 `open_settings`/`proofread` 握手，再次打开同一页证明单例复用，并把文件投递到同一个主实例。只有本轮已取得的精确 PID 才允许在失败清理时终止；预先存在的用户 GUI 会让门禁安全失败，绝不会被清理。超时竞态与预算仍明确属于源码级对抗证据：测试证明已过期请求不会稍后打开设置，显式 deadline 也不会被默认 15 秒排队预算缩短；打包 IPC 门不会向发布二进制注入延迟，因此不把这两个时序场景冒充为包级证据。`--notification-smoke` 与 `--ocr-smoke` 验证打包应用调用链，但不证明 Windows 通知中心可见性，也不替代目标设备上的完整 OCR 模型与语言验收。

## Successful-warning smoke / 成功但有警告的冒烟

Both packaged verifiers accept `--successful-warning-smoke <input>`. The GUI verifier also accepts `--successful-warning-smoke` without an input and then creates the hash-pinned deterministic Gongwen fixture owned by `scripts/release/successful_warning_fixture.py`. Before starting the GUI, it runs the packaged `DocWenCLI` from the same candidate against the same input, action, target, `zh_CN` locale and isolated configuration; exactly one non-empty `GONGWEN-NEEDS-REVIEW` warning and a non-empty output are required. That canonical message is then required in the GUI warning row, tooltip, persistent warning-tone task summary and warning-row PNG, and the CLI/GUI output bytes must match. `--successful-warning-message` is only an additional business golden and cannot override the packaged CLI result. For an explicit input, use `--successful-warning-input-sha256` to pin its identity. This does not prove notification-center visibility.

CLI 与 GUI 打包验证器都支持 `--successful-warning-smoke <input>`；GUI 还支持省略 input，自动创建由 `scripts/release/successful_warning_fixture.py` 管理且哈希固定的 Gongwen 夹具。GUI 启动前，验证器先用同一候选内的打包 `DocWenCLI`，在同一输入、action、target、`zh_CN` locale 和隔离配置上取得唯一且非空的 `GONGWEN-NEEDS-REVIEW` warning，并要求 CLI 产物非空；随后 GUI 必须原样显示该 canonical message，保留 warning row、tooltip、persistent warning-tone task summary 与有效 PNG，并与 CLI 产物字节一致。`--successful-warning-message` 只能增加业务 golden，不能覆盖 CLI 实际结果；显式输入可用 `--successful-warning-input-sha256` 固定身份。不得外推为通知中心可见性。

## Proofread report smoke / 校对报告冒烟

The optional packaged CLI flag `--proofread-report-smoke` creates its own byte-stable Markdown fixture containing a UTF-8 BOM, CRLF, a non-BMP emoji, a combining sequence, a ZWJ sequence, fullwidth symbols, and an unmatched bracket. It invokes the packaged `DocWenCLI validate` command with an explicit report path and verifies `docwen.proofread_report.v2`, the raw-input SHA-256, zero-based Unicode code-point ranges with exclusive ends, source-slice identity for every issue, and machine-applicable `replace_text` fixes only where a rule supplied an explicit replacement. A second `--check none` invocation must successfully write an empty report; a typed CLI failure cannot satisfy that empty-result gate. The input bytes must remain unchanged throughout. This is packaged CLI contract evidence, not proof of Obsidian positioning or fix application.

可选的打包 CLI 参数 `--proofread-report-smoke` 会自行生成字节固定的 Markdown 夹具，覆盖 UTF-8 BOM、CRLF、非 BMP emoji、组合字符、ZWJ 序列、全角符号和未闭合括号。门禁通过显式 report 路径调用打包 `DocWenCLI validate`，核对 `docwen.proofread_report.v2`、原始输入 SHA-256、零基 Unicode code-point 且 exclusive-end 的坐标、每项 issue 与源文本切片一致，以及只有规则显式给出 replacement 时才出现可应用的 `replace_text` fix。随后以 `--check none` 再次执行，必须成功写出空报告；typed CLI failure 不能冒充该成功空结果。全过程还要求输入原始字节不变。该结果只属于打包 CLI 合同证据，不替代 Obsidian 定位和修复应用验收。

## Optional successful-run evidence / 可选成功证据保留

`verify_packaged_gui.py` removes its isolated temporary run directory after a successful verification by default. Pass `--evidence-dir <ABSOLUTE_NEW_DIRECTORY>` only when the exact successful-run files must be retained for review. The destination must be an absolute normalized path whose parent already exists; it must not exist already, be inside the packaged binary directory, use `..` or another resolved-path alias, or pass through a symbolic link, junction, or other reparse point. Retention starts only after every requested smoke and diagnostic check succeeds. The verifier copies the complete directory without following links, then checks the directory/file inventory and every file's exact bytes, byte count, and SHA-256. It never overwrites an existing destination. A copy or verification failure fails closed and leaves the authoritative temporary run directory in place; a partially created destination is not evidence of success and must not be reused.

`verify_packaged_gui.py` 默认在验证成功后删除隔离临时运行目录。仅在需要保留本轮精确成功文件供复核时，显式传入 `--evidence-dir <绝对且不存在的目录>`。目标必须是父目录已存在的绝对规范化路径；不得预先存在、不得位于打包二进制目录内、不得含 `..` 或其他解析路径别名，也不得经过符号链接、junction 或其他 reparse point。只有全部所选 smoke 与诊断检查通过后才开始保留；验证器不会跟随链接，会完整复制目录，并核对目录/文件清单以及每个文件的精确字节、字节数和 SHA-256。已有目标永不覆盖。复制或复核失败会安全失败并保留权威临时运行目录；部分创建的目标不构成成功证据，也不得复用。

## Source entrypoints / 源码入口

The release-gate source-tree smoke runs the installed `docwen --help` and `docwen-gui` console scripts from `tests/e2e/test_source_tree_entrypoints.py`. It uses `release_gate and (integration or gui_smoke or e2e)` and does not fall back to direct imports.
