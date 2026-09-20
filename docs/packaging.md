# Packaging / 打包

DocWen packages one Windows x64 package and two Ubuntu 24.04 x64 packages built from the bundle composition root. The Windows archive contains both GUI and CLI; Ubuntu has GUI+CLI and CLI-only archives. Package verification runs against the produced directory, not the source tree. macOS has no release asset, and its current capability contract reports conversion, validation, numbering, merge, and split as unavailable.

DocWen 的发布目标为一个 Windows x64 完整包和两个 Ubuntu 24.04 x64 包；Windows 包同时包含 GUI 与 CLI，Ubuntu 分为 GUI+CLI 和仅 CLI 两个压缩包。打包验证面向实际产物目录执行，而不是用源码态结果代替。macOS 没有正式附件，当前 capability 契约明确把转换、校对、编号、合并和拆分报告为不可用。

| Platform / 平台 | Distribution status / 发行状态 | Required validation / 所需验证 |
| --- | --- | --- |
| Windows x64 | Release package / 正式发行包 | Exact packaged candidate and automated checks; selected native acceptance recorded separately / 精确打包候选及自动检查；所选原生验收另行记录 |
| Ubuntu 24.04 x64 | Release package / 正式发行包 | Exact manifest-bound package and post-extract automation; native desktop evidence is recorded separately / 精确清单绑定包与解压后自动化；原生桌面证据另行记录 |
| macOS x64/arm64 | No release asset / 无正式附件 | Source CI and opt-in packaging experiment only; primary document operations are unavailable / 仅源码 CI 与手动打包实验，主要文档操作不可用 |

## Required contents / 必需内容

- `DocWen.exe` or `DocWenCLI.exe` and PyInstaller `_internal` content.
- `configs/`, `templates/`, `models/`, locale files and application assets.
- The complete `pymupdf-layout` distribution resource manifest under `_internal/pymupdf/layout/resources`, including its ONNX models and YAML descriptors.
- License, third-party notices and the supported runtime metadata.

PyInstaller collection targets are import-package names, not necessarily distribution names. In particular, the `pymupdf-layout` distribution installs the `pymupdf.layout` package. The build preflight rejects missing or non-package collection targets, and post-package verification derives the expected Layout resource paths from the installed distribution instead of pinning one model-version filename.

PyInstaller 收集参数使用可导入包名，不一定等于发行包名；`pymupdf-layout` 发行包实际安装为 `pymupdf.layout`。构建前置检查会拒绝不存在或不是包的收集目标；产物检查则从当前已安装发行包清单推导全部 Layout 资源路径，要求 ONNX 模型与 YAML 描述文件完整、非空，而不是只写死某一个模型版本文件名。

`pymupdf-layout` 1.27.2.2 ships the same seven resource paths on every platform covered by the locked resource manifests, but its three YAML descriptors use CRLF bytes in the Windows wheel and LF bytes in the Linux and macOS wheels. DocWen therefore pins two complete raw-byte manifests: one for `win32`, and one shared by `linux` and `darwin`. This resource-integrity coverage is independent of the public package-support boundary. Unknown platforms fail closed with `unsupported_resource_platform`; verification does not normalize line endings, accept alternate hashes, or learn trusted bytes from the installed package. Dependency upgrades must audit every locked platform wheel and update both manifests explicitly.

`pymupdf-layout` 1.27.2.2 在锁定资源清单覆盖的平台上提供相同的七条资源路径，但三个 YAML 描述文件在 Windows wheel 中使用 CRLF 字节，在 Linux 与 macOS wheel 中使用 LF 字节。因此 DocWen 固定两套完整的原始字节清单：`win32` 一套，`linux` 与 `darwin` 共用一套。该资源完整性覆盖与正式发布平台边界相互独立。未知平台以 `unsupported_resource_platform` 失败关闭；校验过程不规范化换行、不接受备选哈希，也不从已安装包自举可信字节。升级依赖时必须审计锁文件中的每个平台 wheel，并显式更新两套清单。

The Windows production builder limits PyInstaller's DLL search path to the clean project environment, the manifest-verified CPython base directory, and the Windows system directories. Host-selected `api-ms-win-*` forwarders and `ucrtbase.dll` are forbidden in the payload; supported Windows supplies those system components. The four packaged MSVC runtime files are instead replaced from the locked `pikepdf` wheel before the actual payload inventory and SHA-256 hashes are recorded.

Windows 生产构建器把 PyInstaller 的 DLL 搜索路径限制为干净项目环境、清单验证过的 CPython 基础目录和 Windows 系统目录。载荷禁止包含由宿主偶然选中的 `api-ms-win-*` 转发 DLL 与 `ucrtbase.dll`，这些系统组件由受支持的 Windows 提供；随包发布的四个 MSVC 运行库文件统一替换为锁定 `pikepdf` wheel 中的副本，再记录本次实际文件清单和 SHA-256。

Windows builds once from the verified source and locked toolchain. The resulting payload manifest records actual file paths, sizes and hashes; it is not a checked-in prediction of future executable bytes. There is no preliminary calibration build or per-version payload-hash update. Required entry points, resource/layout checks, path safety, packaged functionality, candidate provenance and hosted-byte verification remain required.

Windows 从已验证源码和锁定工具链构建一次，随后记录实际文件路径、大小与哈希；成品清单不再作为预先提交的可执行文件哈希预测表。不需要预校准构建或逐版本修改成品哈希。必需入口、资源/布局、路径安全、实际包功能、候选来源和公开字节验证继续执行。

## Release gates / 发布门禁

The Linux production build runs once in the digest-pinned Ubuntu 24.04 container declared in the release workflow and produces both archives. `scripts/release/install_linux_build_dependencies.sh` installs system dependencies from one dated Ubuntu archive snapshot and ignores other package repositories. Update the container digest and archive snapshot deliberately, then repeat the build and both extracted-package checks. The host runner's preinstalled libraries are not production build inputs.

Linux 生产构建在发布工作流声明的固定摘要 Ubuntu 24.04 容器内执行一次，生成两个压缩包。`scripts/release/install_linux_build_dependencies.sh` 从固定日期的 Ubuntu 仓库快照安装系统依赖，并忽略其他软件源。更新容器摘要或仓库快照后，必须重新执行构建及两个压缩包的解压后验证；宿主 runner 预装的库不作为生产构建输入。

Repository settings must enable Immutable Releases and an active no-update/no-delete ruleset for numeric `x.y.z`
tags before a version tag is pushed. 仓库设置必须在推送版本标签前启用 Immutable Releases，并启用禁止更新或
删除数字 `x.y.z` 标签的 ruleset。

The hosted GitHub Release workflow enforces this deterministic baseline before publishing one Windows archive and two Ubuntu 24.04 x64 archives:

1. Shared required source checks (Ruff, test governance, Import-Linter, architecture checks and Pyright for all three targets), one Windows full suite with source/core/GUI coverage reports, and the union of fast and release tests on Linux and macOS. Ordinary CI and release use this same gate.
2. Windows package resource/layout verification.
3. Packaged Windows CLI baseline doctor, conversion, and JSON checks.
4. Packaged Windows GUI settings-page construction smoke.
5. Manifest-bound deterministic Ubuntu archive construction with fixed `DocWen-<version>-linux-x64.tar.gz` and `DocWenCLI-<version>-linux-x64.tar.gz` names.
6. CLI and GUI verification against fresh directories extracted from those exact Ubuntu archives.
7. One build per release platform, transported by the producing job's exact artifact ID with digest mismatch rejection, followed by Release SHA-256 generation over all three archives. Reproducibility comparisons are optional engineering checks, not a prerequisite for every release.

Selected native acceptance uses the exact candidate that will be published. Its scope is chosen for each release and can include: source/package CLI parity; OCR and successful-warning CLI paths; GUI startup, settings, notification/OCR, IPC, and successful-warning paths; plus Office, presentation, and SmartDoc routes on a machine where their external dependencies are actually configured. These environment-sensitive checks are recorded separately and are not falsely attributed to the hosted workflow. Visible notification presentation, target-device rendering, and manual UI inspection remain manual evidence.

托管的 GitHub Release 工作流在发布一个 Windows 压缩包和两个 Ubuntu 24.04 x64 压缩包前执行以下确定性基线：

1. 先运行统一必需源码检查（Ruff、测试治理、Import-Linter、架构检查及三个目标平台的 Pyright）；Windows 完整测试只运行一次并生成源码/core/GUI 覆盖率报告，Linux 与 macOS 各运行 fast 和 release 测试的并集。普通 CI 与发布使用同一门禁。
2. 验证 Windows 打包资源与 Layout 资源。
3. 验证 Windows 打包 CLI 的基础 doctor、转换和 JSON 路径。
4. 验证 Windows 打包 GUI 的设置页构造。
5. 按受版本控制的生产清单确定性构造 Ubuntu 压缩包，并固定为 `DocWen-<version>-linux-x64.tar.gz` 与 `DocWenCLI-<version>-linux-x64.tar.gz`。
6. 从这两个精确 Ubuntu 压缩包解压到全新目录后，再分别验证 CLI 和 GUI。
7. 每个发布平台只构建一次，按构建任务输出的 artifact ID 交接并拒绝传输摘要不匹配，再为三个压缩包生成 Release SHA-256。可复现性比对按需研究，不作为每版前提。

原生验收使用将要发布的同一精确候选，按本版范围选择场景，可包括：源码/打包 CLI 一致性、CLI OCR 与成功警告、GUI 启动、设置、通知/OCR、IPC 与成功警告，以及在外部依赖已真实配置的机器上验证 Office、演示文稿和 SmartDoc 路径。这些环境相关结果单独记录，不能冒充托管工作流已经执行。通知中心可见性、目标设备渲染和人工 UI 检查仍属于人工证据。

Signing and publication are separate release operations. A successfully built unsigned package must not be described as signed or published.

### Candidate reuse and recovery / 候选复用与恢复

Hosted publishing uses `GITHUB_TOKEN` and does not read administration-only repository settings or require an additional PAT. Publication succeeds only after the hosted Release reports a complete immutable state. One independent read-only verification downloads all four hosted assets and compares their exact bytes and provenance; publishing itself checks the platform inventory, sizes and digests without repeating those downloads.

Push the numeric version tag at the accepted source commit on the default branch. This one `Release` run invokes the shared ordinary CI gate (including declared release cases), builds each platform once, verifies and attests the packages, publishes those exact bytes, then performs independent readback. The version comes from `pyproject.toml`; a selected numeric tag must match it. Failed, missing, cancelled or skipped required source/build jobs stop publication; configure `Required checks` as a required branch check. The candidate is identified by the producer job's artifact ID and archive digest, with no search of earlier runs. Artifacts are retained for 30 days.

The `publish` job uses the `release` environment and `GITHUB_TOKEN`. It verifies the candidate's exact run attempt, passed producer jobs, source, inventory and provenance, and confirms that the numeric tag points to that exact source commit. The overall run can still be active, or have failed only during a previous publication attempt. `scripts/release/publication.py` creates a draft, pins every draft asset by size and platform digest, publishes and confirms the complete immutable hosted state. The separate read-only `post-verify` job downloads and verifies all hosted bytes once. `published-awaiting-readback` distinguishes publication from the final `verified` receipt. `candidate.json` is a control record and is not a public Release asset.

When native acceptance must precede merging and publication, select the candidate branch in `Run workflow` and use `operation=verify` with no artifact ID. This runs the same source and package checks and creates an attested candidate without creating a tag or Release. Record its exact artifact ID and inspect it with `publication.py fetch` followed by `publication.py inspect`; the inspection receipt binds its bytes and provenance. After acceptance and merging, select a branch or numeric tag still pointing to that same source commit and run `operation=publish` with that artifact ID. The publisher requires the candidate commit to be on the default branch, an ancestor of it, or have the identical complete Git tree after a squash/rebase merge. It creates the numeric tag at the actual candidate commit, records the accepted default-branch commit and comparison basis, and never relabels or rebuilds the candidate. The manifest records the original build's branch/tag in `origin.sourceRef`, which remains the provenance source ref.

For independent hosted verification, use `operation=verify` with the candidate `artifact_id` and a selected ref pointing to its source commit. For an interrupted publication, use `operation=publish`, that same candidate ID and the saved `docwen-publication-progress-RUN-ATTEMPT` ID in `resume_artifact_id`. The IDs retrieve their transport digests from GitHub; there is no manual version, run lookup or digest copying. The previous progress owner must have completed. Uncertain tag, draft, upload or publish writes are reconciled by reading before further action. Exact published releases are read-only; matching drafts upload only missing assets. Do not rerun build jobs to recover publication. A manual `publish` without an artifact ID performs the normal build chain from the selected ref and still requires default-branch acceptance before publication.

Transient reads use bounded backoff and honor Retry-After; the default 60-second recovery budget and 600-second upload/artifact-download ceilings can be raised for slow transports with `DOCWEN_PUBLICATION_READ_BUDGET`, `DOCWEN_PUBLICATION_DATA_TIMEOUT` and `DOCWEN_PUBLICATION_ARTIFACT_TIMEOUT`. Permission, source, inventory and hash conflicts stop immediately.

在默认分支已接纳的提交推送数字版本标签，同一次运行完成统一 CI、每平台一次构建、候选签证、immutable 发布和一次独立资产下载回验。需先做原生验收时，在候选分支手动运行 `verify` 且不填 artifact ID，只构建候选，不写标签或 Release；用 `fetch` 和 `inspect` 生成身份回执，验收并合并后，以仍指向该候选源码的分支/标签运行 `publish` 并填写同一候选 ID。发布前确认源码已被默认分支接纳：相同提交、祖先提交，或 squash/rebase 后完整 Git tree 相同。标签始终指向实际构建源码，原始 `sourceRef` 保留用于来源验证，不能把旧产物冒充新提交构建。

手动 `verify` 填候选 ID 是公开 Release 的只读回验。发布中断后用 `publish`、同一候选 ID 和原进度 `resume_artifact_id` 恢复；原所有者必须结束，结果不明的写入先读回核实，精确已发布资产不再写入，匹配草稿只补缺项。真实宿主验收按每版变更范围选择，单独记录证据，不增加永久人工审批门。

### Independent MSIX channel / 独立 MSIX 渠道

`MSIX candidate` is a separate manual workflow. Select a ref at the portable candidate's source commit and supply its exact `artifact_id`. The workflow fetches and inspects that candidate, then builds the unsigned MSIX directly from its Windows ZIP; it does not rebuild DocWen or gate the portable Release. MSIX metadata binds the portable archive digest, candidate artifact, source commit and provenance receipt. The Store config's source version must match; its package version follows the independent Store channel. Signing, installation acceptance and Store submission are separate operations.

`build_msix.py` requires a new work directory, creates an ownership lease and removes that directory after success. It never resets an existing caller directory or overwrites an output package. Local raw work belongs below the governed workspace `temp`; failed or interrupted runs follow bounded retention.

`MSIX candidate` 为独立手动流程：选择 portable 候选源码提交对应的分支/标签，填写候选 artifact ID；获取并核验后直接从 Windows ZIP 制作未签名 MSIX，不重新构建产品，也不阻挡 portable 发布。元数据记录 portable 摘要、候选 artifact、源码和来源回执；Store 配置的源码版本必须匹配，包版本按 Store 独立管理。签名、安装验收与商店提交另行执行。工作目录必须全新且带租约，成功清理；不删除调用者原目录，也不覆盖已有包。

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

`verify_packaged_gui.py` places local raw runs below the governed workspace `temp` tree and writes an ownership lease. It removes the raw run after success; failure, interruption, or cleanup failure enters the shared bounded retention policy. Pass `--candidate-id ID --receipt-output <ABSOLUTE_JSON_BELOW_WORKSPACE_ACCEPTANCE>` to close a successful run into a compact receipt containing candidate identity, selected gates, binary hash, run-manifest hash, result, and limitations. Pass `--evidence-dir <ABSOLUTE_NEW_DIRECTORY>` only when the complete successful-run files are explicitly required for review. The full-evidence destination remains fail-closed and byte-verified; it is an intentional retained artifact, not default acceptance state.

`verify_packaged_gui.py` 会把本地原始运行放入受治理的 workspace `temp` 并写入所有权租约。成功后删除原始目录；失败、中断或清理失败进入共享的有界保留策略。使用 `--candidate-id ID --receipt-output <位于 workspace/acceptance 下的绝对 JSON 路径>` 可把成功运行关闭为包含候选身份、所选门禁、二进制哈希、运行清单哈希、结果和限制的紧凑回执。只有明确需要完整现场供复核时才传入 `--evidence-dir <绝对且不存在的目录>`；完整现场属于显式保留物，不是默认验收状态。

## Source entrypoints / 源码入口

The release-gate source-tree smoke runs the installed `docwen --help` and `docwen-gui` console scripts from `tests/e2e/test_source_tree_entrypoints.py`. It uses `release_gate and (integration or gui_smoke or e2e)` and does not fall back to direct imports.
