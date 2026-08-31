# Testing / 测试

Tests are grouped by primary behavior family and execution cost. The default repository run selects only non-slow unit and contract tests; GUI, integration, end-to-end and environment-owned checks run in their explicit lanes.

测试按主要行为族和执行成本分层。默认仓库测试只选择非 slow 的 unit 与 contract；GUI、integration、端到端及环境所有型检查进入各自显式门禁。

## Common commands / 常用命令

```powershell
.\.venv\Scripts\python.exe -m pytest
.\.venv\Scripts\python.exe -m pytest -m "release_gate"
.\.venv\Scripts\python.exe -m ruff check .
.\.venv\Scripts\python.exe -m ruff format --check .
.\.venv\Scripts\pyright.exe
.\.venv\Scripts\python.exe tools\run_import_linter.py
```

## Markers / 标记

- `unit`, `contract`: deterministic logic and public contracts.
- `integration`: real files, external libraries or cross-module chains.
- `gui`, `gui_smoke`: widget behavior and real GUI startup routes.
- `e2e`: user-entry-to-output execution.
- `packaged`, `host`: frozen-package and real-host evidence; source doubles cannot satisfy them.
- `office`, `linux_only`: Office-family and Linux-only execution boundaries.
- `pr_gate`, `release_gate`: explicitly selected high-value gates.
- `golden`: stable semantic or artifact comparison.
- `slow`, `windows_only`, `macos_only`: cost and platform boundaries.

Governed `tools/qa.py` runs place basetemp, cache, structured reports, and process TEMP roots physically below `<workspace>/.workspace/temp/p<random>`. The short physical name is deliberate: security-sensitive code canonicalizes aliases, so path-length headroom must come from the real workspace path instead of changing pytest fixture semantics. Per-test directories use a stable `t<12-hex-digest><counter>` name instead of pytest's longer test-name prefix; this preserves isolation and physical identity while leaving headroom for xdist's worker segment and adversarial output names. QA owns each unique directory, removes it after success, and retains it with a JSON lease after failure or interruption. Workspace selection uses `--workspace-root` first, then `DOCWEN_WORKSPACE_ROOT`, then the single `.workspace` path derived only from the governed `<engineering>/repos/docwen` repository layout; the removed direct `<engineering>/docwen` layout and arbitrary ancestor scanning both fail closed. A locally configured `DOCWEN_PYTEST_RUNTIME_ROOT` / `--pytest-runtime-root` must remain below the governed workspace's `temp` tree; repository-internal roots are rejected. `--own-pytest-runtime` may take ownership of a new explicit workspace-contained root, and `--keep-pytest-runtime` preserves an owned successful run. On Windows, the full suite also maps the physical runtime root to a verified temporary drive letter for I/O endpoints that can safely retain the alias. Pytest is allowed to canonicalize `tmp_path` back to the physical workspace path, preserving the same path semantics used by production code. The mapping is removed in `finally`; the lease records it while active so interrupted work can be recovered safely. Each QA start uses the same saved-plan retention policy as the operator tool: terminal-success scratch is immediately eligible, failed or interrupted scratch keeps the newest two per kind for at most 72 hours, manual and non-terminal states remain retained, and live owners always block deletion. CI without a governed local workspace may still supply its runner-owned ephemeral root.

Run `python tools/workspace_cleanup.py` to create a non-mutating housekeeping preview for leased scratch roots below `.workspace/temp`, `.workspace/build`, and `.workspace/tmp`. Successful scratch is immediately eligible; failed or interrupted scratch is bounded to 72 hours and the two newest roots per kind. Root-level `temp-*` directories are reported but never enter the default plan. A deletion requires an exact saved plan below `.workspace/diagnostics`: generate it with `--plan-output`, review it, then pass that same file to a separate `--apply-plan` invocation. Explicit unleased candidates require one or more `--target` arguments plus `--reason`; repository `.venv` and `node_modules` directories require `--clean-deps`. Apply revalidates the allowed root, complete identity and content fingerprints, leases and owner PIDs, and reparse boundaries before any deletion. Direct ad-hoc pytest remains possible, but it is not a release-cleanliness gate. A green release claim requires explaining all skips and proving that required dependencies were collected.

受治理的 QA 会把 basetemp、缓存、结构化报告和进程 TEMP 的真实文件统一放在 `<workspace>/.workspace/temp/p<随机名>`。短物理名称是刻意设计：安全敏感代码会把别名解析回真实路径，因此路径余量必须来自工作区内的真实短路径，不能靠篡改 pytest 临时夹具的路径语义。每项测试的目录名使用稳定的 `t<12 位十六进制摘要><计数器>`，替代 pytest 较长的测试名称前缀；这样仍保持逐测试隔离和真实物理身份，同时为 xdist 工作进程层级和对抗性输出名留出余量。成功自动删除，失败或中断通过 JSON 租约留证。工作区选择依次采用 `--workspace-root`、`DOCWEN_WORKSPACE_ROOT`，最后只根据受治理的 `<工程>/repos/docwen` 布局推导唯一 `.workspace`；已移除的 `<工程>/docwen` 直连布局和任意祖先扫描都失败关闭。本地显式运行根也必须位于该工作区的 `temp` 树内；`--own-pytest-runtime` 只能接管全新、位于工作区内的显式根。Windows 完整套件运行期间仍会把这个物理目录临时映射为经过身份校验的短盘符，供能够安全保留别名的 I/O 端点使用；pytest 的 `tmp_path` 则允许规范化回工作区内物理路径，与生产代码保持同一语义。测试结束时在 `finally` 中解除映射，不会在盘符根目录创建测试文件夹。租约在映射期间记录短盘符，使异常中断后的回收仍能验证目标身份。每次 QA 启动都采用与运维工具相同的保存计划保留策略：成功终态 scratch 立即可回收；失败或中断按 kind 保留最新两份且最长 72 小时；人工保留和非终态继续保留；活动所有者始终阻止删除。`tools/workspace_cleanup.py` 默认只生成不改变工作区的计划；需要删除时，必须先用 `--plan-output` 把计划保存到 `.workspace/diagnostics` 并完成审阅，再在单独命令中用 `--apply-plan` 消费同一文件。根级 `temp-*` 只报告、不进入默认计划；无租约目标需要 `--target` 与 `--reason`，仓库内 `.venv`/`node_modules` 只能通过 `--clean-deps` 显式纳入。apply 会在任何删除前重验允许根、NTFS/文件系统身份、完整元数据与内容指纹、租约/PID 和 reparse 边界。没有本地治理工作区的 CI 可继续使用 runner 自有的临时根。

The default selection expression is `(unit or contract) and not slow`. GUI, integration, end-to-end and other environment-owning primary families run in their owning lanes instead of entering the fast lane by omission. PR and release expansion use `pr_gate and (integration or gui_smoke or e2e)` and `release_gate and (integration or gui_smoke or e2e)`. The shared pytest options are `-v --tb=short --strict-markers --import-mode=importlib -ra`.

CI owns the exact coverage command, creates its external `docwen-pytest-runtime` parent before pytest initializes `--basetemp`, uses the governed `pytest-xdist`/`loadfile` lane without changing selection, and keeps `fail_under = 81` aligned with `pyproject.toml`. On pull requests, that coverage run is also the Windows fast gate; CI does not execute the same Windows selection a second time, and i18n is covered by the cross-platform fast lanes instead of a duplicate standalone job. The threshold is the conservative floor below three exact-tree whole-product samples (81.7468%, 81.7347%, and 81.7369%), not the rounded line-only rate. The stable governance entry points are:

```text
python tools/qa.py --skip-ruff --skip-pyright --suite pr-integration
python tools/qa.py --skip-ruff --skip-pyright --suite release
python tools/check_coverage_source_manifest.py coverage.xml
python tools/check_core_coverage.py coverage.xml --soft-gate
python tools/check_gui_coverage.py coverage-gui.xml
python tools/check_test_governance_consistency.py
python tools/run_import_linter.py
```

`[tool.coverage.run].source` in `pyproject.toml` is the only whole-product coverage source manifest. The coverage job uses bare `--cov --cov-config=pyproject.toml`; every real package under `packages/**/src/` must appear once in that manifest and in `coverage.xml`. Missing XML, no measured files, an unreported configured package, `module-not-imported`, and `no-data-collected` are hard failures rather than an empty green report. Dedicated GUI coverage may still select `docwen_gui` explicitly because it is a separate domain report, not a second whole-product source manifest.

Parallel execution uses `pytest-xdist` and `loadfile` by default. Governed `auto`/`logical` runs are capped at six workers because the suite contains nested CLI and pytest subprocess probes; `DOCWEN_PYTEST_XDIST_WORKERS` selects another worker mode or exact count, `DOCWEN_PYTEST_XDIST_MAX_PROCESSES` changes the automatic cap, and the `DOCWEN_PYTEST_XDIST` switch explicitly disables parallelism when set to `0`. This avoids oversubscribing a host without making ordinary local QA serial by default. Visibility files are `skip_report.json`, `not_collected_report.json`, `slow_report.json`, `subprocess_report.json` and `missing_marker_report.json`. Fixture ownership is documented in `tests/fixtures/README.md`; stable goldens live under `tests/fixtures/golden/`. Any changed fixture requires 人工审 diff.

Primary-marker debt is closed rather than hidden by mass edits. Existing `gui` tests count as primary GUI tests; governed QA requires every collected node to have exactly one primary marker, with zero missing and zero overlapping classifications. Marker-only migrations preserve node IDs, skip semantics, and failure semantics. Core, runtime, application, CLI, GUI, bundle, document, image, spreadsheet, proofread, Markdown, and markup behavior families are explicitly classified; filesystem, process, lock, or cross-module paths are `integration`, public wire and semantic boundaries are `contract`, widgets are `gui`, and pure state is `unit`. Real OOXML used to verify a stable conversion contract remains `contract`; it is not demoted merely because a fixture is serialized to disk. Runtime classification deliberately moved file-, lock-, or cross-module tests from the default fast selection into the complete suite. Application/CLI/bundle classification likewise reserves real process and destination paths for higher-level lanes, while stable semantic contracts remain in the fast lane. The AST guard discovers every configured `testpaths` root and recursively audits class methods; new missing or overlapping primary classifications fail immediately. Collection size is recorded per exact tree and is never treated as a permanent constant or reduced to make a gate green.

Test-file size governance reads the same configured `testpaths` roots and rejects every test module above 700 lines. There is no oversized-file baseline or exception list. Shared fixtures and non-test helpers live in focused support modules, while collected test modules remain independently reviewable. Legacy monolith path names are collection errors rather than ignored paths, so a reverted large test cannot disappear from CI.

Test subprocesses should use `tests.support.subprocess_runner.run_subprocess`, whose default timeout is 60 seconds and whose structured record distinguishes timeout from child exit. A test may choose a different positive timeout, but disabling the timeout with `None` or a non-positive value is rejected. Long-lived protocol sessions that require `Popen` must own an explicit bounded wait/terminate/kill/reap lifecycle and remain in a serial marker family until their process isolation is proved.

Pyright includes all source and test modules under `packages`; a global `**/tests/**` ignore or exclude is not an acceptable substitute for fixing package-test errors.

Governed QA passes its current Python executable to Pyright explicitly. This keeps the type gate reproducible when the project environment is stored under `.workspace/tools` and avoids requiring a repository-local `.venv` junction or copy.

The packaged successful-warning gate uses `--successful-warning-smoke <input>` or the pinned generated fixture with `--successful-warning-smoke`. The GUI verifier first resolves the unique canonical warning through the packaged CLI from the same candidate and rejects empty outputs, then requires the exact warning row/tooltip, persistent warning-tone task summary, warning-row PNG and byte-identical CLI/GUI artifacts. Evidence must preserve the same JSON structured warning and text stderr warning, and prove final artifact 字节一致. Explicit fixtures should also pass `--successful-warning-input-sha256`; `--successful-warning-message` is an additional golden, not a replacement for CLI truth. This gate 不得外推为通知中心可见性.

Add `--proofread-report-smoke` to the packaged CLI verifier when checking the proofread consumer contract. The verifier owns the BOM/CRLF/Unicode fixture, checks `docwen.proofread_report.v2`, raw-byte identity and authoritative ranges, and separately requires a successful empty report from `--check none`. 任何非零退出或 `success=false` 都是 failure，不能折叠成成功空报告。该门不替代 Obsidian 侧 UTF-16 坐标转换、过期结果或实际修复验收。

## Packaged GUI smoke / GUI 打包冒烟

The default packaged GUI smoke uses `DOCWEN_GUI_TEST_AUTOCLOSE_MS`. Add `--notification-smoke`, `--ocr-smoke` or `--ipc-smoke` only for the boundary being verified. With no GUI already running, the IPC gate additionally verifies the exact packaged `info` declaration for `gui.settings`, direct CLI cold start into the proofread section, the runtime handshake, singleton dialog reuse, and same-primary file delivery. Expiry-without-late-side-effect and explicit deadlines longer than 15 seconds remain source-level adversarial tests rather than package claims. The matching limitations are documented in [Packaging](packaging.md).

默认打包 GUI smoke 使用 `DOCWEN_GUI_TEST_AUTOCLOSE_MS`。仅在需要验证对应边界时增加 `--notification-smoke`、`--ocr-smoke` 或 `--ipc-smoke`。在没有 GUI 预先运行的前提下，IPC 门还会精确验证打包 `info` 中的 `gui.settings` 声明、由 CLI 直接冷启动到校对设置页、运行时握手、单例窗口复用和同一主实例文件投递。已过期请求无迟到副作用和显式超过 15 秒 deadline 仍是源码级对抗测试，不冒充包级证据。各模式不能证明的范围见[打包文档](packaging.md)。

Successful packaged-GUI runs are ephemeral by default. Use `--evidence-dir <ABSOLUTE_NEW_DIRECTORY>` only for an explicitly retained run. The verifier rejects existing, relative, aliased, linked/reparse, in-package, or in-run destinations; copies only after all selected gates pass; and rechecks the complete tree plus exact bytes, byte counts, and SHA-256 values. Copy failure retains the authoritative temporary source and fails the gate rather than silently reporting success.

打包 GUI 成功运行默认不保留临时目录。只有明确需要保存本轮证据时才使用 `--evidence-dir <绝对且不存在的目录>`。验证器拒绝已有、相对、别名、链接/reparse、包内或本轮运行目录内的目标；仅在全部所选门禁通过后复制，并重新核对完整目录树、精确字节、字节数和 SHA-256。复制失败会保留权威临时源并使门禁失败，绝不静默报告成功。
