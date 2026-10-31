# CLI / 命令行

> This page documents the human-facing command tree plus the bounded local GUI-control surface. External
> content integrations use [`serve --stdio` Machine Protocol v2 and Artifact Bundle v3](specs/machine-protocol-v2.md);
> ordinary `--json` remains a CLI presentation mode. The deliberate exception is `gui open|activate|status`, whose
> protocol-3 JSON envelope is the public local desktop-control boundary and does not require a Machine session.
>
> 本页记录面向人的命令树以及有界的本机 GUI 控制面。外部内容集成使用 [`serve --stdio` Machine Protocol v2
> 与 Artifact Bundle v3](specs/machine-protocol-v2.md)；普通 `--json` 仍只是 CLI 展示模式。明确的例外是
> `gui open|activate|status`：其 protocol 3 JSON 信封是公开的本机桌面控制边界，不要求先建立 Machine 会话。

DocWen source and packaged builds use the same `docwen` command tree. Run `docwen --help`, `docwen <command> --help`, or `docwen schema <command>` for the executable contract.

DocWen 的源码态与打包态使用同一套 `docwen` 命令树。精确契约以 `docwen --help`、`docwen <command> --help` 和 `docwen schema <command>` 为准。

## Commands / 命令

Existing file and directory arguments are interpreted literally before glob expansion, including names containing
square brackets. A rejected batch member contributes to the totals and nonzero exit status in every output mode.

真实存在的文件或目录优先按字面路径处理，包括含方括号的名称。批处理中被拒绝的输入在所有输出模式下均计入汇总和非零退出状态。

- `serve --stdio`: run the Content-Length framed Machine Protocol v2 server for external consumers.
- `info`: report product, protocol, platform and capability metadata without starting the runtime.
- `doctor`: check base runtime health and report the canonical capability projection.
- `inspect FILE`: inspect one input's actual format and supported routes.
- `resources list|show TYPE`: discover formats, optimizations, templates and numbering schemes.
- `schema [COMMAND_PATH...]`: export the active parser contract.
- `convert FILE --to FORMAT --output-dir DIR`: convert one file into the chosen parent; required for conversions from or to Markdown.
- `convert FILE --to FORMAT --output PATH`: exact-file output for other single-file conversions.
- `convert MARKDOWN --to FORMAT --output-dir DIR --proofread [--check CHECK]...`: render DOCX privately, proofread the generated DOCX, then publish the final result. `--check` on `convert` requires `--proofread`; omitting `--check` uses the configured proofreading defaults. If all four checks are disabled, conversion still publishes the selected document with a skipped-proofreading diagnostic. Markdown sources support final DOCX/DOC/ODT/RTF/WPS targets; proofreading runs on the intermediate DOCX before any required bridge conversion. Availability of the selected target still depends on its bridge. Non-Markdown sources do not support conversion proofreading.
- `validate FILE [--report PATH]`: validate DOCX, Markdown, or legacy Word-family content
  (`DOC`/`WPS`/`RTF`/`ODT` is pre-converted to DOCX by the Application layer); the default is read-only.
- `number markdown FILE --operation add|remove [--scheme ID] (--output PATH | --in-place)`: explicitly add or remove Markdown heading numbering.
- `merge pdf|tables|images FILE... --output PATH`: create one exact aggregate output.
- `split pdf FILE --pages RANGE --output-dir DIR`: split selected PDF pages.
- `batch convert|validate FILE...`: run an explicit multi-file operation. Batch validation accepts the
  same DOCX, Markdown, and legacy Word-family content as single-file validation.
- `gui open [FILE]|open-settings --section proofread|activate|status`: stable GUI control surface. `gui open` starts or activates the GUI; an optional absolute `FILE` is opened after startup. `gui open-settings` opens or focuses the non-blocking singleton Settings dialog at the semantic proofread section, starting the GUI only when no instance is running. An already running GUI must advertise `open_settings` and `proofread` through `gui status`; an older running GUI fails with typed `capability_unavailable` instead of launching a second instance or falling back to the main window. `gui activate` only targets an already running GUI.
- `config reset GROUP --yes`: reset one configuration group non-interactively.

For templates, `resources list templates` is the source of truth. Pass its exact `resources[].id` value to
`convert --template`; do not reconstruct an ID from the display name or path. Display names, filenames, paths,
blank values, and differently cased IDs are rejected.

`md` is the only Markdown target identifier. `markdown` is rejected rather than normalized.

There is no `run --action` compatibility entry. Internal runtime action names are not public CLI commands.

不提供 `run --action` 兼容入口；运行时内部 action 名称不是公开 CLI 命令。

## Direct Markdown resource resolution / 直接 Markdown 资源解析

Human-facing CLI and GUI conversion of a raw Markdown file remains independent of any editor integration. A short
local image name is resolved only beside the source file, in the sibling same-name document directory, or under the
configured source-local convention directories (by default `.`, `assets`, `images`, and `attachments`). Convention
directories are relative descendants of the source directory; absolute and traversal search directories are ignored.
DocWen never recursively scans a Workspace, Vault, parent tree, CWD, or drive for a matching basename. Filenames with
spaces are supported. An authored absolute path or an explicit relative path containing directory segments is followed
as an exact locator and is not a filename search. Missing short links follow the configured error policy.

面向人的 CLI 与 GUI 直接转换原始 Markdown 时，不依赖任何编辑器集成。短图片文件名只在源文件同目录、
同名文档目录，以及源目录下配置的约定目录中解析；默认约定目录为 `.`、`assets`、`images` 和
`attachments`。约定目录必须是源目录下的相对后代，绝对路径和含路径穿越的搜索目录会被忽略。
DocWen 不会为了匹配文件名递归扫描 Workspace、Vault、父目录树、CWD 或磁盘。带空格文件名正常支持。
Markdown 明确写出的绝对路径或带目录段的相对路径属于精确定位，不属于文件名搜索；未找到的短链接按配置的错误策略处理。

## Output and safety / 输出与安全

Conversions to Markdown publish a result folder inside the parent selected with `--output-dir DIR`, even for one Markdown file. Markdown-to-document/spreadsheet conversions publish one final file directly in that parent; multiple files (such as worksheet CSVs) or an explicitly requested audit sidecar use a result folder. Names include the source name, timestamp and current input format. Use `--output PATH` for one exact final file, including a single file generated from Markdown; grouped results require `--output-dir`.

生成 Markdown 时始终在 `--output-dir DIR` 指定的父目录下创建结果文件夹。Markdown 转文档或表格只生成一个最终文件时直接输出；多个文件（如多个工作表 CSV），或显式要求审计附带文件时保留结果文件夹。名称包含原文件名、时间戳和本次传入格式。单文件可用 `--output PATH` 指定精确路径，包括由 Markdown 生成的单文件；分组结果使用 `--output-dir`。

- JSON mode emits exactly one protocol 3 envelope conforming to [the JSON schema](specs/json-contracts.schema.json).
- Machine meaning comes from stable fields and typed error codes, not localized messages.
- Write commands require an explicit destination.
- `validate` is read-only unless `--report PATH` is present. A Markdown input writes a structured JSON
  report; DOCX and pre-converted DOC/WPS/RTF/ODT inputs write an annotated DOCX report.
- `convert --proofread` is a different intent: Markdown→DOCX/DOC/ODT/RTF/WPS is accepted, the generated DOCX is proofread before the selected document is published, and a proofreading failure prevents the generated document from being published as a successful result.
- Existing targets are rejected by default; `--overwrite` is required to replace one intentionally.
- Single-file commands never silently rename a requested output.
- On Windows, every public input and output path must use ordinary absolute syntax and be at most 259 UTF-16 code units. DocWen rejects `\\?\` / `\\.\` extended-length syntax before a backend starts because the supported conversion backends do not share one reliable extended-path contract. Runtime-generated artifact names may cross that boundary internally; DocWen keeps reported paths ordinary while adapting only its own filesystem calls.
- Global options are limited to language, JSON/text verbosity and timing; batch, parallelism and confirmation flags belong to their commands.

## Deadlines and proofreading / 截止与校对

A batch deadline or Ctrl+C stops further submission and cancels outstanding Application request IDs before
waiting for workers. JSON keeps one result per input, including successful publications and their paths.
Already committed files remain successful when a later input times out. Work that never started is reported as
`details.execution_state=not_started`; if a worker has not stopped within the bounded cooperative wait, its
state is `cancellation_unconfirmed`. Such a worker retains its cancelled reservation until it exits. This is
not a promise to forcibly terminate Python threads; a non-cooperative thread can still delay interpreter exit.
The command reports timeout exit code 8 or interruption exit code 130 without rewriting prior commit facts.

批处理达到总截止或收到 Ctrl+C 后，先停止续提并取消未完成的 Application 请求，再等待 worker 收束。
JSON 保留每个输入的结果及已成功发布的路径；后续项目超时不会把先前提交的文件改成失败。
尚未启动的输入以 `details.execution_state=not_started` 表示；有界协作等待后仍未退出的 worker 以
`cancellation_unconfirmed` 表示，并保留其已取消的预约身份直到真正退出。这不承诺强杀 Python 线程；
不合作的线程仍可能推迟解释器退出。命令以 8 表示超时、130 表示中断，同时保留各项已提交的事实。

Conversion proofreading accepts Markdown workflow input with final `docx`, `doc`, `odt`, `rtf` or `wps` targets.
The CLI validator, help and `schema convert` / `schema batch convert` target constraints share the Application
support set. `--check none` explicitly disables all checks; omitting it keeps configured defaults. Legacy targets
still require an available bridge. Generic text sniffed as `txt` remains `txt` in the inspection and source identity,
while admission into the Markdown workflow permits the same Markdown proofreading route and report contract.

转换后校对接受 Markdown 工作流输入以及上述五种最终文档格式；CLI 校验、帮助和 schema 的目标准入约束
复用 Application 的同一支持集。`--check none` 显式关闭全部检查，未指定时使用配置默认值；旧文档目标仍须
具备可用 bridge。嗅探为 `txt` 的普通文本保留其真实检测格式和源身份；被准入 Markdown 工作流后可使用
同一 Markdown 校对路由及报告合同。

## Runtime discovery / 运行时发现

`docwen resources list formats --json` is initialized from the loaded plugin composition. It returns a
source-to-route matrix containing conversions and action routes, route availability, dependency/platform
gates and limitations. An initialized composition with no routes is a successful empty result. Failure to
initialize or query the runtime is instead a typed `capability_unavailable` error; consumers must not turn
that failure into an empty list.

The same projection includes `security.dependency_egress_guard`. Supported CLI entry points report
`state=enforced`, `scope=docwen_python_process`, the `bootstrap` layer, the preserved local transports, and
`external_processes=not_managed`. This is the machine-readable boundary; consumers must not infer broader
operating-system or external-Office isolation.

`docwen resources show formats SOURCE --json` selects one source entry from the same projection. This is
the supported discovery surface for consumers; extension tables and hard-coded target lists are not.

`docwen resources list optimizations --json` projects an independent `docwen.optimizations` version 1
contract from the optimization declarations and action routes in the loaded manifests. Each resource has
`id`, `name`, `action_name`, canonical `scopes`, `available`, `state`, and `bindings`. Each binding has exactly
`scope`, `route_id`, `source`, `source_category`, `target`, `available`, and `state`. The public resource `id`
and internal `action_name` are independent values; consumers select by `id` and must use the declared binding
instead of assuming equality or constructing scopes such as `${category}_to_md`.

An initialized runtime with no optimization declarations returns the same valid contract with
`resources=[]` and zero counts. A missing, malformed, or contradictory manifest/runtime projection is a typed
`capability_unavailable` failure, never a successful empty list or a static fallback. `resources show
optimizations ID --json` selects one resource from this same projection.

The optimization inventory is capability-first and locale-neutral: `--lang` localizes operator-facing
presentation and supplies execution context, but it never hides a Runtime capability. An English-interface
user can therefore discover and run a Chinese-document optimizer when its route and dependencies are
available. UI locale, document/content language, OCR language, and capability availability are separate
facts; consumers must not reconstruct locale eligibility from manifest `extra` fields.

`docwen doctor --json` consumes that same initialized runtime projection exactly once. Its
`data.capability_summary` exclusively owns dependency gates, routes, availability and limitations.
`data.checks` contains the path-free base-health probes for configuration loading, the temporary directory,
and the active dependency-egress guard; `data.all_ok` summarizes those probes. Missing optional dependencies, unavailable
external Office providers and the routes they limit do not change `data.all_ok` or the process exit code;
consumers read those facts from `data.capability_summary`. A failed configuration or temporary-directory
probe, or a guard that is not enforced, produces `data.all_ok=false` and a non-zero health exit status even though the JSON envelope reports
that the diagnostic command itself executed successfully. Diagnostic output never exposes host configuration
or temporary paths.

`docwen doctor --json` 只读取一次同一套运行时能力投影。`data.capability_summary` 保留规范投影，
并独占依赖门禁、路由、可用性、限制和 `security.dependency_egress_guard` 边界。`data.checks`
包含脱敏后的配置加载、临时目录和第三方依赖出站保护三项基础检查，`data.all_ok` 汇总这三项。缺少可选依赖、缺少外部 Office 提供者以及因此受限或
不可用的路由，都不会改变 `data.all_ok` 或退出码；消费者应从 `data.capability_summary` 读取这些
事实。配置、临时目录检查失败或出站保护未处于 enforced 状态时，返回 `data.all_ok=false` 和非零健康状态；此时 JSON
envelope 的 `success=true` 仍表示诊断命令本身执行成功，且输出不会暴露本机配置或临时路径。

## Examples / 示例

```powershell
docwen info --json
docwen inspect document.docx --json
docwen resources list formats --json
docwen resources list optimizations --json
docwen schema convert --json
docwen convert document.docx --to md --output-dir exports
docwen convert note.md --to docx --output-dir exports --proofread --check typo --check punct
docwen validate document.docx --check typo --check punct --json
docwen merge pdf part-1.pdf part-2.pdf --output combined.pdf
docwen gui status --json
docwen gui open-settings --section proofread --json --quiet
```
