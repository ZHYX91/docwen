# Testing / 测试

Engineering ownership belongs to the caller. `tools/workspace_root.py --init <parent>/.workspace
--repository <checkout>` registers a local checkout without private planning files. Local QA and
`tools/run_source.py` select that root explicitly or via `DOCWEN_WORKSPACE_ROOT`; linked worktrees
can use their primary checkout's governed root. Product processes receive only `DOCWEN_RUNTIME_ROOT`,
with no repository-layout or README dependency. CI uses its existing explicitly owned runtime.
Workspace repository cleanup reads `workspace.json`; legacy roots without registration never grant
cleanup access to neighboring repositories. Registration does not bypass leases, Git tracking or links.

Local QA and CI use `tools/qa.py` for the same mandatory source checks: whole-repository Ruff,
test governance, all Import-Linter dependency contracts, architecture cleanliness, and Pyright
for Windows, Linux and macOS. `python tools/qa.py --skip-pytest` runs these checks before costly
tests or builds. The release workflow reuses this same `tests.yml` gate. Windows runs the full suite once with coverage;
Linux and macOS each run one `platform` selection combining fast contracts and declared release cases.
The full Windows run already includes GUI, PR and release cases; no second coverage or GUI test run is needed.
A passing package preflight alone is insufficient.

Tests are grouped by primary behavior family and execution cost. The default repository run selects only non-slow unit and contract tests; GUI, integration, end-to-end and environment-owned checks run in their explicit lanes.

Focused local runs may pass test paths through `PYTEST_ADDOPTS`. The QA entry point pins the repository's pytest configuration and root fixtures even when every selected path belongs to a single workspace package.

Before successful owned-runtime cleanup, QA captures the existing root and lease identity and pins the directory
object through cleanup and failure recording, without blocking Windows delete sharing. If recursive cleanup
partially removes the marker and fails, the shared run-lease writer may exclusively restore that same known lease
only while the root identity still matches. A replaced marker/root is never adopted. Recording failures keep the
original cleanup error visible, attempt the existing external recovery hold, and still return the failed gate (2);
they do not reset retention deadlines, repeat deletion, or turn cleanup failure into a warning pass.

QA 清理成功运行前捕获现有根目录和租约身份，并在清理及失败记录期间持有目录对象，保留 Windows 的
删除共享。递归删除部分移除 marker 后失败时，共享租约 writer 仅在
根身份保持一致时排他恢复已知租约；不接管被替换的 marker／根目录。记录失败保留原始清理错误并尝试
既有外部 recovery hold，门禁仍返回失败 2；不重置保留期限、不再次强删、不降为警告通过。

Failure runtimes are subject to the bounded housekeeping policy on subsequent QA invocations, even when a
keep-runtime request accompanied the failing run. Export the compact reports and capture specifically needed
original samples with run/lease/path/hash identities before starting another retention scan. Then close owned
raw runs through the existing saved-plan/acceptance-closeout entry. A later rerun is new evidence, never a
replacement for a missing earlier payload; an empty reused path is not authority to reconstruct or delete it.

失败 runtime 在后续 QA 启动时仍受有界 housekeeping 策略管理，包括带 keep-runtime 请求的失败运行。
下一次保留扫描前应导出紧凑报告，并以 run／lease／路径／hash 身份留存确需的原始样本，再使用既有
保存计划或 acceptance-closeout 入口收口有主原始目录。重跑属于新证据，不能替代已缺失的旧 payload；
同路径下后来出现的空目录不构成重建租约或删除授权。

`test_config_durable_transaction_recovery.py` exercises actual journal replacement followed by a failed
directory barrier, real first-file compensation and abrupt child-process exit, then recovery in a new process.
It checks exact bytes, modes, logical symlink identity, trusted state and retained journal ownership; a failed
first recovery is retried independently. `test_config_transaction_safety.py` covers failure to make rollback
`PREPARED` durable before or after replacement. `test_template_manager.py` applies the same crash boundary to
file-plus-catalog identity, and checks committed cleanup warnings using saved/reopened DOCX inputs, final bytes,
IDs, order and defaults. `test_template_management_ui.py` checks actual Qt import counts and separately visible
cleanup warnings. Directory-barrier and trash-service failures are injected boundaries; they do not establish
power-loss behavior or native desktop recycle-bin behavior. Linux/macOS real-symlink source regressions and
Windows symlink-privilege skips remain separate from packaged native acceptance.

上述配置恢复用例执行真实 journal replace、注入目录屏障异常、落盘第一份补偿后以 `os._exit` 终止子进程，
再由新进程恢复，独立检查 bytes、mode、逻辑符号链接、trust 与日志；首次恢复失败另行重试。模板回归把
同一边界应用于文件与目录身份，并保存重开真实 DOCX 后核最终内容、ID、顺序与默认状态；Qt 回归核导入
成功／失败计数与独立清理警告。这些用例的屏障故障和回收站服务故障属于注入边界，不等于断电或原生
桌面回收站验收。真实符号链接与 Windows 权限 skip、源码 Qt 与最终原生包验收分别记录。

`test_markdown_full_source_ownership.py` and `test_source_ownership_consumers.py` assert exact non-heading source
preservation, real embedded section/block content and link removal across container exits. The proofreading report
contract adds visible-text positives and literal-code negatives with independently calculated source ranges and
input byte hashes. These tests include front matter, HTML/comments, unequal fence lengths, multiple backticks,
container scopes, BOM and CRLF; they do not use a second invocation of the implementation to build the expected text.

源码所有权回归独立断言非标题范围保持不变、实际嵌入正文与容器退出后的链接策略。校对同时覆盖应报告的可见
文字及不应报告的代码，按输入直接计算坐标和字节摘要，避免以同一实现生成期望。

`test_markdown_block_lifecycle.py` and `test_heading_block_numbering_contract.py` check paragraph eligibility,
ordered-list interruption versus established siblings, lazy/container literal boundaries, complete Setext titles,
empty ATX syntax and visible autolink heading selection. `test_markdown_block_scope.py` checks the corresponding
proofreading ranges and fixes against authored offsets, with BOM and LF/CRLF variants. Literal continuations and
actual container exits are paired; a dedented line alone is not assumed to end a list-owned literal.

`test_heading_inventory_alignment.py` compares the direct semantic plan, actual parser and visible numbering for
the same ordered heading inventory, including Setext, empty headings and disabled formatting templates. It checks
enabled/disabled H7–H9 inside quotes and lists. `test_markdown_block_heading_contract.py` follows real file
conversion through the writer, then independently inspects native OOXML heading styles, exact body/code text,
REF instructions and displayed caches before using the reader. It checks source-byte preservation and both
addressable and unaddressed extended headings. Parser agreement or a same-pipeline round trip alone does not
establish those final-content expectations; these generated-file checks do not replace installed Office tests.

`test_docx_heading_reader_routing.py` separately writes source-native list headings and resolved heading-only
packages with real semantic sessions. Independent numbering XML and target-map assertions precede reader checks;
ordinary list IDs cannot identify a protocol version, and missing or damaged resolved evidence remains a refusal.
Heading-only recovery does not require unrelated caption bindings. The file conversion lane also covers ordinary
H1/H6 list headings, so enabling H7–H9 cannot hide the shared reader boundary behind a level-specific exception.

`test_markdown_bom_projection.py` distinguishes a complete source's encoding signature from a fragment's leading
U+FEFF. YAML hyperlink labels and note/document OPC parts retain that literal character; the file conversion tests
also retain a body-leading U+FEFF after YAML, both with and without a separate source signature. Container-math
regressions preserve native OMML and ordinary source-anchor metadata without promoting nested direct-profile
caption declarations or creating new Word bookmarks for ordinary anchors.
List continuation regressions separately inspect the physical ordinary-anchor carrier: after a table, code block,
formula, heading or earlier paragraph, the later paragraph's ID owns exactly its own paragraph. The leading
paragraph's existing complete-item ownership has a separate positive control; carrier proof is not weakened.
Real package checks also pair a complete-item anchor with a later paragraph anchor, including an empty list marker;
the authenticated nested SDTs must retain both identities and the exact parent-child relationship.

块生命周期与编号回归分别检查段落资格、有序列表中断与同级项、lazy／字面块的容器边界、完整 Setext、
空 ATX 语法及可见自动链接标题。校对回归按原文独立计算 BOM、LF／CRLF 坐标与替换值，将仍属字面块的
续行与真实退出配对，不仅凭去缩进判定退出。语义回归联合检查真实标题库存、可见编号及引用缓存，包括
空格式模板和容器内 H7–H9 开关。文件集成在 reader 之前独立检查 OOXML 标题样式、完整正文／代码、REF
指令与显示缓存，保留源文件字节；解析器自洽和同管线往返都不能单独替代这些内容断言或真实 Office 验收。
reader 分流回归另用真实语义 session 写出普通列表标题与纯标题 resolved 包，先独立检查编号 XML 和目标
map，再验证读取；不把列表编号 ID 当版本标志，缺失或损坏的 resolved 证据仍拒绝。文件路线同时覆盖
H1／H6 列表标题，避免只对 H7–H9 特判。
完整源 BOM 与片段首 U+FEFF 分别测试：YAML 链接 label、注释／正文 OPC part 和 YAML 之后的正文仍保留
字面字符。容器公式检查原生 OMML 与普通源锚点恢复，不提升 direct 路线的嵌套题注，也不新增普通锚点书签。
列表续段另独立检查普通锚点的实际 carrier：在表格、代码、公式、标题或前段之后，后续段落 ID 只包围自身
段落；首段原有完整列表项所有权保留单独正控，不放宽 carrier 证明。
真实文件还将整项锚点与后续段落锚点配对，覆盖空列表标记；认证后的嵌套 SDT 必须保留两个身份与精确父子关系。

`test_task_anchor_ownership.py` and `test_task_anchor_file_contract.py` pair ordinary, checked and unchecked items
with later paragraph IDs. Actual packages retain task state, every child block and the distinct nested carriers;
duplicate lead markers remain refusals. `test_docx_list_item_carrier_shape.py` accepts a leading paragraph followed
by native tables/paragraphs, while rejecting table-first shapes, unproven wrappers, fields/bookmarks and crossing.
`test_runtime_quote_owner_paths.py` preserves equal-range ancestor depth and rejects terminal self, invalid kinds,
crossing and ambiguous parents. `test_nested_quote_anchor_packages.py` and `test_quote_anchor_recovery_content.py`
inspect native content, maps and physical parent-child edges before fresh reading, including quoted IDs, actual
line breaks versus literal code, and IDs before page/section separators. `test_docx_ordinary_list_topology.py` uses
real v3 sessions and fresh packages to pair recursively malformed referenced numbering with equivalent unused
definitions; valid start-zero, decimal/bullet and low/high ID controls retain exact content and target identity.

上述任务项回归同时检查普通项、已完成／未完成状态、后续段落 ID、完整子块及嵌套载体；重复首段标记仍拒绝。
列表项载体以段落开头并保留表格／段落续块，表格起始、未认证包装、字段／书签及 crossing 分别有负控。
引用回归保留同界祖先的深度，拒绝 terminal self、非法种类、crossing 与歧义父节点；真实包先独立检查内容、
map 和物理父子边，再读取 ID、真实换行与代码字面值，并核对 ID 在分页／分节符之前。编号分流以实际 v3
session 和 fresh package 检查递归坏子树及未使用定义对照，合法零起点、decimal／bullet、低／高编号 ID
保留正文与目标身份。以上均不替代真实 Office 或最终候选验收。

`test_runtime_heading_dialect.py` checks the production source adapter before semantic markers are inserted:
disabled/enabled H7–H9 boundaries, caption and anchor examples inside complete code spans, the following real
caption's independent range and counter, and valid/invalid fenced carriers. Frozen-oracle behavior remains a
separate assertion. The visible-content integration tests independently inspect the resulting DOCX text before
using the reader; equivalent fenced output is accepted only with the exact complete code body.

源适配层回归在插入语义标记前检查实际输入方言、完整代码内的伪题注和伪锚点，以及后续真题注的坐标与计数；
合法围栏载体和非法结束行均有独立断言。冻结 oracle 保留单独对照。集成回归先检查实际 DOCX 正文，再检查
reader 的完整代码正文，不以等价的行内／围栏拼写差异误判内容丢失。

`test_machine_input_role_order.py` drives the real Application/Runtime/Bundle path with source and resource handles
in different positions, then independently checks published DOCX text, image bytes and Bundle hashes. Invalid
roles, kinds and cardinalities still publish nothing. `test_markdown_visible_content.py` checks native package
contents before invoking the real reader: default/disabled/enabled nested pipe tables, rich inline containers,
task-list children and each drawing's own image relationship and description. Explicit `markdown_embed` output is
used when asserting recovered image alt text; the default Wiki image policy does not promise an alias.

`test_resolved_carrier_roundtrip.py` separates successful writing, independent ZIP/XML inventory and real reading.
Its legal nested carriers and standalone captions are paired with unknown-parent, duplicate, misplaced, missing-map
and altered-field refusals. `test_docx_resolved_binding_boundaries.py` rejects missing source-carrier bindings and
post-bind object mutations while preserving an earlier output. Bookmark-allocation regressions measure real
inventory calls alongside exact native heading text and globally unique IDs; call counts alone are not the content
oracle. Request-local allocations and controlled table writers share reservations, while completed/reopened package
proofs rebuild their physical inventory and reject late name/ID/orphan collisions. These generated-file checks do
not establish Office save/reopen, Obsidian integration, platform process behavior or final-candidate acceptance.

上述回归分别核对真实 ConversionService→Runtime→Bundle 路径、独立 ZIP/XML 内容及真实 reader；无效角色和载体仍拒绝。
书签成本测试同时断言真实内容及全局唯一身份，不以替身调用次数替代产物正确性。请求内分配与受控表格写入者
共享预约，但完成／重开证明仍读取新的全包物理库存。生成文件回归不替代 Office 保存重开、Obsidian、各平台
进程或最终候选验收。

`test_table_merger_text_limits.py` saves and reopens actual XLSX files and exercises Runtime publication
at the final merged-string boundary. `test_epub_spine_integrity.py` uses the real EPUB reader with a
missing spine member and checks chapter order, diagnostics and published Markdown. These generated
files are deterministic regressions; they do not replace native Office acceptance.

`test_output_finalizer_collisions.py` covers real output-name contention, dangling symbolic links,
cancellation before publication and TaskManager's partial-publication result. Windows environments
without symbolic-link privilege report that specific case as skipped.

Office macro-policy tests cover bounded BIFF/OLE/ZIP/CSV inspection, COM ownership and call ordering,
failure diagnostics and the disposable LibreOffice profile. They use harmless generated inputs and
host doubles, and never establish that installed Office or WPS actually suppresses VBA or Excel 4.0
macros. Candidate acceptance must separately cover the affected installed backends, safe marker
samples, first-open ownership, ordinary conversion and preservation of existing user processes.

`tests/support/gongwen_template.py` builds a minimal attachment template in code, preserving the conflicting
character-unit and physical indents found in an actual user template. No binary template or local archive is
required by CI. `test_gongwen_indentation.py` covers style/ancestor/default/direct inheritance, body/cell anchors,
one/multiple/two-/three-digit attachments, character-zero cancellation and formatting preservation;
`test_attachment_origin_semantics.py` checks first-line/hanging origins, character-zero length fallbacks
and document defaults. Paired physical lengths retain the actual Word anchor (31.6 pt in the generated
legacy-format case); character-only axes still use the font fallback. Regressions preserve wider authored
hanging indents and compare saved lengths in twips rather than floating-point EMU conversions.
`test_gongwen_conversion.py` exercises both production Markdown input routes with
one/multiple attachments and the 9/10 and 99/100 boundaries. `test_attachment_number_padding.py` saves and
reopens code-generated body/cell templates to check literal space padding, exactly one ordinary space after
the period, inherited formatting and absence of inserted tabs or automatic numbering.
`test_attachment_alignment_modes.py` additionally exercises explicit `spaces`/`tab` layouts and
the configured default through both production routes. It saves/reopens body and cell anchors,
retains justification, and checks one name tab sharing the continuation origin, cancellation of
earlier authored tab stops, preservation of later stops, and rejection of invalid configuration.
The packaged GUI/CLI field gate retains four explicit space-layout cases and adds enabled tab-layout
value/empty cases; its independent oracle rejects missing/extra tabs and misplaced name tab stops.
Space-layout geometry remains an approximation under mixed scripts and justification; native
acceptance must report those differences instead of claiming both modes have identical precision.
Gongwen's `test_field_text_preservation.py` exercises actual DOCX extraction,
including labelled unusual/Latin signers and uncertain recipient prose retained in the body. These structural checks do not
prove pagination: selected native Word checks must also verify name columns and long-name continuation lines.

`test_gongwen_roundtrip_metadata.py` uses generated DOCX files to preserve attachment order and repeated names,
distinguish letterheads after variable-length header blocks from printing authorities, and consume pure imprint
tables while retaining unrecognized content in mixed tables. It covers body/table placement, composite document
numbers and signer names, merged copy-recipient cells, and printer/date pairs that must not cross tables, rows,
headings or body-to-table boundaries. Combined printing lines own their agency names, and an earlier attachment
list must not turn unknown cells in a mixed imprint table into a separate annex. Actual-file reverse conversion
remains a separate acceptance step.

`test_security_retention_context.py` saves and reopens generated DOCX header paragraphs/tables containing
custom or empty classification labels with `★` and a retention duration. It checks original field text,
leading-zero copy IDs, Arabic/full-width/Chinese durations, and protects inline prose, headings, lists
and data tables from being consumed as metadata.

`tests/integration/test_quote_content_mapping.py` validates quote child styles, current native list starts,
continuations, literal code, body order and table data. Markdown spelling and reasonable quote/list nesting
changes are normal mappings and do not require a loss warning. Authored IDs, links and fenced literal sources
retain their separate integrity tests. Native acceptance uses the same candidate with representative mixed
content in the relevant installed Word, WPS and LibreOffice hosts; it checks useful visible content after
save/reopen rather than requiring reconstruction of the source container tree.

`tests/integration/test_table_header_file_roundtrip.py` is a direct `.md` → `.docx` → `.md` file
integration regression for complex header roles, native merges and independent repetition policy. It removes the
source before importing the DOCX and covers conditional-style edits, changed cell text, moved tables, stale
geometry/bindings and ordinary tables whose only extension is repetition. OOXML edits model recovery boundaries;
they do not prove Word save/reopen or physical pagination.
Use the same candidate for the affected native Word scenarios, including turning first-row/first-column styles off
and checking repeated headers across a real page break. Clipboard and neutral-port tests do not replace this file lane.

`packages/plugins/markdown/tests/test_table_borders.py` checks physical multi-row three-line borders, nested and
staggered group spans, unmanaged name-only compatibility, template separator inheritance/attributes and explicit
`none`/`nil` protection. Conflicting inherited `insideH`/cell-top rules and nonzero cell spacing opt out. The direct
file lane also converts `tests/fixtures/markdown/multirow-three-line.md`, covering two/three-row headers, two-dimensional
groups and data merges without changing role/repetition metadata. Native acceptance must check residual internal
lines in row-spanning headings and complete lower borders after Word save/reopen; OOXML assertions alone cannot
establish border-conflict rendering.

测试按主要行为族和执行成本分层。默认仓库测试只选择非 slow 的 unit 与 contract；GUI、integration、端到端及环境所有型检查进入各自显式门禁。

## Common commands / 常用命令

```powershell
.\.venv\Scripts\python.exe tools\qa.py --suite fast
.\.venv\Scripts\python.exe tools\qa.py --suite full --coverage
.\.venv\Scripts\python.exe tools\qa.py --suite pr-integration --skip-ruff --skip-pyright
.\.venv\Scripts\python.exe tools\qa.py --suite release --skip-ruff --skip-pyright
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

Governed `tools/qa.py` runs place basetemp, cache, structured reports, and process TEMP roots physically below `<workspace>/.workspace/temp/p<random>`. The short physical name is deliberate: security-sensitive code canonicalizes aliases, so path-length headroom must come from the real workspace path instead of changing pytest fixture semantics. Per-test directories use a stable `t<12-hex-digest><counter>` name instead of pytest's longer test-name prefix; this preserves isolation and physical identity while leaving headroom for xdist's worker segment and adversarial output names. QA owns each unique directory, removes it after success, and retains it with a JSON lease after failure or interruption. Workspace selection uses `--workspace-root` first, then `DOCWEN_WORKSPACE_ROOT`, then the single `.workspace` path derived only from the governed `<engineering>/repos/docwen` repository layout; the removed direct `<engineering>/docwen` layout and arbitrary ancestor scanning both fail closed. A locally configured `DOCWEN_PYTEST_RUNTIME_ROOT` / `--pytest-runtime-root` must remain below the governed workspace's `temp` tree; repository-internal roots are rejected. `--own-pytest-runtime` may take ownership of a new explicit workspace-contained root, and `--keep-pytest-runtime` preserves an owned successful run. On Windows, the full suite also maps the physical runtime root to a verified temporary drive letter for I/O endpoints that can safely retain the alias. Pytest is allowed to canonicalize `tmp_path` back to the physical workspace path, preserving the same path semantics used by production code. The mapping is removed in `finally`; the lease records it while active so interrupted work can be recovered safely. Each QA start uses the same saved-plan retention policy as the operator tool: terminal-success scratch is immediately eligible; failure, interruption, cleanup failure, and a dead `active` owner keep the newest two per kind for at most 72 hours; explicit manual retention remains until an operator closes it; live owners always block deletion. CI without a governed local workspace supplies a new runner-temp root with `--own-pytest-runtime`, so QA writes its lease and cleans successful raw runs. `--report-output` exports only the five visibility JSON reports, optional coverage XML and a hashed summary to a new directory; local report destinations belong below `.workspace/acceptance`. Missing success reports or export failure fails QA and retains the raw run. `--tests-only` is used by CI only after its required source-check job succeeds; ordinary local QA keeps all source checks enabled.

Run `python tools/workspace_cleanup.py` to create a non-mutating housekeeping preview for leased scratch roots below `.workspace/temp` and `.workspace/build`. Tool probes also use leased directories below `temp`; no second `tmp` root is required or created. Successful scratch is immediately eligible; failed, interrupted, cleanup-failed, or abandoned scratch is bounded to 72 hours and the two newest roots per kind. Root-level `temp-*` directories are reported but never enter the default plan. A deletion requires an exact saved plan below `.workspace/diagnostics`: generate it with `--plan-output`, review it, then pass that same file to a separate `--apply-plan` invocation. Explicit unleased candidates require one or more `--target` arguments plus `--reason`; repository `.venv` and `node_modules` directories require `--clean-deps`. Apply revalidates the allowed root, complete identity and content fingerprints, leases and owner PIDs, and reparse boundaries before any deletion. Local engineering validation must use `tools/qa.py`; a manually assembled pytest command is an exceptional diagnostic action and cannot satisfy the cleanliness gate. A green release claim requires explaining all skips and proving that required dependencies were collected.

Engineering lease producers share `tools/run_lease.py` for identity creation and state transitions. QA reports skipped roots even when no root is eligible; foreign owners and unknown states remain protected and require explicit historical closeout. Windows operators may generate a plan with `--disposition recycle`. The disposition is fingerprint-bound; apply cannot override it, and recycling failure never falls back to deletion. Recycled bytes do not mean disk space has been reclaimed.

Cleanup plans bind the existing interpreter path, binary hash, environment prefixes, Python version and cleanup-tool source identities; apply and its result receipt retain that same environment identity. Use the same absolute interpreter and checkout for planning and apply. Missing identities in older plans, changed environments or changed tool code fail before journal creation or target mutation: inspect any existing progress first, then generate and review a new plan in the intended environment. A new plan does not clear a pending or ambiguous earlier mutation. No environment is created by this mechanism.

Recycle plans additionally bind a relocation-stable payload proof covering file bytes, physical entry identities and raw reparse data without traversing link targets. Recovery confirmation and completed-plan replay verify that exact payload and its `$I` metadata; matching old paths alone is insufficient. A failed or uncertain post-move verification keeps exact observed candidates and the original outcome in the pending progress record, which cannot be blindly retried. Older recycle plans without payload proof require regeneration. These tests use owned disposable fixtures; they do not authorize cleanup of retained user runs.

For source launch and Qt rendering, the managed scope also covers initialization and teardown. If recursive cleanup removes the lease before failing, it can recreate that lease exclusively only after verifying the original directory identity. A replaced marker is never overwritten. When terminal recording or safe restoration fails, a sibling `<run-name>.recovery.json` preserves the original identity, explicitly known outcome, execution error and recording error for manual reconciliation; its lease snapshot is the last in-memory state, not proof that a replacement marker was written. This record does not authorize deletion. Its presence or an inability to inspect it blocks planning and execution, including previously saved plans, regardless of its contents or age. Selecting an unleased parent also scans for descendant recovery records before reaching each outermost owned run with a validated ordinary-file lease marker; nested fixtures inside that owned run remain its payload, not independent cleanup authorities. Reconcile the retained run and record manually before retiring either; removing the hold is a separate operator action, never a consequence of a claimed successful outcome. Secondary recording failures do not hide the original error.

Windows lease liveness uses a read-only process handle and a zero-time wait. It never sends a console signal; access-denied and uncertain query results keep the leased directory protected.

After a local publication directory below `.workspace/artifacts` passes immutable public-byte verification, the publisher generates a retirement plan below `.workspace/diagnostics`. The plan binds the retained verification receipt, exact local inventory and filesystem identity. Apply rechecks these identities; ordinary explicit targets still cannot select protected artifacts. The plan is not applied automatically, so consumers can finish using the candidate. Runner-owned CI artifacts are outside this local lifecycle.

On Windows, newly generated housekeeping plans default to recycling. A saved plan keeps its explicit disposition; a failed recycle never falls back to permanent deletion and never empties the Recycle Bin. Apply reports each preflight/apply target on stderr, preserving JSON stdout. A workspace-wide OS lock prevents concurrent application. The adjacent `*.progress.json` binds the plan fingerprint and records each completed target with file fsync and atomic replacement, including its recoverable payload. Reapplying the same plan verifies completed targets remain absent (and recycled payloads remain recoverable), skips those targets, and revalidates every remaining target. A recreated target or missing recovery payload stops the operation. A pending mutation whose result is unknown, including a recorded refusal, requires inspection of that progress record; it is never blindly retried. Interrupted journal writes also fail closed. This covers process interruption and ordinary I/O failures; it does not promise power-loss durability of directory metadata. This is per-target recovery, not a transaction that rolls back a partially completed batch.

Recycle plans additionally bind a relocation-stable physical payload identity: all root/child File IDs,
relative entry shape, complete ordinary-file content and raw reparse metadata. The owner verifies the
actual `$R` payload and exact `$I` identity before recording completion, without following old absolute
link targets or rewriting lease bytes. A post-operation refusal records known source absence, exact
recovery candidates and observation errors under the existing unconfirmed journal entry; it cannot be
retried automatically. Old recycle plans without this proof require a newly reviewed plan.
On replay, current recovery availability is checked independently of the original successful observation.
A missing, replaced or changed payload refuses that replay and leaves the earlier completed record intact.

Run leases retain `finishedAt` and the original `outcome` separately from a later cleanup failure. QA and production builds record success before cleanup, and cancellation remains interruption rather than success. Explicit successful runtime retention records the responsible owner, reason and a review deadline (72 hours by default). Manual holds without this metadata, or past their deadline, appear as review observations and remain protected; review expiry does not authorize removal. Existing root mismatches, active owners and reparse refusals keep their original error and include a suggested next step. These engineering checks are not GUI, Office or release-package acceptance.

受治理的 QA 会把 basetemp、缓存、结构化报告和进程 TEMP 的真实文件统一放在 `<workspace>/.workspace/temp/p<随机名>`。短物理名称是刻意设计：安全敏感代码会把别名解析回真实路径，因此路径余量必须来自工作区内的真实短路径，不能靠篡改 pytest 临时夹具的路径语义。每项测试的目录名使用稳定的 `t<12 位十六进制摘要><计数器>`，替代 pytest 较长的测试名称前缀；这样仍保持逐测试隔离和真实物理身份，同时为 xdist 工作进程层级和对抗性输出名留出余量。成功自动删除，失败或中断通过 JSON 租约留证。工作区选择依次采用 `--workspace-root`、`DOCWEN_WORKSPACE_ROOT`，最后只根据受治理的 `<工程>/repos/docwen` 布局推导唯一 `.workspace`；已移除的 `<工程>/docwen` 直连布局和任意祖先扫描都失败关闭。本地显式运行根也必须位于该工作区的 `temp` 树内；`--own-pytest-runtime` 只能接管全新、位于工作区内的显式根。Windows 完整套件运行期间仍会把这个物理目录临时映射为经过身份校验的短盘符，供能够安全保留别名的 I/O 端点使用；pytest 的 `tmp_path` 则允许规范化回工作区内物理路径，与生产代码保持同一语义。测试结束时在 `finally` 中解除映射，不会在盘符根目录创建测试文件夹。租约在映射期间记录短盘符，使异常中断后的回收仍能验证目标身份。每次 QA 启动都采用与运维工具相同的保存计划保留策略：成功终态 scratch 立即可回收；失败或中断按 kind 保留最新两份且最长 72 小时；人工保留和非终态继续保留；活动所有者始终阻止删除。`tools/workspace_cleanup.py` 默认只生成不改变工作区的计划；需要删除时，必须先用 `--plan-output` 把计划保存到 `.workspace/diagnostics` 并完成审阅，再在单独命令中用 `--apply-plan` 消费同一文件。根级 `temp-*` 只报告、不进入默认计划；无租约目标需要 `--target` 与 `--reason`，仓库内 `.venv`/`node_modules` 只能通过 `--clean-deps` 显式纳入。apply 会在任何删除前重验允许根、NTFS/文件系统身份、完整元数据与内容指纹、租约/PID 和 reparse 边界。没有本地治理工作区的 CI 通过 `--own-pytest-runtime` 接管全新 runner 临时根；`--report-output` 导出五份可见性 JSON、可选覆盖率 XML 和哈希摘要后，成功立即清理原始运行。Windows 全量测试与覆盖率、GUI 覆盖率使用同一次运行；Linux/macOS 各执行一次 fast 与 release 标记的并集。

The default selection expression is `(unit or contract) and not slow`. GUI, integration, end-to-end and other environment-owning primary families run in their owning lanes instead of entering the fast lane by omission. PR and release expansion use `pr_gate and (integration or gui_smoke or e2e)` and `release_gate and (integration or gui_smoke or e2e)`. The shared pytest options are `-v --tb=short --strict-markers --import-mode=importlib -ra`.

QA owns coverage through `python tools/qa.py --suite full --coverage`. It sets coverage data and XML inside the leased raw runtime, uses the governed `pytest-xdist`/`loadfile` lane, and reads the unchanged `fail_under = 81` threshold directly from `pyproject.toml`. That single Windows full run applies to pull requests, main and release calls. The same XML feeds whole-product source-manifest validation, the core domain summary and the GUI domain summary. The 81% threshold applies to the whole product including GUI; `--coverage` rejects fast or partial suites. The stable governance entry points are:

```text
python tools/qa.py --skip-ruff --skip-pyright --suite pr-integration
python tools/qa.py --skip-ruff --skip-pyright --suite release
python tools/check_coverage_source_manifest.py coverage.xml
python tools/check_core_coverage.py coverage.xml --soft-gate
python tools/check_gui_coverage.py coverage.xml
python tools/check_test_governance_consistency.py
python tools/run_import_linter.py
```

`[tool.coverage.run].source` in `pyproject.toml` is the only whole-product coverage source manifest. The coverage job uses bare `--cov --cov-config=pyproject.toml`; every real package under `packages/**/src/` must appear once in that manifest and in `coverage.xml`. Missing XML, no measured files, an unreported configured package, `module-not-imported`, and `no-data-collected` are hard failures rather than an empty green report. The GUI domain report reads the same full coverage XML; it does not run GUI tests again or maintain another source manifest.

Parallel execution uses `pytest-xdist` and `loadfile` by default. Governed `auto`/`logical` runs are capped at six workers because the suite contains nested CLI and pytest subprocess probes; `DOCWEN_PYTEST_XDIST_WORKERS` selects another worker mode or exact count, `DOCWEN_PYTEST_XDIST_MAX_PROCESSES` changes the automatic cap, and the `DOCWEN_PYTEST_XDIST` switch explicitly disables parallelism when set to `0`. This avoids oversubscribing a host without making ordinary local QA serial by default. Visibility files are `skip_report.json`, `not_collected_report.json`, `slow_report.json`, `subprocess_report.json` and `missing_marker_report.json`. Fixture ownership is documented in `tests/fixtures/README.md`; stable goldens live under `tests/fixtures/golden/`. Any changed fixture requires 人工审 diff.

Primary-marker debt is closed rather than hidden by mass edits. Existing `gui` tests count as primary GUI tests; governed QA requires every collected node to have exactly one primary marker, with zero missing and zero overlapping classifications. Marker-only migrations preserve node IDs, skip semantics, and failure semantics. Core, runtime, application, CLI, GUI, bundle, document, image, spreadsheet, proofread, Markdown, and markup behavior families are explicitly classified; filesystem, process, lock, or cross-module paths are `integration`, public wire and semantic boundaries are `contract`, widgets are `gui`, and pure state is `unit`. Real OOXML used to verify a stable conversion contract remains `contract`; it is not demoted merely because a fixture is serialized to disk. Runtime classification deliberately moved file-, lock-, or cross-module tests from the default fast selection into the complete suite. Application/CLI/bundle classification likewise reserves real process and destination paths for higher-level lanes, while stable semantic contracts remain in the fast lane. The AST guard discovers every configured `testpaths` root and recursively audits class methods; new missing or overlapping primary classifications fail immediately. Collection size is recorded per exact tree and is never treated as a permanent constant or reduced to make a gate green.

Repository tests named `*_projection_contract.py` are current unit/contract/golden guards over checked-in projections or fixtures. They are not live candidate, package, Office, Store, or host evidence. Real acceptance is executed by the owning packaged or host verifier and may emit a compact receipt under `.workspace/acceptance`; source test success cannot substitute for that receipt.

For an acceptance tool that owns a leased raw run but does not integrate receipt generation, first mark the lease `completed-success` after all selected gates pass and the owner process exits, then run `tools/acceptance_closeout.py` with the exact candidate, gate, subject name, subject byte count, and subject SHA-256. The closeout refuses live or non-success runs, writes a recoverable provisional record, retires the raw run through the saved-plan housekeeping path, and publishes one compact receipt below `.workspace/acceptance`.

Test-file size governance reads the same configured `testpaths` roots and rejects every test module above 700 lines. There is no oversized-file baseline or exception list. Shared fixtures and non-test helpers live in focused support modules, while collected test modules remain independently reviewable. Legacy monolith path names are collection errors rather than ignored paths, so a reverted large test cannot disappear from CI.

Test subprocesses should use `tests.support.subprocess_runner.run_subprocess`, whose default timeout is 60 seconds and whose structured record distinguishes timeout from child exit. A test may choose a different positive timeout, but disabling the timeout with `None` or a non-positive value is rejected. Long-lived protocol sessions that require `Popen` must own an explicit bounded wait/terminate/kill/reap lifecycle and remain in a serial marker family until their process isolation is proved.

Pyright includes all source and test modules under `packages`; a global `**/tests/**` ignore or exclude is not an acceptable substitute for fixing package-test errors.

Governed QA passes its current Python executable to Pyright explicitly. This keeps the type gate reproducible when the project environment is stored under `.workspace/tools` and avoids requiring a repository-local `.venv` junction or copy.

The packaged successful-warning gate uses `--successful-warning-smoke <input>` or the pinned generated fixture with `--successful-warning-smoke`. The GUI verifier first resolves the unique canonical warning through the packaged CLI from the same candidate and rejects empty outputs, then requires the exact warning row/tooltip, persistent warning-tone task summary, warning-row PNG and byte-identical CLI/GUI artifacts. Evidence must preserve the same JSON structured warning and text stderr warning, and prove final artifact 字节一致. Explicit fixtures should also pass `--successful-warning-input-sha256`; `--successful-warning-message` is an additional golden, not a replacement for CLI truth. This gate 不得外推为通知中心可见性.

Add `--proofread-report-smoke` to the packaged CLI verifier when checking the proofread consumer contract. The verifier owns the BOM/CRLF/Unicode fixture, checks `docwen.proofread_report.v2`, raw-byte identity and authoritative ranges, and separately requires a successful empty report from `--check none`. 任何非零退出或 `success=false` 都是 failure，不能折叠成成功空报告。该门不替代 Obsidian 侧 UTF-16 坐标转换、过期结果或实际修复验收。

## Packaged GUI smoke / GUI 打包冒烟

The default packaged GUI smoke uses `DOCWEN_GUI_TEST_AUTOCLOSE_MS`. Add `--notification-smoke`, `--ocr-smoke` or `--ipc-smoke` only for the boundary being verified. With no GUI already running, the IPC gate additionally verifies the exact packaged `info` declaration for `gui.settings`, direct CLI cold start into the proofread section, the runtime handshake, singleton dialog reuse, and same-primary file delivery. Expiry-without-late-side-effect and explicit deadlines longer than 15 seconds remain source-level adversarial tests rather than package claims. The matching limitations are documented in [Packaging](packaging.md).

默认打包 GUI smoke 使用 `DOCWEN_GUI_TEST_AUTOCLOSE_MS`。仅在需要验证对应边界时增加 `--notification-smoke`、`--ocr-smoke` 或 `--ipc-smoke`。在没有 GUI 预先运行的前提下，IPC 门还会精确验证打包 `info` 中的 `gui.settings` 声明、由 CLI 直接冷启动到校对设置页、运行时握手、单例窗口复用和同一主实例文件投递。已过期请求无迟到副作用和显式超过 15 秒 deadline 仍是源码级对抗测试，不冒充包级证据。各模式不能证明的范围见[打包文档](packaging.md)。

Successful packaged-GUI runs are ephemeral by default. In the governed local layout, raw runs are created below `.workspace/temp` with a lease; success removes them immediately, while failure or cleanup failure enters the bounded retention policy. Use `--candidate-id ID --receipt-output <ABSOLUTE_JSON_BELOW_WORKSPACE_ACCEPTANCE>` to close a successful run into a compact identity/result/limitation receipt. Use `--evidence-dir <ABSOLUTE_NEW_DIRECTORY>` only when a reviewer explicitly needs the complete run. The verifier rejects unsafe destinations and verifies exact copied bytes. Copy failure retains the authoritative leased source and fails the gate rather than silently reporting success.

打包 GUI 成功运行默认不保留原始目录。本地治理布局会把原始运行放入带租约的 `.workspace/temp`：成功立即删除，失败或清理失败进入有上限的保留策略。使用 `--candidate-id ID --receipt-output <位于 workspace/acceptance 下的绝对 JSON 路径>` 可把成功运行关闭为紧凑回执；只有复核者明确需要完整现场时才使用 `--evidence-dir <绝对且不存在的目录>`。复制失败会保留带租约的权威临时源并使门禁失败，绝不静默报告成功。


## Change-scoped acceptance and closeout

Input toolbar regressions cover right-aligned compact Add/Paste actions, the shared paste/menu border,
trash-icon Clear, and reflow under light/dark themes at all five application scales and standard/large
fonts (`test_input_area_layout.py`). Format overview checks cover whitespace-separated names, centered
pyramid rows, diagonal category colons and narrow-window wrapping. Paste/menu event tests distinguish normal and plain-text requests;
the Clear keyboard test checks selection reset while preserving the source file. Native inspection
also checks the SVG arrow, keyboard focus and representative live theme/scale changes in an isolated
profile. Source rendering and geometry checks do not establish packaged or cross-machine acceptance.

Keep the existing required CI lanes. Select the smallest representative set of additional checks
by the affected behavior and record the selection and reasons for omitted surfaces. For a fix,
cover the original reported case and a relevant boundary; expand for distinct risks rather than
repeating every historical scenario. Documentation-only and host-independent logic changes may
omit host execution when the reason is recorded. Confirmed serious product defects block the
affected release. Missing host/tool evidence stays unverified and is assessed by risk, without an
automatic publication block or extra approval step. Claim acceptance only for observed passing scope.

A local desktop, sandbox or suitable VM may host the real application when the required software,
input and observation are available. Choose one for equivalent desktop coverage; do not repeat all
environments. Static/offscreen checks do not replace affected native interactions. Different
Office engines, platform-specific launchers and installation channels remain distinct risks.

| Change | Additional acceptance |
|---|---|
| Documentation or public screenshots | References, English README / listing language, actual Qt layout; no package rebuild for media-only changes |
| GUI behavior or layout | Changed interaction on the real application; representative light/dark, language and DPI cases |
| Conversion, templates or Office integration | Affected formats with representative real inputs and relevant installed host software |
| Machine protocol or consumers | Producer/consumer contract checks plus affected real consumer operations |
| Packaging, launcher or installer | Exact candidate startup, CLI/GUI identity, upgrade and repeat launch on the affected platform |

Build once per platform, accept those bytes, publish them, then perform one independent public readback. Store certification and UOS prerequisites are channel-specific; missing external access does not turn an untested scope into a pass or block unrelated channels. Retries revalidate identity and reuse the verified candidate.

Cleanup also discovers leased runs under the registered `repos/docwen/build` and `repos/docwen-openclaw/build` directories. It never selects either build root itself, unleased repository directories or tracked files. Source directories such as `scripts/build` remain outside this boundary. Existing owner, retention and content identity checks still apply.

`tools/acceptance_closeout.py` accepts `--result passed|failed|superseded`. Nonpassing closeout requires `--reason` explaining the original failure and, when applicable, its replacement. Original leases remain in the receipt; live owners and active/nonterminal runs are rejected. `--disposition recycle` retains verified recovery locations on Windows. Recycling uses the same implementation for files and directories, fails when unavailable, and never falls back to permanent deletion.

Manual closeout defaults to recycling on Windows (the same platform default as housekeeping). Automated owners that require deletion must select it explicitly. Once run identity and evidence are validated, preflight failures also leave a compact receipt with the failure stage and original result. WSL links are inspected without following them; only relative targets with a plain, existing destination are supported. Absolute guest paths, parent traversal, nested reparse targets and unknown reparse types stop cleanup.

For Linux acceptance, extract packages inside a leased native ext4 filesystem, not directly into a Windows-mounted directory. Export compact evidence, stop guest processes, verify the mount is detached, then retire the owned backing image through the managed cleanup entry point. A mounted image is not eligible for recycling.

Add repeated `--evidence <relative-file>` arguments to preserve UTF-8 failure summaries before raw cleanup (at most eight files, 64 KiB each). Files must be plain files inside the owned run; credential files are excluded. The formal receipt is written before cleanup and records cleanup errors without changing the original result.

For routine local coverage history, use `tools/qa.py --suite full --coverage --compact-coverage --report-output <new-directory>`. All coverage gates still consume the full original XML before export. The compact export retains total/package statistics and the original XML hash; it cannot rerun line-level checks. Omit `--compact-coverage` for the current release or CI evidence that needs full XML. Keep result/identity/limitations and necessary screenshots, not duplicate runtime trees. Public product screenshots use real Qt rendering; they are not evidence of native interaction.

For OCR-table changes, use `verify_packaged_cli.py --table-smoke` on the same final candidate. It checks one ruled and one borderless table, surrounding prose, exact cell rows, no duplicated text and the plain-OCR off switch. See [the implementation contract](specs/ocr-tables.md). Real user screenshots, embedded-document routes, GUI state and Office preconversion remain affected-scope host acceptance, not claims established by this synthetic smoke alone.

Focused rich-input regressions include `test_structured_clipboard_converter_semantics.py` (GUI clipboard projection through Runtime and final XLSX save/reopen at UTF-16 boundaries, default/custom templates and composed paragraphs), `test_structured_clipboard_xlsx_capacity.py` (Image Semantics), and `test_mhtml_resource_identity.py` (real MIME archives, distinct host/query/path-case/CID image bytes, literal body/code URLs, relative Content-Location/base href, conflicting references, and native file URI images). These prove source conversion behavior; POSIX file URI execution still requires Linux/macOS, and native provider/Office/candidate behavior remains a separate acceptance layer.

`test_web_archive_review_boundaries.py` adds final Runtime regressions for MIME parent bases, related start/alternative roots, unique basename aliases, relative HTML base decoys, bounded missing-local warnings, and HTML/MHTML/ENEX authored placeholder text. Layout's final HTML helper and Core placeholder tests prove collision checks and single-pass restoration; the Layout helper currently has no production caller, so those cases do not claim PDF/Office execution. The structured clipboard final-artifact suite preserves an existing same-name XLSX byte-for-byte under a frozen publication identity when a later request overflows. See [web archive contracts](specs/web-archives.md).

Round-two regressions cover final retained template text through real TemplateManager import, canonical ID, Runtime and Finalizer, including sparse last-cell/formula preservation and same-name output bytes. `test_review_round2_boundaries.py` covers literal Content-ID versus one URL decode, nested unsupported alternative fallback without hiding malformed roots, actual Windows drive-qualified PNGs, parsed ENML dependencies versus comment/CDATA/attribute lookalikes and external entities, and omit/OCR I/O and diagnostic behavior. Controlled OCR outcomes prove resource handling, not an OCR engine. Windows short-drive aliases do not prove a separate physical volume.

Markdown XLSX template scans visit existing cells in row/column order rather than materializing their bounding rectangle. A saved sparse template with a last-row/last-column sentinel exercises YAML, row/column and image placeholders, formula preservation and final save/reopen; the guard fails before an accidental dense scan can allocate the empty grid.

`test_xlsx_image_text_plan.py` combines near-limit text with real template/table images and URL-quoted Unicode paths. It checks final ASCII/non-BMP limits, exact drawing bytes, literal marker trust, overwritten/protected cells, placeholder discovery order, duplicate drawings, sparse cells and the internal CSV-template helper. Temporary scope strings never enter openpyxl; output assertions save and reopen real workbooks. `test_mhtml_query_resolution.py` checks actual PNG bytes through Runtime under paired synthetic/absolute bases, empty query/fragment delimiters and relative paths.


## Native content and resource regression / 原生内容与资源回归

The file regressions in `test_native_docx_reading_contract.py`, `test_formula_and_gongwen_source_content.py`,
`test_resource_identity_pipeline.py` and `test_semantic_quote_break_recovery.py` build real minimal OOXML/ENEX/PPTX
inputs and follow production conversion, finalization and Bundle publication. Independent ZIP/XML facts, exact
text/order, mathematical tokens, Markdown table ASTs and destination-to-resource byte hashes are the oracles.
Quote tests inspect physical carriers and break nodes before fresh reading and production source replanning.
`test_resource_relocation_boundaries.py` deliberately supplies declared manifests to isolate finalizer ownership,
ambiguous-alias rejection and exact non-target byte preservation; it is not a native host or Machine session.
The OCR sidecar case substitutes only the typed OCR engine outcome; resource allocation, file writes and Bundle
hash checks remain real. The native attachment fixture contains actual DOCX bytes, not text with an invented MIME.
`test_mathml_text_and_tokens.py` supplies independent MathML leaf and lexical-boundary expectations.

这些回归使用真实最小 OOXML／ENEX／PPTX，经生产转换、finalizer 与 Bundle 发布，按 ZIP/XML 原始事实、完整
正文／次数／顺序、数学 token、独立 Markdown table AST 及最终目标→资源 SHA 核验。quote 在 fresh reader
之前检查物理 carrier 和 break 节点，再调用生产 replan。重定位边界测试显式提供声明 manifest，只证明
finalizer 的所有权、歧义拒绝和非目标字节保留，不冒充宿主或 Machine session。OCR 用例只替换引擎的 typed
outcome，后续命名、实际写入与 Bundle 摘要均为真实链；普通附件使用真实 DOCX 字节，不依赖虚构 MIME。
MathML 叶及控制词边界另用独立预期检查。以上不替代真实 Office、Qt、Obsidian、跨平台或最终候选验收。

The related `test_gongwen_formula_structure.py` and `test_gongwen_request_source_ownership.py` pass the same
production ConfigPort snapshot to ApplicationController and Runtime. They pair shipped heading-cleanup rules
with an explicit empty-rule request; configuring only the runtime loader does not establish the request snapshot.
`test_visible_semantics_and_literal_code.py` checks exact visible paragraph and code AST content around real
legacy/v3 references, while `test_nested_quote_separator_ownership.py` checks complete parent/child source ranges
across trailing page/section separators. `test_heading_list_numbering_identity.py` independently checks native
Heading text and effective list numbering before asserting list/Heading structure and counter progression.

公文结构与源归属回归把同一生产 ConfigPort 快照交给 ApplicationController 与 Runtime，同时检查已发布的
标题清理规则和显式空规则请求；仅配置 runtime loader 不代表请求已携带该快照。可见语义／字面代码回归
在真实旧版与 v3 引用周边检查精确正文和代码 AST，嵌套 quote 回归核验尾部页／节分隔两侧完整父子范围；
列表内标题回归先独立检查原生 Heading 正文及有效编号，再检查列表／标题结构和实际计数推进。

`test_clipboard_wiki_image_resource_identity.py` uses generated PNGs and an explicit typed ClipboardDocument,
then the real input store, Runtime, finalizer and Bundle. It independently checks every Wiki occurrence's target,
empty/nonempty alias, source hash and surrounding text, including repeated resources and real output-root
contention. It does not claim a native OS clipboard capture. Shared formatter regressions distinguish raw resource
paths from already encoded references, preserving literal percent identity, URI query/fragment and existing embeds.

结构化剪贴板 Wiki 回归使用生成 PNG 与显式 typed ClipboardDocument，经真实 input store、Runtime、finalizer
及 Bundle，逐 occurrence 独立检查目标、空／非空 alias、原资源摘要与周围正文，并覆盖重复资源和实际输出根
冲突；不把该结构输入称作原生系统剪贴板抓取。共享格式化回归区分原始资源路径与已编码引用，分别保留字面
百分号身份、URI query／fragment 及既有 embed 语义。


## Native content and publication regressions / 原生内容与最终发布回归

`test_spreadsheet_final_content.py`, `test_spreadsheet_projection_boundaries.py` and
`test_spreadsheet_hub_delimiter.py` create saved/reopened CSV/TSV/XLSX packages and compare
final independent Markdown AST or complete typed matrices, worksheet order and resource hashes.
Sparse/merge cases additionally inspect actual stored cells and cancellation/close boundaries;
allocation instrumentation supplements the content assertions. The hub substitutes only the
external Office backend after independently reopening the real XLSX hub.

`test_xlsx_ods_typed_publication.py` and `test_xlsx_ods_delivery_ownership.py` use a complete
externalLink relationship graph and typed ODS packages. Source and final facts come from separate
XML readers; moved values, missing cells, wrong types, malformed ancestry and repeated-cell
identity are tested alongside valid strings, empty strings, numbers and booleans. The backend
substitute delivers real ODS bytes to the actual policy, Runtime, finalizer and Bundle; it does
not establish Excel/LibreOffice conversion behavior.

`test_pdf_split_selection_publication.py` checks actual saved PDF page text and dimensions after
CLI/Runtime publication, and checks zero publication for whole-request refusal.
`test_pdf_selection_bounds.py` is an offscreen widget/model test, not a physical Qt observation.
`test_clipboard_table_budget_publication.py` uses frozen synthetic HTML/plain captures through the
real Store and request builder, then checks full DOCX body, merge, nested order and header metadata.
It covers both native-representable headers and associations preserved in custom XML.

Invoice native PDF/OFD regressions compare signed YAML strings and every item-table cell against
independent input facts, including physical-page identity and uncertainty/conflict diagnostics.
No OCR engine is exercised by these native-text/package tests. Existing OCR substitute tests
retain their separate evidence level.

`test_caption_resource_path_identity.py` checks each saved DOCX caption image occurrence against
its final Markdown target, Bundle logical path and original SHA. The native reference boundary
suites inspect authenticated bookmark/field XML before checking full reader content, formatting,
link identity, actual breaks and production replan. `test_native_textbox_carriers.py` supplies
complete DrawingML/VML/Choice/Fallback envelopes; older reduced textbox XML tests remain projection
checks. Saving a ZIP and inspecting these envelopes does not establish real Word save/reopen.

上述回归全部使用新合成文件，不冒充不可用的原报告样本。最终内容、次数／顺序、坐标／类型及目标→SHA
以独立消费核对；自身往返、子串、产物数量或替身调用次数只作辅助。ODS／hub 仅在外部 Office 边界提供
替身；PDF、Invoice 原生文本和 OFD 容器检查不代表 OCR 引擎或真实桌面。文本框完整包裹结构与早期简化
XML 投影测试分开记录，真实 Word 保存重开、Office／LibreOffice／Qt／Obsidian、跨平台进程、最终包候选
及 hosted CI 另行验收。QA 原失败、夹具纠正及显式保留 runtime 的结果各自记录，不将保留运行称作自动
清理通过。


Focused GUI regressions use real `QMimeData` text/URL drops, ConfigPort persistence followed by a fresh reload and TextValidator observations, and actual Qt DeferredDelete delivery after moving/sorting a mixture of attached and queued cards. The successful-warning hook is checked against a real ActivityRecordsModel and details widget, including operation/source negatives and complete local detail text. Offscreen Qt checks establish these source-level behaviors; native desktop drag/drop, idle card interactions and packaged screenshots remain separate acceptance scopes.

定向 GUI 回归使用真实 `QMimeData` 文本／URL 拖入、ConfigPort 保存后新实例重载与 TextValidator 观察，并在已挂／待创建卡片混合排序后投递真实 Qt DeferredDelete 事件。成功警告 hook 使用真实 ActivityRecordsModel 与详情控件，覆盖 operation／source 负控和完整本地详情。离屏 Qt 仅证明这些源码行为，原生桌面拖入、空闲后卡片交互与最终包截图仍属独立验收范围。

MSIX and macOS publication regressions inject bounded copy/hash/link/report failures into real filesystem artifacts and inspect retained lease/publication identities, final bytes and unknown-file preservation. MSIX SDK construction is explicitly substituted in that fault lane; macOS publication tests reuse saved and reopened tar archives without building or signing. Native SDK and arm64 candidate verification still run only in their platform acceptance lanes.

MSIX／macOS 提交回归对真实文件注入有界 copy／hash／link／report 故障，核对留存租约、提交身份、最终字节及未知文件保护。故障测试中的 MSIX SDK 构建明确为替身；macOS 测试复用已保存重开的 tar，不构建或签名。原生 SDK 与 arm64 候选验证继续由各平台验收通道完成。


`test_office_linux_process_ownership.py` compiles a bounded native C launcher/child fixture inside the owned test runtime. The child creates a separate session and emits an activity heartbeat; tests capture its actual PID/start ticks and profile/input/output arguments before cancellation or the explicit 0.8-second timeout. The normal case is released only after observation. A TERM-ignoring child exercises KILL, while an unrelated native sentinel continues to emit heartbeats. Checks require the owner’s kernel wait/reap receipt, a joined supervisor, stable stopped heartbeat and independent identity absence. Saved/reopened DOCX input and final copy bytes are checked separately from process lifetime.

`test_office_linux_process_ownership.py` 在测试自有运行根编译有界 native C launcher／child。孩子建立独立 session 并持续写心跳；取消或显式 0.8 秒超时前，测试取得实际 PID／start ticks 及 profile／输入／输出参数，正常用例在完成观察后才放行。忽略 TERM 的孩子验证 KILL，无关 native sentinel 心跳继续。断言结合监督者的内核 wait／reap 回执、join、停止后的稳定心跳与独立身份消失；保存重开的 DOCX 输入与最终复制字节单独核验。

The missing-confirmation regression drops only the final IPC receipt after the real owner has run and reaped its native child. Core and the real Runtime/PrintPlugin route must still fail with zero final artifacts and retain the independent process state after TaskManager cleanup. This is a deterministic receipt fault; it is not a D-state observation or a real LibreOffice conversion. Actual installed LibreOffice version, normal conversion, cancellation, hard timeout, macro behavior and packaged-host acceptance remain independent gates.

缺失确认回归只丢弃真实 owner 已运行并回收原生孩子后的最终 IPC 回执；Core 与真实 Runtime／PrintPlugin 路线仍须失败、零最终产物，TaskManager 清理后独立进程现场继续保留。这是确定性回执故障，不是 D-state 观察或真实 LibreOffice 转换。实际安装的 LibreOffice 版本、正常转换、取消、硬超时、宏行为和最终包宿主验收继续作为独立门禁。

A separate native KILL-refusal case delivers TERM to a genuinely active TERM-ignoring child, then injects `PermissionError` only at the production KILL send. After the bounded bridge failure, independent PID/start-ticks and an advancing heartbeat must still identify that child; profile/private input/output remain and no final artifact is published. Only after these assertions does a test-owned release permit rescue through the original retained pidfd, another kernel wait-child check and reap to ECHILD. This verifies active-but-unconfirmed retention separately from a lost receipt; the rescue is not a production cleanup success or D-state evidence.

独立的原生 KILL 拒绝用例向真实活动且忽略 TERM 的孩子发送 TERM，仅在生产 KILL 发送处注入 `PermissionError`。有界桥接失败后，独立 PID／start ticks 及继续增长的心跳仍须指向同一孩子；profile／私有输入／输出保留且不发布最终产物。断言结束后测试才放行收尾，用事先保留的原 pidfd、再次确认的内核 wait-child 关系及 ECHILD 回收。这与只丢回执的用例分别证明“仍在活动但未确认”和“缺少回执”；测试救援不算生产清理成功，也不是 D-state 证据。

Discovery-layer controls distinguish readable task files, missing interfaces, vanished secondary threads, namespace mismatches and finite entry/time/text budgets. Native controls may supply a test-owned `children` file populated from the C child's independently recorded procfs PID, or intentionally omit that interface to exercise the real status/PPid scan. Such supplied files prove the discovery parser and actual downstream kernel ownership; they do not assert that the host exposes a usable kernel `children` file.

发现层对照区分可读任务文件、接口缺失、次线程退出、命名空间身份冲突及有限的条目／时间／文本预算。原生对照可由测试提供 `children` 文件，并填入 C 孩子独立记录的 procfs PID，或故意省略接口以执行真实 status／PPid 扫描。这些测试文件证明发现解析和后续真实内核归属，不宣称宿主实际提供可用的内核 `children` 文件。


`test_spreadsheet_literal_final_docx.py` saves and reopens native XLSX inputs, then publishes Markdown and DOCX
through Runtime/finalizer and Artifact Bundle. An independent Markdown-to-HTML consumer and direct OOXML cell
matrices distinguish literal `<`, `>`, `&`, authored `&lt;`, caret, backslash/pipe, source line breaks and literal
`<br>` text. Ordinary and Structural Tables configurations retain the same text, while horizontal/vertical native
merges retain their separate geometry. The fresh DOCX package checks do not claim a real Word save/reopen.

上述真实文件回归先保存重开 XLSX，再经 Runtime/finalizer 与 Bundle 发布 Markdown 和 DOCX。独立 Markdown
消费者及直接 OOXML 单元格矩阵区分原字符、原本就是 entity 拼写的文字、反斜线／pipe、真实行内换行和
字面 `<br>`；普通／Structural 配置的正文相同，实际横向／纵向合并另核几何，不冒称真实 Word 保存重开。

### CLI cancellation, publication and admitted workflows / CLI 取消、发布与准入工作流

`test_cli_deadline_v3.py` drives the real composed CLI in child processes with bounded phase gates around
worker start, CSV conversion and actual OutputFinalizer publication. Saved XLSX matrices and aligned per-input
results establish earlier success, later cancellation and unstarted inputs. Linux SIGINT is measured separately
from Windows Ctrl+C, which these tests do not establish. A separate delayed-worker test checks that bounded
drain does not release a cancelled Application scope before the worker exits. The instrumentation calls the
production controller, converters and finalizer; it does not replace them with a mock controller.

CLI 截止回归使用真实组合入口的子进程，仅在 worker 启动、CSV 转换及真实最终发布前后设置有界时序门；
重开的 XLSX 矩阵与逐输入结果证明先前成功、后续取消及未启动状态。Linux SIGINT 与 Windows Ctrl+C 分开，
本组不建立后者的宿主证据。延迟 worker 对照检查有界等待结束后仍保留取消身份，直至真正退出。

`test_validate_execute_02.py` checks inspection, admission and real Runtime report publication for neutral Markdown,
inline code, a single list, headings and plain text, including original-byte hashes and independent source ranges.
`test_cli_document_proofread_publication.py` examines real private DOCX content and comments before a final public
commit. Its ODT writer is explicitly an Office-boundary double with an independently inspected ODF package;
it proves pipeline/publication behavior, not an installed Office conversion. Missing-bridge cases cover all four
legacy target admissions. Native Office save/reopen remains separate. Logging namespace checks use actual file
and console sinks in a child process, so no logging-global state leaks back into the test process.

校对回归对普通 Markdown、行内代码、单列表、标题与纯文本走实际检查、准入和 Runtime 报告发布，独立核对
源字节 hash 与 range。转换后校对先检查实际私有 DOCX 的正文和批注，再检查一次公开提交；ODT writer 明确是
Office 边界替身，以独立 ODF 包检查输出，只证明管线及发布行为，不证明安装的 Office 转换能力。缺 bridge
对照覆盖四种旧文档目标准入；原生保存重开另验。日志命名空间用子进程中的真实文件及 console sink 检查，
不向测试主进程泄漏 logging 全局状态。

The cancellation tests also exercise real SIGINT inside executor submission, after submission returns, while
Application returns a reservation and while Deadline returns its captured handle. The real controller then reuses
the same request ID and source to publish a fresh XLSX, proving the old cancelled scope no longer owns that ID.
Application-only contract cases pair pre-alias capture and partial-alias rollback with refusal to acquire a prior
owner and rejection of stale exact cancellation/release against a later same-ID scope. These are in-process
ownership checks, separate from native process termination.

取消回归还在 executor 提交内部、提交返回后、Application 返回预约及 Deadline 返回已捕获句柄时发送真实 SIGINT，
随后以同一 controller、request ID 和源文件重新发布真实 XLSX，证明旧取消 scope 已释放该身份。Application 合同
对照覆盖先交付句柄后挂 alias、部分 alias 回滚、拒绝接管既有 owner，以及旧句柄不得取消或释放同 ID 新 scope；
这些属于进程内所有权检查，与原生进程终止证据分开。
