# Runtime artifacts / 运行产物

Plugins produce typed results and declare artifacts. The runtime output finalizer owns final placement, naming, collision handling and recovery.

插件返回类型化结果并声明产物；最终位置、命名、冲突处理和恢复由 Runtime Output Finalizer 统一负责。

This page describes the current in-process runtime contract. Internal `ArtifactManifest` and `OutputManifest`
objects are not the external wire format. The implemented external process boundary is
[Machine Protocol v1 and Artifact Bundle v2](specs/machine-protocol-v1.md). The Application Service maps runtime
results into a BundleDraft; Bundle commit then validates locators, graph semantics, size and SHA-256 before
publishing the task Bundle. CLI JSON presentation objects are not a compatibility layer for external consumers.

本页描述当前进程内 Runtime 契约。内部 `ArtifactManifest`、`OutputManifest` 不是外部 wire format。新进程
边界已由 [Machine Protocol v1 与 Artifact Bundle v2](specs/machine-protocol-v1.md) 实现。Application Service
先把 runtime 结果映射为 BundleDraft，Bundle commit 再校验 locator、图语义、size 与 SHA-256 后发布任务
Bundle。CLI JSON 展示对象不是外部消费者兼容层。

## Contract / 契约

- Source files are never output targets and are not modified in place.
- Staging and intermediate files live in request-owned workspaces.
- Final paths are resolved inside the selected output root and checked against traversal and reparse escapes.
- Name collisions use deterministic suffixes and remain safe across threads and processes.
- Multi-artifact publication is transactional where the route declares a group.
- Success, partial success, failure and cancellation report only artifacts that actually survive finalization.
- Physical-page OCR keeps `P` pages/frames separate from `K` exported images. Runtime metadata is mapped into the
  closed Bundle `page_fragment`/`page_resource` relation payloads and then revalidated; it is never exposed as an
  open metadata bag or inferred from filenames.
- A diagnostic may bind to an output `artifact_id`. The Bundle mapping must preserve that binding, and a dangling
  diagnostic reference fails closed.

## Output policy / 输出策略

Directory mode, custom output path, date subdirectory, automatic folder opening and intermediate-file retention are configured through `configs/output.toml`. A request snapshot freezes these choices before conversion.

目录模式、自定义输出路径、日期子目录、自动打开目录和中间文件保留策略由 `configs/output.toml` 控制，并在转换开始前冻结为请求快照。

Best-effort conversions must attach warnings when the target format cannot preserve a supported source feature. Warnings are part of the result contract and must remain visible in CLI JSON and GUI task details.

The normative physical-page mapping, one-based page/source numbers, zero-based relation ordinals, OCR status set,
and unresolved-resource rule are in [Physical-page OCR and artifact relations](specs/physical-page-ocr.md).
