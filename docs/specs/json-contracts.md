# JSON contracts / JSON 契约

> This document describes the CLI `--json` presentation envelope. It is not the stable external process
> boundary and is not a compatibility target. Cross-product consumers use
> [Machine Protocol v1 and Artifact Bundle v2](machine-protocol-v1.md).
>
> 本文描述 CLI `--json` 展示信封；它不是稳定的外部进程边界，也不承担兼容目标。跨产品消费者使用
> [Machine Protocol v1 与 Artifact Bundle v2](machine-protocol-v1.md)。

The CLI presentation envelope conforms to [json-contracts.schema.json](json-contracts.schema.json). The only
public process-integration contract is `docwen.machine.v1`; protocol 2/3 compatibility modes are not provided.

CLI 展示信封遵循 [json-contracts.schema.json](json-contracts.schema.json)。唯一公共进程集成契约是
`docwen.machine.v1`；不提供 protocol 2/3 兼容模式。

## Envelope / 信封结构

Every JSON response contains exactly these top-level fields:

- `protocol_version`: protocol major, currently `3`;
- `product_version`: DocWen product version;
- `success`: terminal success flag;
- `command`: normalized command path such as `convert` or `merge pdf`;
- `data`: command-specific result data;
- `error`: `null` on success, otherwise a typed error object;
- `warnings`: structured, non-terminal diagnostics;
- `meta`: optional execution metadata such as `duration_ms`.

## Rules / 规则

- JSON mode writes exactly one valid JSON document to stdout.
- Runtime logs and human diagnostics never corrupt stdout.
- Error identity comes from `category` and `code`, never localized `message` text.
- `success=false` always carries a typed top-level `error` object. A mixed batch uses
  `batch_partial_failure` and process exit code `9`; an interrupted batch uses
  `operation_cancelled`.
- Successful file commands report absolute `inputs`, `output`, and `artifacts` paths.
- A failed command does not report artifacts that were not created.
- Nested-command failures retain the normalized leaf command path (for example `number markdown`), including parser, path-policy and unexpected-error envelopes.
- `--timing` adds `meta.duration_ms`; otherwise `meta` remains an empty object.
- Required-field, error-category, or semantic changes require a new protocol decision.

## Proofread report / 校对报告

Successful Markdown validation projects the report into `data.details.proofread` and, when
`--report PATH` is supplied, writes the same report as JSON. The report follows
[docwen.proofread_report.v2.schema.json](../../contracts/schemas/docwen.proofread_report.v2.schema.json),
identified by `docwen.proofread_report.v2`.
Its authoritative `range` uses zero-based Unicode-code-point offsets, lines, and columns with
an exclusive end; `source.content_sha256` binds the report to the exact input bytes. A `fix` is present only when a rule provides an explicit,
machine-applicable replacement; presentation text in `suggestion` is never an edit contract.

Markdown 校对成功时，报告会投影到 `data.details.proofread`；显式传入 `--report PATH` 时，
同一报告还会写入 JSON 文件。报告遵循
[docwen.proofread_report.v2.schema.json](../../contracts/schemas/docwen.proofread_report.v2.schema.json)
（当前为 `2.0`）。权威 `range` 使用从零开始的 Unicode code point 偏移、行号和列号，结束位置
为 exclusive；`source.content_sha256` 绑定精确输入字节。只有规则明确给出机器可应用替换时才存在 `fix`，
展示用 `suggestion` 不能被解释为替换文本。

## Runtime capability security status / 运行时能力安全状态

`resources list formats` and `doctor.data.capability_summary` expose the same runtime capability object.
Its `security.dependency_egress_guard` member reports `state`, `installed`, `active`, `scope`, `policy`,
`mechanism`, `bootstrap`, `local_transports`, and `external_processes`. A supported CLI process reports `state=enforced`,
`scope=docwen_python_process`, and `external_processes=not_managed`; the field must not be interpreted as an
operating-system sandbox or as control over Office/WPS/LibreOffice.

`resources list formats` 与 `doctor.data.capability_summary` 返回同一份运行时能力对象。其中
`security.dependency_egress_guard` 明确给出状态、作用域、机制、启动层、本地传输和外部进程边界。受支持的
CLI 进程必须报告 `state=enforced`；`external_processes=not_managed` 表示不能把该状态扩大解释为
操作系统沙箱或对 Office/WPS/LibreOffice 的控制。

## Optimization discovery / 优化资源发现

`resources list optimizations` returns an independent `docwen.optimizations` version 1 object inside the
protocol 3 envelope's `data`. Its exact top-level fields are `resource`, `contract`, `runtime`, `resources`,
and `counts`. Resource fields are `id`, `name`, `action_name`, `scopes`, `available`, `state`, and `bindings`.
Binding fields are `scope`, `route_id`, `source`, `source_category`, `target`, `available`, and `state`.

The command line spelling and protocol 3 envelope do not expose an internal action as a command. A caller
passes the public optimization `id`; DocWen resolves the declared `action_name` from the same runtime
projection. A successful empty runtime uses `resources=[]`; discovery or declaration failure uses the typed
`capability_unavailable` error.

`resources list optimizations` 在 protocol 3 的 `data` 中返回独立的 `docwen.optimizations` v1 对象。
公开资源 ID 与内部 action 名称彼此独立；消费者不得按扩展名、类别或字符串拼接推断 action/scope。
成功的空组合返回空资源数组，发现失败或清单冲突返回 `capability_unavailable`。

## Template discovery / 模板资源发现

`resources list templates` returns `{type, resources, total}`. Every template resource has the exact fields
`id`, `name`, `target`, `description`, `path`, `size_bytes`, and `modified_ns`. The canonical `id` is the token
that a protocol 3 consumer passes unchanged to `convert --template`; consumers must not derive an ID from
`name`, `path`, or a filename.

Core computes the stable ID as
`template.<target>.<sha256(target + NUL + NFC(name).strip().casefold())>`. Absolute paths, file content,
size, and modification time are deliberately excluded, so the same logical named resource keeps its ID when
installed elsewhere or updated in place. A rename changes the resource identity. If two discovered templates
for the same target collide after NFC/casefold normalization, discovery fails closed instead of selecting one
by directory order or fuzzy name matching.

`resources list templates` 返回 `{type, resources, total}`；每项精确包含 `id`、`name`、`target`、
`description`、`path`、`size_bytes` 和 `modified_ns`。protocol 3 消费者必须把公开 `id` 原样传给
`convert --template`，不得从显示名、路径或文件名推导；这些非规范选择值会被直接拒绝。同一 target 下
NFC/casefold 后的身份冲突会使发现安全失败。

## Success example / 成功示例

```json
{
  "protocol_version": 3,
  "product_version": "0.9.1",
  "success": true,
  "command": "convert",
  "data": {
    "inputs": ["C:\\work\\report.docx"],
    "output": "C:\\work\\report.md",
    "artifacts": [
      {
        "path": "C:\\work\\report.md",
        "kind": "markdown",
        "media_type": "text/markdown",
        "primary": true
      }
    ],
    "details": {}
  },
  "error": null,
  "warnings": [],
  "meta": {}
}
```

## Error example / 错误示例

```json
{
  "protocol_version": 3,
  "product_version": "0.9.1",
  "success": false,
  "command": "convert",
  "data": {},
  "error": {
    "category": "conflict",
    "code": "output_exists",
    "message": "Output target already exists.",
    "details": {"path": "C:\\work\\report.md"},
    "hint": "Use --overwrite only when replacing the target is intentional."
  },
  "warnings": [],
  "meta": {}
}
```
