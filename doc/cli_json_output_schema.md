# CLI JSON 输出契约（Schema v2）

本文件描述 DocWen CLI 在 `--json` 模式下的输出契约。此契约通过两类机制共同锁定：

- 黄金样例：`tests/fixtures/golden/*.json`
- 字段 Schema：`doc/cli_json_output_schema.json`

## 1. 顶层 Envelope（所有命令统一）

```json
{
  "schema_version": 2,
  "success": true,
  "command": "convert",
  "data": {},
  "error": null,
  "warnings": [],
  "timing": null
}
```

字段说明：

- schema_version: 固定为 2
- success: boolean，命令语义上的成功/失败
- command: string，子命令名或动作名（例如 convert / formats / templates / doctor / inspect）
- data: object，成功时的业务数据（按 command 不同而不同）
- error: null|object，失败时的结构化错误（`error_code/message/details`）
- warnings: array，警告列表（默认空数组）
- timing: null|object，扩展字段（当前默认 null）

## 2. convert（单文件）

对应入口：`docwen.cli.executor.execute_action(..., json_mode=True)`

```json
{
  "schema_version": 2,
  "success": true,
  "command": "convert",
  "data": {
    "action": "convert",
    "input_file": "path/to/input.docx",
    "output_file": "path/to/output.md",
    "message": "ok",
    "error_code": null,
    "details": null,
    "duration": 0.0,
    "metadata": {}
  },
  "error": null,
  "warnings": [],
  "timing": null
}
```

## 3. convert（批量汇总）

对应入口：`docwen.cli.executor.execute_batch(..., json_mode=True)`

```json
{
  "schema_version": 2,
  "success": false,
  "command": "convert",
  "data": {
    "action": "convert",
    "total": 2,
    "processed_count": 1,
    "success_count": 1,
    "failed_count": 1,
    "interrupted": false,
    "results": []
  },
  "error": {
    "error_code": null,
    "message": "1/2 文件处理失败",
    "details": null
  },
  "warnings": [],
  "timing": null
}
```
