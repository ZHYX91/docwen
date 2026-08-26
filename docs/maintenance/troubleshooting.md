# Troubleshooting / 故障排查

## Environment / 环境

```powershell
.\.venv\Scripts\python.exe --version
docwen doctor --json
git status --short --branch
```

Use the workspace virtual environment. Confirm that collection-critical dependencies are installed before interpreting missing tests or routes.

使用 workspace 虚拟环境。解释测试或 route 缺失前，先确认 collection-critical 依赖已安装。

## Conversion failures / 转换失败

- Re-run with JSON output and inspect typed diagnostics.
- Confirm the source extension and detected format agree.
- Check output containment, permissions and name collisions.
- For Office routes, run `docwen doctor` and verify the selected backend.
- For OCR, confirm the requested language model exists under `models/rapidocr/`.
- Preserve the source and a minimal reproducible copy; do not edit the original during diagnosis.

## GUI / GUI 问题

- Start from `docwen-gui` in a terminal to retain startup diagnostics.
- Verify Qt/PySide6 versions and the active display platform.
- Use focused `pytest-qt` tests before physical desktop checks.
- Close only processes owned by the current test; do not terminate unrelated Office/WPS sessions.

Logs are runtime artifacts and may contain paths. Redact private paths before sharing.
