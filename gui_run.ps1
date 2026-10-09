$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSCommandPath
$python = Join-Path $repoRoot ".venv/Scripts/python.exe"
if (-not (Test-Path -LiteralPath $python)) { $python = (Get-Command python -ErrorAction Stop).Source }
& $python (Join-Path $repoRoot "tools/run_source.py") gui @args
exit $LASTEXITCODE
