$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSCommandPath
$srcPath = Join-Path $repoRoot "src"

if (-not (Test-Path $srcPath)) {
  Write-Error "src 目录不存在: $srcPath"
  exit 1
}

$env:PYTHONPATH = $srcPath
python -m docwen.gui_run @args

