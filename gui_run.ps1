$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSCommandPath
$packagesPath = Join-Path $repoRoot "packages"

# Build PYTHONPATH with new packages/*/src directories.
$pkgSrcPaths = @()
Get-ChildItem -Path $packagesPath -Directory -ErrorAction SilentlyContinue | ForEach-Object {
  $pkgSrc = Join-Path $_.FullName "src"
  if (Test-Path $pkgSrc) {
    $pkgSrcPaths += $pkgSrc
  }
}

$currentPythonPath = $env:PYTHONPATH
if ([string]::IsNullOrWhiteSpace($currentPythonPath)) {
  $env:PYTHONPATH = ($pkgSrcPaths -join ";")
} else {
  $pythonPathEntries = $currentPythonPath -split ";" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
  foreach ($entry in $pkgSrcPaths) {
    if ($pythonPathEntries -notcontains $entry) {
      $pythonPathEntries += $entry
    }
  }
  $env:PYTHONPATH = ($pythonPathEntries -join ";")
}

$venvPython = Join-Path $repoRoot ".venv\Scripts\python.exe"
if (Test-Path $venvPython) {
  & $venvPython -m docwen_gui @args
} else {
  $systemPython = Get-Command python -ErrorAction SilentlyContinue
  if (-not $systemPython) {
    Write-Error '未找到可用的 Python 解释器。请先创建 ".venv" 或安装 Python。'
    exit 1
  }

  & $systemPython.Source -m docwen_gui @args
}
