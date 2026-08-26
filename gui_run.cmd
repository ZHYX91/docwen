@echo off
setlocal enabledelayedexpansion

set "REPO_ROOT=%~dp0"
set "PKGS_PATH=%REPO_ROOT%packages"
set "VENV_PY=%REPO_ROOT%.venv\Scripts\python.exe"

REM Build PYTHONPATH with new packages/*/src directories.
set "PKG_SRC="
for /d %%p in ("%PKGS_PATH%\*") do (
  if exist "%%p\src" set "PKG_SRC=%%p\src;!PKG_SRC!"
)

if defined PYTHONPATH (
  set "PYTHONPATH=!PKG_SRC!;%PYTHONPATH%"
) else (
  set "PYTHONPATH=!PKG_SRC!"
)

if exist "%VENV_PY%" (
  "%VENV_PY%" -m docwen_gui %*
) else (
  where python >nul 2>nul
  if errorlevel 1 (
    >&2 echo No Python interpreter found. Create ".venv" or install Python first.
    exit /b 1
  )
  python -m docwen_gui %*
)
