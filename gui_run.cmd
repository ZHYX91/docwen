@echo off
setlocal
set "REPO_ROOT=%~dp0"
set "VENV_PY=%~dp0.venv/Scripts/python.exe"
if exist "%VENV_PY%" (
  "%VENV_PY%" "%REPO_ROOT%tools/run_source.py" gui %*
) else (
  python "%REPO_ROOT%tools/run_source.py" gui %*
)
exit /b %errorlevel%
