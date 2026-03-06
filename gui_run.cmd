@echo off
setlocal

set "REPO_ROOT=%~dp0"
set "PYTHONPATH=%REPO_ROOT%src"
python -m docwen.gui_run %*

